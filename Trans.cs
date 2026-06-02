// Trans.cs — convert a T1 ASSET CSV "template" export into a thin,
// import-ready shape. The source CSV groups every asset across multiple
// rows (one ASSET row + N ATTRIBUTE rows); Trans thins that out (one row
// per asset) for editing and walks it back to bulk-import shape.
//
// Source CSV format (T1 download template):
//   Row 1  — system-reserved literal "FORMAT ASSET, STANDARD 1.0, …"
//   Row 2  — column header. Column A is always "LineType"; the rest are
//            direct fields (AssetRegisterName, AssetNumber, …) and the
//            attribute slot columns (AttributeCode, SearchPath, LevelNumber,
//            AssetAttributeUserfield1, AssetAttributeSelectionType1, …).
//   Row 3+ — data rows, dispatched on LineType:
//              "ASSET"     — exactly one per asset record; carries that
//                            asset's direct-field values.
//              "ATTRIBUTE" — zero or more per asset, immediately following
//                            the ASSET row. Each one fills a single
//                            (AttributeCode, LevelNumber) slot of the most
//                            recent ASSET above.
//
// Trans is a pure CSV-in / CSV-out pipeline — no Excel involvement. The
// class is stateless apart from the nominated-direct-fields list loaded
// from config.json; every method takes the source CSV path as its first
// argument. Public methods:
//   ParseMeta(sourceCsv) / SaveMetaToJson(sourceCsv, jsonPath, nodeName) /
//   SaveMetaToCsv(sourceCsv, outputCsv)
//     Source CSV → hierarchical meta dict, then on disk as JSON or as a
//     6-row-header CSV (kind / code / level / suffix / dataType / header).
//   Template2Thin(sourceCsv, outputCsv, assetTypeOnly = false)
//     Source CSV → thin CSV. Collapses each asset's ASSET + N ATTRIBUTE
//     rows into a single row, with one column per AttributeCode (cell
//     value = that attribute's SearchPath) plus one column per nominated
//     direct field. Output is a plain CSV: one column-header row + one
//     payload row per asset; no LineType column. assetTypeOnly=true keeps
//     only the ASSET_TYPE attribute column.
//   Thin2Import(sourceCsv, outputCsv)
//     Thin CSV → T1 bulk-import CSV. Reverses Template2Thin: re-adds the
//     LineType column and emits one "ASSET" row plus one "ATTRIBUTE" row
//     per non-empty AttributeCode value, in the shape T1's bulk-import
//     accepts.
//   Template2Flat(sourceCsv, outputCsv)
//     Like Template2Thin, but also exposes the ASSET_TYPE captioned
//     sub-fields (AssetAttributeUserfield<N>, AssetAttributeSelectionType<N>)
//     as their own columns whenever any asset has a value there. Columns
//     are named "ASSET_TYPE/<level>/<suffix>". Payload rows are sorted by
//     the ASSET_TYPE SearchPath.
//   Flat2Import(sourceCsv, outputCsv)
//     Reverses Template2Flat. Behaves like Thin2Import for plain attribute
//     columns; for ASSET_TYPE captioned columns it re-folds them into the
//     ATTRIBUTE row's AssetAttributeUserfield<N> / SelectionType<N> cells,
//     emitting one ATTRIBUTE row per ASSET_TYPE level with values.
//   Csv2Xlsx(csvPath, sheetName)
//     Convenience: load a CSV as a worksheet in the same-named xlsx
//     (csvPath with .xlsx extension). Existing workbook is reused. If the
//     proposed `sheetName` is already taken, "1", "2"… is appended until
//     a free name is found.
//
// Nominated direct fields are loaded once from the top-level
// "nominated_fields" array in config.json — see Trans.FromConfig.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using ClosedXML.Excel;

namespace T1Sync
{
    public class Trans
    {
        private readonly List<string> _nominatedFields = new();

        public Trans(params string[] nominatedFields)
        {
            if (nominatedFields != null) _nominatedFields.AddRange(nominatedFields);
        }

        /// <summary>
        /// Build a Trans with `nominated_fields` loaded from the
        /// top-level `nominated_fields` array in config.json. Shared across
        /// services — no service parameter required.
        /// </summary>
        public static Trans FromConfig(string configPath = T1Client_Interop.DefaultConfigPath)
        {
            var fields = LoadNominatedFromConfig(configPath);
            return new Trans(fields.ToArray());
        }

        private static List<string> LoadNominatedFromConfig(string configPath)
        {
            var fields = new List<string>();
            using var stream = File.OpenRead(configPath);
            using var doc = JsonDocument.Parse(stream);
            var root = doc.RootElement;

            if (root.TryGetProperty("nominated_fields", out var nf) && nf.ValueKind == JsonValueKind.Array)
            {
                foreach (var f in nf.EnumerateArray())
                {
                    if (f.ValueKind == JsonValueKind.String)
                    {
                        var s = f.GetString();
                        if (!string.IsNullOrEmpty(s)) fields.Add(s);
                    }
                }
            }
            return fields;
        }

        // ---------- CSV → hierarchical meta dict ----------

        public Dictionary<string, object> ParseMeta(string sourceCsvPath)
        {
            var rows = ReadCsv(sourceCsvPath);
            if (rows.Count < 2)
            {
                return new Dictionary<string, object>
                {
                    ["fields"] = new Dictionary<string, object>(),
                    ["attributes"] = new Dictionary<string, object>(),
                };
            }

            // Skip the FORMAT line; row index 1 is the header row.
            var header = rows[1];
            var headerIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < header.Count; i++)
            {
                if (!string.IsNullOrEmpty(header[i]) && !headerIndex.ContainsKey(header[i]))
                    headerIndex[header[i]] = i;
            }

            var dataRows = rows.Skip(2).ToList();

            int lineTypeIdx = headerIndex.GetValueOrDefault("LineType", -1);
            int attrCodeIdx = headerIndex.GetValueOrDefault("AttributeCode", -1);
            int levelNumIdx = headerIndex.GetValueOrDefault("LevelNumber", -1);

            var fields = new Dictionary<string, object>();
            var attributes = new Dictionary<string, object>();

            // --- Direct (nominated) fields from the ASSET row ---
            var assetRow = dataRows.FirstOrDefault(r =>
                lineTypeIdx >= 0 && lineTypeIdx < r.Count &&
                r[lineTypeIdx].Equals("ASSET", StringComparison.OrdinalIgnoreCase));

            if (assetRow != null)
            {
                foreach (var fieldName in _nominatedFields)
                {
                    if (!headerIndex.TryGetValue(fieldName, out var idx)) continue;
                    if (idx >= assetRow.Count) continue;
                    fields[fieldName] = InferDataType(assetRow[idx]);
                }
            }

            // --- ATTRIBUTE rows → attributes[code] = { dataType, levels: { lvl: { suffix: [caption, dt] } } } ---
            if (attrCodeIdx >= 0 && levelNumIdx >= 0)
            {
                foreach (var row in dataRows)
                {
                    if (lineTypeIdx < 0 || lineTypeIdx >= row.Count) continue;
                    if (!row[lineTypeIdx].Equals("ATTRIBUTE", StringComparison.OrdinalIgnoreCase)) continue;
                    if (attrCodeIdx >= row.Count) continue;

                    var attrCode = row[attrCodeIdx];
                    if (string.IsNullOrEmpty(attrCode)) continue;

                    var levelStr = (levelNumIdx < row.Count ? row[levelNumIdx] : "").Trim();

                    if (!attributes.TryGetValue(attrCode, out var attrObj) || attrObj is not Dictionary<string, object> attrNode)
                    {
                        attrNode = new Dictionary<string, object> { ["dataType"] = "A" };
                        attributes[attrCode] = attrNode;
                    }

                    AddCaptionedFor(attrNode, row, headerIndex, attrCode, levelStr, "Userfield");
                    AddCaptionedFor(attrNode, row, headerIndex, attrCode, levelStr, "SelectionType");
                }
            }

            return new Dictionary<string, object>
            {
                ["fields"] = fields,
                ["attributes"] = attributes,
            };
        }

        private static void AddCaptionedFor(
            Dictionary<string, object> attrNode, List<string> row, Dictionary<string, int> headerIndex,
            string attrCode, string levelStr, string family)
        {
            Dictionary<string, object>? levels = null;
            Dictionary<string, object>? levelDict = null;

            for (int n = 1; n <= 20; n++)
            {
                var colName = $"AssetAttribute{family}{n}";
                if (!headerIndex.TryGetValue(colName, out var idx)) continue;
                if (idx >= row.Count) continue;
                var value = row[idx];
                if (string.IsNullOrEmpty(value)) continue;

                if (levels == null)
                {
                    if (!attrNode.TryGetValue("levels", out var lvlObj) || lvlObj is not Dictionary<string, object> lvlDict)
                    {
                        lvlDict = new Dictionary<string, object>();
                        attrNode["levels"] = lvlDict;
                    }
                    levels = lvlDict;

                    if (!levels.TryGetValue(levelStr, out var ldObj) || ldObj is not Dictionary<string, object> ld)
                    {
                        ld = new Dictionary<string, object>();
                        levels[levelStr] = ld;
                    }
                    levelDict = ld;
                }

                var suffix = $"{family}{n}";
                if (levelDict!.ContainsKey(suffix)) continue;
                // CSV doesn't carry T1's caption — leave it blank for now.
                levelDict[suffix] = new object[] { "", InferDataType(value) };
            }
        }

        private static string InferDataType(string value)
        {
            if (string.IsNullOrEmpty(value)) return "A";
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out _)) return "N";
            // Heuristic: ISO-ish date detection — treat as 'D'.
            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out _)) return "D";
            return "A";
        }

        // ---------- meta dict → JSON file ----------

        public string SaveMetaToJson(string sourceCsvPath, string jsonPath, string nodeName)
        {
            var meta = ParseMeta(sourceCsvPath);
            var wrapped = new Dictionary<string, object> { [nodeName] = meta };
            var options = new JsonSerializerOptions { WriteIndented = true };
            var dir = Path.GetDirectoryName(jsonPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            File.WriteAllText(jsonPath, JsonSerializer.Serialize(wrapped, options), Encoding.UTF8);
            return jsonPath;
        }

        // ---------- meta dict → 6-row-header CSV ----------

        public string SaveMetaToCsv(string sourceCsvPath, string outputCsvPath)
        {
            var dir = Path.GetDirectoryName(outputCsvPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            var columns = BuildMetaColumns(ParseMeta(sourceCsvPath));
            using var sw = new StreamWriter(outputCsvPath, false, Encoding.UTF8);
            sw.WriteLine(EscapeCsvRow(columns.Select(c => c.Item1)));  // kind
            sw.WriteLine(EscapeCsvRow(columns.Select(c => c.Item2)));  // AttributeCode
            sw.WriteLine(EscapeCsvRow(columns.Select(c => c.Item3)));  // level
            sw.WriteLine(EscapeCsvRow(columns.Select(c => c.Item4)));  // suffix
            sw.WriteLine(EscapeCsvRow(columns.Select(c => c.Item5)));  // dataType
            sw.WriteLine(EscapeCsvRow(columns.Select(c => c.Item6)));  // header (caption)
            return outputCsvPath;
        }

        // ---------- Thin2Import: thin CSV → T1 bulk-import CSV ----------
        //
        // Reverses Template2Thin. Each input row holds one asset's data in
        // a thin shape (AttributeCode columns first, then direct fields);
        // T1's bulk import wants that asset split back into one ASSET row
        // (carrying the direct fields) plus one ATTRIBUTE row per non-empty
        // AttributeCode (carrying that code + its SearchPath value).
        //
        // Input header (saved-as CSV from Template2Thin):
        //   <AttributeCode 1> … <AttributeCode N> | AssetRegisterName, AssetNumber, …
        //   ─── leading attribute columns ───       ─── nominated direct fields ───
        //
        // Output:
        //   Row 1     — "FORMAT ASSET, STANDARD 1.0, DEFINITION $DEFAULT"
        //   Row 2     — LineType, <nominated direct fields…>, AttributeCode, SearchPath
        //   Row 3+    — per input asset:
        //                 row a:   LineType="ASSET", <direct values…>, blank, blank
        //                 row b…:  one per non-empty (non-"NULL") AttributeCode cell,
        //                          LineType="ATTRIBUTE", <blanks…>,
        //                          AttributeCode=<column name>, SearchPath=<cell value>
        //
        // Column matching is case-insensitive; the leading/nominated boundary
        // is the position of `AssetRegisterName` in the input header.
        public string Thin2Import(string sourceCsvPath, string outputCsvPath)
        {
            const string formatStr = "FORMAT ASSET, STANDARD 1.0, DEFINITION $DEFAULT";

            var rows = ReadCsv(sourceCsvPath);

            bool hasFormat = rows.Count > 0 && rows[0].Count > 0 && rows[0][0] == formatStr;
            var headerRow = hasFormat
                ? (rows.Count > 1 ? new List<string>(rows[1]) : new List<string>())
                : (rows.Count > 0 ? new List<string>(rows[0]) : new List<string>());
            var dataRows = hasFormat
                ? rows.Skip(2).ToList()
                : rows.Skip(1).ToList();

            // Boundary = index of AssetRegisterName in the source header
            // (case-insensitive). Everything before it is a "leading attribute
            // column"; its value becomes an extra ATTRIBUTE row per asset.
            int boundary = 0;
            for (int i = 0; i < headerRow.Count; i++)
            {
                if (string.Equals(headerRow[i], "AssetRegisterName", StringComparison.OrdinalIgnoreCase))
                {
                    boundary = i;
                    break;
                }
            }

            // Snapshot leading column names BEFORE removing them from the header.
            var leadingNames = headerRow.GetRange(0, boundary);

            // Drop leading cols, place key column names immediately after nominated fields.
            if (boundary > 0) headerRow.RemoveRange(0, boundary);
            headerRow.Insert(0, "LineType");
            
            int exIdx = headerRow.Count;
            int eyIdx = exIdx + 1;
            
            headerRow.Add("AttributeCode");
            headerRow.Add("SearchPath");

            var outRows = new List<List<string>>
            {
                new List<string> { formatStr },
                headerRow,
            };

            foreach (var srcRow in dataRows)
            {
                // Snapshot the leading values BEFORE removing them from row a.
                var leadingVals = new List<string>(boundary);
                for (int i = 0; i < boundary; i++)
                    leadingVals.Add(i < srcRow.Count ? srcRow[i] : "");

                // Row a: source row with the leading cells dropped, padded to match header length
                var rowA = new List<string>(srcRow);
                if (boundary > 0)
                    rowA.RemoveRange(0, Math.Min(boundary, rowA.Count));
                rowA.Insert(0, "ASSET");
                while (rowA.Count < headerRow.Count) rowA.Add("");
                outRows.Add(rowA);

                // One row b per non-empty (non-"NULL") leading cell.
                for (int j = 0; j < leadingNames.Count; j++)
                {
                    var v = leadingVals[j];
                    if (string.IsNullOrEmpty(v)) continue;
                    if (string.Equals(v, "NULL", StringComparison.OrdinalIgnoreCase)) continue;
                    var rowB = new List<string>(new string[headerRow.Count]);
                    for (int i = 0; i < rowB.Count; i++) rowB[i] = "";
                    rowB[0]     = "ATTRIBUTE";
                    rowB[exIdx] = leadingNames[j];
                    rowB[eyIdx] = v;
                    outRows.Add(rowB);
                }
            }

            var dir = Path.GetDirectoryName(outputCsvPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            using var sw = new StreamWriter(outputCsvPath, false, Encoding.UTF8);
            foreach (var row in outRows)
                sw.WriteLine(EscapeCsvRow(row));

            return outputCsvPath;
        }

        // ---------- Flat2Import: flat CSV (with ASSET_TYPE captions) → T1 bulk-import CSV ----------
        //
        // Reverses Template2Flat. Same shape as Thin2Import for plain
        // attribute columns; captioned ASSET_TYPE columns (header pattern
        // "ASSET_TYPE/<level>/<suffix>") are re-folded into the
        // AssetAttributeUserfield<N> / AssetAttributeSelectionType<N> cells of
        // an ATTRIBUTE row carrying LevelNumber=<level>. One ATTRIBUTE row is
        // emitted per ASSET_TYPE level that has at least one non-empty value.
        //
        // Output header: LineType, <nominated direct fields…>, AttributeCode,
        // SearchPath, LevelNumber, AssetAttributeUserfield1..N,
        // AssetAttributeSelectionType1..N (the last three groups appear only
        // when the input has ASSET_TYPE captioned columns).
        public string Flat2Import(string sourceCsvPath, string outputCsvPath)
        {
            const string formatStr = "FORMAT ASSET, STANDARD 1.0, DEFINITION $DEFAULT";

            var rows = ReadCsv(sourceCsvPath);
            bool hasFormat = rows.Count > 0 && rows[0].Count > 0 && rows[0][0] == formatStr;
            var headerRow = hasFormat
                ? (rows.Count > 1 ? new List<string>(rows[1]) : new List<string>())
                : (rows.Count > 0 ? new List<string>(rows[0]) : new List<string>());
            var dataRows = hasFormat ? rows.Skip(2).ToList() : rows.Skip(1).ToList();

            // Boundary = index of AssetRegisterName. Everything before it is a
            // "leading attribute column" — either a scalar code, or a captioned
            // ASSET_TYPE column matching "<code>/<level>/<suffix>".
            int boundary = 0;
            for (int i = 0; i < headerRow.Count; i++)
            {
                if (string.Equals(headerRow[i], "AssetRegisterName", StringComparison.OrdinalIgnoreCase))
                { boundary = i; break; }
            }

            var leadingCols = new List<(string Header, bool IsCaption, string Code, int Level, string Suffix)>();
            int maxUserfieldN = 0, maxSelectionTypeN = 0;
            for (int i = 0; i < boundary; i++)
            {
                var h = headerRow[i];
                var parts = h.Split('/');
                if (parts.Length == 3 && int.TryParse(parts[1], out int lvl))
                {
                    leadingCols.Add((h, true, parts[0], lvl, parts[2]));
                    if (parts[2].StartsWith("Userfield", StringComparison.OrdinalIgnoreCase) &&
                        int.TryParse(parts[2].Substring("Userfield".Length), out int un))
                        maxUserfieldN = Math.Max(maxUserfieldN, un);
                    else if (parts[2].StartsWith("SelectionType", StringComparison.OrdinalIgnoreCase) &&
                             int.TryParse(parts[2].Substring("SelectionType".Length), out int sn))
                        maxSelectionTypeN = Math.Max(maxSelectionTypeN, sn);
                }
                else
                {
                    leadingCols.Add((h, false, h, 0, ""));
                }
            }

            bool hasCaptions = maxUserfieldN > 0 || maxSelectionTypeN > 0;

            // Build output header.
            var outHeader = new List<string> { "LineType" };
            for (int i = boundary; i < headerRow.Count; i++) outHeader.Add(headerRow[i]);
            int attrCodeIdx = outHeader.Count; outHeader.Add("AttributeCode");
            int searchPathIdx = outHeader.Count; outHeader.Add("SearchPath");
            int levelNumIdx = -1;
            var userfieldIdx = new int[maxUserfieldN + 1];      // index 1..N (0 unused)
            var selectionIdx = new int[maxSelectionTypeN + 1];
            if (hasCaptions)
            {
                levelNumIdx = outHeader.Count; outHeader.Add("LevelNumber");
                for (int n = 1; n <= maxUserfieldN; n++)
                {
                    userfieldIdx[n] = outHeader.Count;
                    outHeader.Add($"AssetAttributeUserfield{n}");
                }
                for (int n = 1; n <= maxSelectionTypeN; n++)
                {
                    selectionIdx[n] = outHeader.Count;
                    outHeader.Add($"AssetAttributeSelectionType{n}");
                }
            }

            var outRows = new List<List<string>>
            {
                new List<string> { formatStr },
                outHeader,
            };

            List<string> BlankRow() { var r = new List<string>(new string[outHeader.Count]); for (int i = 0; i < r.Count; i++) r[i] = ""; return r; }

            foreach (var srcRow in dataRows)
            {
                // Snapshot leading + direct values.
                var leadingVals = new List<string>(boundary);
                for (int i = 0; i < boundary; i++)
                    leadingVals.Add(i < srcRow.Count ? srcRow[i] : "");

                // Row a: LineType=ASSET, direct fields filled.
                var rowA = BlankRow();
                rowA[0] = "ASSET";
                for (int i = 0; i < headerRow.Count - boundary; i++)
                {
                    int srcI = boundary + i;
                    rowA[1 + i] = srcI < srcRow.Count ? srcRow[srcI] : "";
                }
                outRows.Add(rowA);

                // Scalar attribute rows + ASSET_TYPE caption collection.
                string assetTypeSp = "";
                var atByLevel = new SortedDictionary<int, Dictionary<string, string>>();
                for (int j = 0; j < leadingCols.Count; j++)
                {
                    var col = leadingCols[j];
                    var v = leadingVals[j];
                    if (string.IsNullOrEmpty(v) || string.Equals(v, "NULL", StringComparison.OrdinalIgnoreCase))
                        continue;
                    bool isAssetType = string.Equals(col.Code, "ASSET_TYPE", StringComparison.OrdinalIgnoreCase);
                    if (col.IsCaption && isAssetType)
                    {
                        if (!atByLevel.TryGetValue(col.Level, out var lvlDict))
                        {
                            lvlDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            atByLevel[col.Level] = lvlDict;
                        }
                        lvlDict[col.Suffix] = v;
                    }
                    else if (!col.IsCaption && isAssetType)
                    {
                        assetTypeSp = v;
                    }
                    else if (!col.IsCaption)
                    {
                        var rowB = BlankRow();
                        rowB[0] = "ATTRIBUTE";
                        rowB[attrCodeIdx] = col.Code;
                        rowB[searchPathIdx] = v;
                        outRows.Add(rowB);
                    }
                }

                // Emit ASSET_TYPE row(s). If captioned values exist, one row per
                // level; otherwise just a single scalar row (matching Thin2Import).
                if (atByLevel.Count > 0)
                {
                    foreach (var (lvl, lvlDict) in atByLevel)
                    {
                        var rowB = BlankRow();
                        rowB[0] = "ATTRIBUTE";
                        rowB[attrCodeIdx] = "ASSET_TYPE";
                        rowB[searchPathIdx] = assetTypeSp;
                        if (levelNumIdx >= 0) rowB[levelNumIdx] = lvl.ToString();
                        foreach (var (sfx, val) in lvlDict)
                        {
                            if (sfx.StartsWith("Userfield", StringComparison.OrdinalIgnoreCase) &&
                                int.TryParse(sfx.Substring("Userfield".Length), out int un) &&
                                un >= 1 && un <= maxUserfieldN)
                                rowB[userfieldIdx[un]] = val;
                            else if (sfx.StartsWith("SelectionType", StringComparison.OrdinalIgnoreCase) &&
                                     int.TryParse(sfx.Substring("SelectionType".Length), out int sn) &&
                                     sn >= 1 && sn <= maxSelectionTypeN)
                                rowB[selectionIdx[sn]] = val;
                        }
                        outRows.Add(rowB);
                    }
                }
                else if (!string.IsNullOrEmpty(assetTypeSp))
                {
                    var rowB = BlankRow();
                    rowB[0] = "ATTRIBUTE";
                    rowB[attrCodeIdx] = "ASSET_TYPE";
                    rowB[searchPathIdx] = assetTypeSp;
                    outRows.Add(rowB);
                }
            }

            var dir = Path.GetDirectoryName(outputCsvPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            using var sw = new StreamWriter(outputCsvPath, false, Encoding.UTF8);
            foreach (var row in outRows)
                sw.WriteLine(EscapeCsvRow(row));

            return outputCsvPath;
        }

        // ---------- Csv2Xlsx: load CSV as a worksheet in the same-named xlsx ----------
        //
        // Output path is csvPath with the .xlsx extension; if the workbook
        // already exists it's reused and the CSV is appended as a new sheet.
        // If `sheetName` is already taken, "1", "2"… is appended until a
        // free name is found.
        //
        // Each data row's first cell also gets a hyperlink pointing at the
        // T1 AssetMyMaintenance page for that asset, parameterised by the
        // row's AssetRegisterName + AssetNumber values (header lookup is
        // case-insensitive). Rows missing either value are left as plain
        // text.
        public string Csv2Xlsx(string csvPath, string sheetName)
        {
            const string urlTemplate =
                "https://maroondah-build.t1cloud.com/T1Default/CiAnywhere/Web/MAROONDAH-build/"
                + "AssetsCore/AssetMyMaintenance?f=$ASC.ASSET.MNT&suite=CES"
                + "&SK.AssetRegisterName={assetregistername}&SK.KeyedAssetNumber={assetnumber}";

            var xlsxPath = Path.ChangeExtension(csvPath, ".xlsx");
            var rows = ReadCsv(csvPath);

            using var wb = File.Exists(xlsxPath) ? new XLWorkbook(xlsxPath) : new XLWorkbook();
            if (wb.Worksheets.Contains("Sheet")) wb.Worksheets.Delete("Sheet");

            var actualName = sheetName;
            int n = 1;
            while (wb.Worksheets.Contains(actualName))
            {
                actualName = sheetName + n;
                n++;
            }

            // Locate AssetRegisterName / AssetNumber columns in the header row.
            int arnCol = -1, anCol = -1;
            if (rows.Count > 0)
            {
                var header = rows[0];
                for (int i = 0; i < header.Count; i++)
                {
                    if (arnCol < 0 && string.Equals(header[i], "AssetRegisterName", StringComparison.OrdinalIgnoreCase))
                        arnCol = i;
                    if (anCol < 0 && string.Equals(header[i], "AssetNumber", StringComparison.OrdinalIgnoreCase))
                        anCol = i;
                }
            }

            var ws = wb.Worksheets.Add(actualName);
            for (int r = 0; r < rows.Count; r++)
            {
                var row = rows[r];
                for (int c = 0; c < row.Count; c++)
                    ws.Cell(r + 1, c + 1).Value = row[c];

                if (r == 0 || arnCol < 0 || anCol < 0) continue;
                var arn = arnCol < row.Count ? row[arnCol] : "";
                var an  = anCol  < row.Count ? row[anCol]  : "";
                if (string.IsNullOrEmpty(arn) || string.IsNullOrEmpty(an)) continue;

                var url = urlTemplate
                    .Replace("{assetregistername}", Uri.EscapeDataString(arn))
                    .Replace("{assetnumber}",      Uri.EscapeDataString(an));
                ws.Cell(r + 1, 1).SetHyperlink(new XLHyperlink(url));
            }

            wb.SaveAs(xlsxPath);
            return xlsxPath;
        }

        // Minimal CSV escaper: wrap field in double quotes if it contains
        // a comma, double quote, or newline; double up internal quotes.
        private static string EscapeCsvRow(IEnumerable<string> fields)
        {
            return string.Join(",", fields.Select(f =>
            {
                if (f == null) return "";
                if (f.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0)
                    return "\"" + f.Replace("\"", "\"\"") + "\"";
                return f;
            }));
        }

        // ---------- Template2Thin: source CSV → thin CSV ----------
        //
        // The source template stores one asset across many CSV rows (one
        // LineType=ASSET row + 0..N LineType=ATTRIBUTE rows). This method
        // collapses that into a single row per asset:
        //
        //   <AttributeCode 1> … <AttributeCode N> | AssetRegisterName, AssetNumber, …
        //   ─── one column per distinct code ───    ─── nominated direct fields ───
        //
        // The cell under each AttributeCode column holds that attribute's
        // SearchPath. Captioned sub-fields (Userfield1, SelectionType2, …)
        // are NOT exploded into columns — this is the "brief" view.
        //
        // Output is a CSV with exactly two sections: row 1 = column header
        // (AttributeCode names followed by nominated direct field names),
        // row 2+ = one payload row per asset. No LineType column —
        // Thin2Import re-adds it. assetTypeOnly=true keeps only the
        // ASSET_TYPE attribute column (case-insensitive match).
        public string Template2Thin(string sourceCsvPath, string outputCsvPath, bool assetTypeOnly = false)
        {
            var (assets, attrCodesOrdered) = ReadCsvBrief(sourceCsvPath);

            if (assetTypeOnly)
            {
                attrCodesOrdered = attrCodesOrdered
                    .Where(c => string.Equals(c, "ASSET_TYPE", StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            // Column order: AttributeCode columns lead (one per distinct code
            // from ATTRIBUTE rows, regardless of LevelNumber), then nominated
            // direct fields with AssetRegisterName/AssetNumber forced to the
            // front. No LineType — Thin2Import re-adds it.
            var briefCols = new List<(string Kind, string Code, string Level, string Suffix, string DataType, string Header)>();
            foreach (var code in attrCodesOrdered)
                briefCols.Add(("Attribute", "", "", "", "A", code));
            foreach (var f in OrderNominatedFields(_nominatedFields))
                briefCols.Add(("", "", "", "", "A", f));

            var dir = Path.GetDirectoryName(outputCsvPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            using var sw = new StreamWriter(outputCsvPath, false, Encoding.UTF8);

            // Single column-header row.
            sw.WriteLine(EscapeCsvRow(briefCols.Select(c => c.Header)));

            // One payload row per asset:
            //   kind=""        + header → direct field         (asset.Fields)
            //   kind="Attribute" + level=""  → AttributeCode scalar (asset.Attributes)
            foreach (var asset in assets)
            {
                var row = new List<string>(briefCols.Count);
                foreach (var c in briefCols)
                {
                    string? val = null;
                    if (c.Kind == "" && !string.IsNullOrEmpty(c.Header))
                        asset.Fields.TryGetValue(c.Header, out val);
                    else if (c.Kind == "Attribute" && string.IsNullOrEmpty(c.Level))
                        asset.Attributes.TryGetValue(c.Header, out val);
                    row.Add(val ?? "");
                }
                sw.WriteLine(EscapeCsvRow(row));
            }

            return outputCsvPath;
        }

        // ---------- Template2Flat: source CSV → flat CSV with ASSET_TYPE captions ----------
        //
        // Same shape as Template2Thin, plus extra columns for the ASSET_TYPE
        // captioned sub-fields whenever any asset has a non-empty value at
        // some (level, suffix). Those columns sit right after the ASSET_TYPE
        // scalar column and are named "ASSET_TYPE/<level>/<suffix>", e.g.
        // "ASSET_TYPE/1/Userfield1". Payload rows are sorted by the
        // ASSET_TYPE SearchPath (case-insensitive, ascending; empty last).
        public string Template2Flat(string sourceCsvPath, string outputCsvPath)
        {
            var (assets, attrCodesOrdered) = ReadCsvBrief(sourceCsvPath);

            // Find (level, suffix) combos that have a non-empty value in at
            // least one asset. Sort by numeric level then suffix.
            var slotSet = new HashSet<(int Level, string Suffix)>();
            foreach (var asset in assets)
            {
                foreach (var (levelStr, suffixDict) in asset.AssetTypeCaptions)
                {
                    if (!int.TryParse(levelStr, out int level)) continue;
                    foreach (var (suffix, val) in suffixDict)
                    {
                        if (!string.IsNullOrEmpty(val)) slotSet.Add((level, suffix));
                    }
                }
            }
            var captionSlots = slotSet.OrderBy(s => s.Level).ThenBy(s => s.Suffix, StringComparer.OrdinalIgnoreCase).ToList();

            // Locate the canonical ASSET_TYPE code in the source (preserves its
            // case). If it's missing entirely, fall back to the literal string.
            string assetTypeCode = attrCodesOrdered.FirstOrDefault(
                c => string.Equals(c, "ASSET_TYPE", StringComparison.OrdinalIgnoreCase)) ?? "ASSET_TYPE";

            // Build column layout. Tagged so we can both render the header and
            // look up each cell's value.
            //   "Attr" with Code=<code>                    → asset.Attributes[code]
            //   "Caption" with Code=ASSET_TYPE, Level, Sfx → asset.AssetTypeCaptions[level][suffix]
            //   "Field" with Header=<field name>           → asset.Fields[name]
            var cols = new List<(string Kind, string Header, string Code, string Level, string Suffix)>();
            bool hasAssetType = attrCodesOrdered.Any(c => string.Equals(c, "ASSET_TYPE", StringComparison.OrdinalIgnoreCase));
            if (hasAssetType)
            {
                cols.Add(("Attr", assetTypeCode, assetTypeCode, "", ""));
                foreach (var (lvl, sfx) in captionSlots)
                {
                    var header = $"{assetTypeCode}/{lvl}/{sfx}";
                    cols.Add(("Caption", header, assetTypeCode, lvl.ToString(), sfx));
                }
            }
            foreach (var code in attrCodesOrdered)
            {
                if (string.Equals(code, "ASSET_TYPE", StringComparison.OrdinalIgnoreCase)) continue;
                cols.Add(("Attr", code, code, "", ""));
            }
            foreach (var f in OrderNominatedFields(_nominatedFields))
                cols.Add(("Field", f, "", "", ""));

            // Sort assets by ASSET_TYPE SearchPath. Empty/missing sorts last
            // and ties break on original input order.
            var sorted = assets
                .Select((a, i) => (Asset: a, Index: i,
                                    Sp: a.Attributes.TryGetValue(assetTypeCode, out var s) ? s : ""))
                .OrderBy(t => string.IsNullOrEmpty(t.Sp))
                .ThenBy(t => t.Sp, StringComparer.OrdinalIgnoreCase)
                .ThenBy(t => t.Index)
                .Select(t => t.Asset)
                .ToList();

            var dir = Path.GetDirectoryName(outputCsvPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            using var sw = new StreamWriter(outputCsvPath, false, Encoding.UTF8);
            sw.WriteLine(EscapeCsvRow(cols.Select(c => c.Header)));

            foreach (var asset in sorted)
            {
                var row = new List<string>(cols.Count);
                foreach (var c in cols)
                {
                    string val = "";
                    if (c.Kind == "Attr")
                    {
                        asset.Attributes.TryGetValue(c.Code, out var v);
                        val = v ?? "";
                    }
                    else if (c.Kind == "Caption")
                    {
                        if (asset.AssetTypeCaptions.TryGetValue(c.Level, out var lvlDict) &&
                            lvlDict.TryGetValue(c.Suffix, out var v))
                            val = v;
                    }
                    else if (c.Kind == "Field")
                    {
                        asset.Fields.TryGetValue(c.Header, out var v);
                        val = v ?? "";
                    }
                    row.Add(val);
                }
                sw.WriteLine(EscapeCsvRow(row));
            }

            return outputCsvPath;
        }

        // Walks the source CSV, grouping each LineType=ASSET row with its
        // following LineType=ATTRIBUTE rows into one AssetRecord. Returns
        // (assets, ordered list of distinct AttributeCodes by first appearance).
        private (List<AssetRecord> Assets, List<string> AttrCodes) ReadCsvBrief(string sourceCsvPath)
        {
            var rows = ReadCsv(sourceCsvPath);
            var assets = new List<AssetRecord>();
            var attrCodesOrdered = new List<string>();
            var seenCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (rows.Count < 2) return (assets, attrCodesOrdered);

            var header = rows[1];
            var headerIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < header.Count; i++)
            {
                if (!string.IsNullOrEmpty(header[i]) && !headerIndex.ContainsKey(header[i]))
                    headerIndex[header[i]] = i;
            }

            int lineTypeIdx   = headerIndex.GetValueOrDefault("LineType",      -1);
            int attrCodeIdx   = headerIndex.GetValueOrDefault("AttributeCode", -1);
            int searchPathIdx = headerIndex.GetValueOrDefault("SearchPath",    -1);
            int levelNumIdx   = headerIndex.GetValueOrDefault("LevelNumber",   -1);

            AssetRecord? current = null;

            foreach (var row in rows.Skip(2))
            {
                if (lineTypeIdx < 0 || lineTypeIdx >= row.Count) continue;
                var lineType = row[lineTypeIdx];

                if (lineType.Equals("ASSET", StringComparison.OrdinalIgnoreCase))
                {
                    current = new AssetRecord();
                    foreach (var f in _nominatedFields)
                    {
                        if (!headerIndex.TryGetValue(f, out var idx)) continue;
                        if (idx >= row.Count) continue;
                        current.Fields[f] = row[idx];
                    }
                    assets.Add(current);
                }
                else if (lineType.Equals("ATTRIBUTE", StringComparison.OrdinalIgnoreCase) && current != null)
                {
                    if (attrCodeIdx < 0 || attrCodeIdx >= row.Count) continue;
                    var code = row[attrCodeIdx];
                    if (string.IsNullOrEmpty(code)) continue;
                    var sp = (searchPathIdx >= 0 && searchPathIdx < row.Count) ? row[searchPathIdx] : "";

                    // Multiple ATTRIBUTE rows per code (one per level) all carry the
                    // same SearchPath in T1's CSV export, so first-write-wins is fine.
                    if (!current.Attributes.ContainsKey(code))
                        current.Attributes[code] = sp;

                    if (seenCodes.Add(code)) attrCodesOrdered.Add(code);

                    // For ASSET_TYPE, also stash any captioned sub-fields at this
                    // level — Template2Flat exposes them as their own columns.
                    if (string.Equals(code, "ASSET_TYPE", StringComparison.OrdinalIgnoreCase))
                    {
                        var levelStr = (levelNumIdx >= 0 && levelNumIdx < row.Count) ? row[levelNumIdx].Trim() : "";
                        if (!string.IsNullOrEmpty(levelStr))
                        {
                            CollectAssetTypeCaptions(current, row, headerIndex, levelStr, "Userfield");
                            CollectAssetTypeCaptions(current, row, headerIndex, levelStr, "SelectionType");
                        }
                    }
                }
            }

            return (assets, attrCodesOrdered);
        }

        private static void CollectAssetTypeCaptions(AssetRecord asset, List<string> row,
            Dictionary<string, int> headerIndex, string levelStr, string family)
        {
            Dictionary<string, string>? levelDict = null;
            for (int n = 1; n <= 20; n++)
            {
                var colName = $"AssetAttribute{family}{n}";
                if (!headerIndex.TryGetValue(colName, out var idx)) continue;
                if (idx >= row.Count) continue;
                var value = row[idx];
                if (string.IsNullOrEmpty(value)) continue;
                if (levelDict == null)
                {
                    if (!asset.AssetTypeCaptions.TryGetValue(levelStr, out levelDict))
                    {
                        levelDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        asset.AssetTypeCaptions[levelStr] = levelDict;
                    }
                }
                levelDict[$"{family}{n}"] = value;
            }
        }

        private class AssetRecord
        {
            public Dictionary<string, string> Fields     = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> Attributes = new(StringComparer.OrdinalIgnoreCase);
            // ASSET_TYPE captioned values keyed [levelStr][suffix] -> value.
            // Populated by ReadCsvBrief whenever an ATTRIBUTE row for ASSET_TYPE
            // carries non-empty AssetAttributeUserfield<N>/SelectionType<N> cells.
            // Used by Template2Flat; ignored by Template2Thin.
            public Dictionary<string, Dictionary<string, string>> AssetTypeCaptions
                = new(StringComparer.OrdinalIgnoreCase);
        }

        // Promote AssetRegisterName and AssetNumber to the front of the
        // nominated direct-field list; preserve original order for the rest.
        private static List<string> OrderNominatedFields(IEnumerable<string> nominated)
        {
            var leading = new[] { "AssetRegisterName", "AssetNumber" };
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var ordered = new List<string>();
            foreach (var name in leading)
            {
                var match = nominated.FirstOrDefault(f => string.Equals(f, name, StringComparison.OrdinalIgnoreCase));
                if (match != null && seen.Add(match)) ordered.Add(match);
            }
            foreach (var f in nominated)
                if (seen.Add(f)) ordered.Add(f);
            return ordered;
        }

        // Thin pass-through to the shared MetaSchema.BuildColumns helper —
        // guarantees byte-identical column layout to T1Client_Interop / T1Client_ClosedXML.
        private static List<(string, string, string, string, string, string)> BuildMetaColumns(object nodeMetaObj)
        {
            var typed = MetaSchema.BuildColumns(nodeMetaObj);
            var columns = new List<(string, string, string, string, string, string)>(typed.Count);
            foreach (var c in typed) columns.Add((c.Kind, c.Code, c.Level, c.Suffix, c.DataType, c.Header));
            return columns;
        }

        // ---------- minimal CSV reader (handles quoted fields with embedded commas / quotes) ----------

        private static List<List<string>> ReadCsv(string path)
        {
            var rows = new List<List<string>>();
            var fields = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;

            using var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            int ch;
            while ((ch = reader.Read()) != -1)
            {
                char c = (char)ch;
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (reader.Peek() == '"') { current.Append('"'); reader.Read(); }
                        else inQuotes = false;
                    }
                    else current.Append(c);
                }
                else
                {
                    if (c == '"') inQuotes = true;
                    else if (c == ',') { fields.Add(current.ToString()); current.Clear(); }
                    else if (c == '\r') { /* ignore */ }
                    else if (c == '\n')
                    {
                        fields.Add(current.ToString()); current.Clear();
                        rows.Add(fields); fields = new List<string>();
                    }
                    else current.Append(c);
                }
            }
            if (current.Length > 0 || fields.Count > 0)
            {
                fields.Add(current.ToString());
                rows.Add(fields);
            }
            return rows;
        }
    }
}
