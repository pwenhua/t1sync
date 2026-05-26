// Trans.cs — read a T1 ASSET CSV export and produce the same meta
// shape as workshop-TP.json (built by T1Client.ParseAssetsMeta), then save
// it to an Excel workbook the same way T1Client.SaveMetaToExcel does.
//
// CSV layout (T1 standard export):
//   line 1: "FORMAT ASSET, STANDARD 1.0, …"        ← skipped
//   line 2: header row (LineType, AssetRegisterName, …, AttributeCode,
//                       SearchPath, LevelNumber, AssetAttributeUserfield1, …)
//   line 3: LineType=ASSET row — direct field values
//   line 4+: LineType=ATTRIBUTE rows — one per (AttributeCode, LevelNumber)
//
// Meta shape produced (matches T1Client.ParseAssetsMeta hierarchical schema):
//   {
//     "Tree/Street Tree": {
//        "fields": { "AssetRegisterName": "A", "AssetNumber": "A", ... },
//        "attributes": {
//          "ASSET_TYPE": {
//            "dataType": "A",
//            "levels": {
//              "1": { "Userfield1": [caption, "N"], "SelectionType1": [caption, "A"], ... },
//              "2": { "Userfield1": [caption, "A"] }
//            }
//          },
//          "LOCATION":    { "dataType": "A" },
//          "SERVICEAREA": { "dataType": "A" }
//        }
//     }
//   }
// CSV exports don't carry T1's captions, so the caption slot is left empty
// (""). The Excel that comes out has a blank row-6 cell for each captioned
// attribute, ready for the user to fill in.
//
// Usage:
//   var t = new Trans("ASSET_Export_25052026-011611.csv",
//                              "AssetRegisterName", "AssetNumber", "Description",
//                              "ShortDescription", "Status", "OperatingStatus");
//   var meta = t.ParseMeta();
//   t.SaveMetaToJson(@"c:\temp\csv-meta.json", "Tree/Street Tree");
//   t.SaveMetaToExcel(@"c:\temp\csv-meta.xlsx", "Tree/Street Tree");

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
        private readonly string _csvPath;
        private readonly List<string> _nominatedFields = new();

        public Trans(string csvPath, params string[] nominatedFields)
        {
            _csvPath = csvPath;
            if (nominatedFields != null) _nominatedFields.AddRange(nominatedFields);
        }

        /// <summary>
        /// Build a Trans with `nominated_fields` loaded from the
        /// top-level `nominated_fields` array in config.json. Shared across
        /// services — no service parameter required.
        /// </summary>
        public static Trans FromConfig(
            string csvPath,
            string configPath = T1Client_Interop.DefaultConfigPath)
        {
            var fields = LoadNominatedFromConfig(configPath);
            return new Trans(csvPath, fields.ToArray());
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

        // ---------- Step 1: CSV → meta dict ----------

        public Dictionary<string, object> ParseMeta()
        {
            var rows = ReadCsv(_csvPath);
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

        // ---------- Step 2: meta dict → JSON file ----------

        public string SaveMetaToJson(string jsonPath, string nodeName)
        {
            var meta = ParseMeta();
            var wrapped = new Dictionary<string, object> { [nodeName] = meta };
            var options = new JsonSerializerOptions { WriteIndented = true };
            File.WriteAllText(jsonPath, JsonSerializer.Serialize(wrapped, options), Encoding.UTF8);
            return jsonPath;
        }

        // ---------- Step 3: meta dict → Excel workbook ----------
        // Mirrors T1Client_Interop.SaveMetaToExcel (ClosedXML path).

        public string SaveMetaToExcel(string xlsxPath, string nodeName)
        {
            var meta = ParseMeta();
            var wrapped = new Dictionary<string, object> { [nodeName] = meta };
            return SaveMetaToExcelStatic(xlsxPath, wrapped);
        }

        // ---------- Step 4: CSV (template) → flat brief Excel ----------
        //
        // SaveMetaToExcel produces a metadata-only workbook (6 header rows + no
        // data). Template2Flat produces a data workbook:
        //   • The CSV is a "template" — one logical asset spans many CSV rows:
        //       one ASSET line + 0..N following ATTRIBUTE lines.
        //   • The output is "flat" — one row per asset, one column per
        //     (nominated direct field) and one column per (AttributeCode).
        //   • "Brief" — captioned sub-fields inside each attribute (Userfield1,
        //     SelectionType2, …) are NOT exploded into their own columns.
        //     The column for each AttributeCode just holds the SearchPath
        //     (i.e. the "value" of that attribute).
        //
        // Header layout is the same 6-row scheme T1Client uses, so the output
        // is consumable by extract_asset / sync_asset_from_excel unchanged.
        // ---------- Step 5: source CSV → well-formatted T1 import CSV ----------
        //
        // Rules:
        //   Row 1 — must start with the literal FORMAT line
        //           "FORMAT ASSET, STANDARD 1.0, DEFINITION $DEFAULT".
        //           If the source CSV already has it at row 1, keep it; otherwise
        //           a fresh FORMAT row is prepended.
        //   Row 2 — column header, with `AttributeCode` placed at Excel column EX
        //           (the 154th column) and `SearchPath` at column EY (155th).
        //           If the source header is narrower it's padded with empty cells.
        //   Row 3+ — for each source data record (row a):
        //              • write row a verbatim (with the leading attribute columns
        //                stripped, padded to ≥ column EY)
        //              • for every "leading attribute column" (any source column
        //                appearing BEFORE `AssetRegisterName` in the header)
        //                whose row-a cell is non-empty and not "NULL", emit an
        //                extra ATTRIBUTE row b with
        //                    col 0  = "ATTRIBUTE"
        //                    col EX = the leading column's header name
        //                    col EY = the row-a value for that column
        //                This generalises the previous Asset_Type-only handling
        //                to match the new Template2Flat layout, where every
        //                AttributeCode is a leading column ahead of the nominated
        //                direct fields. Column matching is case-insensitive.
        public string Flat2Import(string sourceCsvPath, string outputCsvPath)
        {
            const string formatStr = "FORMAT ASSET, STANDARD 1.0, DEFINITION $DEFAULT";
            const int    exIdx     = 153;  // 0-based; column 154; Excel column EX
            const int    eyIdx     = 154;  // 0-based; column 155; Excel column EY

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

            // Drop leading cols, pad to ≥ eyIdx + 1, place key column names.
            if (boundary > 0) headerRow.RemoveRange(0, boundary);
            headerRow.Insert(0, "LineType");
            while (headerRow.Count <= eyIdx) headerRow.Add("");
            headerRow[exIdx] = "AttributeCode";
            headerRow[eyIdx] = "SearchPath";

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

                // Row a: source row with the leading cells dropped, padded to ≥ EY.
                var rowA = new List<string>(srcRow);
                if (boundary > 0)
                    rowA.RemoveRange(0, Math.Min(boundary, rowA.Count));
                rowA.Insert(0, "ASSET");
                while (rowA.Count <= eyIdx) rowA.Add("");
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

        // ---------- Step 6: CSV → T1Sync 6-row header + flat data per asset ----------

        public string Template2Flat(string xlsxPath, string sheet, bool assetTypeOnly = false)
        {
            var (assets, attrCodesOrdered) = ReadCsvBrief();

            // When assetTypeOnly is true, drop every AttributeCode column except
            // ASSET_TYPE (matched case-insensitively so "Asset_Type"/"asset_type"
            // also count). Default false keeps the original behaviour: every
            // distinct AttributeCode in the source gets its own column.
            if (assetTypeOnly)
            {
                attrCodesOrdered = attrCodesOrdered
                    .Where(c => string.Equals(c, "ASSET_TYPE", StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            var dir = Path.GetDirectoryName(xlsxPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            using var wb = File.Exists(xlsxPath) ? new XLWorkbook(xlsxPath) : new XLWorkbook();
            if (wb.Worksheets.Contains("Sheet")) wb.Worksheets.Delete("Sheet");

            var sheetName = MetaSchema.SanitizeSheetName(sheet);
            IXLWorksheet ws;
            List<(string Kind, string Level, string Header)> layout;

            if (wb.Worksheets.Contains(sheetName))
            {
                // Sheet already exists (e.g. previously initialized by SaveMetaToExcel).
                // Reuse its 6-row header as the column layout — captioned-attribute
                // columns get left blank below ("brief" = ignore internal levels).
                ws = wb.Worksheet(sheetName);
                var lastCol = ws.LastColumnUsed();
                int maxCol = lastCol?.ColumnNumber() ?? 0;
                layout = new List<(string, string, string)>(maxCol);
                for (int i = 1; i <= maxCol; i++)
                {
                    layout.Add((
                        ws.Cell(1, i).GetString() ?? "",   // kind
                        ws.Cell(3, i).GetString() ?? "",   // level
                        ws.Cell(6, i).GetString() ?? ""    // header / caption
                    ));
                }
            }
            else
            {
                // No existing sheet — create one with the brief layout.
                // Column order: AttributeCode columns lead (one per distinct code
                // collected from ATTRIBUTE rows, regardless of LevelNumber);
                // followed by the nominated direct fields with AssetRegisterName
                // and AssetNumber forced to the front.
                ws = wb.Worksheets.Add(sheetName);
                var briefCols = new List<(string Kind, string Code, string Level, string Suffix, string DataType, string Header)>();
                foreach (var code in attrCodesOrdered)
                    briefCols.Add(("Attribute", "", "", "", "A", code));
                foreach (var f in OrderNominatedFields(_nominatedFields))
                    briefCols.Add(("", "", "", "", "A", f));

                for (int i = 0; i < briefCols.Count; i++)
                {
                    var c = briefCols[i];
                    ws.Cell(1, i + 1).Value = c.Kind;
                    ws.Cell(2, i + 1).Value = c.Code;
                    ws.Cell(3, i + 1).Value = c.Level;
                    ws.Cell(4, i + 1).Value = c.Suffix;
                    ws.Cell(5, i + 1).Value = c.DataType;
                    ws.Cell(6, i + 1).Value = c.Header;
                    ws.Column(i + 1).Style.NumberFormat.Format = MetaSchema.NumberFormatFor(c.DataType);
                }
                layout = briefCols.Select(c => (c.Kind, c.Level, c.Header)).ToList();
            }

            // Data rows start at row 7 — one per asset.
            //   kind=""        + header → direct field         → fill from asset.Fields
            //   kind="Attribute" + level=""  → AttributeCode scalar → fill from asset.Attributes
            //   kind="Attribute" + level="level_N"             → captioned, leave blank
            for (int r = 0; r < assets.Count; r++)
            {
                var asset = assets[r];
                int rowNum = 7 + r;
                for (int i = 0; i < layout.Count; i++)
                {
                    var (kind, level, header) = layout[i];
                    string? val = null;
                    if (kind == "" && !string.IsNullOrEmpty(header))
                        asset.Fields.TryGetValue(header, out val);
                    else if (kind == "Attribute" && string.IsNullOrEmpty(level))
                        asset.Attributes.TryGetValue(header, out val);
                    // else: captioned attribute → leave blank ("brief")

                    if (!string.IsNullOrEmpty(val)) ws.Cell(rowNum, i + 1).Value = val;
                }
            }

            wb.SaveAs(xlsxPath);
            return xlsxPath;
        }

        // ---------- Step 7: highlight leaf-node rows in the asset_type column ----------
        //
        // Opens an existing Template2Flat workbook, finds the asset_type column
        // by scanning row 6 (the header row, case-insensitive), then yellow-fills
        // every data-row cell whose value is a leaf in the hierarchical path.
        //
        // A value V is a "leaf" if no other value in the column extends it with
        // a path separator — i.e. no V' starts with V + '\' or V + '/'. So with
        //   Tree
        //   Tree\Street Tree
        //   Tree\Street Tree\Example Tree
        // only the third row is highlighted.
        public string HighlightLeaf(string xlsxPath, string sheet)
        {
            var sheetName = MetaSchema.SanitizeSheetName(sheet);
            using var wb = new XLWorkbook(xlsxPath);
            if (!wb.Worksheets.Contains(sheetName))
                throw new ArgumentException($"Sheet '{sheetName}' not found in '{xlsxPath}'.");
            var ws = wb.Worksheet(sheetName);

            int maxCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            int colIdx = -1;
            for (int c = 1; c <= maxCol; c++)
            {
                if (string.Equals(ws.Cell(6, c).GetString(), "asset_type", StringComparison.OrdinalIgnoreCase))
                {
                    colIdx = c;
                    break;
                }
            }
            if (colIdx < 0) return xlsxPath;

            int maxRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            if (maxRow < 7) { wb.Save(); return xlsxPath; }

            var rowValues = new List<(int Row, string Value)>();
            for (int r = 7; r <= maxRow; r++)
            {
                var v = ws.Cell(r, colIdx).GetString();
                if (!string.IsNullOrEmpty(v)) rowValues.Add((r, v));
            }

            var allValues = new HashSet<string>(
                rowValues.Select(rv => rv.Value),
                StringComparer.OrdinalIgnoreCase);

            foreach (var (r, v) in rowValues)
            {
                if (IsLeafPath(v, allValues))
                    ws.Cell(r, colIdx).Style.Fill.BackgroundColor = XLColor.Yellow;
            }

            wb.Save();
            return xlsxPath;
        }

        private static bool IsLeafPath(string value, HashSet<string> all)
        {
            foreach (var other in all)
            {
                if (other.Length <= value.Length) continue;
                if (!other.StartsWith(value, StringComparison.OrdinalIgnoreCase)) continue;
                var sep = other[value.Length];
                if (sep == '\\' || sep == '/') return false;
            }
            return true;
        }

        // Walks the CSV row-by-row, grouping each ASSET line with its following
        // ATTRIBUTE lines into one AssetRecord. Returns (assets, ordered list of
        // distinct AttributeCodes encountered, by first appearance).
        private (List<AssetRecord> Assets, List<string> AttrCodes) ReadCsvBrief()
        {
            var rows = ReadCsv(_csvPath);
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
                }
            }

            return (assets, attrCodesOrdered);
        }

        private class AssetRecord
        {
            public Dictionary<string, string> Fields     = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> Attributes = new(StringComparer.OrdinalIgnoreCase);
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

        private static string SaveMetaToExcelStatic(string xlsxPath, Dictionary<string, object> meta)
        {
            var dir = Path.GetDirectoryName(xlsxPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            using var wb = File.Exists(xlsxPath) ? new XLWorkbook(xlsxPath) : new XLWorkbook();

            if (wb.Worksheets.Contains("Sheet"))
                wb.Worksheets.Delete("Sheet");

            foreach (var kvp in meta)
            {
                var nodeName = kvp.Key;
                var sheetName = UniqueSheetName(wb, nodeName);
                var ws = wb.Worksheets.Add(sheetName);
                var columns = BuildMetaColumns(kvp.Value);

                for (int i = 0; i < columns.Count; i++)
                {
                    var col = columns[i];
                    ws.Cell(1, i + 1).Value = col.Item1;   // kind
                    ws.Cell(2, i + 1).Value = col.Item2;   // AttributeCode
                    ws.Cell(3, i + 1).Value = col.Item3;   // level
                    ws.Cell(4, i + 1).Value = col.Item4;   // suffix
                    ws.Cell(5, i + 1).Value = col.Item5;   // dataType
                    ws.Cell(6, i + 1).Value = col.Item6;   // header (caption)

                    var format = col.Item5 switch
                    {
                        "N" => "General",
                        "D" => "yyyy-mm-dd",
                        _   => "@",
                    };
                    ws.Column(i + 1).Style.NumberFormat.Format = format;
                }
            }

            if (!wb.Worksheets.Any()) wb.Worksheets.Add("Sheet");

            wb.SaveAs(xlsxPath);
            return xlsxPath;
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

        // Thin pass-through to MetaSchema (single source of truth for sheet naming).
        private static string SanitizeSheetName(string name) => MetaSchema.SanitizeSheetName(name);

        private static string UniqueSheetName(XLWorkbook wb, string baseName)
        {
            var name = SanitizeSheetName(baseName);
            if (!wb.Worksheets.Contains(name)) return name;
            var stem = name.Length > 29 ? name.Substring(0, 29) : name;
            for (int i = 1; i < 100; i++)
            {
                var candidate = stem + i.ToString("D2");
                if (!wb.Worksheets.Contains(candidate)) return candidate;
            }
            throw new InvalidOperationException($"Could not allocate unique sheet name for '{baseName}'");
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
