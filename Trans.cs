// Trans.cs — convert a T1 ASSET CSV "template" export into flat,
// import-ready shapes. The source CSV groups every asset across multiple
// rows (one ASSET row + N ATTRIBUTE rows); Trans flattens that for editing
// and walks it back to bulk-import shape.
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
//   Template2Flat(sourceCsv, outputCsv, assetTypeOnly = false)
//     Source CSV → flat CSV. Collapses each asset's ASSET + N ATTRIBUTE
//     rows into a single row, with one column per AttributeCode (cell
//     value = that attribute's SearchPath) plus one column per nominated
//     direct field. Output is a plain CSV: one column-header row + one
//     payload row per asset; no LineType column. assetTypeOnly=true keeps
//     only the ASSET_TYPE attribute column.
//   Flat2Import(sourceCsv, outputCsv)
//     Flat CSV → T1 bulk-import CSV. Reverses Template2Flat: re-adds the
//     LineType column and emits one "ASSET" row plus one "ATTRIBUTE" row
//     per non-empty AttributeCode value, in the shape T1's bulk-import
//     accepts.
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

        // ---------- Flat2Import: flat CSV → T1 bulk-import CSV ----------
        //
        // Reverses Template2Flat. Each input row holds one asset's data in
        // a flat shape (AttributeCode columns first, then direct fields);
        // T1's bulk import wants that asset split back into one ASSET row
        // (carrying the direct fields) plus one ATTRIBUTE row per non-empty
        // AttributeCode (carrying that code + its SearchPath value).
        //
        // Input header (saved-as CSV from Template2Flat):
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
        public string Flat2Import(string sourceCsvPath, string outputCsvPath)
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

        // ---------- Template2Flat: source CSV → flat CSV ----------
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
        // Flat2Import re-adds it. assetTypeOnly=true keeps only the
        // ASSET_TYPE attribute column (case-insensitive match).
        public string Template2Flat(string sourceCsvPath, string outputCsvPath, bool assetTypeOnly = false)
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
            // front. No LineType — Flat2Import re-adds it.
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
