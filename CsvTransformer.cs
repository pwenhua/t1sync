// CsvTransformer.cs — read a T1 ASSET CSV export and produce the same meta
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
//   var t = new CsvTransformer("ASSET_Export_25052026-011611.csv",
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
    public class CsvTransformer
    {
        private readonly string _csvPath;
        private readonly List<string> _nominatedFields = new();

        public CsvTransformer(string csvPath, params string[] nominatedFields)
        {
            _csvPath = csvPath;
            if (nominatedFields != null) _nominatedFields.AddRange(nominatedFields);
        }

        /// <summary>
        /// Build a CsvTransformer with `nominated_fields` loaded from
        /// config.json (under t1ws.&lt;service&gt;.nominated_fields).
        /// </summary>
        public static CsvTransformer FromConfig(
            string csvPath,
            string service,
            string configPath = T1Client_Interop.DefaultConfigPath)
        {
            var fields = LoadNominatedFromConfig(service, configPath);
            return new CsvTransformer(csvPath, fields.ToArray());
        }

        private static List<string> LoadNominatedFromConfig(string service, string configPath)
        {
            var fields = new List<string>();
            using var stream = File.OpenRead(configPath);
            using var doc = JsonDocument.Parse(stream);
            var root = doc.RootElement;

            // Prefer t1ws.<service>.nominated_fields; fall back to top-level <service>.nominated_fields
            JsonElement svc;
            if (root.TryGetProperty("t1ws", out var t1ws) && t1ws.TryGetProperty(service, out svc))
            {
                // ok
            }
            else if (!root.TryGetProperty(service, out svc))
            {
                return fields;
            }

            if (svc.TryGetProperty("nominated_fields", out var nf) && nf.ValueKind == JsonValueKind.Array)
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

        // 6-row column tuples (kind, AttributeCode, level, suffix, dataType, header)
        // from the hierarchical node_meta shape. Matches T1Client_Interop.BuildMetaColumns.
        private static List<(string, string, string, string, string, string)> BuildMetaColumns(object nodeMetaObj)
        {
            var columns = new List<(string, string, string, string, string, string)>();
            JsonElement root;

            if (nodeMetaObj is JsonElement je) root = je;
            else
            {
                var json = JsonSerializer.Serialize(nodeMetaObj);
                root = JsonDocument.Parse(json).RootElement;
            }

            if (root.TryGetProperty("fields", out var fieldsNode) && fieldsNode.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in fieldsNode.EnumerateObject())
                {
                    columns.Add(("", "", "", "", prop.Value.ToString() ?? "", prop.Name));
                }
            }

            if (root.TryGetProperty("attributes", out var attrsNode) && attrsNode.ValueKind == JsonValueKind.Object)
            {
                foreach (var attrProp in attrsNode.EnumerateObject())
                {
                    var attrCode = attrProp.Name;
                    var attrNode = attrProp.Value;
                    if (attrNode.ValueKind != JsonValueKind.Object) continue;

                    var dataType = attrNode.TryGetProperty("dataType", out var dtProp) && dtProp.ValueKind == JsonValueKind.String
                        ? dtProp.GetString()! : "A";
                    columns.Add(("Attribute", "", "", "", dataType, attrCode));

                    if (!attrNode.TryGetProperty("levels", out var levelsNode) || levelsNode.ValueKind != JsonValueKind.Object)
                        continue;

                    var sortedLevels = levelsNode.EnumerateObject()
                        .OrderBy(p => int.TryParse(p.Name, out var lvl) ? lvl : 0);

                    foreach (var levelProp in sortedLevels)
                    {
                        var levelKey = levelProp.Name;
                        if (levelProp.Value.ValueKind != JsonValueKind.Object) continue;
                        foreach (var suffixProp in levelProp.Value.EnumerateObject())
                        {
                            var suffix = suffixProp.Name;
                            var leaf = suffixProp.Value;
                            string caption = "", leafDt = "";
                            if (leaf.ValueKind == JsonValueKind.Array && leaf.GetArrayLength() >= 2)
                            {
                                caption = leaf[0].ToString();
                                leafDt = leaf[1].ToString();
                            }
                            else if (leaf.ValueKind == JsonValueKind.Array && leaf.GetArrayLength() == 1)
                            {
                                caption = leaf[0].ToString();
                            }
                            columns.Add(("Attribute", attrCode, $"level_{levelKey}", suffix, leafDt, caption));
                        }
                    }
                }
            }
            return columns;
        }

        private static string SanitizeSheetName(string name)
        {
            var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
            var sb = new StringBuilder(name.Length);
            foreach (var c in name) sb.Append(invalid.Contains(c) ? '_' : c);
            var result = sb.ToString();
            if (result.Length > 31) result = result.Substring(0, 31);
            return string.IsNullOrEmpty(result) ? "Sheet" : result;
        }

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
