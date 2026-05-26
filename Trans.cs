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
        // data). Template2FlatBrief produces a data workbook:
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
        // ---------- Step 7: CSV → cleaned template (nominated columns, compacted) ----------
        //
        // Output layout:
        //   Row 1   — verbatim CSV line 1 (the FORMAT line); typically only A1
        //             carries content so the cell ends up in A1 either way.
        //   Row 2   — only the nominated column names, COMPACTED into columns
        //             A, B, C… (no gap columns, even if the original CSV had
        //             non-nominated fields scattered between them).
        //   Row 3+  — one row per ASSET LineType in the source CSV, values
        //             aligned with row 2's compact positions.
        // The output is narrow (≈ #nominated_fields + 1 columns).
        public string Simplify0(string xlsxPath, string sheet)
        {
            var rows = ReadCsv(_csvPath);
            if (rows.Count < 2) return xlsxPath;

            var formatLine = rows[0];   // CSV line 1
            var headerLine = rows[1];   // CSV line 2 — full column header

            var headerIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headerLine.Count; i++)
            {
                if (!string.IsNullOrEmpty(headerLine[i]) && !headerIndex.ContainsKey(headerLine[i]))
                    headerIndex[headerLine[i]] = i;
            }

            int lineTypeIdx = headerIndex.GetValueOrDefault("LineType", -1);

            // Ordered list of CSV column indices to keep: LineType first, then
            // nominated fields in their declared order. Output cell order
            // follows this list (column 1 = LineType, column 2 = first nominated, …).
            var keepIndices = new List<int>();
            if (lineTypeIdx >= 0) keepIndices.Add(lineTypeIdx);
            foreach (var f in _nominatedFields)
            {
                if (headerIndex.TryGetValue(f, out var idx) && !keepIndices.Contains(idx))
                    keepIndices.Add(idx);
            }

            // Filter to ASSET rows only.
            var assets = new List<List<string>>();
            if (lineTypeIdx >= 0)
            {
                foreach (var row in rows.Skip(2))
                {
                    if (lineTypeIdx < row.Count &&
                        row[lineTypeIdx].Equals("ASSET", StringComparison.OrdinalIgnoreCase))
                    {
                        assets.Add(row);
                    }
                }
            }

            var dir = Path.GetDirectoryName(xlsxPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            using var wb = File.Exists(xlsxPath) ? new XLWorkbook(xlsxPath) : new XLWorkbook();
            if (wb.Worksheets.Contains("Sheet")) wb.Worksheets.Delete("Sheet");

            var sheetName = UniqueSheetName(wb, sheet);
            var ws = wb.Worksheets.Add(sheetName);

            // Row 1 — verbatim original first row. CSV line 1 normally only has
            // content in cell A1 (the FORMAT string), so it lands in A1.
            for (int c = 0; c < formatLine.Count; c++)
                if (!string.IsNullOrEmpty(formatLine[c]))
                    ws.Cell(1, c + 1).Value = formatLine[c];

            // Row 2 — nominated column names, compacted into A, B, C…
            for (int i = 0; i < keepIndices.Count; i++)
            {
                int srcIdx = keepIndices[i];
                if (srcIdx < headerLine.Count && !string.IsNullOrEmpty(headerLine[srcIdx]))
                    ws.Cell(2, i + 1).Value = headerLine[srcIdx];
            }

            // Row 3+ — ASSET data, same compacted positions as row 2.
            for (int r = 0; r < assets.Count; r++)
            {
                var asset = assets[r];
                int rowNum = 3 + r;
                for (int i = 0; i < keepIndices.Count; i++)
                {
                    int srcIdx = keepIndices[i];
                    if (srcIdx < asset.Count && !string.IsNullOrEmpty(asset[srcIdx]))
                        ws.Cell(rowNum, i + 1).Value = asset[srcIdx];
                }
            }

            wb.SaveAs(xlsxPath);
            return xlsxPath;
        }

        // ---------- Step 6: CSV → ultra-compact flat Excel (nominated fields only) ----------
        //
        // Even simpler than Flat1: the output has ONLY the nominated
        // direct-field columns — no CSV-template padding, no T1Sync 6-row
        // header. Just one header row + one data row per ASSET.
        //
        //   Row 1   — nominated field names (e.g. AssetRegisterName, AssetNumber, …)
        //   Row 2+  — one row per ASSET LineType in the source CSV.
        public string Flat2(string xlsxPath, string sheet)
        {
            var rows = ReadCsv(_csvPath);
            if (rows.Count < 2) return xlsxPath;

            var headerLine = rows[1];
            var headerIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headerLine.Count; i++)
            {
                if (!string.IsNullOrEmpty(headerLine[i]) && !headerIndex.ContainsKey(headerLine[i]))
                    headerIndex[headerLine[i]] = i;
            }

            int lineTypeIdx = headerIndex.GetValueOrDefault("LineType", -1);

            var assets = new List<List<string>>();
            if (lineTypeIdx >= 0)
            {
                foreach (var row in rows.Skip(2))
                {
                    if (lineTypeIdx < row.Count &&
                        row[lineTypeIdx].Equals("ASSET", StringComparison.OrdinalIgnoreCase))
                    {
                        assets.Add(row);
                    }
                }
            }

            var dir = Path.GetDirectoryName(xlsxPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            using var wb = File.Exists(xlsxPath) ? new XLWorkbook(xlsxPath) : new XLWorkbook();
            if (wb.Worksheets.Contains("Sheet")) wb.Worksheets.Delete("Sheet");

            var sheetName = UniqueSheetName(wb, sheet);
            var ws = wb.Worksheets.Add(sheetName);

            // Row 1: just the nominated field names. Columns formatted as text
            // by default so values like "0100038" keep their leading zeros.
            for (int i = 0; i < _nominatedFields.Count; i++)
            {
                ws.Cell(1, i + 1).Value = _nominatedFields[i];
                ws.Column(i + 1).Style.NumberFormat.Format = "@";
            }

            // Row 2+: one row per ASSET.
            for (int r = 0; r < assets.Count; r++)
            {
                var asset = assets[r];
                int rowNum = 2 + r;
                for (int i = 0; i < _nominatedFields.Count; i++)
                {
                    var f = _nominatedFields[i];
                    if (headerIndex.TryGetValue(f, out var idx) &&
                        idx < asset.Count && !string.IsNullOrEmpty(asset[idx]))
                    {
                        ws.Cell(rowNum, i + 1).Value = asset[idx];
                    }
                }
            }

            wb.SaveAs(xlsxPath);
            return xlsxPath;
        }

        // ---------- Step 5: CSV → simple Excel that mirrors the CSV template ----------
        //
        // Output layout:
        //   Row 1     — verbatim copy of CSV line 1 (the FORMAT line).
        //   Row 2     — verbatim copy of CSV line 2 (the full column header,
        //               every CSV column included so the template shape is
        //               preserved for round-trip with T1's CSV importer).
        //   Row 3+    — one row per ASSET LineType in the source CSV.
        //               Only the nominated direct-field columns are populated;
        //               every other column is left blank. ATTRIBUTE rows are
        //               dropped entirely.
        public string Flat1(string xlsxPath, string sheet)
        {
            var rows = ReadCsv(_csvPath);
            if (rows.Count < 2) return xlsxPath;

            var formatLine = rows[0];   // CSV row 1
            var headerLine = rows[1];   // CSV row 2 — full column header

            // Map header name → column index (CSV column position).
            var headerIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headerLine.Count; i++)
            {
                if (!string.IsNullOrEmpty(headerLine[i]) && !headerIndex.ContainsKey(headerLine[i]))
                    headerIndex[headerLine[i]] = i;
            }

            int lineTypeIdx = headerIndex.GetValueOrDefault("LineType", -1);

            // Indices of columns to populate in data rows: LineType + nominated.
            var keepIndices = new HashSet<int>();
            if (lineTypeIdx >= 0) keepIndices.Add(lineTypeIdx);
            foreach (var f in _nominatedFields)
                if (headerIndex.TryGetValue(f, out var idx)) keepIndices.Add(idx);

            // Filter to ASSET rows only.
            var assets = new List<List<string>>();
            if (lineTypeIdx >= 0)
            {
                foreach (var row in rows.Skip(2))
                {
                    if (lineTypeIdx < row.Count &&
                        row[lineTypeIdx].Equals("ASSET", StringComparison.OrdinalIgnoreCase))
                    {
                        assets.Add(row);
                    }
                }
            }

            var dir = Path.GetDirectoryName(xlsxPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            using var wb = File.Exists(xlsxPath) ? new XLWorkbook(xlsxPath) : new XLWorkbook();
            if (wb.Worksheets.Contains("Sheet")) wb.Worksheets.Delete("Sheet");

            var sheetName = UniqueSheetName(wb, sheet);
            var ws = wb.Worksheets.Add(sheetName);

            // Row 1 — verbatim FORMAT line.
            for (int c = 0; c < formatLine.Count; c++)
                if (!string.IsNullOrEmpty(formatLine[c]))
                    ws.Cell(1, c + 1).Value = formatLine[c];

            // Row 2 — verbatim full CSV column header.
            for (int c = 0; c < headerLine.Count; c++)
                if (!string.IsNullOrEmpty(headerLine[c]))
                    ws.Cell(2, c + 1).Value = headerLine[c];

            // Row 3+ — ASSET data, only nominated columns populated.
            for (int r = 0; r < assets.Count; r++)
            {
                var asset = assets[r];
                int rowNum = 3 + r;
                foreach (int c in keepIndices)
                {
                    if (c < asset.Count && !string.IsNullOrEmpty(asset[c]))
                        ws.Cell(rowNum, c + 1).Value = asset[c];
                }
            }

            wb.SaveAs(xlsxPath);
            return xlsxPath;
        }

        public string Template2FlatBrief(string xlsxPath, string sheet)
        {
            var (assets, attrCodesOrdered) = ReadCsvBrief();

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
                // No existing sheet — create one with the brief layout
                // (nominated direct fields + one column per AttributeCode).
                ws = wb.Worksheets.Add(sheetName);
                var briefCols = new List<(string Kind, string Code, string Level, string Suffix, string DataType, string Header)>();
                foreach (var f in _nominatedFields)
                    briefCols.Add(("", "", "", "", "A", f));
                foreach (var code in attrCodesOrdered)
                    briefCols.Add(("Attribute", "", "", "", "A", code));

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
