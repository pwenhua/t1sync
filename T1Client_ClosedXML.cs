// T1Client_ClosedXML.cs - drop-in alternative to T1Client_Interop that uses
// ClosedXML only (no Microsoft.Office.Interop.Excel).
//
// The 'file' parameter must be a local path; SharePoint/OneDrive URLs are not
// supported here (download/upload them separately if needed).
//
// Usage: replace `new T1Client_Interop("workshop-TP")` with
//                `new T1Client_ClosedXML("workshop-TP")`.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

namespace T1Sync
{
    public class T1Client_ClosedXML : T1Client_Interop
    {
        public T1Client_ClosedXML(string service, string configPath = DefaultConfigPath)
            : base(service, configPath) { }

        public new string SaveMetaToExcel(string file, Dictionary<string, object>? meta = null)
        {
            // Pure ClosedXML version — local file only, no Excel needed.
            // Use this instead of the base T1Client_Interop.SaveMetaToExcel when
            // `file` is a local path (much faster, no COM round-trip).
            if (meta == null)
            {
                var metaPath = MetaPath;
                if (!File.Exists(metaPath))
                    throw new FileNotFoundException($"Metadata file not found: {metaPath}");
                var jsonText = File.ReadAllText(metaPath, Encoding.UTF8);
                meta = JsonSerializer.Deserialize<Dictionary<string, object>>(jsonText);
            }

            var dir = Path.GetDirectoryName(file);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            using var wb = File.Exists(file) ? new XLWorkbook(file) : new XLWorkbook();

            if (wb.Worksheets.Contains("Sheet"))
                wb.Worksheets.Delete("Sheet");

            foreach (var kvp in meta!)
            {
                var sheetName = UniqueSheetNameLocal(wb, kvp.Key);
                var ws = wb.Worksheets.Add(sheetName);
                var columns = BuildMetaColumnsLocal(kvp.Value);

                for (int i = 0; i < columns.Count; i++)
                {
                    var c = columns[i];
                    ws.Cell(1, i + 1).Value = c.Item1;
                    ws.Cell(2, i + 1).Value = c.Item2;
                    ws.Cell(3, i + 1).Value = c.Item3;
                    ws.Cell(4, i + 1).Value = c.Item4;
                    ws.Cell(5, i + 1).Value = c.Item5;
                    ws.Cell(6, i + 1).Value = c.Item6;

                    var fmt = c.Item5 switch
                    {
                        "N" => "General",
                        "D" => "yyyy-mm-dd",
                        _   => "@",
                    };
                    ws.Column(i + 1).Style.NumberFormat.Format = fmt;
                }
            }

            if (!wb.Worksheets.Any()) wb.Worksheets.Add("Sheet");

            wb.SaveAs(file);
            Debug.WriteLine($"Saved spreadsheet to {file}");
            return file;
        }

        public new string SyncAssetFromExcel(string file, string sheet, int firstRow, int lastRow, bool dryrun = false)
        {
            // Same control flow as T1Client_Interop.SyncAssetFromExcel, but `file`
            // is opened directly with ClosedXML (no Excel COM round-trip).
            var xlsxPath = file;
            using var wb = new XLWorkbook(xlsxPath);

            var sheetName = SanitizeSheetNameLocal(sheet);
            if (!wb.Worksheets.Contains(sheetName))
            {
                Debug.WriteLine($"  -> Sheet {sheetName} not found in workbook.");
                return xlsxPath;
            }

            // Sheet name like "Tree_Street Tree" → class name "Tree/Street Tree"
            // (first underscore → '/'); used to look up asset_classes config.
            string trueAssetType = sheet;
            int underscoreIdx = trueAssetType.IndexOf('_');
            if (underscoreIdx >= 0)
            {
                trueAssetType = trueAssetType.Substring(0, underscoreIdx) + "/" + trueAssetType.Substring(underscoreIdx + 1);
            }

            var ws = wb.Worksheet(sheetName);
            var lastUsed = ws.LastColumnUsed();
            var maxCol = lastUsed?.ColumnNumber() ?? 0;

            var headers = new List<(string Kind, string Code, string Level, string Suffix, string Header)>();
            for (int col = 1; col <= maxCol; col++)
            {
                headers.Add((
                    ws.Cell(1, col).GetString() ?? "",
                    ws.Cell(2, col).GetString() ?? "",
                    ws.Cell(3, col).GetString() ?? "",
                    ws.Cell(4, col).GetString() ?? "",
                    ws.Cell(6, col).GetString() ?? ""
                ));
            }

            int? assetNumCol = null;
            int? assetRegCol = null;
            for (int i = 0; i < headers.Count; i++)
            {
                if (headers[i].Header == "AssetNumber") assetNumCol = i + 1;
                else if (headers[i].Header == "AssetRegisterName") assetRegCol = i + 1;
            }

            string? assetRegister = null;
            if (SvcConfig.TryGetProperty("asset_register", out var arProp) ||
                SvcConfig.TryGetProperty("asset register", out arProp))
            {
                assetRegister = arProp.GetString();
            }

            string? templateId = null;
            string? seedId = null;
            if (SvcConfig.TryGetProperty("asset_classes", out var assetClasses) && assetClasses.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in assetClasses.EnumerateObject())
                {
                    if (string.Equals(prop.Name, trueAssetType, StringComparison.OrdinalIgnoreCase))
                    {
                        templateId = prop.Value.TryGetProperty("template", out var tmpProp) ? tmpProp.GetString() : null;
                        seedId = prop.Value.TryGetProperty("seed", out var seedProp) ? seedProp.GetString() : null;
                        break;
                    }
                }
            }

            for (int row = firstRow; row <= lastRow; row++)
            {
                try
                {
                    var assetCell = ws.Cell(row, 2);
                    var assetNumber = assetCell.Value.IsText ? assetCell.GetString() : "";

                    JsonObject node;
                    if (!string.IsNullOrWhiteSpace(assetNumber))
                    {
                        Debug.WriteLine($"  -> {sheetName} row {row}: updating asset {assetNumber}");
                        var asset = FetchAsset(assetNumber);
                        node = JsonNode.Parse(asset.GetRawText())!.AsObject();
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(templateId))
                        {
                            ws.Cell(row, "AA").Value = $"Missing 'template' for class '{trueAssetType}'.";
                            continue;
                        }
                        if (string.IsNullOrEmpty(seedId))
                        {
                            ws.Cell(row, "AA").Value = $"Missing 'seed' for class '{trueAssetType}'.";
                            continue;
                        }

                        Debug.WriteLine($"  -> {sheetName} row {row}: creating asset from template {templateId}");

                        var seedAsset = FetchAsset(seedId);
                        var seedNode = JsonNode.Parse(seedAsset.GetRawText())!.AsObject();

                        seedNode["AssetRegisterName"] = assetRegister;
                        seedNode["TemplateAssetNumberInternal"] = templateId;
                        seedNode["AssetNumber"] = null;

                        string? newAssetNumber = null;
                        string? newAssetRegister = assetRegister;

                        if (!dryrun)
                        {
                            var result = SaveAsset(seedNode.ToJsonString(), "ep_asset_create");
                            newAssetNumber = result.TryGetProperty("AssetNumber", out var anProp) ? anProp.GetString() : null;
                            if (string.IsNullOrEmpty(newAssetNumber))
                            {
                                ws.Cell(row, "AA").Value = "Create returned no AssetNumber.";
                                continue;
                            }
                            newAssetRegister = result.TryGetProperty("AssetRegisterName", out var arNameProp) ? arNameProp.GetString() : assetRegister;
                            if (assetNumCol.HasValue) SetCellValueLocal(ws.Cell(row, assetNumCol.Value), newAssetNumber);
                            if (assetRegCol.HasValue && !string.IsNullOrEmpty(newAssetRegister))
                            {
                                SetCellValueLocal(ws.Cell(row, assetRegCol.Value), newAssetRegister);
                            }
                        }
                        else
                        {
                            newAssetNumber = $"DRYRUN_NEW_ROW_{row}";
                        }

                        // Build the update payload from the original seed with seed_id retargeted.
                        var seedStr = seedAsset.GetRawText().Replace(seedId, newAssetNumber);
                        node = JsonNode.Parse(seedStr)!.AsObject();
                        node["AssetNumber"] = newAssetNumber;
                        if (!string.IsNullOrEmpty(newAssetRegister)) node["AssetRegisterName"] = newAssetRegister;
                    }

                    for (int colIdx = 0; colIdx < headers.Count; colIdx++)
                    {
                        var (kind, code, level, suffix, header) = headers[colIdx];
                        var cell = ws.Cell(row, colIdx + 1);
                        var cellValue = ReadCellValueLocal(cell);

                        // Only the top-level ASSET_TYPE column (row-6 header == "ASSET_TYPE")
                        // gets forced to trueAssetType. Captioned columns have code == "ASSET_TYPE"
                        // but their headers are captions.
                        bool isAssetType = string.Equals(header, "asset_type", StringComparison.OrdinalIgnoreCase) ||
                                           string.Equals(header, "AssetType", StringComparison.OrdinalIgnoreCase);

                        if (isAssetType)
                        {
                            var cellValueStr = cellValue?.ToString() ?? "";
                            if (!string.Equals(cellValueStr, trueAssetType, StringComparison.OrdinalIgnoreCase))
                            {
                                cell.Style.Fill.BackgroundColor = XLColor.Yellow;
                            }
                            cellValue = trueAssetType;
                        }

                        if (cellValue == null || (cellValue is string s && string.IsNullOrEmpty(s))) continue;

                        if (kind == "Attribute")
                        {
                            if (!string.IsNullOrEmpty(code) && !string.IsNullOrEmpty(suffix) && level.StartsWith("level_"))
                            {
                                SetAttributeValueLocal(node, code, level, suffix, cellValue);
                            }
                        }
                        else if (!string.IsNullOrEmpty(header))
                        {
                            if (header.Equals("AssetRegisterName", StringComparison.OrdinalIgnoreCase)) continue;
                            node[header] = JsonSerializer.SerializeToNode(cellValue);
                        }
                    }

                    if (!string.IsNullOrEmpty(assetRegister))
                    {
                        node["AssetRegisterName"] = assetRegister;
                    }

                    if (dryrun)
                    {
                        Directory.CreateDirectory(@"c:\temp");
                        var dumpPath = @"c:\temp\payload.txt";
                        File.WriteAllText(dumpPath, node.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
                        ws.Cell(row, "AA").Value = $"Dry run: Payload saved to {Path.GetFileName(dumpPath)}";
                    }
                    else
                    {
                        SaveAsset(node.ToJsonString());
                        ws.Cell(row, "AA").Value = "";
                    }
                }
                catch (Exception ex)
                {
                    ws.Cell(row, "AA").Value = ex.Message;
                }
            }

            wb.Save();
            Debug.WriteLine($"Synced spreadsheet at {xlsxPath}");
            return xlsxPath;
        }

        // Pulls live values from T1 into the named sheet, one row per asset.
        // Geometry (the AssetMap WKT) is always extracted; the WKT lands in
        // a row-6 header named "WKT" (case-insensitive). No DB.
        public new string ExtractAssetToExcel(string file, string sheet, int firstRow, int lastRow)
            => ExtractAssetCore(file, sheet, firstRow, lastRow, databaseInstance: null);

        // Same as ExtractAssetToExcel for the scalar fields, plus the
        // geometry (WKT + sp_geometry) gets upserted to the SQL table
        // configured under config.database[databaseInstance]. Always extracts
        // geometry; throws if the database instance can't be opened.
        public new string ExtractAssetToDB(string file, string sheet, int firstRow, int lastRow, string databaseInstance)
            => ExtractAssetCore(file, sheet, firstRow, lastRow, databaseInstance);

        private string ExtractAssetCore(string file, string sheet, int firstRow, int lastRow, string? databaseInstance)
        {
            var xlsxPath = file;
            using var wb = new XLWorkbook(xlsxPath);

            var sheetName = SanitizeSheetNameLocal(sheet);
            if (!wb.Worksheets.Contains(sheetName))
            {
                Debug.WriteLine($"  -> Sheet {sheetName} not found in workbook.");
                return xlsxPath;
            }

            var ws = wb.Worksheet(sheetName);
            var lastUsed = ws.LastColumnUsed();
            var maxCol = lastUsed?.ColumnNumber() ?? 0;

            var headers = new List<(string, string, string, string, string)>();
            for (int col = 1; col <= maxCol; col++)
            {
                headers.Add((
                    ws.Cell(2, col).GetString() ?? "",
                    ws.Cell(3, col).GetString() ?? "",
                    ws.Cell(4, col).GetString() ?? "",
                    ws.Cell(5, col).GetString() ?? "",
                    ws.Cell(6, col).GetString() ?? ""
                ));
            }

            int? assetNumCol = null;
            int? wktCol = null;
            for (int i = 0; i < headers.Count; i++)
            {
                if (assetNumCol == null && headers[i].Item5.Equals("AssetNumber", StringComparison.OrdinalIgnoreCase))
                    assetNumCol = i + 1;
                if (wktCol == null && headers[i].Item5.Equals("WKT", StringComparison.OrdinalIgnoreCase))
                    wktCol = i + 1;
            }

            if (!assetNumCol.HasValue)
            {
                Debug.WriteLine($"  -> No 'AssetNumber' header found in sheet {sheetName}.");
                return xlsxPath;
            }

            // ExtractAssetToDB requires a usable DB; ExtractAssetToExcel passes null.
            SqlConnection? dbConn = null;
            string? dbTable = null;
            if (!string.IsNullOrEmpty(databaseInstance))
            {
                using var stream = File.OpenRead(ConfigPath);
                using var doc = JsonDocument.Parse(stream);
                if (doc.RootElement.TryGetProperty("database", out var dbRoot) &&
                    dbRoot.TryGetProperty(databaseInstance, out var dbInst) &&
                    dbInst.TryGetProperty("connection_string", out var connStrProp) &&
                    dbInst.TryGetProperty("table", out var tableProp))
                {
                    dbTable = tableProp.GetString();
                    dbConn = new SqlConnection(connStrProp.GetString());
                    dbConn.Open();
                    Debug.WriteLine($"  -> DB '{databaseInstance}' connected; geometry → {dbTable}");
                }
                else
                {
                    throw new InvalidOperationException(
                        $"DB instance '{databaseInstance}' not found / incomplete in config.database.");
                }
            }

            try
            {
                for (int row = firstRow; row <= lastRow; row++)
                {
                    var assetNumber = ws.Cell(row, assetNumCol.Value).GetString();
                    if (string.IsNullOrEmpty(assetNumber)) continue;

                    try
                    {
                        Debug.WriteLine($"  -> {ws.Name} row {row}: fetching asset {assetNumber}");
                        var asset = FetchAsset(assetNumber);

                        for (int colIdx = 0; colIdx < headers.Count; colIdx++)
                        {
                            var (attrCode, level, suffix, _, header) = headers[colIdx];
                            if (header.Equals("WKT", StringComparison.OrdinalIgnoreCase)) continue;

                            var val = ExtractValueLocal(asset, attrCode, level, suffix, header);
                            if (val != null)
                            {
                                SetCellValueLocal(ws.Cell(row, colIdx + 1), val);
                            }
                        }

                        // Geometry is always extracted; destination depends on the mode.
                        var wkt = ExtractGeometryWkt(asset);
                        if (wkt != null && wktCol.HasValue)
                            SetCellValueLocal(ws.Cell(row, wktCol.Value), wkt);

                        string? dbError = null;
                        if (dbConn != null && !string.IsNullOrEmpty(dbTable) && wkt != null)
                        {
                            try
                            {
                                WriteGeometryToDb(wkt, assetNumber, dbConn, dbTable);
                            }
                            catch (Exception dbEx)
                            {
                                dbError = "DB: " + dbEx.Message;
                            }
                        }

                        ws.Cell(row, "AA").Value = dbError ?? "";
                    }
                    catch (Exception ex)
                    {
                        ws.Cell(row, "AA").Value = ex.Message;
                    }
                }
            }
            finally
            {
                dbConn?.Dispose();
            }

            wb.Save();
            Debug.WriteLine($"Updated spreadsheet at {xlsxPath}");
            return xlsxPath;
        }

        // Collect ready-made WKT strings from AssetMap.MapLayers[*].Geometries[*].WKT.
        // Multiple geometries are wrapped in a GEOMETRYCOLLECTION. Returns null
        // when the asset has no geometry.
        private static string? ExtractGeometryWkt(JsonElement asset)
        {
            var wkts = new List<string>();
            if (asset.TryGetProperty("AssetMap", out var assetMap) &&
                assetMap.TryGetProperty("MapLayers", out var mapLayers) &&
                mapLayers.ValueKind == JsonValueKind.Array)
            {
                foreach (var layer in mapLayers.EnumerateArray())
                {
                    if (!layer.TryGetProperty("Geometries", out var geoms) || geoms.ValueKind != JsonValueKind.Array)
                        continue;
                    foreach (var g in geoms.EnumerateArray())
                    {
                        if (g.TryGetProperty("WKT", out var wktProp) && wktProp.ValueKind == JsonValueKind.String)
                        {
                            var w = wktProp.GetString();
                            if (!string.IsNullOrWhiteSpace(w)) wkts.Add(w);
                        }
                    }
                }
            }
            if (wkts.Count == 0) return null;
            return wkts.Count == 1
                ? wkts[0]
                : "GEOMETRYCOLLECTION(" + string.Join(", ", wkts) + ")";
        }

        // Idempotent upsert: delete the existing row for this compkey, then
        // insert wkt + sp_geometry (parsed via geometry::STGeomFromText, SRID 4326).
        private static void WriteGeometryToDb(string wkt, string assetNumber, SqlConnection conn, string table)
        {
            using (var del = new SqlCommand($"DELETE FROM {table} WHERE compkey = @compkey", conn))
            {
                del.Parameters.AddWithValue("@compkey", assetNumber);
                del.ExecuteNonQuery();
            }
            using (var ins = new SqlCommand(
                $"INSERT INTO {table} (compkey, wkt, sp_geometry) " +
                "VALUES (@compkey, @wkt, geometry::STGeomFromText(@wkt, 4326))", conn))
            {
                ins.Parameters.AddWithValue("@compkey", assetNumber);
                ins.Parameters.AddWithValue("@wkt", wkt);
                ins.ExecuteNonQuery();
            }
        }

        // ------- Local copies of the private helpers from T1Client_Interop (so this file is self-contained) -------

        // Thin pass-through to the shared MetaSchema.BuildColumns helper —
        // guarantees byte-identical column layout to T1Client_Interop / Trans.
        private static List<(string, string, string, string, string, string)> BuildMetaColumnsLocal(object nodeMetaObj)
        {
            var typed = MetaSchema.BuildColumns(nodeMetaObj);
            var columns = new List<(string, string, string, string, string, string)>(typed.Count);
            foreach (var c in typed) columns.Add((c.Kind, c.Code, c.Level, c.Suffix, c.DataType, c.Header));
            return columns;
        }

        private static string UniqueSheetNameLocal(XLWorkbook wb, string baseName)
        {
            var name = SanitizeSheetNameLocal(baseName);
            if (!wb.Worksheets.Contains(name)) return name;
            var stem = name.Length > 29 ? name.Substring(0, 29) : name;
            for (int i = 1; i < 100; i++)
            {
                var cand = stem + i.ToString("D2");
                if (!wb.Worksheets.Contains(cand)) return cand;
            }
            throw new InvalidOperationException($"Could not allocate unique sheet name for '{baseName}'");
        }

        // Thin pass-through to MetaSchema (single source of truth for sheet naming).
        private static string SanitizeSheetNameLocal(string name) => MetaSchema.SanitizeSheetName(name);

        private static void SetAttributeValueLocal(JsonObject asset, string attrCode, string level, string suffix, object? value)
        {
            if (!level.StartsWith("level_")) return;
            if (!int.TryParse(level.Substring("level_".Length), out var targetLevel)) return;

            if (asset["AssetAttributes"] is not JsonArray attrs) return;
            var valueKey = "AttributeItem" + suffix;

            foreach (var item in attrs)
            {
                if (item is not JsonObject entry) continue;
                if ((string?)entry["AttributeCode"] != attrCode) continue;
                var sp = (string?)entry["SearchPath"] ?? "";
                var entryLevel = string.IsNullOrEmpty(sp) ? 0 : sp.Split('\\').Length;
                if (entryLevel != targetLevel) continue;
                if (!entry.ContainsKey(valueKey)) continue;

                entry[valueKey] = value == null ? null : JsonSerializer.SerializeToNode(value);
                return;
            }
        }

        private static object? ReadCellValueLocal(IXLCell cell)
        {
            var v = cell.Value;
            if (v.IsBlank) return null;
            if (v.IsText) return v.GetText();
            if (v.IsNumber) return v.GetNumber();
            if (v.IsBoolean) return v.GetBoolean();
            if (v.IsDateTime) return v.GetDateTime().ToString("o");
            return cell.GetString();
        }

        private static void SetCellValueLocal(IXLCell cell, object? value)
        {
            switch (value)
            {
                case null: cell.Value = ""; break;
                case string s: cell.Value = s; break;
                case bool b: cell.Value = b; break;
                case int i: cell.Value = i; break;
                case long l: cell.Value = l; break;
                case double d: cell.Value = d; break;
                case decimal m: cell.Value = m; break;
                case DateTime dt: cell.Value = dt; break;
                default: cell.Value = value.ToString() ?? ""; break;
            }
        }

        private static object? JsonElementToValueLocal(JsonElement el)
        {
            return el.ValueKind switch
            {
                JsonValueKind.String => el.GetString(),
                JsonValueKind.Number => el.TryGetInt64(out var l) ? (object)l : el.GetDouble(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null => null,
                _ => el.GetRawText(),
            };
        }

        private static object? ExtractValueLocal(JsonElement asset, string attrCode, string level, string suffix, string header)
        {
            // Captioned attribute — attrCode + suffix populated.
            if (!string.IsNullOrEmpty(attrCode) && !string.IsNullOrEmpty(suffix))
            {
                int targetLevel = 0;
                if (level.StartsWith("level_") && int.TryParse(level.Substring("level_".Length), out var lv))
                    targetLevel = lv;
                var valueKey = "AttributeItem" + suffix;

                if (asset.TryGetProperty("AssetAttributes", out var attrs) && attrs.ValueKind == JsonValueKind.Array)
                {
                    foreach (var entry in attrs.EnumerateArray())
                    {
                        if (!entry.TryGetProperty("AttributeCode", out var cp) || cp.GetString() != attrCode)
                            continue;
                        var sp = entry.TryGetProperty("SearchPath", out var spProp) && spProp.ValueKind == JsonValueKind.String
                            ? spProp.GetString()! : "";
                        var entryLevel = string.IsNullOrEmpty(sp) ? 0 : sp.Split('\\').Length;
                        if (entryLevel != targetLevel) continue;
                        if (entry.TryGetProperty(valueKey, out var v))
                            return JsonElementToValueLocal(v);
                    }
                }
                return null;
            }

            // Root field (top-level on the asset payload).
            // Source list: T1Client_Interop.RootFields (loaded from config.json).
            if (Array.IndexOf(T1Client_Interop.RootFields, header) >= 0)
            {
                return asset.TryGetProperty(header, out var v) ? JsonElementToValueLocal(v) : null;
            }

            // AttributeCode top-level scalar (LOCATION, SERVICEAREA, ASSET_TYPE, ...).
            if (asset.TryGetProperty("AssetAttributes", out var attrs2) && attrs2.ValueKind == JsonValueKind.Array)
            {
                foreach (var entry in attrs2.EnumerateArray())
                {
                    if (!entry.TryGetProperty("AttributeCode", out var cp) || cp.GetString() != header)
                        continue;
                    if (!entry.TryGetProperty("IsPrimaryValue", out var pv) || pv.ValueKind != JsonValueKind.True)
                        continue;
                    return entry.TryGetProperty("SearchPath", out var spProp) && spProp.ValueKind == JsonValueKind.String
                        ? spProp.GetString()! : "";
                }
            }
            return null;
        }
    }
}
