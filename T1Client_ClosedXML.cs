// T1Client_ClosedXML.cs - drop-in alternative to T1Client that uses ClosedXML
// only (no Microsoft.Office.Interop.Excel / no OnlineExcelHelper).
//
// The 'file' parameter must be a local path; SharePoint/OneDrive URLs are not
// supported here (download/upload them separately if needed).
//
// Usage: replace `new T1Client("workshop-TP")` with
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

namespace T1Sync
{
    public class T1Client_ClosedXML : T1Client
    {
        public T1Client_ClosedXML(string service, string configPath = DefaultConfigPath)
            : base(service, configPath) { }

        public new string SyncAssetFromExcel(string file, string sheet, int firstRow, int lastRow)
        {
            // Same control flow as T1Client.SyncAssetFromExcel, minus the
            // OnlineExcelHelper round-trip. `file` is opened directly with ClosedXML.
            //
            // Create flow:
            //   1) POST minimal {AssetRegisterName, TemplateAssetNumberInternal} to
            //      ep_asset_create → returns the new AssetNumber.
            //   2) FetchAsset(newAssetNumber) — use the freshly-created asset as the seed.
            //   3) Apply cell values on top.
            //   4) POST via SaveAsset (ep_asset_save) to persist updates.
            //   5) Always export the final payload to c:\temp\payload.txt for inspection.
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
            if (SvcConfig.TryGetProperty("asset_classes", out var assetClasses) && assetClasses.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in assetClasses.EnumerateObject())
                {
                    if (string.Equals(prop.Name, trueAssetType, StringComparison.OrdinalIgnoreCase))
                    {
                        templateId = prop.Value.TryGetProperty("template", out var tmpProp) ? tmpProp.GetString() : null;
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

                        Debug.WriteLine($"  -> {sheetName} row {row}: creating asset from template {templateId}");

                        // 1) Create from template with a minimal payload.
                        var createPayload = new Dictionary<string, string?>
                        {
                            ["AssetRegisterName"] = assetRegister,
                            ["TemplateAssetNumberInternal"] = templateId,
                        };
                        var result = SaveAsset(createPayload, "ep_asset_create");
                        var newAssetNumber = result.TryGetProperty("AssetNumber", out var anProp) ? anProp.GetString() : null;
                        if (string.IsNullOrEmpty(newAssetNumber))
                        {
                            ws.Cell(row, "AA").Value = "Create returned no AssetNumber.";
                            continue;
                        }
                        var newAssetRegister = result.TryGetProperty("AssetRegisterName", out var arNameProp) ? arNameProp.GetString() : assetRegister;
                        if (assetNumCol.HasValue) SetCellValueLocal(ws.Cell(row, assetNumCol.Value), newAssetNumber);
                        if (assetRegCol.HasValue && !string.IsNullOrEmpty(newAssetRegister))
                        {
                            SetCellValueLocal(ws.Cell(row, assetRegCol.Value), newAssetRegister);
                        }

                        // 2) Fetch the just-created asset; use it as the seed payload.
                        var createdAsset = FetchAsset(newAssetNumber);
                        node = JsonNode.Parse(createdAsset.GetRawText())!.AsObject();
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

                    // Always dump the final payload to c:\temp\payload.txt for inspection.
                    Directory.CreateDirectory(@"c:\temp");
                    var dumpPath = @"c:\temp\payload.txt";
                    File.WriteAllText(dumpPath, node.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));

                    SaveAsset(node.ToJsonString());
                    ws.Cell(row, "AA").Value = $"Saved; payload dumped to {Path.GetFileName(dumpPath)}";
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

        public new string ExtractAsset(string file, string sheet, int firstRow, int lastRow)
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
            for (int i = 0; i < headers.Count; i++)
            {
                if (headers[i].Item5.Equals("AssetNumber", StringComparison.OrdinalIgnoreCase))
                {
                    assetNumCol = i + 1;
                    break;
                }
            }

            if (!assetNumCol.HasValue)
            {
                Debug.WriteLine($"  -> No 'AssetNumber' header found in sheet {sheetName}.");
                return xlsxPath;
            }

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
                        var val = ExtractValueLocal(asset, attrCode, level, suffix, header);
                        if (val != null)
                        {
                            SetCellValueLocal(ws.Cell(row, colIdx + 1), val);
                        }
                    }
                    ws.Cell(row, "AA").Value = "";
                }
                catch (Exception ex)
                {
                    ws.Cell(row, "AA").Value = ex.Message;
                }
            }

            wb.Save();
            Debug.WriteLine($"Updated spreadsheet at {xlsxPath}");
            return xlsxPath;
        }

        // ------- Local copies of the private helpers from T1Client (so this file is self-contained) -------

        private static string SanitizeSheetNameLocal(string name)
        {
            var invalidChars = new[] { ':', '\\', '/', '?', '*', '[', ']' };
            var cleaned = new StringBuilder(name.Length);
            foreach (var ch in name)
            {
                cleaned.Append(invalidChars.Contains(ch) ? '_' : ch);
            }
            var result = cleaned.ToString();
            if (result.Length > 31) result = result.Substring(0, 31);
            return string.IsNullOrEmpty(result) ? "Sheet" : result;
        }

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
            var rootFields = new[] { "AssetRegisterName", "AssetNumber", "Description", "ShortDescription", "Status", "OperatingStatus" };
            if (Array.IndexOf(rootFields, header) >= 0)
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
