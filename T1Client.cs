#define USE_CLOSEDXML

// T1Client.cs - C# port of T1Client.py (synchronous)
//
// Loads config.json, caches the OAuth2 access token, and exposes
// FetchAsset / SaveAsset / ParseAssetMeta / GetMetaLookup that mirror
// the Python T1Client.
//
// SSL verification is disabled to match Python's verify=False — fine for the
// internal workshop tenant, remove the ServerCertificateCustomValidationCallback
// for any externally trusted host.

using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
#if USE_CLOSEDXML
using ClosedXML.Excel;
#endif

namespace T1Sync
{
    public class T1Client
    {
        public const string DefaultConfigPath = @"..\..\..\config.json";
        public const string DefaultMetaPath = @"..\..\..\t1_meta.json";

        private static readonly HttpClient Http = CreateClient();
        private static readonly Regex MetaKeyRegex = new Regex(
            @"^(AttributeItem(?:Userfield|SelectionType)\d+)_META_$",
            RegexOptions.Compiled);

        private static readonly string[] RootFields =
        {
            "AssetRegisterName", "AssetNumber", "Description", "ShortDescription", "Status","OperatingStatus"
        };

        public string Service { get; }
        public string ConfigPath { get; }
        public JsonElement SvcConfig => _svcConfig;

        private readonly JsonElement _config;
        private readonly JsonElement _svcConfig;
        private string? _token;

        public T1Client(string service = "t1ws_workshop", string configPath = DefaultConfigPath)
        {
            Service = service;
            ConfigPath = configPath;
            _config = LoadConfig(configPath);
            _svcConfig = _config.GetProperty(service);
        }

        private static HttpClient CreateClient()
        {
            var handler = new HttpClientHandler
            {
                ServerCertificateCustomValidationCallback = (_, _, _, _) => true,
            };
            return new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(30) };
        }

        private static JsonElement LoadConfig(string path)
        {
            using var stream = File.OpenRead(path);
            using var doc = JsonDocument.Parse(stream);
            return doc.RootElement.Clone();
        }

        private static string ReadString(HttpContent content)
        {
            using var stream = content.ReadAsStream();
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }

        public string GetToken(bool forceRefresh = false)
        {
            if (_token != null && !forceRefresh) return _token;

            var tokenCfg = _svcConfig.GetProperty("ep_get_token");
            var baseUrl = _svcConfig.GetProperty("base_url").GetString()!.TrimEnd('/') + "/";
            var url = baseUrl + tokenCfg.GetProperty("url").GetString()!.TrimStart('/');
            var method = (tokenCfg.TryGetProperty("method", out var m) ? m.GetString() : "POST")!.ToUpperInvariant();

            var form = new Dictionary<string, string>
            {
                ["client_id"] = tokenCfg.GetProperty("client_id").GetString()!,
                ["client_secret"] = tokenCfg.GetProperty("client_secret").GetString()!,
                ["grant_type"] = tokenCfg.GetProperty("grant_type").GetString()!,
            };

            using var req = new HttpRequestMessage(new HttpMethod(method), url)
            {
                Content = new FormUrlEncodedContent(form),
            };
            using var resp = Http.Send(req);
            resp.EnsureSuccessStatusCode();

            var body = ReadString(resp.Content);
            using var doc = JsonDocument.Parse(body);
            _token = doc.RootElement.GetProperty("access_token").GetString();
            return _token!;
        }

        public JsonElement FetchAsset(string testId, string endpoint = "ep_asset_get")
        {
            var ep = _svcConfig.GetProperty(endpoint);
            var baseUrl = _svcConfig.GetProperty("base_url").GetString()!.TrimEnd('/') + "/";
            var path = ep.GetProperty("url").GetString()!.TrimStart('/');
            var url = baseUrl + path + testId;

            HttpResponseMessage Request(string token)
            {
                var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                return Http.Send(req);
            }

            var resp = Request(GetToken());
            if (resp.StatusCode == System.Net.HttpStatusCode.Unauthorized)
            {
                resp.Dispose();
                resp = Request(GetToken(forceRefresh: true));
            }

            try
            {
                resp.EnsureSuccessStatusCode();
                var body = ReadString(resp.Content);
                using var doc = JsonDocument.Parse(body);
                return doc.RootElement.Clone();
            }
            finally
            {
                resp.Dispose();
            }
        }

        public JsonElement SaveAsset(object payload, string endpoint = "ep_asset_save")
        {
            var ep = _svcConfig.GetProperty(endpoint);
            var baseUrl = _svcConfig.GetProperty("base_url").GetString()!.TrimEnd('/') + "/";
            var path = ep.GetProperty("url").GetString()!.TrimStart('/');
            var url = baseUrl + path;
            var method = (ep.TryGetProperty("method", out var m) ? m.GetString() : "POST")!.ToUpperInvariant();

            var json = payload is string s ? s : JsonSerializer.Serialize(payload);

            HttpResponseMessage Request(string token)
            {
                var req = new HttpRequestMessage(new HttpMethod(method), url)
                {
                    Content = new StringContent(json, Encoding.UTF8, "application/json"),
                };
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                return Http.Send(req);
            }

            var resp = Request(GetToken());
            if (resp.StatusCode == System.Net.HttpStatusCode.Unauthorized)
            {
                resp.Dispose();
                resp = Request(GetToken(forceRefresh: true));
            }

            try
            {
                resp.EnsureSuccessStatusCode();
                var body = ReadString(resp.Content);
                if (string.IsNullOrWhiteSpace(body))
                {
                    using var empty = JsonDocument.Parse("{}");
                    return empty.RootElement.Clone();
                }
                using var doc = JsonDocument.Parse(body);
                return doc.RootElement.Clone();
            }
            finally
            {
                resp.Dispose();
            }
        }

        private static string InferDataType(JsonElement value)
        {
            // 'N' for numeric, 'A' for everything else (string/null/bool/object/array).
            return value.ValueKind == JsonValueKind.Number ? "N" : "A";
        }

        public static Dictionary<string, object> ParseAssetMeta(JsonElement asset)
        {
            // Flat schema: each AttributeCode becomes a top-level scalar (its
            // primary SearchPath's inferred dataType), and each captioned
            // attribute becomes a top-level entry whose value is
            // [AttributeCode, "level_<n>", <field-suffix>, <dataType>].
            var result = new Dictionary<string, object>();

            // Root fields — store just the inferred dataType.
            foreach (var field in RootFields)
            {
                if (asset.TryGetProperty(field, out var fieldProp))
                {
                    result[field] = InferDataType(fieldProp);
                }
            }

            if (!asset.TryGetProperty("AssetAttributes", out var attributes) || attributes.ValueKind != JsonValueKind.Array)
            {
                return result;
            }

            // Group AssetAttributes by AttributeCode, preserving first-seen order.
            var groups = new Dictionary<string, List<JsonElement>>();
            var groupOrder = new List<string>();
            foreach (var entry in attributes.EnumerateArray())
            {
                if (!entry.TryGetProperty("AttributeCode", out var codeProp) || codeProp.ValueKind != JsonValueKind.String)
                    continue;
                var code = codeProp.GetString();
                if (string.IsNullOrEmpty(code)) continue;
                if (!groups.TryGetValue(code, out var list))
                {
                    list = new List<JsonElement>();
                    groups[code] = list;
                    groupOrder.Add(code);
                }
                list.Add(entry);
            }

            foreach (var code in groupOrder)
            {
                var entries = groups[code];

                // AttributeCode → inferred dataType of the primary entry's SearchPath.
                JsonElement? primary = null;
                foreach (var e in entries)
                {
                    if (e.TryGetProperty("IsPrimaryValue", out var pv) &&
                        (pv.ValueKind == JsonValueKind.True ||
                         (pv.ValueKind == JsonValueKind.String && pv.GetString()?.ToLower() == "true")))
                    {
                        primary = e;
                        break;
                    }
                }
                if (primary.HasValue)
                {
                    var sp = primary.Value.TryGetProperty("SearchPath", out var spProp) ? spProp : default;
                    result[code] = new object[] { "Attribute", InferDataType(sp) };
                }

                // Captioned attributes — collect, sort by level, then add.
                var items = new List<(int level, string caption, object[] value)>();
                foreach (var entry in entries)
                {
                    var searchPath = entry.TryGetProperty("SearchPath", out var spProp) && spProp.ValueKind == JsonValueKind.String
                        ? spProp.GetString()! : "";
                    int level = string.IsNullOrEmpty(searchPath) ? 0 : searchPath.Split('\\').Length;

                    foreach (var prop in entry.EnumerateObject())
                    {
                        var match = MetaKeyRegex.Match(prop.Name);
                        if (!match.Success || prop.Value.ValueKind != JsonValueKind.String)
                            continue;
                        var valueKey = match.Groups[1].Value;
                        if (!entry.TryGetProperty(valueKey, out _)) continue;

                        try
                        {
                            using var metaDoc = JsonDocument.Parse(prop.Value.GetString()!);
                            var meta = metaDoc.RootElement;
                            if (!meta.TryGetProperty("Caption", out var capProp) || capProp.ValueKind != JsonValueKind.String)
                                continue;
                            var caption = capProp.GetString()!;
                            if (string.IsNullOrEmpty(caption)) continue;
                            var dataType = meta.TryGetProperty("DataType", out var dtProp) && dtProp.ValueKind == JsonValueKind.String
                                ? dtProp.GetString()! : "";
                            var suffix = valueKey.StartsWith("AttributeItem")
                                ? valueKey.Substring("AttributeItem".Length)
                                : valueKey;
                            items.Add((level, caption, new object[] { code, $"level_{level}", suffix, dataType }));
                        }
                        catch (JsonException)
                        {
                            // ignore invalid json
                        }
                    }
                }

                items.Sort((a, b) => a.level.CompareTo(b.level));
                foreach (var (_, caption, value) in items)
                {
                    result[caption] = value;
                }
            }

            return result;
        }

        public Dictionary<string, object> GetMetaLookup(bool forceRefresh = false)
        {
            var lookupPath = DefaultMetaPath;

            if (!forceRefresh && File.Exists(lookupPath))
            {
                var json = File.ReadAllText(lookupPath, Encoding.UTF8);
                return JsonSerializer.Deserialize<Dictionary<string, object>>(json)!;
            }

            Debug.WriteLine("Fetching metadata from API...");
            var lookup = new Dictionary<string, object>();
            if (File.Exists(lookupPath))
            {
                try
                {
                    var json = File.ReadAllText(lookupPath, Encoding.UTF8);
                    var existing = JsonSerializer.Deserialize<Dictionary<string, object>>(json);
                    if (existing != null) lookup = existing;
                }
                catch
                {
                    // ignore malformed file
                }
            }

            if (_svcConfig.TryGetProperty("asset_parse_meta", out var items))
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                foreach (var prop in items.EnumerateObject())
                {
                    var name = prop.Name;
                    var testId = prop.Value.GetString();
                    if (testId == null) continue;

                    Debug.WriteLine($"  -> Processing node: '{name}'");
                    var asset = FetchAsset(testId);
                    lookup[name] = ParseAssetMeta(asset);

                    File.WriteAllText(lookupPath, JsonSerializer.Serialize(lookup, options), Encoding.UTF8);
                }
            }

            Debug.WriteLine($"Saved metadata lookup to {lookupPath}");
            return lookup;
        }

        private static string SanitizeSheetName(string name)
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

        private static List<(string, string, string, string, string, string)> BuildMetaColumns(object nodeMetaObj)
        {
            // Produce 6-row column tuples (kind, AttributeCode, level, suffix, dataType, header)
            // from the flat node_meta dict. kind is "Attribute" for Attribute fields, "" for root fields.
            var columns = new List<(string, string, string, string, string, string)>();
            JsonElement root;

            if (nodeMetaObj is JsonElement je)
            {
                root = je;
            }
            else
            {
                var json = JsonSerializer.Serialize(nodeMetaObj);
                root = JsonDocument.Parse(json).RootElement;
            }

            foreach (var prop in root.EnumerateObject())
            {
                var key = prop.Name;
                var value = prop.Value;

                if (value.ValueKind == JsonValueKind.Array && value.GetArrayLength() == 4)
                {
                    // Captioned attribute: [AttributeCode, "level_X", suffix, dataType]
                    columns.Add((
                        "Attribute",
                        value[0].ToString(),
                        value[1].ToString(),
                        value[2].ToString(),
                        value[3].ToString(),
                        key
                    ));
                }
                else if (value.ValueKind == JsonValueKind.Array && value.GetArrayLength() == 2)
                {
                    // Top-level Attribute scalar: ["Attribute", dataType]
                    columns.Add((
                        value[0].ToString(),
                        "",
                        "",
                        "",
                        value[1].ToString(),
                        key
                    ));
                }
                else
                {
                    // Root field — bare dataType char.
                    columns.Add(("", "", "", "", value.ToString() ?? "", key));
                }
            }
            return columns;
        }

        public string SaveMetaToExcel(Dictionary<string, object>? meta = null)
        {
#if USE_CLOSEDXML
            if (meta == null)
            {
                var metaPath = DefaultMetaPath;
                if (!File.Exists(metaPath))
                    throw new FileNotFoundException($"Metadata file not found: {metaPath}");

                var jsonText = File.ReadAllText(metaPath, Encoding.UTF8);
                meta = JsonSerializer.Deserialize<Dictionary<string, object>>(jsonText);
            }

            var xlsxPath = _svcConfig.GetProperty("asset_meta_file").GetString()!;
            var dir = Path.GetDirectoryName(xlsxPath);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }

            using var wb = File.Exists(xlsxPath) ? new XLWorkbook(xlsxPath) : new XLWorkbook();
            
            if (wb.Worksheets.Contains("Sheet"))
            {
                wb.Worksheets.Delete("Sheet");
            }

            foreach (var kvp in meta!)
            {
                var nodeName = kvp.Key;
                var sheetName = SanitizeSheetName(nodeName);

                if (wb.Worksheets.Contains(sheetName))
                {
                    wb.Worksheets.Delete(sheetName);
                }

                var ws = wb.Worksheets.Add(sheetName);
                var columns = BuildMetaColumns(kvp.Value);

                for (int i = 0; i < columns.Count; i++)
                {
                    var colData = columns[i];
                    ws.Cell(1, i + 1).Value = colData.Item1;
                    ws.Cell(2, i + 1).Value = colData.Item2;
                    ws.Cell(3, i + 1).Value = colData.Item3;
                    ws.Cell(4, i + 1).Value = colData.Item4;
                    ws.Cell(5, i + 1).Value = colData.Item5;
                    ws.Cell(6, i + 1).Value = colData.Item6;

                    var format = colData.Item5 switch
                    {
                        "N" => "General",
                        "D" => "yyyy-mm-dd",
                        _   => "@",
                    };
                    ws.Column(i + 1).Style.NumberFormat.Format = format;
                }
            }

            if (!wb.Worksheets.Any())
            {
                wb.Worksheets.Add("Sheet");
            }

            wb.SaveAs(xlsxPath);
            Debug.WriteLine($"Saved spreadsheet to {xlsxPath}");
            return xlsxPath;
#else
            throw new NotImplementedException(
                "SaveMetaToExcel requires the ClosedXML NuGet package. " +
                "To enable it, install ClosedXML and uncomment '#define USE_CLOSEDXML' " +
                "at the top of T1Client.cs (or add it to your project properties)."
            );
#endif
        }

        private static object? JsonElementToValue(JsonElement el)
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

        private static object? ExtractValue(JsonElement asset, string attrCode, string level, string suffix, string header)
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
                            return JsonElementToValue(v);
                    }
                }
                return null;
            }

            // Root field (top-level on the asset payload).
            if (Array.IndexOf(RootFields, header) >= 0)
            {
                return asset.TryGetProperty(header, out var v) ? JsonElementToValue(v) : null;
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
         
        private static void SetCellValue(IXLCell cell, object? value)
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

        private static object? ReadCellValue(IXLCell cell)
        {
            var v = cell.Value;
            if (v.IsBlank) return null;
            if (v.IsText) return v.GetText();
            if (v.IsNumber) return v.GetNumber();
            if (v.IsBoolean) return v.GetBoolean();
            if (v.IsDateTime) return v.GetDateTime().ToString("o");
            return cell.GetString();
        }

        private static void SetAttributeValue(JsonObject asset, string attrCode, string level, string suffix, object? value)
        {
            // Mutate asset["AssetAttributes"] in place: overwrite AttributeItem<suffix>
            // on the entry matching attrCode + level.
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
                if (entry.ContainsKey(valueKey))
                {
                    entry[valueKey] = value == null ? null : JsonSerializer.SerializeToNode(value);
                    return;
                }
            }
        }

        public string SaveAssetFromExcel(string endpoint = "save_asset")
        {
            // For each row in [first_row, last_row] of every sheet:
            //  - Read AssetNumber from column 2; skip unless it's a non-empty text cell.
            //  - FetchAsset() to get the full asset JSON.
            //  - For every direct-field column (row 1 blank), overwrite the top-level
            //    field named in row 6 with the cell value.
            //  - For every captioned-attribute column (row 1 = "Attribute" with a
            //    level_N indicator in row 3), find the matching AssetAttributes entry
            //    and overwrite AttributeItem<suffix>.
            //  - POST the modified JSON via SaveAsset() (uses asset_save endpoint, adds
            //    'Authorization: Bearer <token>').
            var cfg = _svcConfig.GetProperty(endpoint);
            var xlsxPath = cfg.GetProperty("file").GetString()!;
            var firstRow = cfg.GetProperty("first_row").GetInt32();
            var lastRow = cfg.GetProperty("last_row").GetInt32();

            using var wb = new XLWorkbook(xlsxPath);
            foreach (var ws in wb.Worksheets)
            {
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

                for (int row = firstRow; row <= lastRow; row++)
                {
                    var assetCell = ws.Cell(row, 2);
                    if (!assetCell.Value.IsText) continue;
                    var assetNumber = assetCell.GetString();
                    if (string.IsNullOrWhiteSpace(assetNumber)) continue;

                    Debug.WriteLine($"  -> {ws.Name} row {row}: saving asset {assetNumber}");
                    var asset = FetchAsset(assetNumber);

                    var node = JsonNode.Parse(asset.GetRawText())!.AsObject();

                    for (int colIdx = 0; colIdx < headers.Count; colIdx++)
                    {
                        var (kind, code, level, suffix, header) = headers[colIdx];
                        var cellValue = ReadCellValue(ws.Cell(row, colIdx + 1));
                        var valueNode = cellValue == null ? null : JsonSerializer.SerializeToNode(cellValue);

                        if (kind == "Attribute")
                        {
                            // Captioned attribute: needs AttributeCode + level_N + suffix.
                            if (!string.IsNullOrEmpty(code) && !string.IsNullOrEmpty(suffix) && level.StartsWith("level_"))
                            {
                                SetAttributeValue(node, code, level, suffix, cellValue);
                            }
                            // Top-level scalar Attribute (no level) — skip.
                        }
                        else if (!string.IsNullOrEmpty(header))
                        {
                            node[header] = valueNode;
                        }
                    }

                    SaveAsset(node.ToJsonString());
                }
            }

            Debug.WriteLine($"Saved spreadsheet rows from {xlsxPath}");
            return xlsxPath;
        }

        public string ExtractAsset(string endpoint = "extract_asset")
        {
            var cfg = _svcConfig.GetProperty(endpoint);
            var xlsxPath = cfg.GetProperty("file").GetString()!;
            var firstRow = cfg.GetProperty("first_row").GetInt32();
            var lastRow = cfg.GetProperty("last_row").GetInt32();

            using var wb = new XLWorkbook(xlsxPath);
            foreach (var ws in wb.Worksheets)
            {
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

                for (int row = firstRow; row <= lastRow; row++)
                {
                    var assetNumber = ws.Cell(row, 2).GetString();
                    if (string.IsNullOrEmpty(assetNumber)) continue;

                    Debug.WriteLine($"  -> {ws.Name} row {row}: fetching asset {assetNumber}");
                    var asset = FetchAsset(assetNumber);

                    for (int colIdx = 0; colIdx < headers.Count; colIdx++)
                    {
                        var (attrCode, level, suffix, _, header) = headers[colIdx];
                        var val = ExtractValue(asset, attrCode, level, suffix, header);
                        if (val != null)
                        {
                            SetCellValue(ws.Cell(row, colIdx + 1), val);
                        }
                    }
                }
            }

            wb.Save();
            Debug.WriteLine($"Updated spreadsheet at {xlsxPath}");
            return xlsxPath;
        }

        public string CreateAsset(string endpoint = "create_asset")
        {
            var cfg = _svcConfig.GetProperty(endpoint);
            var xlsxPath = cfg.GetProperty("file").GetString()!;
            var firstRow = cfg.GetProperty("first_row").GetInt32();
            var lastRow = cfg.GetProperty("last_row").GetInt32();

            string? assetRegister = null;
            if (_svcConfig.TryGetProperty("asset_register", out var arProp) ||
                _svcConfig.TryGetProperty("asset register", out arProp))
            {
                assetRegister = arProp.GetString();
            }

            using var wb = new XLWorkbook(xlsxPath);

            var sheetKey = cfg.TryGetProperty("sheet", out var shProp) ? shProp.GetString() : null;
            var templateId = cfg.TryGetProperty("template", out var tmplProp) ? tmplProp.GetString() : null;

            if (string.IsNullOrEmpty(sheetKey) || string.IsNullOrEmpty(templateId))
            {
                Debug.WriteLine("  -> Missing 'sheet' or 'template' in create_asset config.");
                return xlsxPath;
            }

            var sheetName = SanitizeSheetName(sheetKey);
            if (!wb.Worksheets.Contains(sheetName))
            {
                Debug.WriteLine($"  -> Sheet {sheetName} not found in workbook.");
                return xlsxPath;
            }

            var ws = wb.Worksheet(sheetName);
            var lastUsed = ws.LastColumnUsed();
            var maxCol = lastUsed?.ColumnNumber() ?? 0;

            var headers = new List<string>();
            for (int col = 1; col <= maxCol; col++)
            {
                headers.Add(ws.Cell(6, col).GetString() ?? "");
            }

            int? assetNumCol = null;
            int? assetRegCol = null;
            for (int i = 0; i < headers.Count; i++)
            {
                if (headers[i] == "AssetNumber") assetNumCol = i + 1;
                else if (headers[i] == "AssetRegisterName") assetRegCol = i + 1;
            }

            for (int row = firstRow; row <= lastRow; row++)
            {
                Debug.WriteLine($"  -> {sheetName} row {row}: creating asset from template {templateId}");

                var payload = new Dictionary<string, string?>
                {
                    ["AssetRegisterName"] = assetRegister,
                    ["TemplateAssetNumberInternal"] = templateId
                };

                var result = SaveAsset(payload, "ep_asset_create");

                if (assetNumCol.HasValue && result.TryGetProperty("AssetNumber", out var anProp))
                {
                    SetCellValue(ws.Cell(row, assetNumCol.Value), JsonElementToValue(anProp));
                }
                if (assetRegCol.HasValue && result.TryGetProperty("AssetRegisterName", out var arNameProp))
                {
                    SetCellValue(ws.Cell(row, assetRegCol.Value), JsonElementToValue(arNameProp));
                }
            }

            wb.Save();
            Debug.WriteLine($"Updated spreadsheet at {xlsxPath}");
            return xlsxPath;
        }
    }
}
