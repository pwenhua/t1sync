#define USE_CLOSEDXML

// T1Client_Interop.cs - T1 client that uses Microsoft.Office.Interop.Excel
// (Excel COM) for spreadsheet I/O. Use this when the workbook lives on
// SharePoint/OneDrive (http(s) URL). For local files, prefer T1Client_ClosedXML.
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
using Application = Microsoft.Office.Interop.Excel.Application;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Data.SqlClient;

#if USE_CLOSEDXML
using ClosedXML.Excel;
#endif

namespace T1Sync
{
    public class T1Client_Interop
    {
        public const string DefaultConfigPath = @"..\..\..\config.json";

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
        public string MetaPath => $@"..\..\..\{Service}.json";
        public JsonElement SvcConfig => _svcConfig;

        private readonly JsonElement _config;
        private readonly JsonElement _svcConfig;
        private string? _token;

        public T1Client_Interop(string service, string configPath = DefaultConfigPath)
        {
            if (string.IsNullOrEmpty(service))
                throw new ArgumentException("Service name must be provided.", nameof(service));

            ConfigPath = configPath;
            _config = LoadConfig(configPath);
            Service = service;
            _svcConfig = _config.TryGetProperty("t1ws", out var t1ws) && t1ws.TryGetProperty(Service, out var svc) ? svc : _config.GetProperty(Service);
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

            // Insert the service's asset register (e.g. "TP_AR") between path and id.
            var assetRegister = "";
            if (_svcConfig.TryGetProperty("asset register", out var arProp) ||
                _svcConfig.TryGetProperty("asset_register", out arProp))
            {
                assetRegister = arProp.GetString()?.Trim('/') ?? "";
            }
            var registerSegment = string.IsNullOrEmpty(assetRegister) ? "" : assetRegister + "/";

            var url = baseUrl + path + registerSegment + testId;

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

        public static Dictionary<string, object> ParseAssetItemMeta(JsonElement asset)
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

        public Dictionary<string, object> ParseAssetsMeta()
        {
            Debug.WriteLine("Fetching metadata from API for all configured asset types...");
            var lookup = new Dictionary<string, object>();

            if (_svcConfig.TryGetProperty("asset_classes", out var items) && items.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in items.EnumerateArray())
                {
                    var name = item.TryGetProperty("class", out var clsProp) ? clsProp.GetString() : null;
                    var testId = item.TryGetProperty("seed", out var anProp) ? anProp.GetString() : null;
                    if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(testId)) continue;

                    Debug.WriteLine($"  -> Processing node: '{name}'");
                    var asset = FetchAsset(testId);
                    lookup[name] = ParseAssetItemMeta(asset);
                }
            }

            var options = new JsonSerializerOptions { WriteIndented = true };
            File.WriteAllText(MetaPath, JsonSerializer.Serialize(lookup, options), Encoding.UTF8);
            Debug.WriteLine($"Saved parsed metadata to {MetaPath}");

            return lookup;
        }

        public Dictionary<string, object> GetMetaLookup(bool forceRefresh = false)
        {
            var lookupPath = MetaPath;

            if (!forceRefresh && File.Exists(lookupPath))
            {
                try
                {
                    var json = File.ReadAllText(lookupPath, Encoding.UTF8);
                    return JsonSerializer.Deserialize<Dictionary<string, object>>(json)!;
                }
                catch (Exception ex) when (ex is JsonException || ex is IOException)
                {
                    Debug.WriteLine($"Could not read meta lookup cache, forcing refresh: {ex.Message}");
                }
            }

            return ParseAssetsMeta();
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

        private static string UniqueSheetName(XLWorkbook wb, string baseName)
        {
            // Sanitized sheet name; if it already exists, append '01', '02', ...
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

        public string SaveMetaToExcel(string file, Dictionary<string, object>? meta = null)
        {
#if USE_CLOSEDXML
            if (meta == null)
            {
                var metaPath = MetaPath;
                if (!File.Exists(metaPath))
                    throw new FileNotFoundException($"Metadata file not found: {metaPath}");

                var jsonText = File.ReadAllText(metaPath, Encoding.UTF8);
                meta = JsonSerializer.Deserialize<Dictionary<string, object>>(jsonText);
            }

            var xlsxPath = file;
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
                var sheetName = UniqueSheetName(wb, nodeName);

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
                "at the top of T1Client_Interop.cs (or add it to your project properties)."
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
            // Mutate asset["AssetAttributes"] in place: replace AttributeItem<suffix>
            // on the single entry whose AttributeCode + SearchPath-level match.
            //
            // Example for "Near Power Line" with meta ["ASSET_TYPE", "level_2", "Userfield1", "A"]:
            //   - targetLevel = 2  (level_N → integer)
            //   - valueKey    = "AttributeItemUserfield1"
            //   - scan AssetAttributes for the entry where AttributeCode == "ASSET_TYPE"
            //     and SearchPath splits into 2 segments on '\' (e.g. "Tree\Street Tree")
            //   - replace entry["AttributeItemUserfield1"] with the cell value.
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

        private class OnlineExcelHelper : IDisposable
        {
            private Application? _app;
            private string _originalFile;
            public string LocalFilePath { get; }
            public bool IsOnline { get; }

            public OnlineExcelHelper(string file)
            {
                _originalFile = file;
                IsOnline = file.StartsWith("http", StringComparison.OrdinalIgnoreCase);
                if (IsOnline)
                {
                    _app = new Application();
                    // For debugging — flip these to see Excel and any error/auth dialogs it raises.
                    _app.Visible = true;
                    _app.DisplayAlerts = true;
                    var wb = _app.Workbooks.Open(file);
                    LocalFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                    wb.SaveAs(LocalFilePath);
                    wb.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                }
                else
                {
                    LocalFilePath = file;
                }
            }

            public void SaveBackToOnline()
            {
                if (!IsOnline || _app == null) return;

                Microsoft.Office.Interop.Excel.Workbook? wb = null;
                try
                {
                    wb = _app.Workbooks.Open(LocalFilePath); 
                    wb.Save();
                }
                finally
                { 
                    if (wb != null)
                    {
                        wb.Save();
                        wb.Close();
                    }

                    _app.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);


                }
            }

            public void Dispose()
            {
                if (IsOnline)
                {
                    if (_app != null)
                    {
                        _app.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);
                        _app = null;
                    }
                    if (File.Exists(LocalFilePath))
                    {
                        try { File.Delete(LocalFilePath); } catch { }
                    }
                }
            }
        }

        public void SyncAssetFromExcel(string excelFilePath, string sheet, int firstRow, int lastRow, bool dryrun = false)
        {
            // Sync a sheet to T1: update existing assets, create new ones for blank rows.
            //  - If column 2 (AssetNumber) is a non-empty text cell → fetch that asset.
            //  - If blank → POST to ep_asset_create with the sheet's template (from
            //    svc_config['asset_classes']), then fetch the *seed* asset (also from
            //    asset_classes) and use it as the JSON shape, patching in the new
            //    AssetNumber/AssetRegisterName so SaveAsset targets the new asset.
            //  - Apply each non-blank cell value:
            //      * row-1 "Attribute" with level_N → mutate AssetAttributes.
            //      * row-1 blank → overwrite the top-level field named in row 6.
            //  - POST the modified JSON via SaveAsset().

            var maxCol = 26;
            Application xlApp = new Application();
            xlApp.DisplayAlerts = false; 
            var wb = xlApp.Workbooks.Open(excelFilePath, 0, false);
            try
            {
                string trueAssetType = sheet;
                int underscoreIdx = trueAssetType.IndexOf('_');
                if (underscoreIdx >= 0)
                {
                    trueAssetType = trueAssetType.Substring(0, underscoreIdx) + "/" + trueAssetType.Substring(underscoreIdx + 1);
                }

                var ws = wb.Worksheets[sheet];              

                var headers = new List<(string Kind, string Code, string Level, string Suffix, string Header)>();
                for (int col = 1; col <= maxCol; col++)
                {
                    headers.Add((
                        ws.Cells[1, col].Value ?? "",
                        ws.Cells[2, col].Value ?? "",
                        ws.Cells[3, col].Value ?? "",
                        ws.Cells[4, col].Value ?? "",
                        ws.Cells[6, col].Value ?? ""
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
                if (_svcConfig.TryGetProperty("asset_register", out var arProp) ||
                    _svcConfig.TryGetProperty("asset register", out arProp))
                {
                    assetRegister = arProp.GetString();
                }

                string? templateId = null;
                string? seedId = null;
                if (_svcConfig.TryGetProperty("asset_classes", out var assetClasses) && assetClasses.ValueKind == JsonValueKind.Object)
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
                        var assetCell = ws.Cells[row, 2];
                        var assetNumber = Convert.ToString(assetCell.Value) ?? "";

                        JsonObject node;
                        if (!string.IsNullOrWhiteSpace(assetNumber))
                        {
                            Debug.WriteLine($"  -> {sheet} row {row}: updating asset {assetNumber}");
                            var asset = FetchAsset(assetNumber);
                            node = JsonNode.Parse(asset.GetRawText())!.AsObject();
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(templateId))
                            {
                                ws.Cells[row, 27].Value = $"Missing 'template' for class '{trueAssetType}'.";
                                continue;
                            }
                            if (string.IsNullOrEmpty(seedId))
                            {
                                ws.Cells[row, 27].Value = $"Missing 'seed' for class '{trueAssetType}'.";
                                continue;
                            }

                            Debug.WriteLine($"  -> {sheet} row {row}: creating asset from template {templateId}");

                            // Fetch the seed asset to use as the template payload
                            var seedAsset = FetchAsset(seedId);
                            var seedNode = JsonNode.Parse(seedAsset.GetRawText())!.AsObject();

                            // Patch the necessary fields for creation
                            seedNode["AssetRegisterName"] = assetRegister;
                            seedNode["TemplateAssetNumberInternal"] = templateId;
                            // Clear the seed's AssetNumber so it gets a new one on creation
                            seedNode["AssetNumber"] = null;

                            string? newAssetNumber = null;
                            string? newAssetRegister = assetRegister;

                            if (!dryrun)
                            {
                                var result = SaveAsset(seedNode.ToJsonString(), "ep_asset_create");
                                newAssetNumber = result.TryGetProperty("AssetNumber", out var anProp) ? anProp.GetString() : null;
                                if (string.IsNullOrEmpty(newAssetNumber))
                                {
                                    ws.Cells[row, 27].Value = "Create returned no AssetNumber.";
                                    continue;
                                }
                                newAssetRegister = result.TryGetProperty("AssetRegisterName", out var arNameProp) ? arNameProp.GetString() : assetRegister;
                                if (assetNumCol.HasValue) ws.Cells[row, assetNumCol.Value].Value = newAssetNumber;
                                if (assetRegCol.HasValue && !string.IsNullOrEmpty(newAssetRegister))
                                {
                                    ws.Cells[row, assetRegCol.Value].Value = newAssetRegister;
                                }
                            }
                            else
                            {
                                newAssetNumber = $"DRYRUN_NEW_ROW_{row}";
                            }

                            // Use the seed string replacement for further local updates
                            var seedStr = seedAsset.GetRawText().Replace(seedId, newAssetNumber);
                            node = JsonNode.Parse(seedStr)!.AsObject();

                            node["AssetNumber"] = newAssetNumber;
                            if (!string.IsNullOrEmpty(newAssetRegister)) node["AssetRegisterName"] = newAssetRegister;
                        }

                        for (int colIdx = 0; colIdx < headers.Count; colIdx++)
                        {
                            var (kind, code, level, suffix, header) = headers[colIdx];
                            var cell = ws.Cells[row, colIdx + 1];
                            var cellValue = cell.Value;

                            // Only the top-level ASSET_TYPE column (where row-6 header == "ASSET_TYPE")
                            // gets forced to trueAssetType. Captioned attributes have code == "ASSET_TYPE"
                            // but their headers are captions ("Near Power Line", "Height(m)", ...).
                            bool isAssetType = string.Equals(header, "asset_type", StringComparison.OrdinalIgnoreCase) ||
                                               string.Equals(header, "AssetType", StringComparison.OrdinalIgnoreCase);

                            if (isAssetType)
                            {
                                var cellValueStr = cellValue?.ToString() ?? "";
                                if (!string.Equals(cellValueStr, trueAssetType, StringComparison.OrdinalIgnoreCase))
                                {
                                    cell.Interior.ColorIndex = 6; // Yellow
                                }
                                cellValue = trueAssetType;
                            }

                            if (cellValue == null || (cellValue is string s && string.IsNullOrEmpty(s))) continue;

                            if (kind == "Attribute")
                            {
                                if (!string.IsNullOrEmpty(code) && !string.IsNullOrEmpty(suffix) && level.StartsWith("level_"))
                                {
                                    SetAttributeValue(node, code, level, suffix, cellValue);
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
                            var dumpPath = @"c:\temp\payload.txt";
                            Directory.CreateDirectory(@"c:\temp");
                            File.WriteAllText(dumpPath, node.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
                            ws.Cells[row, 27].Value = $"Dry run: Payload saved to {Path.GetFileName(dumpPath)}";
                        }
                        else
                        {
                            SaveAsset(node.ToJsonString());
                            ws.Cells[row, 27].Value = "";
                        }
                    }
                    catch (Exception ex)
                    {
                        ws.Cells[row, 27].Value = ex.Message;
                    }
                }

                wb.Save(); 
                Debug.WriteLine($"Updated spreadsheet at {excelFilePath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                if (wb != null)
                {
                    try { wb.Close(); } catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                }

                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            }
        }

        public string ExtractAsset(string file, string sheet, int firstRow, int lastRow, string? databaseInstance = null)
        {
            var maxCol = 50;
            Application xlApp = new Application();
            xlApp.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Workbook? wb = null;
            SqlConnection? dbConn = null;
            string? dbTable = null;
            try
            {
                wb = xlApp.Workbooks.Open(file, 0, false);
                var sheetName = SanitizeSheetName(sheet);

                Microsoft.Office.Interop.Excel.Worksheet ws;
                try
                {
                    ws = wb.Worksheets[sheetName];
                }
                catch
                {
                    Debug.WriteLine($"  -> Sheet {sheetName} not found in workbook.");
                    return file;
                }

                var headers = new List<(string, string, string, string, string)>();
                for (int col = 1; col <= maxCol; col++)
                {
                    headers.Add((
                        Convert.ToString(ws.Cells[2, col].Value) ?? "",
                        Convert.ToString(ws.Cells[3, col].Value) ?? "",
                        Convert.ToString(ws.Cells[4, col].Value) ?? "",
                        Convert.ToString(ws.Cells[5, col].Value) ?? "",
                        Convert.ToString(ws.Cells[6, col].Value) ?? ""
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
                    return file;
                }

                // Optional: open a SQL connection if a valid databaseInstance is configured.
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
                        Debug.WriteLine($"  -> DB instance '{databaseInstance}' not found / incomplete in config.database.");
                    }
                }

                for (int row = firstRow; row <= lastRow; row++)
                {
                    var assetNumber = Convert.ToString(ws.Cells[row, assetNumCol.Value].Value);
                    if (string.IsNullOrEmpty(assetNumber)) continue;

                    try
                    {
                        Debug.WriteLine($"  -> {ws.Name} row {row}: fetching asset {assetNumber}");
                        var asset = FetchAsset(assetNumber);

                        for (int colIdx = 0; colIdx < headers.Count; colIdx++)
                        {
                            var (attrCode, level, suffix, _, header) = headers[colIdx];
                            if (header.Equals("AssetNumber", StringComparison.OrdinalIgnoreCase)) continue;

                            var val = ExtractValue(asset, attrCode, level, suffix, header);
                            if (val != null)
                            {
                                ws.Cells[row, colIdx + 1].Value = val;
                            }
                        }

                        string? dbError = null;
                        if (dbConn != null && !string.IsNullOrEmpty(dbTable))
                        {
                            try
                            {
                                ExtractGeometryToDb(asset, assetNumber, dbConn, dbTable);
                            }
                            catch (Exception dbEx)
                            {
                                dbError = "DB: " + dbEx.Message;
                            }
                        }

                        ws.Cells[row, 27].Value = dbError ?? "";
                    }
                    catch (Exception ex)
                    {
                        ws.Cells[row, 27].Value = ex.Message;
                    }
                }

                wb.Save();
                Debug.WriteLine($"Updated spreadsheet at {file}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                dbConn?.Dispose();

                if (wb != null)
                {
                    try { wb.Close(); } catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                }

                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            }
            return file;
        }

        private static void ExtractGeometryToDb(JsonElement asset, string assetNumber, SqlConnection conn, string table)
        {
            // Navigate to AssetMap.MapLayers[0]; bail unless it's a POINT.
            if (!asset.TryGetProperty("AssetMap", out var assetMap)) return;
            if (!assetMap.TryGetProperty("MapLayers", out var mapLayers)) return;
            if (mapLayers.ValueKind != JsonValueKind.Array || mapLayers.GetArrayLength() == 0) return;

            var firstLayer = mapLayers[0];
            if (!firstLayer.TryGetProperty("GeometryType", out var gt)) return;
            if (!string.Equals(gt.GetString(), "POINT", StringComparison.OrdinalIgnoreCase)) return;

            if (!firstLayer.TryGetProperty("Points", out var points) || points.ValueKind != JsonValueKind.Array) return;

            var coords = new List<(double Lat, double Lon)>();
            foreach (var pt in points.EnumerateArray())
            {
                if (!pt.TryGetProperty("PointLocation", out var loc)) continue;
                if (!loc.TryGetProperty("Latitude", out var latP) || !latP.TryGetDouble(out var lat)) continue;
                if (!loc.TryGetProperty("Longitude", out var lonP) || !lonP.TryGetDouble(out var lon)) continue;
                coords.Add((lat, lon));
            }
            if (coords.Count == 0) return;

            // WKT uses (longitude latitude) order.
            string wkt = coords.Count == 1
                ? $"POINT ({coords[0].Lon} {coords[0].Lat})"
                : "MULTIPOINT (" + string.Join(", ", coords.Select(c => $"({c.Lon} {c.Lat})")) + ")";

            using (var del = new SqlCommand($"DELETE FROM {table} WHERE compkey = @compkey", conn))
            {
                del.Parameters.AddWithValue("@compkey", assetNumber);
                del.ExecuteNonQuery();
            }
            using (var ins = new SqlCommand($"INSERT INTO {table} (compkey, wkt) VALUES (@compkey, @wkt)", conn))
            {
                ins.Parameters.AddWithValue("@compkey", assetNumber);
                ins.Parameters.AddWithValue("@wkt", wkt);
                ins.ExecuteNonQuery();
            }
        }

    }
}
