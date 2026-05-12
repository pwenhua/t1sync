// test_t1client.cs - C# port of test_t1client.py (synchronous)
//
// Smoke tests for T1Client. AssetSaveTest does a real round-trip
// (fetch then POST the same payload back), so only run it when writing
// to the workshop tenant is intended.

using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace T1Sync
{
    public static class TestT1Client
    {
        public static void AssetGetTest(string testId)
        {
            Debug.WriteLine($"Fetching asset {testId} using T1Client...");
            var client = new T1Client();
            var response = client.FetchAsset(testId);
            Debug.WriteLine("Response fetched successfully.");

            var topLevelKeys = response.EnumerateObject().Select(p => p.Name).ToList();
            Debug.WriteLine($"Top-level keys in response: [{string.Join(", ", topLevelKeys)}]");

            if (response.TryGetProperty("AssetAttributes", out var attrs) &&
                attrs.ValueKind == JsonValueKind.Array)
            {
                Debug.WriteLine($"Number of AssetAttributes: {attrs.GetArrayLength()}");
            }
        }

        public static void AssetSaveTest(string testId)
        {
            Debug.WriteLine($"Round-trip save test for asset {testId}...");
            var client = new T1Client();

            var asset = client.FetchAsset(testId);
            var fetchedKeys = asset.EnumerateObject().Select(p => p.Name).Take(8);
            Debug.WriteLine($"Fetched asset, top-level keys: {string.Join(", ", fetchedKeys)}...");

            var response = client.SaveAsset(asset);
            Debug.WriteLine("Save response received.");

            if (response.ValueKind == JsonValueKind.Object)
            {
                var saveKeys = response.EnumerateObject().Select(p => p.Name).ToList();
                Debug.WriteLine($"Top-level keys in save response: [{string.Join(", ", saveKeys)}]");

                if (response.TryGetProperty("Messages", out var msgs))
                {
                    Debug.WriteLine($"Messages: {msgs.GetRawText()}");
                }
            }
            else
            {
                Debug.WriteLine($"Save response kind: {response.ValueKind}");
            }
        }

        public static void ParseMetaTest()
        {
            Debug.WriteLine("Testing ParseAssetMeta for all items in asset_parse_meta...");
            var client = new T1Client();

            if (client.SvcConfig.TryGetProperty("asset_parse_meta", out var items))
            {
                foreach (var prop in items.EnumerateObject())
                {
                    var name = prop.Name;
                    var testId = prop.Value.GetString();
                    if (testId == null) continue;

                    Debug.WriteLine($"Processing {name} (asset {testId})...");
                    var asset = client.FetchAsset(testId);
                    var parsed = T1Client.ParseAssetMeta(asset);
                    Debug.WriteLine($"Parsed metadata into {parsed.Count} entries.");
                    foreach (var kvp in parsed)
                    {
                        switch (kvp.Value)
                        {
                            case object[] arr:
                                Debug.WriteLine($"  - {kvp.Key} = [{string.Join(", ", arr)}]");
                                break;
                            case JsonElement je when je.ValueKind == JsonValueKind.Array:
                                Debug.WriteLine($"  - {kvp.Key} = [{string.Join(", ", je.EnumerateArray().Select(e => e.ToString()))}]");
                                break;
                            default:
                                Debug.WriteLine($"  - {kvp.Key} = {kvp.Value}");
                                break;
                        }
                    }
                }
            }
        }

        public static void MetaLookupTest()
        {
            Debug.WriteLine("Testing GetMetaLookup...");
            var client = new T1Client();
            var lookup = client.GetMetaLookup(forceRefresh: true);
            Debug.WriteLine($"Lookup generated with {lookup.Count} nodes.");
            foreach (var kvp in lookup)
            {
                int count = kvp.Value switch
                {
                    Dictionary<string, object> d => d.Count,
                    JsonElement je when je.ValueKind == JsonValueKind.Object => je.EnumerateObject().Count(),
                    _ => 0,
                };
                Debug.WriteLine($"  - Node: {kvp.Key}, Entries: {count}");
            }
        }
        public static void SaveMetaTest()
        {
            Debug.WriteLine("Testing SaveMetaToExcel...");
            var client = new T1Client();
            try
            {
                var path = client.SaveMetaToExcel();
                Debug.WriteLine($"Spreadsheet written: {path}");
            }
            catch (NotImplementedException ex)
            {
                Debug.WriteLine($"SaveMetaToExcel skipped: {ex.Message}");
            }
        }

        public static void ExtractAssetTest()
        {
            Debug.WriteLine("Testing ExtractAsset...");
            var client = new T1Client();
            try
            {
                var path = client.ExtractAsset();
                Debug.WriteLine($"Spreadsheet updated: {path}");
            }
            catch (NotImplementedException ex)
            {
                Debug.WriteLine($"ExtractAsset skipped: {ex.Message}");
            }
        }

        public static void SaveAssetFromExcelTest()
        {
            Debug.WriteLine("Testing SaveAssetFromExcel...");
            var client = new T1Client();
            try
            {
                var path = client.SaveAssetFromExcel();
                Debug.WriteLine($"Pushed rows from spreadsheet: {path}");
            }
            catch (NotImplementedException ex)
            {
                Debug.WriteLine($"SaveAssetFromExcel skipped: {ex.Message}");
            }
        }
    }
}
