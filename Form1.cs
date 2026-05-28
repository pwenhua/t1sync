using System.Diagnostics;

namespace T1Sync
{
    // Test harness for T1Sync. Demos that share variables are grouped into one
    // method; comment out lines inside a demo to skip individual operations.
    //   button1 ("T1WS")           → T1Client operations
    //   button2 ("Trans") → Trans operations
    public partial class Form1 : Form
    {
        public Form1() => InitializeComponent();

        // ===== button1 — T1Client demos. Uncomment ONE per click. =====
        private void button1_Click(object sender, EventArgs e)
        {
            //T1HttpDemo();
            T1ExcelDemo();
        }

        // ===== button2 — Trans demos. =====
        private void button2_Click(object sender, EventArgs e)
        {
            //CsvToMetaDemo();
            CsvDemo();
        }

        // ---------------- T1Client demos ----------------

        // HTTP-only operations against T1 — no Excel touched.
        // All three share `service`; the first two also share `testAssetId`.
        // SaveAsset is a real write (round-trip — fetched payload posted back).
        // ParseAssetsMeta walks every entry in svc_config.asset_classes and writes <service>.json.
        private static void T1HttpDemo()
        {
            const string service     = "workshop-TP";
            const string testAssetId = "0100017";

            var client = new T1Client_Interop(service);

            var asset = client.FetchAsset(testAssetId);
            Debug.WriteLine($"FetchAsset {testAssetId}: top-level keys = " +
                            string.Join(", ", asset.EnumerateObject().Select(p => p.Name)));
            if (asset.TryGetProperty("AssetAttributes", out var attrs))
                Debug.WriteLine($"  AssetAttributes count = {attrs.GetArrayLength()}");

            var resp = client.SaveAsset(asset);
            Debug.WriteLine($"SaveAsset round-trip OK (response kind = {resp.ValueKind}).");

            var lookup = client.ParseAssetsMeta();   // writes <service>.json next to config.json
            Debug.WriteLine($"ParseAssetsMeta: {lookup.Count} asset types.");
        }

        // Spreadsheet-aware operations. All share service + xlsxFile (+ sheet/rows for sync/extract).
        // Factory auto-routes to T1Client_ClosedXML (local path) or T1Client_Interop (http URL).
        // The three ops are mutually impacting on the same file — comment out the ones you
        // don't want before clicking, especially SaveMetaToExcel (it rewrites the headers).
        private static void T1ExcelDemo()
        {
            const string service  = "workshop-TP";
            const string xlsxFile = @"c:/temp/workshop-TP_AR.xlsx";
            // Online alternative:
            //const string xlsxFile = @"https://maroondahcc.sharepoint.com/sites/AssetsWorkspace/Shared%20Documents/Asset%20Management/AssetRegister/AssetRegisterTest.xlsx?web=1";
            const string sheet    = "Tree_Street Tree";
            const int    firstRow = 7;
            const int    lastRow  = 7;

            //T1ClientFactory.SaveMetaToExcel(service, xlsxFile);
            //T1ClientFactory.SyncAssetFromExcel(service, xlsxFile, sheet, firstRow, lastRow, dryrun: true);
            //T1ClientFactory.ExtractAssetToExcel(service, xlsxFile, sheet, firstRow, lastRow);
            T1ClientFactory.ExtractAssetToDB(service, xlsxFile, sheet, firstRow, lastRow, databaseInstance: "mcc");
        }

        // ---------------- Trans demos ----------------

        private static void CsvToMetaDemo()
        {
            // nominated_fields comes from the top-level "nominated_fields" array in config.json.
            var t = Trans.FromConfig();
            t.SaveMetaToJson(@"c:\temp\template.csv", @"c:\temp\csv-meta.json", "Tree/Street Tree");
            t.SaveMetaToCsv(@"c:\temp\template.csv", @"c:\temp\csv-meta.csv");
        }

        // CSV → CSV via Template2Flat: 6-row header + one row per asset, every
        // AttributeCode its own column (cell value = SearchPath). assetTypeOnly:
        // true keeps only the ASSET_TYPE column. Flat2Import walks it back to
        // the T1 bulk-import shape.
        private static void CsvDemo()
        {
            var t = Trans.FromConfig();
            //t.Template2Flat(@"c:\temp\template.csv", @"c:\temp\flat.csv");
            t.Template2Flat(@"c:\temp\template.csv", @"c:\temp\flat.csv", assetTypeOnly: true);
            //t.Flat2Import(@"c:\temp\flat.csv", @"c:\temp\import.csv");
        }
    }
}
