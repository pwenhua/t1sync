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
            const string csvSource = "ASSET_Export_25052026-011611.csv";
            const string nodeName  = "Tree/Street Tree";
            const string metaJson  = @"c:\temp\csv-meta.json";
            const string metaXlsx  = @"c:\temp\csv-meta.xlsx";

            // nominated_fields comes from the top-level "nominated_fields" array in config.json.
            var t = Trans.FromConfig(csvSource);
            t.SaveMetaToJson(metaJson, nodeName);
            t.SaveMetaToExcel(metaXlsx, nodeName);
            Debug.WriteLine($"CSV → meta: {metaJson}  +  {metaXlsx}");
        }

        // CSV → Excel transform via Template2Flat (6-row header, every AttributeCode
        // → its own column). Pass assetTypeOnly: true to keep only the ASSET_TYPE
        // attribute column. Flat2Import is the CSV → CSV variant.
        private static void CsvDemo()
        { 

            var t = Trans.FromConfig(@"c:\temp\template.csv");
            //t.Template2Flat(outXlsx, sheet);
            t.Template2Flat(@"c:\temp\flat.xlsx", "09", assetTypeOnly: true);
            //t.Flat2Import(@"c:\temp\flat_office.csv", @"c:\temp\import_office.csv");
             
        }
    }
}
