using System.Diagnostics;

namespace T1Sync
{
    // Test harness for T1Sync. Demos that share variables are grouped into one
    // method; comment out lines inside a demo to skip individual operations.
    //   button1 ("T1WS")           → T1Client operations
    //   button2 ("CsvTransformer") → CsvTransformer operations
    public partial class Form1 : Form
    {
        public Form1() => InitializeComponent();

        // ===== button1 — T1Client demos. Uncomment ONE per click. =====
        private void button1_Click(object sender, EventArgs e)
        {
            //T1HttpDemo();
            T1ExcelDemo();
        }

        // ===== button2 — CsvTransformer demos. =====
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
            T1ClientFactory.ExtractAsset(service, xlsxFile, sheet, firstRow, lastRow, databaseInstance: "mcc");
        }

        // ---------------- CsvTransformer demos ----------------

        private static void CsvToMetaDemo()
        {
            const string csvSource = "ASSET_Export_25052026-011611.csv";
            const string nodeName  = "Tree/Street Tree";
            const string metaJson  = @"c:\temp\csv-meta.json";
            const string metaXlsx  = @"c:\temp\csv-meta.xlsx";

            // nominated_fields comes from the top-level "nominated_fields" array in config.json.
            var t = CsvTransformer.FromConfig(csvSource);
            t.SaveMetaToJson(metaJson, nodeName);
            t.SaveMetaToExcel(metaXlsx, nodeName);
            Debug.WriteLine($"CSV → meta: {metaJson}  +  {metaXlsx}");
        }

        // Runs all four CSV → Excel transforms in one click. They all write
        // into the SAME workbook, each as a separate sheet (UniqueSheetName
        // auto-appends "01", "02"… when a sheet with that base name exists).
        //   Sheet 1 ("Tree_Street Tree")   ← Template2FlatBrief (6-row header)
        //   Sheet 2 ("Tree_Street Tree01") ← TemplateSimple1    (2-row CSV-shape)
        //   Sheet 3 ("Tree_Street Tree02") ← TemplateSimple2    (1-row compact)
        //   Sheet 4 ("Tree_Street Tree03") ← TemplateSimple0    (no header, data only)
        private static void CsvDemo()
        {
            const string csvSource = "ASSET_Export_25052026-011611.csv";
            const string outXlsx   = @"c:\temp\csv-demo.xlsx";
            const string sheet     = "Tree/Street Tree";

            if (File.Exists(outXlsx)) File.Delete(outXlsx);

            var t = CsvTransformer.FromConfig(csvSource);
            t.Template2FlatBrief(outXlsx, sheet);
            t.TemplateSimple1(outXlsx, sheet);
            t.TemplateSimple2(outXlsx, sheet);
            t.TemplateSimple0(outXlsx, sheet);

            Debug.WriteLine($"CSV → {outXlsx} (4 sheets)");
        }
    }
}
