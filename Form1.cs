using System.Diagnostics;

namespace T1Sync
{
    // Test harness for T1Sync. Each demo method is self-contained — its inputs
    // live as local consts inside the method so you can read/edit one demo
    // without scrolling. The two button handlers just toggle which demo runs.
    //   button1 ("T1WS")           → T1Client operations
    //   button2 ("CsvTransformer") → CsvTransformer operations
    public partial class Form1 : Form
    {
        public Form1() => InitializeComponent();

        // ===== button1 — T1Client demos. Uncomment ONE per click. =====
        private void button1_Click(object sender, EventArgs e)
        {
            //FetchAssetDemo();
            //SaveAssetRoundTripDemo();
            //ParseAssetsMetaDemo();
            //SaveMetaToExcelDemo();
            //SyncAssetFromExcelDemo(dryrun: true);
            ExtractAssetDemo(databaseInstance: "mcc");
        }

        // ===== button2 — CsvTransformer demos. Uncomment ONE per click. =====
        private void button2_Click(object sender, EventArgs e)
        {
            //CsvToMetaDemo();
            CsvToFlatBriefDemo();
        }

        // ---------------- T1Client demos ----------------

        private static void FetchAssetDemo()
        {
            const string service     = "workshop-TP";
            const string testAssetId = "0100017";

            var client = new T1Client_Interop(service);
            var asset = client.FetchAsset(testAssetId);
            Debug.WriteLine($"FetchAsset {testAssetId}: top-level keys = " +
                            string.Join(", ", asset.EnumerateObject().Select(p => p.Name)));
            if (asset.TryGetProperty("AssetAttributes", out var attrs))
                Debug.WriteLine($"  AssetAttributes count = {attrs.GetArrayLength()}");
        }

        private static void SaveAssetRoundTripDemo()
        {
            const string service     = "workshop-TP";
            const string testAssetId = "0100017";

            var client = new T1Client_Interop(service);
            var asset = client.FetchAsset(testAssetId);
            var resp = client.SaveAsset(asset);
            Debug.WriteLine($"SaveAsset round-trip OK (response kind = {resp.ValueKind}).");
        }

        private static void ParseAssetsMetaDemo()
        {
            const string service = "workshop-TP";

            var client = new T1Client_Interop(service);
            var lookup = client.ParseAssetsMeta();   // writes <service>.json next to config.json
            Debug.WriteLine($"ParseAssetsMeta: {lookup.Count} asset types.");
        }

        private static void SaveMetaToExcelDemo()
        {
            const string service  = "workshop-TP";
            const string xlsxFile = @"c:/temp/workshop-TP_AR.xlsx";
            // For online use:
            //const string xlsxFile = @"https://maroondahcc.sharepoint.com/sites/AssetsWorkspace/Shared%20Documents/Asset%20Management/AssetRegister/AssetRegisterTest.xlsx?web=1";

            // Factory auto-routes to T1Client_ClosedXML (local) or T1Client_Interop (URL).
            T1ClientFactory.SaveMetaToExcel(service, xlsxFile);
        }

        private static void SyncAssetFromExcelDemo(bool dryrun)
        {
            const string service  = "workshop-TP";
            const string xlsxFile = @"c:/temp/workshop-TP_AR.xlsx";
            const string sheet    = "Tree_Street Tree";
            const int    firstRow = 7;
            const int    lastRow  = 7;

            T1ClientFactory.SyncAssetFromExcel(service, xlsxFile, sheet, firstRow, lastRow, dryrun);
        }

        private static void ExtractAssetDemo(string? databaseInstance = null)
        {
            const string service  = "workshop-TP";
            const string xlsxFile = @"c:/temp/workshop-TP_AR.xlsx";
            const string sheet    = "Tree_Street Tree";
            const int    firstRow = 7;
            const int    lastRow  = 7;

            T1ClientFactory.ExtractAsset(service, xlsxFile, sheet, firstRow, lastRow, databaseInstance);
        }

        // ---------------- CsvTransformer demos ----------------

        private static void CsvToMetaDemo()
        {
            const string service   = "workshop-TP";
            const string csvSource = "ASSET_Export_25052026-011611.csv";
            const string nodeName  = "Tree/Street Tree";
            const string metaJson  = @"c:\temp\csv-meta.json";
            const string metaXlsx  = @"c:\temp\csv-meta.xlsx";

            // nominated_fields comes from config.json (t1ws.<service>.nominated_fields)
            var t = CsvTransformer.FromConfig(csvSource, service);
            t.SaveMetaToJson(metaJson, nodeName);
            t.SaveMetaToExcel(metaXlsx, nodeName);
            Debug.WriteLine($"CSV → meta: {metaJson}  +  {metaXlsx}");
        }

        private static void CsvToFlatBriefDemo()
        {
            const string csvSource = "ASSET_Export_25052026-011611.csv";
            const string sheet     = "Tree/Street Tree";
            const string flatXlsx  = @"c:\temp\csv-flat.xlsx";

            // CSV (template, multi-row per asset) → flat brief Excel.
            // Nominated direct fields are hardcoded here — no service config needed.
            // Captioned attribute sub-fields are not extracted (brief output).
            // Template2FlatBrief auto-creates a brief-layout sheet when the target
            // doesn't exist: one column per nominated field + one per AttributeCode,
            // one row per asset.
            if (File.Exists(flatXlsx)) File.Delete(flatXlsx);

            var t = new CsvTransformer(csvSource,
                "AssetRegisterName", "AssetNumber", "Description",
                "ShortDescription", "Status", "OperatingStatus");
            t.Template2FlatBrief(flatXlsx, sheet);
            Debug.WriteLine($"CSV → flat brief: {flatXlsx}");
        }
    }
}
