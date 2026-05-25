using System.Diagnostics;

namespace T1Sync
{
    // Test harness for T1Sync. Each demo method runs one operation end-to-end;
    // the two button handlers just toggle which demo(s) to execute.
    //   button1 ("T1WS")           → T1Client operations
    //   button2 ("CsvTransformer") → CsvTransformer operations
    public partial class Form1 : Form
    {
        // ---------- Shared test configuration ----------
        private const string Service       = "workshop-TP";
        private const string Sheet         = "Tree_Street Tree";
        private const int    FirstRow      = 7;
        private const int    LastRow       = 7;
        private const string TestAssetId   = "0100017";

        // Excel paths — toggle XlsxFile between the two to switch local/online.
        private const string XlsxLocal     = @"c:/temp/workshop-TP_AR.xlsx";
        private const string XlsxOnline    = @"https://maroondahcc.sharepoint.com/sites/AssetsWorkspace/Shared%20Documents/Asset%20Management/AssetRegister/AssetRegisterTest.xlsx?web=1";
        private const string XlsxFile      = XlsxLocal;

        // CsvTransformer paths.
        private const string CsvSource     = "ASSET_Export_25052026-011611.csv";
        private const string CsvMetaJson   = @"c:\temp\csv-meta.json";
        private const string CsvMetaXlsx   = @"c:\temp\csv-meta.xlsx";

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

        // ===== button2 — CsvTransformer demos. =====
        private void button2_Click(object sender, EventArgs e)
        {
            CsvToMetaDemo();
        }

        // ---------------- T1Client demos ----------------

        private static void FetchAssetDemo()
        {
            var client = new T1Client_Interop(Service);
            var asset = client.FetchAsset(TestAssetId);
            Debug.WriteLine($"FetchAsset {TestAssetId}: top-level keys = " +
                            string.Join(", ", asset.EnumerateObject().Select(p => p.Name)));
            if (asset.TryGetProperty("AssetAttributes", out var attrs))
                Debug.WriteLine($"  AssetAttributes count = {attrs.GetArrayLength()}");
        }

        private static void SaveAssetRoundTripDemo()
        {
            var client = new T1Client_Interop(Service);
            var asset = client.FetchAsset(TestAssetId);
            var resp = client.SaveAsset(asset);
            Debug.WriteLine($"SaveAsset round-trip OK (response kind = {resp.ValueKind}).");
        }

        private static void ParseAssetsMetaDemo()
        {
            var client = new T1Client_Interop(Service);
            var lookup = client.ParseAssetsMeta();   // writes <service>.json next to config.json
            Debug.WriteLine($"ParseAssetsMeta: {lookup.Count} asset types.");
        }

        private static void SaveMetaToExcelDemo()
        {
            // Auto-routes to T1Client_ClosedXML (local) or T1Client_Interop (URL).
            T1ClientFactory.SaveMetaToExcel(Service, XlsxFile);
        }

        private static void SyncAssetFromExcelDemo(bool dryrun)
        {
            T1ClientFactory.SyncAssetFromExcel(Service, XlsxFile, Sheet, FirstRow, LastRow, dryrun);
        }

        private static void ExtractAssetDemo(string? databaseInstance = null)
        {
            T1ClientFactory.ExtractAsset(Service, XlsxFile, Sheet, FirstRow, LastRow, databaseInstance);
        }

        // ---------------- CsvTransformer demos ----------------

        private static void CsvToMetaDemo()
        {
            // nominated_fields comes from config.json (t1ws.<service>.nominated_fields)
            var t = CsvTransformer.FromConfig(CsvSource, Service);
            t.SaveMetaToJson(CsvMetaJson, "Tree/Street Tree");
            t.SaveMetaToExcel(CsvMetaXlsx, "Tree/Street Tree");
            Debug.WriteLine($"CSV → meta: {CsvMetaJson}  +  {CsvMetaXlsx}");
        }
    }
}
