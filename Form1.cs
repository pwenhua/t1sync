using System.Diagnostics;

namespace T1Sync
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //var xlsxFile = @"https://maroondahcc.sharepoint.com/sites/AssetsWorkspace/Shared%20Documents/Asset%20Management/AssetRegister/AssetRegisterTest.xlsx?web=1";
            var xlsxFile = @"c:/temp/workshop-TP_AR.xlsx";

            // Factory picks T1Client_Interop (URL) or T1Client_ClosedXML (local) automatically.
            var service = "workshop-TP";
            var sheet = "Tree_Street Tree";
            var firstRow = 7;
            var lastRow = 7;

            // Single-shot operations (don't need factory; create a client explicitly):
            //var client = new T1Client_Interop(service);
            //var testId = "0100017";
            //var response = client.FetchAsset(testId);
            //Debug.WriteLine($"Top-level keys: {string.Join(", ", response.EnumerateObject().Select(p => p.Name))}");
            //var asset = client.FetchAsset(testId);
            //var response = client.SaveAsset(asset);
            //Debug.WriteLine($"Save response received.");

            //var lookup = client.ParseAssetsMeta();
            //Debug.WriteLine($"Parsed metadata for {lookup.Count} asset types.");

            //T1ClientFactory.SyncAssetFromExcel(service, xlsxFile, sheet, firstRow, lastRow, dryrun: true);

            T1ClientFactory.ExtractAsset(service, xlsxFile, sheet, firstRow, lastRow, databaseInstance: "mcc");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // nominated_fields is read from config.json (t1ws.<service>.nominated_fields)
            var t = CsvTransformer.FromConfig("ASSET_Export_25052026-011611.csv", "workshop-TP");
            var meta = t.ParseMeta();
            t.SaveMetaToJson(@"c:\temp\csv-meta.json", "Tree/Street Tree");
            t.SaveMetaToExcel(@"c:\temp\csv-meta.xlsx", "Tree/Street Tree");
        }
    }
}
