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
            var client = new T1Client("workshop-TP");

            //var xlsxFile = @"c:/temp/workshop-TP_AR.xlsx";
            var xlsxFile = @"https://maroondahcc.sharepoint.com/sites/AssetsWorkspace/Shared%20Documents/Asset%20Management/AssetRegister/AssetRegisterTest.xlsx?web=1";
            var sheet = "Tree_Street Tree";
            var firstRow = 8;
            var lastRow = 8;

            //var testId = "0100017";
            //var response = client.FetchAsset(testId);
            //Debug.WriteLine($"Top-level keys: {string.Join(", ", response.EnumerateObject().Select(p => p.Name))}");

            //var asset = client.FetchAsset(testId);
            //var response = client.SaveAsset(asset);
            //Debug.WriteLine($"Save response received.");

            //var lookup = client.ParseAssetsMeta();
            //Debug.WriteLine($"Parsed metadata for {lookup.Count} asset types.");

            //var path = client.SaveMetaToExcel(xlsxFile);
            //var path = client.ExtractAsset(xlsxFile, sheet, firstRow, lastRow);

            var path = client.SyncAssetFromExcel(xlsxFile, sheet, firstRow, lastRow, dryrun: true);
        }
    }
}
