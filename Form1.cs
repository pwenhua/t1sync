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

            // var testId = "0100017";
            // var response = client.FetchAsset(testId);
            // Debug.WriteLine($"Top-level keys: {string.Join(", ", response.EnumerateObject().Select(p => p.Name))}");

            // var asset = client.FetchAsset(testId);
            // var response = client.SaveAsset(asset);
            // Debug.WriteLine($"Save response received.");

            var lookup = client.ParseAssetsMeta(); 

            // var path = client.SaveMetaToExcel();

            // var path = client.ExtractAsset();

            // var path = client.SaveAssetFromExcel();

            // var path = client.CreateAsset();
        }
    }
}
