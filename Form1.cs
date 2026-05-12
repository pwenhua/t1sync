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
            // TestT1Client.AssetGetTest("0100017");
            // TestT1Client.AssetSaveTest("0100017");
            // TestT1Client.ParseMetaTest();
            // TestT1Client.MetaLookupTest();
            // TestT1Client.SaveMetaTest();
            TestT1Client.ExtractAssetTest();
        }
    }
}
