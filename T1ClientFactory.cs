// T1ClientFactory.cs — auto-route spreadsheet operations to the right client.
//
//   • If the Excel path is an http(s) URL → T1Client_Interop (Excel COM).
//   • Otherwise (local file path)         → T1Client_ClosedXML (ClosedXML only).
//
// Usage:
//   T1ClientFactory.ExtractAsset("workshop-TP", xlsxFile, sheet, firstRow, lastRow,
//                                databaseInstance: "local");
//   T1ClientFactory.SyncAssetFromExcel("workshop-TP", xlsxFile, sheet, firstRow, lastRow,
//                                      dryrun: true);
//   T1ClientFactory.SaveMetaToExcel("workshop-TP", xlsxFile);

using System;
using System.Collections.Generic;

namespace T1Sync
{
    public static class T1ClientFactory
    {
        public static bool IsOnline(string file) =>
            !string.IsNullOrEmpty(file) && file.StartsWith("http", StringComparison.OrdinalIgnoreCase);

        public static string SaveMetaToExcel(
            string service,
            string file,
            Dictionary<string, object>? meta = null,
            string configPath = T1Client_Interop.DefaultConfigPath)
        {
            if (IsOnline(file))
            {
                var client = new T1Client_Interop(service, configPath);
                return client.SaveMetaToExcel(file, meta);
            }
            else
            {
                var client = new T1Client_ClosedXML(service, configPath);
                return client.SaveMetaToExcel(file, meta);
            }
        }

        public static string ExtractAsset(
            string service,
            string file,
            string sheet,
            int firstRow,
            int lastRow,
            string? databaseInstance = null,
            string configPath = T1Client_Interop.DefaultConfigPath)
        {
            if (IsOnline(file))
            {
                var client = new T1Client_Interop(service, configPath);
                return client.ExtractAsset(file, sheet, firstRow, lastRow, databaseInstance);
            }
            else
            {
                var client = new T1Client_ClosedXML(service, configPath);
                return client.ExtractAsset(file, sheet, firstRow, lastRow, databaseInstance);
            }
        }

        public static void SyncAssetFromExcel(
            string service,
            string file,
            string sheet,
            int firstRow,
            int lastRow,
            bool dryrun = false,
            string configPath = T1Client_Interop.DefaultConfigPath)
        {
            if (IsOnline(file))
            {
                var client = new T1Client_Interop(service, configPath);
                client.SyncAssetFromExcel(file, sheet, firstRow, lastRow, dryrun);
            }
            else
            {
                var client = new T1Client_ClosedXML(service, configPath);
                client.SyncAssetFromExcel(file, sheet, firstRow, lastRow, dryrun);
            }
        }
    }
}
