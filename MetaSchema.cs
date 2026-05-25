// MetaSchema.cs — single source of truth for the meta-spreadsheet layout.
//
// All three writers — T1Client_Interop.SaveMetaToExcel (Excel COM),
// T1Client_ClosedXML.SaveMetaToExcel (ClosedXML), and
// CsvTransformer.SaveMetaToExcel (ClosedXML) — call MetaSchema.BuildColumns
// to turn the hierarchical node-meta JSON into a sequence of column tuples,
// then write them with their own library's cell API. Sharing this code
// guarantees the row/column structure of the output is byte-identical
// regardless of which writer produced it.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;

namespace T1Sync
{
    internal static class MetaSchema
    {
        /// <summary>
        /// 6-row column tuples (Kind, AttributeCode, Level, Suffix, DataType, Header)
        /// from the hierarchical node-meta dict (see T1Client.parse_assetitem_meta).
        /// Column order: direct fields first, then per attribute (top-level scalar,
        /// then its captioned sub-fields sorted by integer level).
        /// </summary>
        public static List<(string Kind, string Code, string Level, string Suffix, string DataType, string Header)>
            BuildColumns(object nodeMetaObj)
        {
            var columns = new List<(string, string, string, string, string, string)>();
            JsonElement root;
            if (nodeMetaObj is JsonElement je) root = je;
            else
            {
                var json = JsonSerializer.Serialize(nodeMetaObj);
                root = JsonDocument.Parse(json).RootElement;
            }

            // 1. Direct fields
            if (root.TryGetProperty("fields", out var fieldsNode) && fieldsNode.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in fieldsNode.EnumerateObject())
                    columns.Add(("", "", "", "", prop.Value.ToString() ?? "", prop.Name));
            }

            // 2. Attributes
            if (root.TryGetProperty("attributes", out var attrsNode) && attrsNode.ValueKind == JsonValueKind.Object)
            {
                foreach (var attrProp in attrsNode.EnumerateObject())
                {
                    var attrCode = attrProp.Name;
                    var attrNode = attrProp.Value;
                    if (attrNode.ValueKind != JsonValueKind.Object) continue;

                    var dataType = attrNode.TryGetProperty("dataType", out var dtProp) && dtProp.ValueKind == JsonValueKind.String
                        ? dtProp.GetString()! : "A";
                    columns.Add(("Attribute", "", "", "", dataType, attrCode));

                    if (!attrNode.TryGetProperty("levels", out var levelsNode) || levelsNode.ValueKind != JsonValueKind.Object)
                        continue;

                    var sortedLevels = levelsNode.EnumerateObject()
                        .OrderBy(p => int.TryParse(p.Name, out var lvl) ? lvl : 0);

                    foreach (var levelProp in sortedLevels)
                    {
                        var levelKey = levelProp.Name;
                        if (levelProp.Value.ValueKind != JsonValueKind.Object) continue;
                        foreach (var suffixProp in levelProp.Value.EnumerateObject())
                        {
                            var suffix = suffixProp.Name;
                            var leaf = suffixProp.Value;
                            string caption = "", leafDt = "";
                            if (leaf.ValueKind == JsonValueKind.Array && leaf.GetArrayLength() >= 2)
                            {
                                caption = leaf[0].ToString();
                                leafDt = leaf[1].ToString();
                            }
                            else if (leaf.ValueKind == JsonValueKind.Array && leaf.GetArrayLength() == 1)
                            {
                                caption = leaf[0].ToString();
                            }
                            columns.Add(("Attribute", attrCode, $"level_{levelKey}", suffix, leafDt, caption));
                        }
                    }
                }
            }
            return columns;
        }

        /// <summary>Excel sheet-name rules: max 31 chars, no ":\\/?*[]".</summary>
        public static string SanitizeSheetName(string name)
        {
            var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
            var sb = new StringBuilder(name?.Length ?? 0);
            foreach (var c in name ?? "") sb.Append(invalid.Contains(c) ? '_' : c);
            var result = sb.ToString();
            if (result.Length > 31) result = result.Substring(0, 31);
            return string.IsNullOrEmpty(result) ? "Sheet" : result;
        }

        /// <summary>Excel column NumberFormat for the per-column dataType ('A'/'N'/'D').</summary>
        public static string NumberFormatFor(string dataType) => dataType switch
        {
            "N" => "General",
            "D" => "yyyy-mm-dd",
            _   => "@",
        };
    }
}
