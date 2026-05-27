from pathlib import Path

from T1Client import T1Client
from Trans import Trans

# Python mirror of Form1.cs demos.
#   t1client_demo() → T1Client operations (HTTP + spreadsheet)
#   trans_demo()    → Trans operations (CSV → meta / Excel)


def t1client_demo() -> None:
    service = "workshop-TP"

    # xlsx_file = "https://maroondahcc.sharepoint.com/sites/AssetsWorkspace/Shared%20Documents/Asset%20Management/AssetRegister/AssetRegisterTest.xlsx?web=1"
    xlsx_file = "c:/temp/workshop-TP_AR.xlsx"
    sheet = "Tree_Street Tree"
    first_row = 8
    last_row = 8

    # The Python client only supports local files. If you point it at a URL,
    # download to a local path first (this snippet just bails with a clear error).
    if T1Client.is_online(xlsx_file):
        raise RuntimeError(
            f"Python T1Client requires a local file. Got URL: {xlsx_file}\n"
            "Download it first (e.g. via requests or sharepoint API), then point xlsx_file at the local path."
        )

    client = T1Client(service)

    # test_id = "0100017"
    # response = client.fetch_asset(test_id)
    # print(f"Top-level keys in response: {list(response.keys())}")
    # if "AssetAttributes" in response:
    #     print(f"Number of AssetAttributes: {len(response['AssetAttributes'])}")

    # asset = client.fetch_asset(test_id)
    # response = client.save_asset(asset)
    # print(f"Save response: {response}")

    # lookup = client.parse_assets_meta()
    # path = client.save_meta_to_excel(xlsx_file)

    # path = client.sync_asset_from_excel(xlsx_file, sheet, first_row, last_row, dryrun=True)

    path = client.extract_asset(xlsx_file, sheet, first_row, last_row, database_instance="local")
    print(f"extract_asset done: {path}")


def trans_demo() -> None:
    csv_source = r"c:\temp\flat_4_upload.csv"
    out_xlsx   = r"c:\temp\upload.xlsx"
    out_csv    = r"c:\temp\upload.csv"
    sheet      = "s0"
    node_name  = "Tree/Street Tree"
    meta_json  = r"c:\temp\csv-meta.json"
    meta_xlsx  = r"c:\temp\csv-meta.xlsx"

    for p in (out_xlsx, out_csv):
        try:
            Path(p).unlink()
        except FileNotFoundError:
            pass

    t = Trans.from_config(csv_source)

    # CSV → meta (JSON + 6-row-header Excel sheet).
    # t.save_meta_to_json(meta_json, node_name)
    # t.save_meta_to_excel(meta_xlsx, node_name)
    # print(f"CSV → meta: {meta_json}  +  {meta_xlsx}")

    # CSV → flat Excel (every AttributeCode → its own column).
    t.template2_flat(out_xlsx, sheet)
    # t.template2_flat(out_xlsx, sheet, asset_type_only=True)
    # t.flat2import(csv_source, out_csv)
    print(f"CSV → {out_xlsx}")


if __name__ == "__main__":
    t1client_demo()
    # trans_demo()
