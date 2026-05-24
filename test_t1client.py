from T1Client import T1Client

if __name__ == "__main__":
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
