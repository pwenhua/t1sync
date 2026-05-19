from T1Client import T1Client

if __name__ == "__main__":
    client = T1Client("workshop-TP")

    xlsx_file = "c:/temp/workshop-TP_AR.xlsx"
    sheet = "Tree\\Street Tree"
    first_row = 7
    last_row = 7

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

    # path = client.extract_asset(xlsx_file, sheet, first_row, last_row)

    path = client.sync_asset_from_excel(xlsx_file, sheet, first_row, last_row)
