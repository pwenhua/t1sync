from T1Client import T1Client

if __name__ == "__main__":
    client = T1Client("workshop-TP")

    # test_id = "0100017"
    # response = client.fetch_asset(test_id)
    # print(f"Top-level keys in response: {list(response.keys())}")
    # if "AssetAttributes" in response:
    #     print(f"Number of AssetAttributes: {len(response['AssetAttributes'])}")

    # asset = client.fetch_asset(test_id)
    # response = client.save_asset(asset)
    # print(f"Save response: {response}")

    lookup = client.parse_assets_meta()

    # path = client.save_meta_to_excel()

    # path = client.extract_asset()

    # path = client.update_asset_from_excel()

    # path = client.create_asset()
