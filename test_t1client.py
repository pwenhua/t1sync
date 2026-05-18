import json
from T1Client import T1Client

def asset_get_test(test_id: str):
    print(f"Fetching asset {test_id} using T1Client...")
    client = T1Client()
    response = client.fetch_asset(test_id)
    print("Response fetched successfully.")

    # Print out the keys of the response to see the structure
    print(f"Top-level keys in response: {list(response.keys())}")

    # If AssetAttributes exists, print how many there are
    if "AssetAttributes" in response:
        print(f"Number of AssetAttributes: {len(response['AssetAttributes'])}")


def asset_save_test(test_id: str):
    print(f"Round-trip save test for asset {test_id}...")
    client = T1Client()

    asset = client.fetch_asset(test_id)
    print(f"Fetched asset, top-level keys: {list(asset.keys())[:8]}...")

    response = client.save_asset(asset)
    print("Save response received.")
    print(f"Top-level keys in save response: {list(response.keys()) if isinstance(response, dict) else type(response)}")

    if isinstance(response, dict) and "Messages" in response:
        print(f"Messages: {response['Messages']}")


def parse_meta_test():
    print("Testing parse_assets_meta...")
    client = T1Client()
    lookup = client.parse_assets_meta()

    print(f"Parsed metadata for {len(lookup)} asset types.")
    for name, parsed in lookup.items():
        print(f"Parsed metadata into {len(parsed)} entries.")
        for code, bucket in parsed.items():
            if isinstance(bucket, dict):
                print(f"  - {code}: {len(bucket)} attributes parsed.")
            else:
                print(f"  - {code} = {bucket!r}")
    
    meta_path = client.config_path.parent / f"{client.service}_meta.json"
    with meta_path.open("w", encoding="utf-8") as f:
        json.dump(lookup, f, indent=2, ensure_ascii=False)
    print(f"Saved parsed metadata to {meta_path}")


def save_meta_to_excel_test():
    print("Testing save_meta_to_excel...")
    client = T1Client()
    path = client.save_meta_to_excel()
    print(f"Spreadsheet written: {path}")


def extract_asset_test():
    print("Testing extract_asset...")
    client = T1Client()
    path = client.extract_asset()
    print(f"Spreadsheet updated: {path}")


def update_asset_from_excel_test():
    print("Testing update_asset_from_excel...")
    client = T1Client()
    path = client.update_asset_from_excel()
    print(f"Pushed rows from spreadsheet: {path}")


def create_asset_test():
    print("Testing create_asset...")
    client = T1Client()
    path = client.create_asset()
    print(f"Created assets from spreadsheet: {path}")


if __name__ == "__main__":
    # asset_get_test('0100017')
    # asset_save_test('0100017')
    parse_meta_test()
    # save_meta_to_excel_test()
    extract_asset_test()
    # update_asset_from_excel_test()
    # create_asset_test()
