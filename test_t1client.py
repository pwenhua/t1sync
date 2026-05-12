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
    print("Testing parse_asset_meta for all items in asset_parse_meta...")
    client = T1Client()
    items = client.svc_config.get("asset_parse_meta", {})

    for name, test_id in items.items():
        print(f"Processing {name} (asset {test_id})...")
        asset = client.fetch_asset(test_id)
        parsed = client.parse_asset_meta(asset)
        print(f"Parsed metadata into {len(parsed)} entries.")
        for code, bucket in parsed.items():
            if isinstance(bucket, dict):
                print(f"  - {code}: {len(bucket)} attributes parsed.")
            else:
                print(f"  - {code} = {bucket!r}")


def meta_lookup_test():
    print("Testing get_meta_lookup...")
    client = T1Client()
    lookup = client.get_meta_lookup(force_refresh=True)
    print(f"Lookup generated with {len(lookup)} nodes.")
    for key, val in lookup.items():
        if isinstance(val, dict):
            print(f"  - Node: {key}, Entries: {len(val)}")
        else:
            print(f"  - Node: {key} = {val!r}")


def save_meta_test():
    print("Testing save_meta_to_excel...")
    client = T1Client()
    path = client.save_meta_to_excel()
    print(f"Spreadsheet written: {path}")


def extract_asset_test():
    print("Testing extract_asset...")
    client = T1Client()
    path = client.extract_asset()
    print(f"Spreadsheet updated: {path}")


def save_asset_from_excel_test():
    print("Testing save_asset_from_excel...")
    client = T1Client()
    path = client.save_asset_from_excel()
    print(f"Pushed rows from spreadsheet: {path}")


if __name__ == "__main__":
    # asset_get_test('0100017')
    # asset_save_test('0100017')
    # parse_meta_test()
    # meta_lookup_test()
    # save_meta_test()
    # extract_asset_test()
    save_asset_from_excel_test()
