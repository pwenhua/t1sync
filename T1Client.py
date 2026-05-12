import json
from pathlib import Path
import requests
import re

CONFIG_PATH = Path(__file__).parent / "config.json"
META_KEY = re.compile(r"^(AttributeItem(?:Userfield|SelectionType)\d+)_META_$")
ROOT_FIELDS = ["AssetRegisterName", "AssetNumber", "Description", "ShortDescription", "Status"]
INVALID_SHEET_CHARS = set(":\\/?*[]")

class T1Client:
    def __init__(self, service: str = "t1ws_workshop", config_path: Path = CONFIG_PATH):
        self.service = service
        self.config_path = config_path
        self._token: str | None = None
        self.config = self._load_config()
        self.svc_config = self.config.get(self.service, {})

    def _load_config(self) -> dict:
        with self.config_path.open("r", encoding="utf-8") as f:
            return json.load(f)

    def get_token(self, force_refresh: bool = False) -> str:
        if self._token is None or force_refresh:
            token_cfg = self.svc_config.get("get_token", {})
            url = self.svc_config["base_url"].rstrip("/") + "/" + token_cfg["url"].lstrip("/")
            method = token_cfg.get("method", "POST").upper()

            payload = {
                "client_id": token_cfg["client_id"],
                "client_secret": token_cfg["client_secret"],
                "grant_type": token_cfg["grant_type"],
            }
            headers = {"Content-Type": "application/x-www-form-urlencoded"}

            response = requests.request(method, url, data=payload, headers=headers, timeout=30, verify=False)
            response.raise_for_status()

            data = response.json()
            self._token = data["access_token"]
            
        return self._token

    def fetch_asset(self, asset_number: str, endpoint: str = "asset_get") -> dict:
        ep = self.svc_config[endpoint]
        base = self.svc_config["base_url"].rstrip("/") + "/"
        path = ep["url"].lstrip("/")
        url = base + path + asset_number

        def _request(token: str) -> requests.Response:
            headers = {"Authorization": f"Bearer {token}"}
            return requests.get(url, headers=headers, timeout=30, verify=False)

        response = _request(self.get_token())

        if response.status_code == 401:
            response = _request(self.get_token(force_refresh=True))

        response.raise_for_status()
        return response.json()

    def save_asset(self, payload: dict, endpoint: str = "asset_save") -> dict:
        ep = self.svc_config[endpoint]
        base = self.svc_config["base_url"].rstrip("/") + "/"
        path = ep["url"].lstrip("/")
        url = base + path
        method = ep.get("method", "POST").upper()

        def _request(token: str) -> requests.Response:
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            }
            return requests.request(method, url, json=payload, headers=headers, timeout=30, verify=False)

        response = _request(self.get_token())

        if response.status_code == 401:
            response = _request(self.get_token(force_refresh=True))

        response.raise_for_status()
        return response.json() if response.content else {}

    @staticmethod
    def _infer_data_type(value) -> str:
        # 'N' for numeric (int/float), 'A' for everything else (alpha/string/null/bool).
        if isinstance(value, bool):
            return "A"
        if isinstance(value, (int, float)):
            return "N"
        return "A"

    @staticmethod
    def parse_asset_meta(asset: dict) -> dict:
        """Flat schema: each AttributeCode becomes a top-level scalar (its SearchPath),
        and each captioned attribute becomes a top-level entry whose value is
        [AttributeCode, "level_<n>", <field-suffix>, <dataType>]."""
        result: dict = {}

        # Root fields — store just the inferred dataType.
        for field in ROOT_FIELDS:
            if field in asset:
                result[field] = T1Client._infer_data_type(asset[field])

        # Group AssetAttributes by AttributeCode so each group's captions land
        # right after that group's top-level scalar.
        groups: dict[str, list] = {}
        for entry in asset.get("AssetAttributes", []):
            code = entry.get("AttributeCode")
            if not code:
                continue
            groups.setdefault(code, []).append(entry)

        for code, entries in groups.items():
            # AttributeCode → inferred dataType of the primary entry's SearchPath.
            primary = next((e for e in entries if e.get("IsPrimaryValue")), None)
            if primary is not None:
                result[code] = ["Attribute", T1Client._infer_data_type(primary.get("SearchPath", ""))]

            # Captioned attributes — collect, sort by level, then add.
            items: list[tuple[int, str, list]] = []
            for entry in entries:
                search_path = entry.get("SearchPath", "")
                level = len(search_path.split("\\")) if search_path else 0
                for key, meta_str in entry.items():
                    m = META_KEY.match(key)
                    if not m or not isinstance(meta_str, str):
                        continue
                    value_key = m.group(1)
                    if value_key not in entry:
                        continue
                    try:
                        meta = json.loads(meta_str)
                    except json.JSONDecodeError:
                        continue
                    caption = meta.get("Caption")
                    if not caption:
                        continue
                    suffix = value_key.removeprefix("AttributeItem")
                    items.append((level, caption, [code, f"level_{level}", suffix, meta.get("DataType", "")]))

            items.sort(key=lambda x: x[0])
            for _, caption, value in items:
                result[caption] = value

        return result

    def get_meta_lookup(self, force_refresh: bool = False) -> dict:
        """
        Returns the metadata lookup table.
        If t1_meta.json exists and force_refresh is False, loads from it.
        Otherwise, iterates through the config, updates the metadata incrementally,
        and saves to t1_meta.json.
        """
        lookup_path = Path(__file__).parent / "t1_meta.json"
        
        if not force_refresh and lookup_path.exists():
            with lookup_path.open("r", encoding="utf-8") as f:
                return json.load(f)

        print("Fetching metadata from API...")
        lookup = {}
        if lookup_path.exists():
            with lookup_path.open("r", encoding="utf-8") as f:
                try:
                    lookup = json.load(f)
                except json.JSONDecodeError:
                    pass

        items = self.svc_config.get("asset_parse_meta", {})

        for name, test_id in items.items():
            print(f"  -> Processing node: '{name}'")
            lookup[name] = self.parse_asset_meta(self.fetch_asset(test_id))
            
            # Save incrementally after each node is fetched
            with lookup_path.open("w", encoding="utf-8") as f:
                json.dump(lookup, f, indent=2, ensure_ascii=False)
                
        print(f"Saved metadata lookup to {lookup_path}")
        return lookup

    @staticmethod
    def _sanitize_sheet_name(name: str) -> str:
        cleaned = "".join("_" if ch in INVALID_SHEET_CHARS else ch for ch in name)
        return cleaned[:31] or "Sheet"

    @staticmethod
    def _build_meta_columns(node_meta: dict) -> list[tuple[str, str, str, str, str, str]]:
        """Produce 6-row column tuples (kind, AttributeCode, level, suffix, dataType, header)
        from the flat node_meta dict. `kind` is "Attribute" for Attribute fields, "" for
        root fields."""
        columns: list[tuple[str, str, str, str, str, str]] = []
        for key, value in node_meta.items():
            if isinstance(value, list) and len(value) == 4:
                # Captioned attribute: [AttributeCode, "level_X", suffix, dataType]
                columns.append(("Attribute", str(value[0]), str(value[1]), str(value[2]), str(value[3]), key))
            elif isinstance(value, list) and len(value) == 2:
                # Top-level Attribute scalar: ["Attribute", dataType]
                columns.append((str(value[0]), "", "", "", str(value[1]), key))
            else:
                # Root field — bare dataType char.
                columns.append(("", "", "", "", str(value), key))
        return columns

    def save_meta_to_excel(self, meta: dict | None = None) -> Path:
        """Build a spreadsheet from the meta lookup.
        If `meta` is None, loads from t1_meta.json next to this file.
        Output path comes from svc_config['meta_file']."""
        from openpyxl import Workbook, load_workbook  # lazy import
        from openpyxl.utils import get_column_letter  # lazy import

        if meta is None:
            meta_path = Path(__file__).parent / "t1_meta.json"
            with meta_path.open("r", encoding="utf-8") as f:
                meta = json.load(f)

        xlsx_path = Path(self.svc_config["asset_meta_file"])
        xlsx_path.parent.mkdir(parents=True, exist_ok=True)

        if xlsx_path.exists():
            wb = load_workbook(xlsx_path)
        else:
            wb = Workbook()
            if "Sheet" in wb.sheetnames:
                del wb["Sheet"]

        for node_name, node_meta in meta.items():
            sheet_name = self._sanitize_sheet_name(node_name)
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
            ws = wb.create_sheet(title=sheet_name)
            for col_idx, col_data in enumerate(self._build_meta_columns(node_meta), start=1):
                ws.cell(row=1, column=col_idx, value=col_data[0])
                ws.cell(row=2, column=col_idx, value=col_data[1])
                ws.cell(row=3, column=col_idx, value=col_data[2])
                ws.cell(row=4, column=col_idx, value=col_data[3])
                ws.cell(row=5, column=col_idx, value=col_data[4])
                ws.cell(row=6, column=col_idx, value=col_data[5])

                fmt = {"N": "General", "D": "yyyy-mm-dd"}.get(col_data[4], "@")
                ws.column_dimensions[get_column_letter(col_idx)].number_format = fmt

        if not wb.sheetnames:
            wb.create_sheet(title="Sheet")

        wb.save(xlsx_path)
        print(f"Saved spreadsheet to {xlsx_path}")
        return xlsx_path

    @staticmethod
    def _extract_value(asset: dict, attr_code: str, level: str, suffix: str, header: str):
        """Pull a single cell value out of a fetched asset using the column's
        5-row metadata (AttributeCode, level, suffix, dataType, header)."""
        # Captioned attribute — attr_code + suffix populated.
        if attr_code and suffix:
            target_level = 0
            if level.startswith("level_"):
                try:
                    target_level = int(level[len("level_"):])
                except ValueError:
                    pass
            value_key = "AttributeItem" + suffix
            for entry in asset.get("AssetAttributes", []):
                if entry.get("AttributeCode") != attr_code:
                    continue
                sp = entry.get("SearchPath", "")
                entry_level = len(sp.split("\\")) if sp else 0
                if entry_level != target_level:
                    continue
                if value_key in entry:
                    return entry[value_key]
            return None

        # Root field (top-level on the asset payload).
        if header in ROOT_FIELDS:
            return asset.get(header)

        # AttributeCode top-level scalar (LOCATION, SERVICEAREA, ASSET_TYPE, ...).
        for entry in asset.get("AssetAttributes", []):
            if entry.get("AttributeCode") == header and entry.get("IsPrimaryValue"):
                return entry.get("SearchPath", "")
        return None

    def save_asset_from_excel(self, endpoint: str = "save_asset") -> Path:
        """Push spreadsheet edits back to T1.

        For each row in [first_row, last_row] of every sheet:
        - Read AssetNumber from column 2; skip unless it's a non-empty string.
        - fetch_asset() to get the full asset JSON.
        - For every direct-field column (row 1 blank, not 'Attribute'), overwrite
          the top-level field named in row 6 with the cell value.
        - POST the modified JSON via save_asset() (uses asset_save endpoint, adds
          'Authorization: Bearer <token>')."""
        from openpyxl import load_workbook  # lazy import

        cfg = self.svc_config[endpoint]
        xlsx_path = Path(cfg["file"])
        first_row = int(cfg["first_row"])
        last_row = int(cfg["last_row"])

        wb = load_workbook(xlsx_path)

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            max_col = ws.max_column

            headers = [
                (
                    str(ws.cell(row=1, column=col).value or ""),
                    str(ws.cell(row=6, column=col).value or ""),
                )
                for col in range(1, max_col + 1)
            ]

            for row in range(first_row, last_row + 1):
                asset_number = ws.cell(row=row, column=2).value
                if not isinstance(asset_number, str) or not asset_number.strip():
                    continue

                print(f"  -> {sheet_name} row {row}: saving asset {asset_number}")
                asset = self.fetch_asset(asset_number)

                for col_idx, (kind, header) in enumerate(headers, start=1):
                    if kind == "Attribute" or not header:
                        continue
                    asset[header] = ws.cell(row=row, column=col_idx).value

                self.save_asset(asset)

        return xlsx_path

    def extract_asset(self, endpoint: str = "extract_asset") -> Path:
        """Populate spreadsheet rows with live asset values.

        For each row in [first_row, last_row] of every sheet in the workbook,
        read the AssetNumber from column 2, fetch the asset, and write a value
        into every column based on its 6-row header tuple (row 1 = kind)."""
        from openpyxl import load_workbook  # lazy import

        cfg = self.svc_config[endpoint]
        xlsx_path = Path(cfg["file"])
        first_row = int(cfg["first_row"])
        last_row = int(cfg["last_row"])

        wb = load_workbook(xlsx_path)

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            max_col = ws.max_column

            headers = [
                (
                    str(ws.cell(row=2, column=col).value or ""),
                    str(ws.cell(row=3, column=col).value or ""),
                    str(ws.cell(row=4, column=col).value or ""),
                    str(ws.cell(row=5, column=col).value or ""),
                    str(ws.cell(row=6, column=col).value or ""),
                )
                for col in range(1, max_col + 1)
            ]

            for row in range(first_row, last_row + 1):
                asset_number = ws.cell(row=row, column=2).value
                if asset_number in (None, ""):
                    continue
                print(f"  -> {sheet_name} row {row}: fetching asset {asset_number}")
                asset = self.fetch_asset(str(asset_number))

                for col_idx, (attr_code, level, suffix, _data_type, header) in enumerate(headers, start=1):
                    value = T1Client._extract_value(asset, attr_code, level, suffix, header)
                    if value is not None:
                        ws.cell(row=row, column=col_idx, value=value)

        wb.save(xlsx_path)
        print(f"Updated spreadsheet at {xlsx_path}")
        return xlsx_path

