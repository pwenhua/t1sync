# Trans.py — Python mirror of the C# Trans class.
#
# Reads a T1 ASSET CSV export and produces:
#   - parse_meta()            → hierarchical meta dict (same shape as T1Client.parse_assetitem_meta)
#   - save_meta_to_json(...)  → write the meta dict to a JSON file
#   - save_meta_to_excel(...) → write the meta as a 6-row-header workbook sheet
#   - template2_flat_brief()  → T1Sync 6-row header + brief data (one row per asset)
#   - simplify0()             → row 1 = CSV FORMAT verbatim; row 2 = nominated col
#                               names compacted; row 3+ = ASSET data compacted
#   - flat1()                 → rows 1-2 = CSV header verbatim (~290 cols);
#                               row 3+ = ASSET data populated only in nominated cells
#   - flat2()                 → row 1 = nominated names; row 2+ = ASSET data
#                               (ultra-compact, no CSV decoration)

from __future__ import annotations

import csv
import json
from pathlib import Path

CONFIG_PATH = Path(__file__).parent / "config.json"
INVALID_SHEET_CHARS = set(":\\/?*[]")


class Trans:
    def __init__(self, csv_path: str | Path, *nominated_fields: str):
        self._csv_path = Path(csv_path)
        self._nominated_fields: list[str] = [f for f in nominated_fields if f]

    # ---- factory: read nominated_fields from config.json top-level ----
    @classmethod
    def from_config(cls, csv_path: str | Path, config_path: str | Path = CONFIG_PATH) -> "Trans":
        fields: list[str] = []
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            nf = cfg.get("nominated_fields")
            if isinstance(nf, list):
                fields = [s for s in nf if isinstance(s, str) and s]
        except Exception:
            pass
        return cls(csv_path, *fields)

    # ---- common CSV reading ----
    def _read_csv(self) -> list[list[str]]:
        with open(self._csv_path, "r", encoding="utf-8", newline="") as f:
            return [row for row in csv.reader(f)]

    @staticmethod
    def _header_idx(header_line: list[str]) -> dict[str, int]:
        idx: dict[str, int] = {}
        for i, name in enumerate(header_line):
            if name and name not in idx:
                idx[name] = i
        return idx

    def _asset_rows(self, rows: list[list[str]], line_type_idx: int) -> list[list[str]]:
        out = []
        if line_type_idx < 0:
            return out
        for row in rows[2:]:
            if line_type_idx < len(row) and row[line_type_idx].upper() == "ASSET":
                out.append(row)
        return out

    # ---------- parse_meta (hierarchical) ----------

    def parse_meta(self) -> dict:
        rows = self._read_csv()
        if len(rows) < 2:
            return {"fields": {}, "attributes": {}}

        header_line = rows[1]
        hi = self._header_idx(header_line)
        line_type_idx = hi.get("LineType", -1)
        attr_code_idx = hi.get("AttributeCode", -1)
        level_num_idx = hi.get("LevelNumber", -1)

        fields: dict = {}
        attributes: dict = {}

        # Direct (nominated) fields from the first ASSET row.
        asset_row = next(
            (r for r in rows[2:]
             if 0 <= line_type_idx < len(r) and r[line_type_idx].upper() == "ASSET"),
            None
        )
        if asset_row is not None:
            for f in self._nominated_fields:
                if f in hi and hi[f] < len(asset_row):
                    fields[f] = self._infer_data_type(asset_row[hi[f]])

        # Attributes from ATTRIBUTE rows.
        if attr_code_idx >= 0 and level_num_idx >= 0:
            for row in rows[2:]:
                if line_type_idx < 0 or line_type_idx >= len(row):
                    continue
                if row[line_type_idx].upper() != "ATTRIBUTE":
                    continue
                if attr_code_idx >= len(row):
                    continue
                code = row[attr_code_idx]
                if not code:
                    continue
                level_str = row[level_num_idx].strip() if level_num_idx < len(row) else ""
                attr_node = attributes.setdefault(code, {"dataType": "A"})
                self._add_captioned(attr_node, row, hi, level_str, "Userfield")
                self._add_captioned(attr_node, row, hi, level_str, "SelectionType")

        return {"fields": fields, "attributes": attributes}

    @staticmethod
    def _add_captioned(attr_node: dict, row: list[str], hi: dict[str, int],
                       level_str: str, family: str) -> None:
        levels = None
        level_dict = None
        for n in range(1, 21):
            col_name = f"AssetAttribute{family}{n}"
            idx = hi.get(col_name, -1)
            if idx < 0 or idx >= len(row):
                continue
            value = row[idx]
            if not value:
                continue
            if levels is None:
                levels = attr_node.setdefault("levels", {})
                level_dict = levels.setdefault(level_str, {})
            suffix = f"{family}{n}"
            if suffix in level_dict:
                continue
            # CSV doesn't carry T1's caption — leave it blank.
            level_dict[suffix] = ["", Trans._infer_data_type(value)]

    @staticmethod
    def _infer_data_type(value: str) -> str:
        if not value:
            return "A"
        try:
            float(value)
            return "N"
        except ValueError:
            pass
        try:
            from datetime import datetime
            datetime.fromisoformat(value.replace("Z", "+00:00"))
            return "D"
        except ValueError:
            pass
        return "A"

    # ---------- save_meta_to_json ----------

    def save_meta_to_json(self, json_path: str | Path, node_name: str) -> Path:
        meta = self.parse_meta()
        wrapped = {node_name: meta}
        path = Path(json_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(wrapped, f, indent=2, ensure_ascii=False)
        return path

    # ---------- save_meta_to_excel (full 6-row header, captioned cols shown blank) ----------

    def save_meta_to_excel(self, xlsx_path: str | Path, node_name: str) -> Path:
        from openpyxl import Workbook, load_workbook
        from openpyxl.utils import get_column_letter

        meta = self.parse_meta()
        path = Path(xlsx_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        wb = load_workbook(path) if path.exists() else Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        sheet_name = self._unique_sheet_name(wb, node_name)
        ws = wb.create_sheet(sheet_name)
        columns = self._build_meta_columns(meta)
        for i, c in enumerate(columns, start=1):
            kind, code, level, suffix, dt, header = c
            ws.cell(row=1, column=i, value=kind)
            ws.cell(row=2, column=i, value=code)
            ws.cell(row=3, column=i, value=level)
            ws.cell(row=4, column=i, value=suffix)
            ws.cell(row=5, column=i, value=dt)
            ws.cell(row=6, column=i, value=header)
            ws.column_dimensions[get_column_letter(i)].number_format = self._number_format_for(dt)

        wb.save(path)
        return path

    # ---------- template2_flat_brief (T1Sync 6-row header + brief data per asset) ----------

    def template2_flat_brief(self, xlsx_path: str | Path, sheet: str) -> Path:
        from openpyxl import Workbook, load_workbook
        from openpyxl.utils import get_column_letter

        # Walk CSV: each ASSET line opens a new asset; following ATTRIBUTE lines
        # add (code → SearchPath) to that asset.
        rows = self._read_csv()
        if len(rows) < 2:
            return Path(xlsx_path)

        header_line = rows[1]
        hi = self._header_idx(header_line)
        line_type_idx = hi.get("LineType", -1)
        attr_code_idx = hi.get("AttributeCode", -1)
        search_path_idx = hi.get("SearchPath", -1)

        assets: list[dict] = []
        attr_codes_ordered: list[str] = []
        seen: set[str] = set()
        current: dict | None = None

        for row in rows[2:]:
            if line_type_idx < 0 or line_type_idx >= len(row):
                continue
            lt = row[line_type_idx].upper()
            if lt == "ASSET":
                current = {"fields": {}, "attributes": {}}
                for f in self._nominated_fields:
                    if f in hi and hi[f] < len(row):
                        current["fields"][f] = row[hi[f]]
                assets.append(current)
            elif lt == "ATTRIBUTE" and current is not None:
                if attr_code_idx < 0 or attr_code_idx >= len(row):
                    continue
                code = row[attr_code_idx]
                if not code:
                    continue
                sp = row[search_path_idx] if 0 <= search_path_idx < len(row) else ""
                if code not in current["attributes"]:
                    current["attributes"][code] = sp
                if code not in seen:
                    seen.add(code)
                    attr_codes_ordered.append(code)

        path = Path(xlsx_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        wb = load_workbook(path) if path.exists() else Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        sheet_name = self._sanitize_sheet_name(sheet)
        if sheet_name in wb.sheetnames:
            # Reuse layout from row 1 / row 3 / row 6.
            ws = wb[sheet_name]
            max_col = ws.max_column
            layout = []
            for c in range(1, max_col + 1):
                kind = str(ws.cell(row=1, column=c).value or "")
                level = str(ws.cell(row=3, column=c).value or "")
                header = str(ws.cell(row=6, column=c).value or "")
                layout.append((kind, level, header))
        else:
            # Fresh sheet — brief layout: nominated fields + AttributeCode scalars.
            ws = wb.create_sheet(sheet_name)
            brief_cols = []
            for f in self._nominated_fields:
                brief_cols.append(("", "", "", "", "A", f))
            for code in attr_codes_ordered:
                brief_cols.append(("Attribute", "", "", "", "A", code))
            for i, c in enumerate(brief_cols, start=1):
                kind, code, level, suffix, dt, header = c
                ws.cell(row=1, column=i, value=kind)
                ws.cell(row=2, column=i, value=code)
                ws.cell(row=3, column=i, value=level)
                ws.cell(row=4, column=i, value=suffix)
                ws.cell(row=5, column=i, value=dt)
                ws.cell(row=6, column=i, value=header)
                ws.column_dimensions[get_column_letter(i)].number_format = self._number_format_for(dt)
            layout = [(c[0], c[2], c[5]) for c in brief_cols]

        # Data starts at row 7 — one row per asset.
        for r, asset in enumerate(assets):
            for i, (kind, level, header) in enumerate(layout, start=1):
                val = None
                if kind == "" and header:
                    val = asset["fields"].get(header)
                elif kind == "Attribute" and not level:
                    val = asset["attributes"].get(header)
                if val:
                    ws.cell(row=7 + r, column=i, value=val)

        wb.save(path)
        return path

    # ---------- simplify0 (row 1 full FORMAT, row 2 nominated compacted) ----------

    def simplify0(self, xlsx_path: str | Path, sheet: str) -> Path:
        from openpyxl import Workbook, load_workbook

        rows = self._read_csv()
        if len(rows) < 2:
            return Path(xlsx_path)
        format_line = rows[0]
        header_line = rows[1]
        hi = self._header_idx(header_line)
        line_type_idx = hi.get("LineType", -1)

        keep_indices: list[int] = []
        if line_type_idx >= 0:
            keep_indices.append(line_type_idx)
        for f in self._nominated_fields:
            if f in hi and hi[f] not in keep_indices:
                keep_indices.append(hi[f])

        assets = self._asset_rows(rows, line_type_idx)

        path = Path(xlsx_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        wb = load_workbook(path) if path.exists() else Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
        sheet_name = self._unique_sheet_name(wb, sheet)
        ws = wb.create_sheet(sheet_name)

        # Row 1 — verbatim CSV line 1 (FORMAT line; typically only A1 has content).
        for c, val in enumerate(format_line, start=1):
            if val:
                ws.cell(row=1, column=c, value=val)

        # Row 2 — nominated col names, compacted into A, B, C…
        for i, src_idx in enumerate(keep_indices, start=1):
            if src_idx < len(header_line) and header_line[src_idx]:
                ws.cell(row=2, column=i, value=header_line[src_idx])

        # Row 3+ — ASSET data at the same compacted positions.
        for r, asset in enumerate(assets):
            for i, src_idx in enumerate(keep_indices, start=1):
                if src_idx < len(asset) and asset[src_idx]:
                    ws.cell(row=3 + r, column=i, value=asset[src_idx])

        wb.save(path)
        return path

    # ---------- flat1 (verbatim CSV header rows 1-2 + sparse ASSET data) ----------

    def flat1(self, xlsx_path: str | Path, sheet: str) -> Path:
        from openpyxl import Workbook, load_workbook

        rows = self._read_csv()
        if len(rows) < 2:
            return Path(xlsx_path)
        format_line = rows[0]
        header_line = rows[1]
        hi = self._header_idx(header_line)
        line_type_idx = hi.get("LineType", -1)

        keep_indices: set[int] = set()
        if line_type_idx >= 0:
            keep_indices.add(line_type_idx)
        for f in self._nominated_fields:
            if f in hi:
                keep_indices.add(hi[f])

        assets = self._asset_rows(rows, line_type_idx)

        path = Path(xlsx_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        wb = load_workbook(path) if path.exists() else Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
        sheet_name = self._unique_sheet_name(wb, sheet)
        ws = wb.create_sheet(sheet_name)

        # Row 1 — verbatim CSV line 1.
        for c, val in enumerate(format_line, start=1):
            if val:
                ws.cell(row=1, column=c, value=val)

        # Row 2 — verbatim CSV line 2 (full column header, all ~290 columns).
        for c, val in enumerate(header_line, start=1):
            if val:
                ws.cell(row=2, column=c, value=val)

        # Row 3+ — ASSET data, only nominated cells populated at their original CSV positions.
        for r, asset in enumerate(assets):
            for c in keep_indices:
                if c < len(asset) and asset[c]:
                    ws.cell(row=3 + r, column=c + 1, value=asset[c])

        wb.save(path)
        return path

    # ---------- flat2 (1-row plain header + compact ASSET data) ----------

    def flat2(self, xlsx_path: str | Path, sheet: str) -> Path:
        from openpyxl import Workbook, load_workbook
        from openpyxl.utils import get_column_letter

        rows = self._read_csv()
        if len(rows) < 2:
            return Path(xlsx_path)
        header_line = rows[1]
        hi = self._header_idx(header_line)
        line_type_idx = hi.get("LineType", -1)
        assets = self._asset_rows(rows, line_type_idx)

        path = Path(xlsx_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        wb = load_workbook(path) if path.exists() else Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
        sheet_name = self._unique_sheet_name(wb, sheet)
        ws = wb.create_sheet(sheet_name)

        # Row 1 — just the nominated field names. Columns formatted as text so
        # values like "0100038" keep their leading zeros.
        for i, f in enumerate(self._nominated_fields, start=1):
            ws.cell(row=1, column=i, value=f)
            ws.column_dimensions[get_column_letter(i)].number_format = "@"

        # Row 2+ — one row per asset.
        for r, asset in enumerate(assets):
            for i, f in enumerate(self._nominated_fields, start=1):
                if f in hi:
                    idx = hi[f]
                    if idx < len(asset) and asset[idx]:
                        ws.cell(row=2 + r, column=i, value=asset[idx])

        wb.save(path)
        return path

    # ---------- shared helpers (mirror MetaSchema / C# Trans privates) ----------

    @staticmethod
    def _build_meta_columns(node_meta: dict) -> list[tuple]:
        """6-row column tuples (kind, code, level, suffix, dataType, header).
        Mirrors C# MetaSchema.BuildColumns."""
        columns: list[tuple] = []
        for f, dt in node_meta.get("fields", {}).items():
            columns.append(("", "", "", "", str(dt), f))
        for attr_code, attr_node in node_meta.get("attributes", {}).items():
            if not isinstance(attr_node, dict):
                continue
            dt = str(attr_node.get("dataType", "A"))
            columns.append(("Attribute", "", "", "", dt, attr_code))
            levels = attr_node.get("levels", {})
            for level_key in sorted(levels.keys(), key=lambda k: int(k) if str(k).isdigit() else 0):
                level_dict = levels[level_key]
                if not isinstance(level_dict, dict):
                    continue
                for suffix, leaf in level_dict.items():
                    caption, leaf_dt = "", ""
                    if isinstance(leaf, list) and len(leaf) >= 2:
                        caption, leaf_dt = str(leaf[0]), str(leaf[1])
                    elif isinstance(leaf, list) and len(leaf) == 1:
                        caption = str(leaf[0])
                    columns.append(("Attribute", attr_code, f"level_{level_key}", suffix, leaf_dt, caption))
        return columns

    @staticmethod
    def _sanitize_sheet_name(name: str) -> str:
        cleaned = "".join("_" if c in INVALID_SHEET_CHARS else c for c in (name or ""))
        return cleaned[:31] or "Sheet"

    @staticmethod
    def _unique_sheet_name(wb, base_name: str) -> str:
        name = Trans._sanitize_sheet_name(base_name)
        if name not in wb.sheetnames:
            return name
        stem = name[:29]
        for i in range(1, 100):
            candidate = f"{stem}{i:02d}"
            if candidate not in wb.sheetnames:
                return candidate
        raise ValueError(f"Could not allocate unique sheet name for {base_name!r}")

    @staticmethod
    def _number_format_for(dt: str) -> str:
        return {"N": "General", "D": "yyyy-mm-dd"}.get(dt, "@")
