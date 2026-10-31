#!/usr/bin/env python3
"""Export the MIMU Myanmar P-codes .xlsm into per-level JSON files and a manifest.

Outputs under data/:
  data/manifest.json                    list of SRs, townships, and file references
  data/01_states_regions.json           18 active States/Regions
  data/02_townships/<SR_Pcode>.json     townships for each SR
  data/03_villages/<Tsp_Pcode>.json     villages for each township
"""

import json
import warnings
from pathlib import Path

import openpyxl

warnings.simplefilter("ignore")

SRC = Path(__file__).parent / "Myanmar_PCodes_Release_9.7_Jan2026_Countrywide.xlsm"
DATA = Path(__file__).parent / "data"

SR_SHEET = "01_SR"
SR_HEADER_ROW = 1
TSP_SHEET = "03_Township"
TSP_HEADER_ROW = 3
VLG_SHEET = "_07_Village"
VLG_HEADER_ROW = 6


def clean(value):
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    return value


def load_sheet(wb, name, header_row):
    """Return (headers, rows) skipping blank/instruction rows below the header."""
    ws = wb[name]
    rows = list(ws.iter_rows(min_row=header_row, values_only=True))
    headers = [clean(h) for h in rows[0]]
    ncol = len(headers)
    header_key = headers[0] if headers else None
    data = []
    for row in rows[1:]:
        row = list(row[:ncol])
        first = row[0]
        if first is None or (isinstance(first, str) and first.strip().startswith("-")):
            continue
        if header_key is not None and first == header_key:
            continue
        data.append([clean(c) for c in row])
    return headers, data


def write_json(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(obj, fh, ensure_ascii=False, separators=(",", ":"))


def main():
    wb = openpyxl.load_workbook(SRC, read_only=True, keep_vba=True, data_only=True)

    sr_headers, sr_rows = load_sheet(wb, SR_SHEET, SR_HEADER_ROW)
    tsp_headers, tsp_rows = load_sheet(wb, TSP_SHEET, TSP_HEADER_ROW)
    vlg_headers, vlg_rows = load_sheet(wb, VLG_SHEET, VLG_HEADER_ROW)

    tsp_sr_idx = tsp_headers.index("SR_Pcode")
    tsp_code_idx = tsp_headers.index("Tsp_Pcode")
    vlg_tsp_idx = vlg_headers.index("Tsp_Pcode")

    # Active SRs = those referenced by the township sheet, in SR order.
    active_sr = []
    seen = set()
    for row in tsp_rows:
        code = row[tsp_sr_idx]
        if code not in seen:
            seen.add(code)
            active_sr.append(code)
    active_sr_set = set(active_sr)

    sr_code_idx = sr_headers.index("SR_Pcode")
    sr_name_idx = sr_headers.index("SR_Name_Eng")
    sr_name_mmr_idx = sr_headers.index("SR_Name_MMR")
    active_sr_rows = [r for r in sr_rows if r[sr_code_idx] in active_sr_set]

    tsp_name_idx = tsp_headers.index("Township_Name_Eng")
    tsp_name_mmr_idx = tsp_headers.index("Township_Name_MMR")

    vlg_code_idx = vlg_headers.index("Village_Pcode")
    vlg_name_idx = vlg_headers.index("Village_Name_Eng")
    vlg_name_mmr_idx = vlg_headers.index("Village_Name_MMR")
    vlg_vt_name_idx = vlg_headers.index("Village_Tract_Name_Eng")
    vlg_lng_idx = vlg_headers.index("Longitude")
    vlg_lat_idx = vlg_headers.index("Latitude")

    # Group townships by SR, villages by township.
    tsp_by_sr = {code: [] for code in active_sr}
    for row in tsp_rows:
        tsp_by_sr.setdefault(row[tsp_sr_idx], []).append(row)

    vlg_by_tsp = {}
    for row in vlg_rows:
        vlg_by_tsp.setdefault(row[vlg_tsp_idx], []).append(row)

    def sr_obj(row):
        return {
            "pcode": row[sr_code_idx],
            "name_eng": row[sr_name_idx],
            "name_mmr": row[sr_name_mmr_idx],
        }

    def tsp_obj(row):
        return {
            "pcode": row[tsp_code_idx],
            "name_eng": row[tsp_name_idx],
            "name_mmr": row[tsp_name_mmr_idx],
        }

    def vlg_obj(row):
        return {
            "pcode": row[vlg_code_idx],
            "name_eng": row[vlg_name_idx],
            "name_mmr": row[vlg_name_mmr_idx],
            "vt_name_eng": row[vlg_vt_name_idx],
            "longitude": row[vlg_lng_idx],
            "latitude": row[vlg_lat_idx],
        }

    # 1. states/regions
    write_json(DATA / "01_states_regions.json", [sr_obj(r) for r in active_sr_rows])

    # 2. townships per SR
    for sr_code in active_sr:
        write_json(
            DATA / "02_townships" / f"{sr_code}.json",
            [tsp_obj(r) for r in tsp_by_sr.get(sr_code, [])],
        )

    # 3. villages per township (all townships, empty ones get [])
    empty = 0
    total_villages = 0
    for row in tsp_rows:
        code = row[tsp_code_idx]
        villages = vlg_by_tsp.get(code, [])
        total_villages += len(villages)
        if not villages:
            empty += 1
        write_json(
            DATA / "03_villages" / f"{code}.json",
            [vlg_obj(r) for r in villages],
        )

    # manifest for the viewer
    sr_by_code = {r[sr_code_idx]: r for r in active_sr_rows}
    manifest = {
        "source": SRC.name,
        "levels": {
            "states_regions": "01_states_regions.json",
            "townships_dir": "02_townships",
            "villages_dir": "03_villages",
        },
        "states_regions": [
            {
                "pcode": code,
                "name_eng": sr_by_code[code][sr_name_idx],
                "name_mmr": sr_by_code[code][sr_name_mmr_idx],
                "townships_file": f"{code}.json",
                "township_count": len(tsp_by_sr.get(code, [])),
            }
            for code in active_sr
        ],
        "townships": [
            {
                "pcode": row[tsp_code_idx],
                "name_eng": row[tsp_name_idx],
                "name_mmr": row[tsp_name_mmr_idx],
                "sr_pcode": row[tsp_sr_idx],
                "villages_file": f"{row[tsp_code_idx]}.json",
                "village_count": len(vlg_by_tsp.get(row[tsp_code_idx], [])),
            }
            for row in tsp_rows
        ],
    }
    DATA.mkdir(parents=True, exist_ok=True)
    with (DATA / "manifest.json").open("w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)

    print(f"States/Regions : {len(active_sr_rows)}")
    print(f"Township files : {len(tsp_rows)}")
    print(f"Village files  : {len(tsp_rows)} ({empty} empty)")
    print(f"Villages       : {total_villages}")
    print(f"Output         : {DATA}")


if __name__ == "__main__":
    main()
