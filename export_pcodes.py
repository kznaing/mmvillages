#!/usr/bin/env python3
"""Export the MIMU Myanmar P-codes .xlsm into per-level CSV files and a manifest.

Outputs under data/:
  data/manifest.json                    list of SRs, townships, and file references
  data/01_states_regions.csv            18 active States/Regions (all columns)
  data/02_townships/<SR_Pcode>.csv      townships for each SR (all columns)
  data/03_villages/<Tsp_Pcode>.csv      villages for each township (all columns)
"""

import csv
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


def write_csv(path, headers, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh, lineterminator="\n")
        writer.writerow(headers)
        writer.writerows(rows)


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
    active_sr_rows = [r for r in sr_rows if r[sr_code_idx] in active_sr_set]

    # Group townships by SR, villages by township.
    tsp_by_sr = {code: [] for code in active_sr}
    for row in tsp_rows:
        tsp_by_sr.setdefault(row[tsp_sr_idx], []).append(row)

    vlg_by_tsp = {}
    for row in vlg_rows:
        vlg_by_tsp.setdefault(row[vlg_tsp_idx], []).append(row)

    # 1. states/regions
    write_csv(DATA / "01_states_regions.csv", sr_headers, active_sr_rows)

    # 2. townships per SR
    for sr_code in active_sr:
        write_csv(
            DATA / "02_townships" / f"{sr_code}.csv",
            tsp_headers,
            tsp_by_sr.get(sr_code, []),
        )

    # 3. villages per township (all townships, empty ones get header-only files)
    empty = 0
    total_villages = 0
    for row in tsp_rows:
        code = row[tsp_code_idx]
        villages = vlg_by_tsp.get(code, [])
        total_villages += len(villages)
        if not villages:
            empty += 1
        write_csv(
            DATA / "03_villages" / f"{code}.csv",
            vlg_headers,
            villages,
        )

    # manifest for the viewer
    sr_name_by_code = {r[sr_code_idx]: r[sr_name_idx] for r in active_sr_rows}
    manifest = {
        "source": SRC.name,
        "levels": {
            "states_regions": "01_states_regions.csv",
            "townships_dir": "02_townships",
            "villages_dir": "03_villages",
        },
        "village_columns": {
            "pcode": vlg_headers.index("Village_Pcode"),
            "name_eng": vlg_headers.index("Village_Name_Eng"),
            "name_mmr": vlg_headers.index("Village_Name_MMR"),
            "vt_pcode": vlg_headers.index("VT_Pcode"),
            "vt_name_eng": vlg_headers.index("Village_Tract_Name_Eng"),
            "longitude": vlg_headers.index("Longitude"),
            "latitude": vlg_headers.index("Latitude"),
        },
        "states_regions": [
            {
                "pcode": code,
                "name_eng": sr_name_by_code.get(code, code),
                "townships_file": f"{code}.csv",
                "township_count": len(tsp_by_sr.get(code, [])),
            }
            for code in active_sr
        ],
        "townships": [
            {
                "pcode": row[tsp_code_idx],
                "name_eng": row[tsp_headers.index("Township_Name_Eng")],
                "name_mmr": row[tsp_headers.index("Township_Name_MMR")],
                "sr_pcode": row[tsp_sr_idx],
                "villages_file": f"{row[tsp_code_idx]}.csv",
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
