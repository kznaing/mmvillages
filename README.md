# Myanmar Villages P-Codes Browser

A mobile-first web browser for Myanmar administrative place codes (P-codes) with a
**States/Regions → Townships → Villages** drill-down, plus Google Maps links for
villages that have coordinates.

**Live site (GitHub Pages):** https://kznaing.github.io/mmvillages/

## Data

Source: [MIMU Myanmar P-Codes](https://themimu.info/) — *Release 9.7, Jan 2026 (Countrywide)*.

- 18 States/Regions
- 358 Townships
- 358 Village files (33 townships have no recorded villages)
- 72,571 Villages (~77% with coordinates)

The original workbook is `Myanmar_PCodes_Release_9.7_Jan2026_Countrywide.xlsm`.
`export_pcodes.py` splits it into per-level JSON files.

## Project layout

```
export_pcodes.py              # generates data/ from the .xlsm workbook
index.html                    # single-file viewer (vanilla JS, no build step)
data/
  manifest.json               # states/regions, townships, counts, file references
  01_states_regions.json      # 18 active states/regions (eng + mmr names)
  02_townships/<SR>.json      # townships per state/region (eng + mmr names)
  03_villages/<Tsp>.json      # villages per township (eng + mmr names, tract, lat/lng)
```

## Regenerating the data

Requires Python 3 and `openpyxl`:

```bash
pip install openpyxl
python3 export_pcodes.py
```

## Running locally

Browsers block `fetch()` of local files under `file://`, so serve the folder:

```bash
python3 -m http.server
# open http://localhost:8000/
```

## Deployment

The site is static and published with GitHub Pages from the repository root
(`index.html` + `data/`).

## Credits & disclaimer

Place code data © Myanmar Information Management Unit (MIMU). Not for sale and
not for commercial use. Names and designations do not imply official endorsement
or acceptance.
