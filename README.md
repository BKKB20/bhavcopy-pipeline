# Bhavcopy Pipeline

Automated download, extraction, and processing pipeline for 9,200+ daily Bhavcopy files from NSE and BSE covering 2000–2026 — with resumable downloads, intelligent retry logic, and a full audit log.

## The Problem

Historical Indian market data has no single clean source:
- NSE changed its file format mid-archive (old CSV vs new unified schema)
- BSE throttles heavily and returns timeout errors for holiday dates — indistinguishable from real failures without a BSE-specific holiday set
- Python's standard `holidays` library misses BSE-specific closures (Diwali, Ganesh Chaturthi, Muharram)
- Re-running a failed download blindly wastes the 20,000/day URL fetch quota

## What This Builds

A fully resumable 3-stage pipeline:

**Stage 1 — Download** (`download_bhavcopy.py`)
- Downloads all NSE Bhavcopy ZIPs from Jan 2000 to present (old + new format)
- Downloads all BSE Bhavcopy ZIPs from Jan 2007 to present
- Skips already-downloaded files — safe to re-run at any time
- Handles both NSE formats: `cmDDMMMYYYYbhav.csv.zip` and `BhavCopy_NSE_CM_*.csv.zip`

**Stage 2 — Retry** (`retry_final.py`)
- Classifies 132 failures into: 55 confirmed BSE holidays, 77 genuine throttle timeouts, 1 NSE SSL error
- Retries only real failures — not holidays
- Includes BSE-specific holiday set missing from standard libraries (Diwali, Ganesh Chaturthi, Muharram, COVID closure Apr 6 2020)

**Stage 3 — Extract** (`extract_nse_zips.py`, `extract_nse_new_format.py`, `extract_bse.py`)
- Extracts CSVs from all ZIPs into organised folders
- Handles BSE's uppercase `.ZIP` extension
- Deduplicates across `.zip` and `.ZIP` to avoid double-processing
- Writes 9-sheet Excel audit log: execution summary, per-file log (colour-coded), error log — one set per exchange

## File Structure
```
bhavcopy-pipeline/
├── download_bhavcopy.py        # Stage 1 — download all ZIPs
├── retry_final.py              # Stage 2 — classify and retry failures
├── extract_nse_zips.py         # Stage 3 — extract NSE old format
├── extract_nse_new_format.py   # Stage 3 — extract NSE new format
└── extract_bse.py              # Stage 3 — extract BSE
```

## Output Folder Structure
```
StockData/
├── NSE/
│   ├── Old Format/Extracted/   # ~6,068 CSVs (2000–2024)
│   ├── New Format/Extracted/   # ~300 CSVs (2024–present)
│   └── extract_nse_zips.py
├── BSE/
│   ├── Extracted/              # BSE CSVs
│   └── extract_bse.py
└── log sheet.xlsx              # 9-sheet audit log
```

## Key Lessons Learned

- BSE throttles at 3am — run downloads in the evening
- Never open `log sheet.xlsx` in Excel while scripts are writing — causes file corruption
- BSE returns timeout errors (not 404) for holiday dates — always verify before retrying
- Some pre-2010 BSE dates are genuine archive gaps — BSE never digitised them

## Prerequisites
```bash
pip install requests openpyxl pandas
```

## Built By

Bhavya Khaitan — Second Year B.Tech Textile Technology, VJTI Mumbai
