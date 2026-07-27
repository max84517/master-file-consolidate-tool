# Master File Consolidation Tool

A Python desktop application that consolidates HP Input Device master price tables scattered across multiple supplier folders into a single unified Excel file.

---

## Features

- **Dark-mode UI** built with CustomTkinter
- **Three-segment support** — NB KB, DT KB, Peripheral
- **Smart sheet detection**
  - NB: matches `FY##cNB` / `FY##bNB` (case-insensitive, spaces ignored)
  - DT / Peripheral: matches `FY##`
- **Always reads the newest Excel** in each supplier subfolder (by file mtime)
- **Merged-cell handling** — anchor values are propagated to all secondary cells during ingestion
- **Typo-tolerant column matching** — known header typos (e.g. `IncoTrem` → `IncoTerm`) are normalised via an alias table
- **Column-name alignment** — rows from different sheets/suppliers are aligned by column name, not index position; missing feature columns fill as `N/A`, missing value columns fill as `0`
- **Strict value-column format** — only columns matching `<keyword> <Month> <Year>` (e.g. `HP Cost Nov 2025`) are kept; bare `ODM Cost` / `Rebate` columns without a date suffix are automatically discarded
- **HP FY year auto-correction** — after reading each sheet, value-column dates are validated against HP Fiscal Year rules (FY## = Nov of year ##−1 through Oct of year ##); any wrong calendar year is silently corrected, and a summary of all corrections is printed to the log at the end of consolidation
- **Duplicate row removal** — fully identical rows are automatically dropped in both the by-segment and consolidate-all outputs; only the first occurrence is kept
- **Clean output guarantee** — any column with a blank header is removed from the output before saving; prevents unnamed columns with stray values
- **Check Files** — a pre-ingest dialog shows each supplier's latest Excel file name and last-modified date; files not updated in the current month are flagged with a ⚠ warning
- **FY coverage check** — instantly shows which suppliers are missing a sheet for the selected FY (amber warning, green all-clear)
- **Config persistence** — all folder paths saved to `config.json` and restored on next launch

---

## Project Structure

```
master-file-consolidate-tool/
├── pyproject.toml
├── src/
│   └── master_consolidate/
│       ├── main.py                    # entry point
│       ├── config/
│       │   ├── paths.py               # project-root resolution (frozen/dev)
│       │   └── settings.py            # config.json load / save
│       ├── ingestion/
│       │   └── ingest.py              # file discovery, sheet filtering, copy
│       ├── consolidation/
│       │   └── consolidate.py         # cleaning, column alignment, merging
│       └── ui/
│           └── app.py                 # CustomTkinter dark-mode UI
└── data/                              # created at runtime (gitignored)
    ├── source/
    │   ├── NB/
    │   ├── DT/
    │   └── Peripheral/
    ├── result_by_segment/
    │   ├── NB/
    │   ├── DT/
    │   └── Peripheral/
    └── consolidate_all/
```

---

## Requirements

- Python 3.11+
- [Poetry](https://python-poetry.org/)

---

## Setup

```bash
git clone https://github.com/max84517/master-file-consolidate-tool.git
cd master-file-consolidate-tool
poetry install
```

---

## Usage

```bash
poetry run master-consolidate
```

### Workflow

1. **Set Source Folders** — point NB KB, DT KB, and Peripheral to the top-level folders that contain the `Master price table_<Segment>` subfolder.
2. **Ingest Files** — scans each supplier subfolder, picks the newest Excel, extracts only the relevant FY sheets, and saves them to `data/source/<Segment>/`.
3. **Select FY** — the dropdown auto-populates from the ingested files (default: highest FY). A coverage banner immediately shows any suppliers missing the selected FY sheet.
4. **Select Value Columns** — choose any combination of HP Cost, ODM Cost, Rebate (all selected by default).
5. **Consolidate**
   - **Consolidate by Segment** — produces one Excel per segment:
     `Consolidated Master price table_<Segment>_<YYYYMMDD>.xlsx`
   - **Consolidate All** — runs by-segment first, then merges all three into:
     `Final Consolidated Master price table_<YYYYMMDD>.xlsx`

---

## Column Handling

| Column type | Empty cell | Rule |
|---|---|---|
| Feature (Segment, Series, Color, …) | `N/A` | Preserved as-is |
| Value (HP Cost, ODM Cost, Rebate, …) | `0` | Substring-matched on keyword |

To add a new mandatory feature column, append it to `MANDATORY_COLS` in [consolidate.py](src/master_consolidate/consolidation/consolidate.py).

To register a header typo/alias, add an entry to `_MANDATORY_ALIASES` in the same file.

---

## Dependencies

| Package | Version |
|---|---|
| openpyxl | ≥ 3.1.5 |
| customtkinter | ≥ 5.2.2 |
