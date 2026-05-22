# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app (opens at http://localhost:8501)
streamlit run cpi_calculator.py
```

## Architecture

Two-module structure: a parser and a Streamlit UI.

**`abs_cpi_parser.py`** — stateless parser for ABS 6401.0 Table 1 Excel files. The key constraint is the fixed column layout of the `Data1` sheet: cols 1–9 are CPI index values, 10–18 are YoY%, 19–27 are MoM%, each group covering the same 9 cities in order (`CITIES` list). Data rows begin at row index 10 (Excel row 11). `load_abs_file()` accepts a file path, raw bytes, or file-like object and returns a clean DataFrame with columns `Date`, `Period`, `{City}_Index`, `{City}_YoY`, `{City}_MoM`.

**`cpi_calculator.py`** — Streamlit app. Data flows through three stages:
1. **Load** — `load_bundled_data()` (cached) opens the local `640101.xlsx`; sidebar upload calls `try_load_upload()` as an alternative.
2. **Calculate** — `calc_custom_change()` from the parser computes `((end − start) / start) × 100` for the user-selected city and period range.
3. **Render** — KPI cards, pricing signal box (thresholds: ≥4% red, 2–4% yellow, <2% green), Plotly charts, results table, regional snapshot, and export buttons, all in `main()`.

**`build_excel_report()`** produces a branded two-sheet Excel export (styled report + raw data) using `openpyxl` directly — not the `xlsxwriter` engine.

## Updating Data

Replace `640101.xlsx` with the latest file from ABS (Table 1 — All Groups CPI). The parser requires the `Data1` sheet with the exact column layout described above. No code changes are needed; the sidebar also supports live upload without restarting the app.
