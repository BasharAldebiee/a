# 📮 Australia Post — CPI Calculator v2.0
**Data & Pricing Executive Tool**  
Source: ABS 6401.0 Consumer Price Index, Australia (Monthly)  
Base Period: September 2025 = 100.00 | Latest Data: March 2026

---

## What's in the App

- 24 months of real ABS data (Apr-2024 → Mar-2026)
- 9 regions: Australia, Sydney, Melbourne, Brisbane, Adelaide, Perth, Hobart, Darwin, Canberra
- CPI Index trend chart with annotated start/end points
- YoY % change bar chart with 2% and 4% pricing thresholds
- Automated Pricing Signal (🔴🟡🟢)
- All-regions snapshot table for latest month
- Export to branded Excel report or CSV

## Key Results (Mar-2026)

| Region     | CPI Index | YoY %  | MoM %  |
|-----------|-----------|--------|--------|
| Australia  | 102.44    | **4.6%** | 1.1%  |
| Sydney     | 102.41    | 4.4%   | 1.0%   |
| Adelaide   | 102.70    | **4.9%** | 1.3%  |
| Hobart     | 102.82    | **5.1%** | 1.3%  |
| Darwin     | 102.08    | 4.2%   | 1.2%   |

*Source: ABS, released 29 Apr 2026*

---

## Run in VS Code

```bash
# 1. Place these files in one folder:
#    cpi_calculator.py | abs_cpi_parser.py | 640101.xlsx | requirements.txt

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run
streamlit run cpi_calculator.py
# Opens at http://localhost:8501
```

---

## Update CPI Data Monthly

ABS releases monthly CPI approximately 4 weeks after month end.

1. Visit: https://www.abs.gov.au/statistics/economy/price-indexes-and-inflation/consumer-price-index-australia/latest-release#data-downloads
2. Download **640101.xlsx** (Table 1 — All Groups)
3. Replace the existing `640101.xlsx` in your project folder **OR** use the sidebar upload
4. The app auto-parses the new file — no code changes needed

---

## Formula

```
Custom % Change = ((End CPI − Start CPI) / Start CPI) × 100
```

Note: The ABS official YoY % uses the same formula but always compares to the same month of the prior year. Your custom calculator lets you compare any two periods.

---

## Pricing Decision Guide

| CPI Change | Signal   | Action |
|-----------|---------|--------|
| ≥ 4.0%    | 🔴 Strong Review | Submit pricing adjustment proposal to leadership |
| 2.0–3.9%  | 🟡 Review Warranted | Selective rate review for CPI-exposed products |
| < 2.0%    | 🟢 Stable | Monitor quarterly, no immediate action |

---

## File Structure

```
cpi_calculator/
├── cpi_calculator.py    ← Main Streamlit app (run this)
├── abs_cpi_parser.py    ← ABS file parser module
├── 640101.xlsx          ← Real ABS data (replace with latest)
├── requirements.txt
└── README.md
```

*Australia Post — Data & Pricing Executive Tool. Internal use.*
