# 📊 Excel Sales Report Automation

A Python script that automatically cleans raw sales data and generates a fully formatted Excel report — saving hours of manual spreadsheet work.

## What it does

- ✅ Cleans messy raw data (fixes inconsistent casing, missing values, date formats)
- ✅ Calculates revenue automatically
- ✅ Generates summary tables by **Product**, **Month**, and **Region**
- ✅ Produces a beautifully formatted Excel report with:
  - Colour-coded headers
  - Alternating row colours
  - Auto-fitted column widths
  - Currency formatting
  - A clean Summary dashboard sheet

## Output Preview

The script generates a multi-sheet Excel report:

| Sheet | Contents |
|---|---|
| 📊 Summary | High-level KPIs at a glance |
| Clean Data | Cleaned and processed raw data |
| By Product | Revenue & units grouped by product |
| By Month | Monthly sales trend |
| By Region | Revenue breakdown by region |

## Getting Started

### 1. Install dependencies

```bash
pip install pandas openpyxl
```

### 2. Run the script

```bash
python sales_report.py
```

> If no input file is found, the script auto-generates sample data so you can see it in action immediately.

### 3. Use your own data

Replace `raw_sales_data.xlsx` with your own file. Make sure it has these columns:

| Column | Description |
|---|---|
| Date | Sale date (any standard format) |
| Product | Product name |
| Units Sold | Number of units |
| Unit Price | Price per unit |
| Region | Sales region |
| Salesperson | Name of salesperson |
| Notes | Any notes (optional) |

## Tech Stack

- **Python 3.8+**
- **pandas** — data cleaning and aggregation
- **openpyxl** — Excel formatting and styling

## Use Cases

This kind of automation is perfect for:
- Small businesses tracking weekly/monthly sales
- Finance teams generating recurring reports
- Anyone spending hours copying and formatting spreadsheets manually

---

Built by [Mark](https://www.fiverr.com/markpydev) — Python Developer specialising in Excel & Google Sheets automation.

💼 Need something automated? [Hire me on Fiverr](https://www.fiverr.com/markpydev)
