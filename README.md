# Sales Analytics

This repository contains a Python churn analysis script for Odoo sales exports. The main script reads sales orders, removes excluded customers, computes churn risk, and exports charts and Excel reports.

## Files

- `churn_analysis.py`: Main churn analysis script.
- `excluded_customers.py`: Private list of customer names to exclude from analysis.
- `requirements.txt`: Python dependencies required by the script.

## What the script does

- Loads sales data from `.xlsx`, `.xls`, or `.csv` files.
- Validates required columns: `Creation Date`, `Customer`, `Order Reference`, `Salesperson`, `Status`, and `Total`.
- Cleans the dataset and removes excluded customers.
- Filters to confirmed sales orders (`Sales Order`).
- Computes churn metrics per customer, including inactive months, total revenue, and regularity.
- Scores churn risk and ranks churn candidates.
- Generates a bar chart and a monthly sales heatmap.
- Exports a full churn table and a top customer report to Excel.

## Usage

Install dependencies and run the script:

```bash
pip install -r requirements.txt
python churn_analysis.py --input combined.xlsx
```

Optional arguments:

- `--bar-out`, `-b`: Output filename for the churn bar chart (default: `churn_top_customers_bar_chart.png`).
- `--heatmap-out`, `-m`: Output filename for the heatmap (default: `churn_monthly_heatmap.png`).
- `--exclude`, `-x`: One or more additional customer names to exclude from analysis.

Example with custom outputs:

```bash
python churn_analysis.py --input combined.xlsx --bar-out churn_bar.png --heatmap-out churn_heatmap.png
```

## Output files

- `churned_customers.xlsx`: Full churn table for all customers.
- `top15_churned_customers.xlsx`: Top 15 churned customers by revenue.
- `churn_top_customers_bar_chart.png`: Bar chart of top churned customers.
- `churn_monthly_heatmap.png`: Heatmap of monthly sales activity.

## Private exclusions

Use `excluded_customers.py` to keep private customer names out of the repository and exclude them from analysis. If that file is missing, the script continues with an empty exclusion list.
