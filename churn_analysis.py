"""
Churn analysis for Odoo sales export.

Creates churn detection, scoring and multi-panel plots.

Dependencies: pandas, numpy, matplotlib, seaborn, openpyxl

Configuration variables are defined at the top for easy tuning.
"""
import argparse
import os
from datetime import datetime

import matplotlib
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

import arabic_reshaper
from bidi.algorithm import get_display

# --------------------
# Font Configuration for Arabic Text
# --------------------
arabic_fonts = ["Arial", "Tahoma", "DejaVu Sans", "FreeSans", "Amiri", "Scheherazade"]
available = [f.name for f in fm.fontManager.ttflist]
chosen_font = next((f for f in arabic_fonts if f in available), None)

if chosen_font:
    matplotlib.rcParams["font.family"] = chosen_font
else:
    print("WARNING: No Arabic-compatible font found. Install Amiri font for best results.")

# --------------------
# Configuration
# --------------------
INPUT_FILE = "combined.xlsx"
CHURN_THRESHOLD_MONTHS = 6
MIN_PURCHASE_MONTHS = 3
TOP_N_CUSTOMERS = 15
HEATMAP_TOP_N = 20
BAR_CHART_OUT = "churn_top_customers_bar_chart.png"
HEATMAP_OUT = "churn_monthly_heatmap.png"
EXCEL_OUT = "churned_customers.xlsx"
TOP15_EXCEL_OUT = "top15_churned_customers.xlsx"

# Customers to exclude from all analysis
# Add customer names exactly as they appear in the data.
# Use excluded_customers.py for private customer exclusion lists.
try:
    from excluded_customers import EXCLUDED_CUSTOMERS
except (ImportError, ModuleNotFoundError):
    EXCLUDED_CUSTOMERS = []


# --------------------
# Arabic Text Helper Functions
# --------------------
def fix_arabic(text):
    """Fix Arabic text for proper display in matplotlib.
    
    Reshapes Arabic characters and applies bidi algorithm for correct rendering.
    """
    if not isinstance(text, str):
        text = str(text)
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)


def fix_arabic_list(lst):
    """Apply fix_arabic to every item in a list.
    
    Returns a new list with all strings processed through fix_arabic().
    """
    return [fix_arabic(item) for item in lst]


def load_data(path):
    """Load CSV or Excel and return a DataFrame.

    - Detects by extension.
    - Ensures required columns exist.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(path, engine="openpyxl")
    elif ext == ".csv":
        df = pd.read_csv(path)
    else:
        raise ValueError(f"Unsupported extension: {ext}")

    required = {"Creation Date", "Customer", "Order Reference", "Salesperson", "Status", "Total"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Parse Creation Date as datetime
    df["Creation Date"] = pd.to_datetime(df["Creation Date"], errors="coerce")
    if df["Creation Date"].isna().any():
        # drop rows with invalid dates
        df = df.dropna(subset=["Creation Date"])

    # Strip whitespace from Customer and Salesperson
    df["Customer"] = df["Customer"].astype(str).str.strip()
    df["Salesperson"] = df["Salesperson"].astype(str).str.strip()

    # Ensure Total is numeric
    df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0.0)

    # Apply customer exclusions
    if EXCLUDED_CUSTOMERS:
        excluded_count = len(df[df["Customer"].isin(EXCLUDED_CUSTOMERS)])
        df = df[~df["Customer"].isin(EXCLUDED_CUSTOMERS)]
        print(f"[EXCLUSION] Removed {excluded_count} rows from {len(EXCLUDED_CUSTOMERS)} excluded customer(s)")
        for cust in EXCLUDED_CUSTOMERS:
            print(f"  - {cust}")

    return df


def prepare_monthly_series(df):
    """Aggregate per-customer monthly totals and return dict of Series.

    Returns:
      monthly_by_customer: dict mapping customer -> pd.Series indexed by Period('M')
      global_month_index: PeriodIndex spanning full dataset min->max
      total_by_month: pd.Series aggregated across all customers (PeriodIndex)
    """
    # Create Year-Month period column
    df["Year-Month"] = df["Creation Date"].dt.to_period("M")

    # Aggregate sums per customer per month
    grouped = (
        df.groupby(["Customer", "Year-Month"])["Total"]
        .sum()
        .reset_index()
    )

    # Build per-customer monthly series, reindexing to fill every month
    monthly_by_customer = {}
    for cust, g in grouped.groupby("Customer"):
        s = g.set_index("Year-Month")["Total"].sort_index()
        # Create full period range for this customer
        full_idx = pd.period_range(start=s.index.min(), end=s.index.max(), freq="M")
        s = s.reindex(full_idx, fill_value=0.0)
        s.index.name = "Year-Month"
        monthly_by_customer[cust] = s

    # Global time index across dataset
    global_min = df["Creation Date"].dt.to_period("M").min()
    global_max = df["Creation Date"].dt.to_period("M").max()
    global_month_index = pd.period_range(start=global_min, end=global_max, freq="M")

    # Total by month across all customers (use global index)
    total_by_month = (
        df.groupby("Year-Month")["Total"].sum().reindex(global_month_index, fill_value=0.0)
    )

    return monthly_by_customer, global_month_index, total_by_month


def compute_churn_table(monthly_by_customer, df, global_latest_period):
    """Compute churn metrics for every customer and return a DataFrame.

    Metrics include: first_purchase, last_purchase (period), months_inactive,
    total_revenue, active_months, avg_monthly_spend, first purchase date
    """
    rows = []
    # Precompute maxs for normalization later
    for cust, series in monthly_by_customer.items():
        total_revenue = float(series.sum())
        active_months = int((series > 0).sum())
        # first and last purchase (period -> take as Period)
        if active_months > 0:
            first_purchase = series[series > 0].index.min()
            last_purchase = series[series > 0].index.max()
        else:
            first_purchase = pd.NaT
            last_purchase = pd.NaT

        # months since last purchase relative to global_latest_period
        if pd.isna(last_purchase):
            months_inactive = None
        else:
            months_inactive = (global_latest_period.year - last_purchase.year) * 12 + (
                global_latest_period.month - last_purchase.month
            )

        # average monthly spend over active months (avoid div by zero)
        avg_monthly_spend = float(total_revenue / active_months) if active_months > 0 else 0.0

        # regularity: fraction of months with purchases in their active period
        if active_months > 0:
            total_months_range = (last_purchase.year - first_purchase.year) * 12 + (
                last_purchase.month - first_purchase.month
            ) + 1
            regularity = active_months / total_months_range
        else:
            total_months_range = 0
            regularity = 0.0

        # Salesperson: choose the most common salesperson for the customer
        sp = df.loc[df["Customer"] == cust, "Salesperson"]
        if not sp.empty:
            try:
                salesperson = sp.mode().iloc[0]
            except Exception:
                salesperson = sp.iloc[0]
        else:
            salesperson = ""

        rows.append(
            {
                "Customer": cust,
                "Salesperson": salesperson,
                "First Purchase": first_purchase.to_timestamp() if not pd.isna(first_purchase) else pd.NaT,
                "Last Purchase": last_purchase.to_timestamp() if not pd.isna(last_purchase) else pd.NaT,
                "Months Inactive": months_inactive,
                "Total Revenue": total_revenue,
                "Active Months": active_months,
                "Avg Monthly Spend": avg_monthly_spend,
                "Regularity": regularity,
            }
        )

    churn_df = pd.DataFrame(rows)
    # Compute churn flag based on rules
    churn_df["Is Churn Candidate"] = churn_df["Active Months"] >= MIN_PURCHASE_MONTHS
    churn_df["Months Inactive"] = churn_df["Months Inactive"].fillna(9999).astype(int)
    churn_df["Churned"] = (churn_df["Is Churn Candidate"]) & (churn_df["Months Inactive"] > CHURN_THRESHOLD_MONTHS)

    return churn_df


def score_churn(churn_df):
    """Assign a risk score to churned customers.

    Score combines normalized revenue, months inactive, and regularity.
    Returns churn_df with added `Risk Score` column and sorted by it.
    """
    # Work on numeric columns; avoid division by zero
    max_rev = churn_df["Total Revenue"].replace(0, np.nan).max()
    max_inact = churn_df["Months Inactive"].replace(0, np.nan).max()

    # Normalized components
    churn_df["rev_norm"] = churn_df["Total Revenue"] / max_rev if pd.notna(max_rev) else 0.0
    churn_df["inact_norm"] = churn_df["Months Inactive"] / max_inact if pd.notna(max_inact) else 0.0
    churn_df["reg_norm"] = churn_df["Regularity"].fillna(0.0)

    # Weights: revenue (0.5), inactivity (0.3), regularity (0.2)
    churn_df["Risk Score"] = (
        0.5 * churn_df["rev_norm"].fillna(0)
        + 0.3 * churn_df["inact_norm"].fillna(0)
        + 0.2 * churn_df["reg_norm"].fillna(0)
    )

    churn_df = churn_df.sort_values("Risk Score", ascending=False).reset_index(drop=True)
    churn_df.index.name = "Rank"
    return churn_df


def make_plots(monthly_by_customer, global_month_index, total_by_month, churn_df):
    """Create and save two separate plots: bar chart and heatmap."""
    sns.set(style="whitegrid")
    # Convert PeriodIndex to strings for plotting columns
    months_str = [p.strftime("%Y-%m") for p in global_month_index]

    # Prepare matrix for heatmap (rows = customers, cols = months)
    # For top customers by revenue
    top_customers = churn_df.sort_values("Total Revenue", ascending=False).head(HEATMAP_TOP_N)["Customer"].tolist()

    heatmap_df = pd.DataFrame(index=top_customers, columns=months_str, data=0.0)
    for cust in top_customers:
        s = monthly_by_customer.get(cust)
        if s is None:
            continue
        # convert periods to strings
        for p, val in s.items():
            pstr = p.strftime("%Y-%m")
            if pstr in heatmap_df.columns:
                heatmap_df.at[cust, pstr] = val

    # Top churned customers for bar chart (top N by revenue but only churned)
    churned_top = churn_df[churn_df["Churned"]].sort_values("Total Revenue", ascending=False).head(TOP_N_CUSTOMERS)

    # ===== PLOT 1: Bar Chart =====
    fig1, ax1 = plt.subplots(figsize=(12, 8))
    
    if not churned_top.empty:
        # Fix Arabic customer names for display
        fixed_customer_names = fix_arabic_list(churned_top["Customer"].tolist())
        y = fixed_customer_names
        x = churned_top["Total Revenue"]
        inact = churned_top["Months Inactive"].clip(lower=CHURN_THRESHOLD_MONTHS, upper=CHURN_THRESHOLD_MONTHS * 3)
        cmap = plt.get_cmap("YlOrRd")
        norm = plt.Normalize(vmin=CHURN_THRESHOLD_MONTHS, vmax=CHURN_THRESHOLD_MONTHS * 3)
        colors = [cmap(norm(v)) for v in inact]
        ax1.barh(y, x, color=colors)
        ax1.invert_yaxis()
        ax1.set_title(fix_arabic(f"Top {TOP_N_CUSTOMERS} Churned Customers by Revenue"), fontsize=14, fontweight="bold")
        ax1.set_xlabel(fix_arabic("Total Revenue"), fontsize=12)
        sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
        sm.set_array([])
        cbar = fig1.colorbar(sm, ax=ax1, orientation="vertical", pad=0.02)
        cbar.set_label(fix_arabic("Months Inactive"), fontsize=11)
    else:
        ax1.text(0.5, 0.5, fix_arabic("No churned customers found"), ha="center")
        ax1.set_axis_off()

    fig1.tight_layout()
    fig1.savefig(BAR_CHART_OUT, dpi=150, bbox_inches="tight")
    plt.close(fig1)

    # ===== PLOT 2: Heatmap =====
    fig2, ax2 = plt.subplots(figsize=(14, 10))
    
    if not heatmap_df.empty:
        sns.heatmap(
            heatmap_df.replace(0, np.nan),
            ax=ax2,
            cmap="YlGnBu",
            cbar_kws={"label": fix_arabic("Sales Amount")},
            linewidths=0.5,
            linecolor="gray",
        )
        ax2.set_title(fix_arabic(f"Monthly Sales Heatmap (Top {HEATMAP_TOP_N} Customers)"), fontsize=14, fontweight="bold", pad=20)
        ax2.set_ylabel(fix_arabic("Customer"), fontsize=12)
        ax2.set_xlabel(fix_arabic("Year-Month"), fontsize=12)
        fixed_yticklabels = fix_arabic_list(list(heatmap_df.index))
        ax2.set_yticklabels(fixed_yticklabels, fontsize=9)
        ax2.set_xticklabels(ax2.get_xticklabels(), rotation=45, ha="right")
    else:
        ax2.text(0.5, 0.5, fix_arabic("No data for heatmap"), ha="center")
        ax2.set_axis_off()

    fig2.tight_layout()
    fig2.savefig(HEATMAP_OUT, dpi=150, bbox_inches="tight")
    plt.close(fig2)


def print_and_export(churn_df):
    """Print top 20 churned customers summary and export full tables to Excel."""
    # Prepare summary
    top20 = churn_df[churn_df["Churned"]].sort_values("Total Revenue", ascending=False).head(20)
    summary_cols = [
        "Customer",
        "Salesperson",
        "First Purchase",
        "Last Purchase",
        "Months Inactive",
        "Total Revenue",
        "Avg Monthly Spend",
    ]
    print("Top 20 churned customers summary:\n")
    print(top20[summary_cols].to_string(index=False))

    # Export full churn table
    churn_df.to_excel(EXCEL_OUT, index=False, engine="openpyxl")

    # Export top 15 churned customers with full details
    top15 = churn_df[churn_df["Churned"]].sort_values("Total Revenue", ascending=False).head(15)
    top15.to_excel(TOP15_EXCEL_OUT, index=False, engine="openpyxl")

    print(f"\nFull churn table exported to: {EXCEL_OUT}")
    print(f"Top 15 churned customers exported to: {TOP15_EXCEL_OUT}")
    print(f"Bar chart saved to: {BAR_CHART_OUT}")
    print(f"Heatmap saved to: {HEATMAP_OUT}")
    print(f"Heatmap saved to: {HEATMAP_OUT}")


def main():
    global BAR_CHART_OUT, HEATMAP_OUT, EXCLUDED_CUSTOMERS
    parser = argparse.ArgumentParser(description="Churn analysis for Odoo sales exports")
    parser.add_argument("--input", "-i", default=INPUT_FILE, help="Input file (.xlsx or .csv)")
    parser.add_argument("--bar-out", "-b", default=BAR_CHART_OUT, help="Bar chart output filename")
    parser.add_argument("--heatmap-out", "-m", default=HEATMAP_OUT, help="Heatmap output filename")
    parser.add_argument("--exclude", "-x", nargs="*", default=[], help="Customer names to exclude (space-separated)")
    args = parser.parse_args()

    BAR_CHART_OUT = args.bar_out
    HEATMAP_OUT = args.heatmap_out
    if args.exclude:
        EXCLUDED_CUSTOMERS = args.exclude

    # Load and clean
    df = load_data(args.input)

    # Display exclusion notice if active
    if EXCLUDED_CUSTOMERS:
        print(f"\n{'='*60}")
        print(f"ℹ️  EXCLUSION NOTICE")
        print(f"{'='*60}")
        print(f"{len(EXCLUDED_CUSTOMERS)} customer(s) excluded from this analysis:")
        for cust in EXCLUDED_CUSTOMERS:
            print(f"  • {cust}")
        print(f"{'='*60}\n")

    # Filter to only confirmed orders
    df = df[df["Status"].str.strip().eq("Sales Order")].copy()

    if df.empty:
        print("No 'Sales Order' rows found after filtering. Exiting.")
        return

    # Prepare monthly series
    monthly_by_customer, global_month_index, total_by_month = prepare_monthly_series(df)

    # Latest period in dataset
    latest_period = global_month_index[-1]

    # Compute churn table
    churn_df = compute_churn_table(monthly_by_customer, df, latest_period)

    # Score churn
    churn_df = score_churn(churn_df)

    # Make plots
    make_plots(monthly_by_customer, global_month_index, total_by_month, churn_df)

    # Print and export
    print_and_export(churn_df)


if __name__ == "__main__":
    # Test Arabic text fixing
    # Load test strings from the private `excluded_customers.py` when available
    # so sensitive names are not stored in this public file.
    if EXCLUDED_CUSTOMERS:
        test_strings = EXCLUDED_CUSTOMERS[:3]
    else:
        # Safe, non-sensitive placeholders for public repo
        test_strings = ["مثال عميل", "مثال مندوب", "اسم تجريبي"]

    print("Arabic fix test:")
    for s in test_strings:
        print(f"  Original : {s}")
        print(f"  Fixed    : {fix_arabic(s)}")
        print()

    main()
