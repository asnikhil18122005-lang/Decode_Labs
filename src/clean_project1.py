from pathlib import Path
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
RAW_FILE = PROJECT_ROOT / "data" / "raw" / "Dataset for Data Analytics.xlsx"
CLEANED_DIR = PROJECT_ROOT / "data" / "cleaned"
REPORTS_DIR = PROJECT_ROOT / "reports"

def main():
    CLEANED_DIR.mkdir(parents=True, exist_ok=True)
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(RAW_FILE, engine="openpyxl")

    original_rows = len(df)
    missing_before = df.isna().sum()

    for column in df.select_dtypes(include=["object"]).columns:
        df[column] = df[column].astype("string").str.strip()

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    for column in ["Quantity", "UnitPrice", "ItemsInCart", "TotalPrice"]:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df["CouponCode"] = df["CouponCode"].fillna("NO_COUPON")

    duplicate_rows = df.duplicated().sum()
    duplicate_order_ids = df.duplicated(subset=["OrderID"]).sum()

    df = df.drop_duplicates()
    df = df.drop_duplicates(subset=["OrderID"], keep="first")

    df["ExpectedTotalPrice"] = (df["Quantity"] * df["UnitPrice"]).round(2)
    price_mismatches = (df["ExpectedTotalPrice"] != df["TotalPrice"].round(2)).sum()
    df["TotalPrice"] = df["ExpectedTotalPrice"]
    df = df.drop(columns=["ExpectedTotalPrice"])

    df["Date"] = df["Date"].dt.strftime("%Y-%m-%d")

    final_duplicate_ids = df.duplicated(subset=["OrderID"]).sum()
    final_bad_dates = pd.to_datetime(df["Date"], errors="coerce").isna().sum()

    cleaned_excel = CLEANED_DIR / "cleaned_orders_project1.xlsx"
    cleaned_csv = CLEANED_DIR / "cleaned_orders_project1.csv"
    report_file = REPORTS_DIR / "cleaning_report_project1.txt"

    df.to_excel(cleaned_excel, index=False)
    df.to_csv(cleaned_csv, index=False)

    report = f"""
Data Analytics Project 1 - Cleaning Report

Original rows: {original_rows}
Final rows: {len(df)}

Missing values before cleaning:
{missing_before}

Missing values after cleaning:
{df.isna().sum()}

Duplicate full rows found: {duplicate_rows}
Duplicate OrderID values found: {duplicate_order_ids}
TotalPrice mismatches corrected: {price_mismatches}

Final duplicate OrderID count: {final_duplicate_ids}
Final incorrectly formatted date count: {final_bad_dates}

Output Excel file: {cleaned_excel}
Output CSV file: {cleaned_csv}
"""

    report_file.write_text(report, encoding="utf-8")

    print("Cleaning complete.")
    print(f"Final duplicate OrderID count: {final_duplicate_ids}")
    print(f"Final incorrectly formatted date count: {final_bad_dates}")

if __name__ == "__main__":
    main()
