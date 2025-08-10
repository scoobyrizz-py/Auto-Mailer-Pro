import pandas as pd

EXCEL_FILE = "sales_data.xlsx"
OUTPUT_FILE = "zip_lookup_template.csv"

def main():
    try:
        df = pd.read_excel(EXCEL_FILE, dtype=str)
    except Exception as e:
        print(f"❌ Failed to read {EXCEL_FILE}: {e}")
        return

    if 'Site Zip Code' not in df.columns:
        print("❌ 'Site Zip Code' column not found in the Excel file.")
        return

    zip_codes = df['Site Zip Code'].dropna().astype(str).str.zfill(5).unique()
    zip_codes.sort()

    zip_df = pd.DataFrame({'zip': zip_codes, 'city': '', 'state': ''})
    zip_df.to_csv(OUTPUT_FILE, index=False)
    print(f"✅ Template created: {OUTPUT_FILE}")
    print("➡️ Please open it and fill in the 'city' and 'state' columns.")

if __name__ == "__main__":
    main()
