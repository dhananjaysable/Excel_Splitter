import sys
import pandas as pd
from pathlib import Path
from datetime import datetime

SEARCH_SETS = [
    ["beneficiary", "account"],       # Case 1 format
    ["account", "nickname"],          # Case 2 format
]


def find_target_column(df):
    """Find header row, header column, and mode type."""
    for keywords in SEARCH_SETS:
        for c in df.columns:
            for r in df.index:
                text = df.at[r, c]
                if isinstance(text, str):
                    t = text.lower().replace(" ", "")
                    if all(k in t for k in keywords):
                        return r, c, keywords
    return None


def process_sheet(df):
    """Process a single sheet and return updated dataframe."""
    result = find_target_column(df)

    if result is None:
        # No matching column → return sheet unchanged
        return df

    header_row, col_index, mode = result

    # Choose column names based on mode
    if mode == ["beneficiary", "account"]:
        col1 = "Beneficiary Name"
        col2 = "Account No"
    else:
        col1 = "Account No"
        col2 = "Nickname"

    # Insert new columns
    df.insert(col_index + 1, col1, None)
    df.insert(col_index + 2, col2, None)

    # Write header
    df.at[header_row, col1] = col1
    df.at[header_row, col2] = col2

    # Split rows
    for r in range(header_row + 1, len(df)):
        value = df.at[r, col_index]
        if isinstance(value, str) and "/" in value:
            p1, p2 = value.split("/", 1)
            df.at[r, col1] = p1.strip()
            df.at[r, col2] = p2.strip()

    return df


def main():
    if len(sys.argv) < 2:
        print('Usage: py account.py "yourfile.xlsx"')
        sys.exit(1)

    file = Path(sys.argv[1])

    if not file.exists():
        print(f"❌ File not found: {file}")
        sys.exit(1)

    # Read all sheets
    xls = pd.read_excel(file, sheet_name=None, header=None)

    processed_sheets = {}

    for sheet_name, df in xls.items():
        processed_sheets[sheet_name] = process_sheet(df)

    # Timestamp in output filename
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output = file.with_name(f"{file.stem}_{timestamp}.xlsx")

    # Write all sheets back
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in processed_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    print("===================================================")
    print("   ✅ Multi-Sheet Excel Processed Successfully!")
    print(f"   📁 Output File: {output.name}")
    print("===================================================")


if __name__ == "__main__":
    main()
