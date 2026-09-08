import pandas as pd
import os
from pathlib import Path

# ================== CONFIGURE YOUR PATHS & COLUMNS HERE ==================
PRIMARY_FILE = r"/Users/maheshshekar/desktop/merger/primary/primary_students.xlsx"
SECONDARY_FOLDER = r"/Users/maheshshekar/desktop/merger/secondary"
OUTPUT_FILE = r"/Users/maheshshekar/desktop/merger/output/merged_result.xlsx"

# Select the two columns (order in primary file matters)
KEY_COLUMN = "Roll_No"      # Starting column (primary key)
VALUE_COLUMN = "Error"      # Ending column
# =================================================================================


def get_engine(filepath: str):
    """Return correct engine based on file extension"""
    if str(filepath).lower().endswith(".xls"):
        return "xlrd"
    return "openpyxl"


def select_sheet(filepath: str):
    """Show available sheets and let user choose one"""
    engine = get_engine(filepath)
    xl = pd.ExcelFile(filepath, engine=engine)
    sheet_names = xl.sheet_names

    print("\nAvailable sheets in the file:")
    for i, name in enumerate(sheet_names):
        print(f"  [{i}] {name}")

    while True:
        choice = input("\nEnter sheet number or sheet name: ").strip()

        # User entered a number
        if choice.isdigit():
            idx = int(choice)
            if 0 <= idx < len(sheet_names):
                return sheet_names[idx]
            else:
                print("Invalid number. Try again.")
        # User entered a sheet name
        elif choice in sheet_names:
            return choice
        else:
            print("Invalid sheet name. Try again.")


def merge_excel_files(primary_path: str, secondary_folder: str, output_path: str,
                      key_col: str, value_col: str):

    # ---------- 1. Load Primary & Select Sheet ----------
    if not os.path.exists(primary_path):
        raise FileNotFoundError(f"Primary file not found: {primary_path}")

    print(f"\nReading Primary file: {primary_path}")
    selected_sheet = select_sheet(primary_path)
    print(f"→ Selected sheet: '{selected_sheet}'")

    engine = get_engine(primary_path)
    primary_df = pd.read_excel(primary_path, sheet_name=selected_sheet, engine=engine)
    all_columns = list(primary_df.columns)

    # Validate both columns exist
    for col in [key_col, value_col]:
        if col not in all_columns:
            raise ValueError(f"Column '{col}' not found in primary file (sheet: {selected_sheet}).\n"
                             f"Available columns: {all_columns}")

    # Find positions and take all columns from key_col to value_col (inclusive)
    start_idx = all_columns.index(key_col)
    end_idx = all_columns.index(value_col)

    if start_idx > end_idx:
        start_idx, end_idx = end_idx, start_idx

    selected_columns = all_columns[start_idx : end_idx + 1]

    print(f"Primary file loaded → {len(primary_df)} rows")
    print(f"Selected columns (from '{key_col}' to '{value_col}'): {selected_columns}")

    primary_subset = primary_df[selected_columns].copy()
    all_dfs = [primary_subset]

    # ---------- 2. Process Secondary Files ----------
    secondary_path = Path(secondary_folder)
    if not secondary_path.exists():
        raise FileNotFoundError(f"Secondary folder not found: {secondary_folder}")

    excel_files = list(secondary_path.glob("*.xlsx")) + list(secondary_path.glob("*.xls"))

    if not excel_files:
        print("No secondary Excel files found.")
    else:
        print(f"\nFound {len(excel_files)} secondary file(s):")

    for file in excel_files:
        print(f"  → Processing: {file.name}")
        try:
            sec_engine = get_engine(file)
            sec_df = pd.read_excel(file, sheet_name=selected_sheet, engine=sec_engine)

            available_cols = [col for col in selected_columns if col in sec_df.columns]

            if key_col not in available_cols:
                print(f"    ⚠ Key column '{key_col}' missing. Skipping this file.")
                continue

            if not available_cols:
                print(f"    ⚠ No matching columns found. Skipping.")
                continue

            filtered = sec_df[available_cols].copy()

            for col in selected_columns:
                if col not in filtered.columns:
                    filtered[col] = pd.NA

            filtered = filtered[selected_columns]
            all_dfs.append(filtered)
            print(f"    ✓ Added {len(filtered)} rows | Columns used: {available_cols}")

        except ValueError:
            print(f"    ⚠ Sheet '{selected_sheet}' not found in {file.name}. Skipping.")
        except Exception as e:
            print(f"    ✗ Error reading {file.name}: {e}")

    # ---------- 3. Merge ----------
    final_df = pd.concat(all_dfs, ignore_index=True)
    final_df = final_df.drop_duplicates(subset=selected_columns, keep="first")
    final_df = final_df.sort_values(by=key_col).reset_index(drop=True)

    print(f"\nFinal merged shape: {final_df.shape}")
    print("\nPreview of merged data:")
    print(final_df.to_string(index=False))

    # ---------- 4. Save ----------
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    final_df.to_excel(output_path, index=False)
    print(f"\n✅ Merged file saved to: {output_path}")


if __name__ == "__main__":
    merge_excel_files(
        PRIMARY_FILE,
        SECONDARY_FOLDER,
        OUTPUT_FILE,
        KEY_COLUMN,
        VALUE_COLUMN
    )