import pandas as pd
import os
from pathlib import Path

# ================== CONFIGURE YOUR PATHS, COLUMNS & SHEET HERE ==================
PRIMARY_FILE = r"/Users/maheshshekar/desktop/merger/primary/primary_students.xlsx"
SECONDARY_FOLDER = r"/Users/maheshshekar/desktop/merger/secondary"
OUTPUT_FILE = r"/Users/maheshshekar/desktop/merger/output/merged_result.xlsx"

# Select the two columns (order in primary file matters)
KEY_COLUMN = "Roll_No"      # Starting column (primary key)
VALUE_COLUMN = "Error"      # Ending column

# ---------- NEW OPTION ----------
# Set the sheet name you want to use from all Excel files
# Examples:
#   SHEET_NAME = None          → uses the first sheet (default)
#   SHEET_NAME = "Sheet1"      → uses the sheet named "Sheet1"
#   SHEET_NAME = "StudentData" → uses the sheet named "StudentData"
#   SHEET_NAME = 0             → uses the first sheet (same as None)
#   SHEET_NAME = 1             → uses the second sheet
SHEET_NAME = None
# =================================================================================


def merge_excel_files(primary_path: str, secondary_folder: str, output_path: str,
                      key_col: str, value_col: str, sheet_name=None):

    # ---------- 1. Load Primary ----------
    if not os.path.exists(primary_path):
        raise FileNotFoundError(f"Primary file not found: {primary_path}")

    try:
        primary_df = pd.read_excel(primary_path, sheet_name=sheet_name)
    except ValueError as e:
        raise ValueError(f"Could not read sheet '{sheet_name}' from primary file.\n{e}")

    all_columns = list(primary_df.columns)

    # Validate both columns exist
    for col in [key_col, value_col]:
        if col not in all_columns:
            raise ValueError(f"Column '{col}' not found in primary file (sheet: {sheet_name}).\n"
                             f"Available columns: {all_columns}")

    # Find positions and take all columns from key_col to value_col (inclusive)
    start_idx = all_columns.index(key_col)
    end_idx = all_columns.index(value_col)

    if start_idx > end_idx:
        start_idx, end_idx = end_idx, start_idx

    selected_columns = all_columns[start_idx : end_idx + 1]

    print(f"Primary file loaded → {len(primary_df)} rows")
    print(f"Using sheet: {sheet_name if sheet_name is not None else 'First sheet'}")
    print(f"Selected columns (from '{key_col}' to '{value_col}'): {selected_columns}")

    # Keep only the selected range of columns
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
            sec_df = pd.read_excel(file, sheet_name=sheet_name)

            # Keep only columns that exist in both selected_columns and this secondary file
            available_cols = [col for col in selected_columns if col in sec_df.columns]

            if key_col not in available_cols:
                print(f"    ⚠ Key column '{key_col}' missing. Skipping this file.")
                continue

            if not available_cols:
                print(f"    ⚠ No matching columns found. Skipping.")
                continue

            filtered = sec_df[available_cols].copy()

            # Add any missing columns (from the selected range) as blank
            for col in selected_columns:
                if col not in filtered.columns:
                    filtered[col] = pd.NA

            # Reorder to match primary order
            filtered = filtered[selected_columns]

            all_dfs.append(filtered)
            print(f"    ✓ Added {len(filtered)} rows | Columns used: {available_cols}")

        except ValueError as e:
            print(f"    ⚠ Sheet '{sheet_name}' not found in {file.name}. Skipping.")
        except Exception as e:
            print(f"    ✗ Error reading {file.name}: {e}")

    # ---------- 3. Merge ----------
    final_df = pd.concat(all_dfs, ignore_index=True)

    # Remove exact duplicate rows
    final_df = final_df.drop_duplicates(subset=selected_columns, keep="first")

    # Sort by the key column
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
        VALUE_COLUMN,
        SHEET_NAME
    )