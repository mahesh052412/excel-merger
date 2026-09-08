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


def get_engine(filepath: str | Path) -> str | None:
    """
    Return the correct pandas engine based on file extension.
    - .xls  → xlrd
    - .xlsx / .xlsm / others → openpyxl
    """
    ext = Path(filepath).suffix.lower()
    if ext == ".xls":
        return "xlrd"
    # .xlsx, .xlsm, and any other modern Excel format
    return "openpyxl"


def select_sheet(filepath: str | Path) -> str:
    """Show available sheets and let user choose one"""
    engine = get_engine(filepath)
    try:
        xl = pd.ExcelFile(filepath, engine=engine)
    except Exception as e:
        # Fallback: try the other common engine
        fallback = "openpyxl" if engine == "xlrd" else "xlrd"
        print(f"  ⚠ Failed with engine '{engine}': {e}")
        print(f"  → Trying fallback engine '{fallback}'...")
        xl = pd.ExcelFile(filepath, engine=fallback)

    sheet_names = xl.sheet_names

    print("\nAvailable sheets in the file:")
    for i, name in enumerate(sheet_names):
        print(f"  [{i}] {name}")

    while True:
        choice = input("\nEnter sheet number or sheet name: ").strip()

        if choice.isdigit():
            idx = int(choice)
            if 0 <= idx < len(sheet_names):
                return sheet_names[idx]
            print("Invalid number. Try again.")
        elif choice in sheet_names:
            return choice
        else:
            print("Invalid sheet name. Try again.")


def read_excel_safe(filepath: str | Path, sheet_name: str | None = None) -> pd.DataFrame:
    """
    Read an Excel file with automatic engine selection + fallback.
    """
    engine = get_engine(filepath)
    try:
        return pd.read_excel(filepath, sheet_name=sheet_name, engine=engine)
    except Exception as e:
        fallback = "openpyxl" if engine == "xlrd" else "xlrd"
        print(f"  ⚠ Failed with '{engine}': {e}")
        print(f"  → Retrying with '{fallback}'...")
        return pd.read_excel(filepath, sheet_name=sheet_name, engine=fallback)


def merge_excel_files(primary_path: str, secondary_folder: str, output_path: str,
                      key_col: str, value_col: str):

    # ---------- 1. Load Primary & Select Sheet ----------
    if not os.path.exists(primary_path):
        raise FileNotFoundError(f"Primary file not found: {primary_path}")

    print(f"\nReading Primary file: {primary_path}")
    selected_sheet = select_sheet(primary_path)
    print(f"→ Selected sheet: '{selected_sheet}'")

    primary_df = read_excel_safe(primary_path, sheet_name=selected_sheet)
    all_columns = list(primary_df.columns)

    # Validate both columns exist
    for col in [key_col, value_col]:
        if col not in all_columns:
            raise ValueError(
                f"Column '{col}' not found in primary file (sheet: {selected_sheet}).\n"
                f"Available columns: {all_columns}"
            )

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

    # Case-insensitive collection of all Excel files
    excel_files = []
    for pattern in ("*.xls", "*.xlsx", "*.xlsm", "*.XLS", "*.XLSX", "*.XLSM"):
        excel_files.extend(secondary_path.glob(pattern))

    # Remove duplicates (can happen on case-insensitive filesystems)
    excel_files = list({f.resolve() for f in excel_files})

    if not excel_files:
        print("No secondary Excel files (.xls / .xlsx / .xlsm) found.")
    else:
        print(f"\nFound {len(excel_files)} secondary file(s):")

    for file in excel_files:
        print(f"  → Processing: {file.name}")
        try:
            sec_df = read_excel_safe(file, sheet_name=selected_sheet)

            available_cols = [col for col in selected_columns if col in sec_df.columns]

            if key_col not in available_cols:
                print(f"    ⚠ Key column '{key_col}' missing. Skipping this file.")
                continue

            if not available_cols:
                print(f"    ⚠ No matching columns found. Skipping.")
                continue

            filtered = sec_df[available_cols].copy()

            # Ensure all selected columns exist (fill missing ones with NA)
            for col in selected_columns:
                if col not in filtered.columns:
                    filtered[col] = pd.NA

            filtered = filtered[selected_columns]
            all_dfs.append(filtered)
            print(f"    ✓ Added {len(filtered)} rows | Columns used: {available_cols}")

        except ValueError as ve:
            # Usually means the sheet does not exist
            print(f"    ⚠ Sheet '{selected_sheet}' not found in {file.name}. Skipping. ({ve})")
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
    # Always save as modern .xlsx
    final_df.to_excel(output_path, index=False, engine="openpyxl")
    print(f"\n✅ Merged file saved to: {output_path}")


if __name__ == "__main__":
    merge_excel_files(
        PRIMARY_FILE,
        SECONDARY_FOLDER,
        OUTPUT_FILE,
        KEY_COLUMN,
        VALUE_COLUMN
    )