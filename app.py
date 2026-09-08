import streamlit as st
import pandas as pd
from io import BytesIO
import tempfile
import subprocess
from pathlib import Path
import warnings

warnings.filterwarnings("ignore")

st.set_page_config(page_title="Excel Merger", page_icon="📊", layout="wide")

st.title("📊 Excel File Merger")
st.markdown(
    "Upload one **Primary** file and multiple **Secondary** files. "
    "Select a sheet + two columns — all columns between them will be merged."
)

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("How to use")
    st.info("""
1. Upload Primary Excel file  
2. Select the Sheet you want  
3. Upload Secondary Excel file(s)  
4. Select Key Column & Value Column  
5. Click **Merge Files**
""")
    st.markdown("---")
    st.caption("Supports `.xlsx`, `.xls`, `.xlsm` + auto conversion via LibreOffice")


def get_preferred_engine(filename: str) -> str:
    """Return preferred engine based on extension."""
    ext = Path(filename).suffix.lower()
    if ext == ".xls":
        return "xlrd"
    return "openpyxl"  # .xlsx, .xlsm, etc.


def convert_with_libreoffice(uploaded_file) -> BytesIO:
    """Convert any spreadsheet to .xlsx using LibreOffice."""
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = Path(tmpdir) / uploaded_file.name
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        try:
            result = subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to", "xlsx",
                    "--outdir", tmpdir,
                    str(input_path),
                ],
                capture_output=True,
                text=True,
                timeout=90,
            )
        except FileNotFoundError:
            raise Exception("LibreOffice (`soffice`) is not installed or not in PATH.")

        if result.returncode != 0:
            raise Exception(f"LibreOffice conversion failed:\n{result.stderr}")

        converted_files = list(Path(tmpdir).glob("*.xlsx"))
        if not converted_files:
            raise Exception("Conversion finished but no .xlsx file was created.")

        with open(converted_files[0], "rb") as f:
            return BytesIO(f.read())


def read_excel_safe(file, sheet_name=None):
    """
    Robust Excel reader with multiple fallbacks.
    Order:
      1. Preferred engine (based on extension)
      2. Alternate engine
      3. pandas automatic detection (engine=None)
      4. LibreOffice conversion (if available)
    """
    preferred = get_preferred_engine(getattr(file, "name", "file.xlsx"))
    engines_to_try = [preferred]

    # Add the other common engine
    if preferred == "openpyxl":
        engines_to_try.append("xlrd")
    else:
        engines_to_try.append("openpyxl")

    # Also try pandas auto-detection
    engines_to_try.append(None)

    last_error = None

    for engine in engines_to_try:
        try:
            if hasattr(file, "seek"):
                file.seek(0)

            if sheet_name is None:
                return pd.ExcelFile(file, engine=engine)
            else:
                return pd.read_excel(file, sheet_name=sheet_name, engine=engine)
        except Exception as e:
            last_error = e
            continue

    # Last resort → LibreOffice
    try:
        if hasattr(file, "seek"):
            file.seek(0)
        converted = convert_with_libreoffice(file)

        if sheet_name is None:
            return pd.ExcelFile(converted, engine="openpyxl")
        else:
            return pd.read_excel(converted, sheet_name=sheet_name, engine="openpyxl")
    except Exception as e:
        last_error = e

    # Final failure message with advice
    raise Exception(
        f"Could not read the file with any method.\n\n"
        f"Tried: openpyxl → xlrd → auto → LibreOffice\n"
        f"Last error: {last_error}\n\n"
        f"**What you can do:**\n"
        f"1. Open the file in Excel / Google Sheets and **Save As → .xlsx**\n"
        f"2. Install LibreOffice (`soffice`) so automatic conversion works\n"
        f"3. Make sure the file is not password-protected or corrupted"
    )


# ---------------- File Uploaders ----------------
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Primary File")
    primary_file = st.file_uploader(
        "Drag & drop or click to upload Primary Excel",
        type=["xlsx", "xls", "xlsm"],
        key="primary",
    )

with col2:
    st.subheader("2️⃣ Secondary Files")
    secondary_files = st.file_uploader(
        "Drag & drop or click to upload Secondary Excel(s)",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key="secondary",
    )

# ---------------- Main Logic ----------------
if primary_file is not None:
    try:
        with st.spinner("Reading / converting Primary file..."):
            xl = read_excel_safe(primary_file)
            sheet_names = xl.sheet_names

        st.success(f"Primary file loaded successfully (`{primary_file.name}`)")

        # Sheet selection
        st.subheader("📄 Select Sheet")
        selected_sheet = st.selectbox(
            "Choose which sheet to use (applied to Primary + all Secondary files)",
            options=sheet_names,
            index=0,
        )

        primary_df = read_excel_safe(primary_file, sheet_name=selected_sheet)
        all_columns = list(primary_df.columns)

        st.write(f"**Sheet selected:** `{selected_sheet}` → **{len(primary_df)}** rows")
        st.write(f"**Available columns:** {all_columns}")

        # Column selection
        st.subheader("3️⃣ Select Columns")
        col_a, col_b = st.columns(2)

        with col_a:
            key_col = st.selectbox("Key Column (starting column)", all_columns, index=0)
        with col_b:
            value_col = st.selectbox(
                "Value Column (ending column)",
                all_columns,
                index=len(all_columns) - 1 if all_columns else 0,
            )

        start_idx = all_columns.index(key_col)
        end_idx = all_columns.index(value_col)
        if start_idx > end_idx:
            start_idx, end_idx = end_idx, start_idx
        selected_columns = all_columns[start_idx : end_idx + 1]

        st.info(f"**Columns that will be merged:** `{selected_columns}`")

        # ---------------- Merge Button ----------------
        if st.button("🚀 Merge Files", type="primary", use_container_width=True):

            if not secondary_files:
                st.warning("Please upload at least one Secondary file.")
            else:
                with st.spinner("Merging files..."):
                    primary_subset = primary_df[selected_columns].copy()
                    all_dfs = [primary_subset]
                    skipped_files = []

                    for file in secondary_files:
                        try:
                            sec_df = read_excel_safe(file, sheet_name=selected_sheet)

                            available_cols = [
                                c for c in selected_columns if c in sec_df.columns
                            ]

                            if key_col not in available_cols:
                                skipped_files.append(
                                    f"`{file.name}` → Key column `{key_col}` missing"
                                )
                                continue

                            filtered = sec_df[available_cols].copy()

                            # Ensure every selected column exists
                            for col in selected_columns:
                                if col not in filtered.columns:
                                    filtered[col] = pd.NA

                            filtered = filtered[selected_columns]
                            all_dfs.append(filtered)

                        except Exception as e:
                            skipped_files.append(f"`{file.name}` → {e}")

                    if skipped_files:
                        st.warning("Some files were skipped:")
                        for msg in skipped_files:
                            st.write(f"- {msg}")

                    final_df = pd.concat(all_dfs, ignore_index=True)
                    final_df = final_df.drop_duplicates(
                        subset=selected_columns, keep="first"
                    )
                    final_df = final_df.sort_values(by=key_col).reset_index(drop=True)

                    st.success(
                        f"✅ Merge complete! "
                        f"**{final_df.shape[0]} rows × {final_df.shape[1]} columns**"
                    )

                    st.subheader("Preview of Merged Data")
                    st.dataframe(final_df, use_container_width=True)

                    # Prepare download
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        final_df.to_excel(writer, index=False, sheet_name="Merged")
                    buffer.seek(0)

                    st.download_button(
                        label="📥 Download Merged Excel",
                        data=buffer,
                        file_name="merged_result.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

    except Exception as e:
        st.error(f"Error reading Primary file:\n{e}")
        st.info(
            "Tips:\n"
            "- Open the file in Excel / Google Sheets and **Save As → .xlsx**\n"
            "- Install LibreOffice (`soffice`) for automatic conversion of difficult files\n"
            "- Make sure the file is not password-protected or corrupted"
        )

else:
    st.info("👆 Please upload a Primary Excel file to begin.")