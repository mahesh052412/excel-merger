import streamlit as st
import pandas as pd
from io import BytesIO
import tempfile
import subprocess
from pathlib import Path
import warnings
import shutil

warnings.filterwarnings("ignore")

st.set_page_config(page_title="Excel Merger", page_icon="📊", layout="wide")

st.title("📊 Excel File Merger")
st.markdown(
    "Upload one **Primary** file and multiple **Secondary** files. "
    "Select a sheet + header row + two columns — all columns between them will be merged."
)

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("How to use")
    st.info("""
1. Upload Primary Excel file  
2. Select the Sheet  
3. Select which **Row** contains the column names  
4. Upload Secondary Excel file(s)  
5. Select Key Column & Value Column  
6. Click **Merge Files**
""")
    st.markdown("---")
    st.success("**Recommended:** Upload `.xlsx` files for best results.")
    st.caption("Supports `.xlsx`, `.xls`, `.xlsm` (powered by calamine)")


def is_libreoffice_available() -> bool:
    return shutil.which("soffice") is not None


def convert_with_libreoffice(uploaded_file) -> BytesIO:
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = Path(tmpdir) / uploaded_file.name
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

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

        if result.returncode != 0:
            raise Exception(f"LibreOffice conversion failed:\n{result.stderr}")

        converted_files = list(Path(tmpdir).glob("*.xlsx"))
        if not converted_files:
            raise Exception("Conversion finished but no .xlsx file was created.")

        with open(converted_files[0], "rb") as f:
            return BytesIO(f.read())


def read_excel_safe(file, sheet_name=None, header=0):
    """
    Robust Excel reader.
    Priority order:
      1. calamine   (best - supports xls + xlsx + xlsm + xlsb)
      2. openpyxl
      3. xlrd
      4. pandas auto
      5. LibreOffice (only if available)
    """
    engines_to_try = ["calamine", "openpyxl", "xlrd", None]
    last_error = None

    for engine in engines_to_try:
        try:
            if hasattr(file, "seek"):
                file.seek(0)

            if sheet_name is None:
                return pd.ExcelFile(file, engine=engine)
            else:
                return pd.read_excel(
                    file,
                    sheet_name=sheet_name,
                    header=header,
                    engine=engine
                )
        except Exception as e:
            last_error = e
            continue

    # Optional LibreOffice fallback (only works on your local machine)
    if is_libreoffice_available():
        try:
            if hasattr(file, "seek"):
                file.seek(0)
            converted = convert_with_libreoffice(file)

            if sheet_name is None:
                return pd.ExcelFile(converted, engine="calamine")
            else:
                return pd.read_excel(
                    converted,
                    sheet_name=sheet_name,
                    header=header,
                    engine="calamine"
                )
        except Exception as e:
            last_error = e

    # User-friendly final error
    raise Exception(
        "Could not read this Excel file.\n\n"
        "Possible reasons:\n"
        "• Old or unusual `.xls` format\n"
        "• File is corrupted or password-protected\n"
        "• Not a real Excel file\n\n"
        "**Solution (recommended):**\n"
        "1. Open the file in Excel or Google Sheets\n"
        "2. File → Save As → **Excel Workbook (.xlsx)**\n"
        "3. Upload the new `.xlsx` file"
    )


# ---------------- File Uploaders ----------------
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Primary File")
    primary_file = st.file_uploader(
        "Upload Primary Excel file",
        type=["xlsx", "xls", "xlsm"],
        key="primary",
    )

with col2:
    st.subheader("2️⃣ Secondary Files")
    secondary_files = st.file_uploader(
        "Upload Secondary Excel file(s)",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key="secondary",
    )

# ---------------- Main Logic ----------------
if primary_file is not None:
    try:
        with st.spinner("Reading Primary file..."):
            xl = read_excel_safe(primary_file)
            sheet_names = xl.sheet_names

        st.success(f"Primary file loaded: `{primary_file.name}`")

        # Sheet selection
        st.subheader("📄 Select Sheet")
        selected_sheet = st.selectbox(
            "Choose sheet (used for Primary + all Secondary files)",
            options=sheet_names,
            index=0,
        )

        # Header row selection
        st.subheader("🔢 Select Header Row")
        st.caption("Select the row that contains the column names (same row will be used in all files)")

        header_row_display = st.number_input(
            "Header Row Number (1 = first row)",
            min_value=1,
            value=1,
            step=1,
        )
        header_row = header_row_display - 1

        # Read primary with selected header
        primary_df = read_excel_safe(
            primary_file,
            sheet_name=selected_sheet,
            header=header_row
        )
        all_columns = list(primary_df.columns)

        st.write(f"**Sheet:** `{selected_sheet}` | **Header Row:** `{header_row_display}` | **Data rows:** {len(primary_df)}")
        st.write(f"**Columns found:** {all_columns}")

        # Column selection
        st.subheader("3️⃣ Select Columns")
        col_a, col_b = st.columns(2)

        with col_a:
            key_col = st.selectbox("Key Column (starting)", all_columns, index=0)
        with col_b:
            value_col = st.selectbox(
                "Value Column (ending)",
                all_columns,
                index=len(all_columns) - 1 if all_columns else 0,
            )

        start_idx = all_columns.index(key_col)
        end_idx = all_columns.index(value_col)
        if start_idx > end_idx:
            start_idx, end_idx = end_idx, start_idx
        selected_columns = all_columns[start_idx : end_idx + 1]

        st.info(f"**Columns to merge:** `{selected_columns}`")

        # Merge button
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
                            sec_df = read_excel_safe(
                                file,
                                sheet_name=selected_sheet,
                                header=header_row
                            )

                            available_cols = [c for c in selected_columns if c in sec_df.columns]

                            if key_col not in available_cols:
                                skipped_files.append(f"`{file.name}` → Key column missing")
                                continue

                            filtered = sec_df[available_cols].copy()

                            for col in selected_columns:
                                if col not in filtered.columns:
                                    filtered[col] = pd.NA

                            filtered = filtered[selected_columns]
                            all_dfs.append(filtered)

                        except Exception as e:
                            skipped_files.append(f"`{file.name}` → Could not read file")

                    if skipped_files:
                        st.warning("Some files were skipped:")
                        for msg in skipped_files:
                            st.write(f"- {msg}")

                    final_df = pd.concat(all_dfs, ignore_index=True)
                    final_df = final_df.drop_duplicates(subset=selected_columns, keep="first")
                    final_df = final_df.sort_values(by=key_col).reset_index(drop=True)

                    st.success(f"✅ Merge complete! **{final_df.shape[0]} rows × {final_df.shape[1]} columns**")

                    st.subheader("Preview")
                    st.dataframe(final_df, use_container_width=True)

                    # Download
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
        st.error("Failed to read the Primary file")
        st.markdown(str(e))
        st.info(
            "**What your friend should do:**\n\n"
            "1. Open the Excel file in Microsoft Excel or Google Sheets\n"
            "2. Click **File → Save As**\n"
            "3. Choose **Excel Workbook (*.xlsx)**\n"
            "4. Upload the new `.xlsx` file"
        )

else:
    st.info("👆 Please upload a Primary Excel file to begin.")