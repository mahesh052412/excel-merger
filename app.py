import streamlit as st
import pandas as pd
from io import BytesIO
import tempfile
import subprocess
import os
from pathlib import Path
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="Excel Merger", page_icon="📊", layout="wide")

st.title("📊 Excel File Merger")
st.markdown("Upload one **Primary** file and multiple **Secondary** files. Select sheet + two columns — all columns between them will be merged.")

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
    st.caption("Supports .xlsx, .xls + auto conversion via LibreOffice")


def convert_with_libreoffice(uploaded_file):
    """
    Convert any spreadsheet to .xlsx using LibreOffice (most reliable method)
    Returns path to the converted .xlsx file
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        # Save uploaded file
        input_path = Path(tmpdir) / uploaded_file.name
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Run LibreOffice conversion
        try:
            result = subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to", "xlsx",
                    "--outdir", tmpdir,
                    str(input_path)
                ],
                capture_output=True,
                text=True,
                timeout=60
            )
        except FileNotFoundError:
            raise Exception("LibreOffice (soffice) is not installed or not in PATH.")

        if result.returncode != 0:
            raise Exception(f"LibreOffice conversion failed:\n{result.stderr}")

        # Find the converted file
        converted_files = list(Path(tmpdir).glob("*.xlsx"))
        if not converted_files:
            raise Exception("Conversion finished but no .xlsx file was created.")

        converted_path = converted_files[0]

        # Read into memory so we can return a file-like object
        with open(converted_path, "rb") as f:
            data = f.read()

        return BytesIO(data)


def read_excel_safe(file, sheet_name=None):
    """
    Try normal engines first, then LibreOffice conversion as last resort.
    """
    # Reset pointer
    if hasattr(file, "seek"):
        file.seek(0)

    engines = ["openpyxl", "xlrd"]
    last_error = None

    # 1. Try normal engines
    for engine in engines:
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

    # 2. Try LibreOffice conversion
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

    raise Exception(f"Could not read the file with any method.\nLast error: {last_error}")


# ---------------- File Uploaders ----------------
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Primary File")
    primary_file = st.file_uploader(
        "Drag & drop or click to upload Primary Excel",
        type=["xlsx", "xls"],
        key="primary"
    )

with col2:
    st.subheader("2️⃣ Secondary Files")
    secondary_files = st.file_uploader(
        "Drag & drop or click to upload Secondary Excel(s)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="secondary"
    )

# ---------------- Main Logic ----------------
if primary_file is not None:
    try:
        with st.spinner("Reading / converting file..."):
            xl = read_excel_safe(primary_file)
            sheet_names = xl.sheet_names

        st.success(f"Primary file loaded successfully (`{primary_file.name}`)")

        # Sheet selection
        st.subheader("📄 Select Sheet")
        selected_sheet = st.selectbox(
            "Choose which sheet to use (will be applied to Primary + all Secondary files)",
            options=sheet_names,
            index=0
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
                index=len(all_columns) - 1 if len(all_columns) > 0 else 0
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
                            skipped_files.append(f"`{file.name}` → {e}")

                    if skipped_files:
                        st.warning("Some files were skipped:")
                        for msg in skipped_files:
                            st.write(f"- {msg}")

                    final_df = pd.concat(all_dfs, ignore_index=True)
                    final_df = final_df.drop_duplicates(subset=selected_columns, keep="first")
                    final_df = final_df.sort_values(by=key_col).reset_index(drop=True)

                    st.success(f"✅ Merge complete! **{final_df.shape[0]} rows × {final_df.shape[1]} columns**")

                    st.subheader("Preview of Merged Data")
                    st.dataframe(final_df, use_container_width=True)

                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        final_df.to_excel(writer, index=False, sheet_name="Merged")
                    buffer.seek(0)

                    st.download_button(
                        label="📥 Download Merged Excel",
                        data=buffer,
                        file_name="merged_result.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

    except Exception as e:
        st.error(f"Error reading Primary file:\n{e}")
        st.info("Make sure LibreOffice is installed. You can also convert the file manually to .xlsx.")

else:
    st.info("👆 Please upload a Primary Excel file to begin.")