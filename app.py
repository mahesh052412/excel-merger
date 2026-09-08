import streamlit as st
import pandas as pd
from io import BytesIO

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
        # First read all sheet names from primary file
        xl = pd.ExcelFile(primary_file)
        sheet_names = xl.sheet_names

        st.success(f"Primary file loaded successfully")
        
        # Sheet selection
        st.subheader("📄 Select Sheet")
        selected_sheet = st.selectbox(
            "Choose which sheet to use (will be applied to Primary + all Secondary files)",
            options=sheet_names,
            index=0
        )

        # Now read the selected sheet
        primary_df = pd.read_excel(primary_file, sheet_name=selected_sheet)
        all_columns = list(primary_df.columns)

        st.write(f"**Sheet selected:** `{selected_sheet}` → {len(primary_df)} rows")
        st.write(f"**Available columns:** {all_columns}")

        # Column selection
        st.subheader("3️⃣ Select Columns")
        col_a, col_b = st.columns(2)

        with col_a:
            key_col = st.selectbox("Key Column (starting column)", all_columns, index=0)
        with col_b:
            value_col = st.selectbox("Value Column (ending column)", all_columns, index=len(all_columns)-1)

        # Calculate selected columns range
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

                    # Primary data
                    primary_subset = primary_df[selected_columns].copy()
                    all_dfs = [primary_subset]

                    skipped_files = []

                    # Process secondary files
                    for file in secondary_files:
                        try:
                            # Try to read the same sheet name
                            sec_df = pd.read_excel(file, sheet_name=selected_sheet)
                            
                            available_cols = [c for c in selected_columns if c in sec_df.columns]

                            if key_col not in available_cols:
                                skipped_files.append(f"`{file.name}` → Key column '{key_col}' missing")
                                continue

                            if not available_cols:
                                skipped_files.append(f"`{file.name}` → No matching columns")
                                continue

                            filtered = sec_df[available_cols].copy()

                            # Add missing columns as blank
                            for col in selected_columns:
                                if col not in filtered.columns:
                                    filtered[col] = pd.NA

                            filtered = filtered[selected_columns]
                            all_dfs.append(filtered)

                        except ValueError:
                            skipped_files.append(f"`{file.name}` → Sheet '{selected_sheet}' not found")
                        except Exception as e:
                            skipped_files.append(f"`{file.name}` → Error: {e}")

                    # Show skipped files if any
                    if skipped_files:
                        st.warning("Some files were skipped:")
                        for msg in skipped_files:
                            st.write(f"- {msg}")

                    # Combine everything
                    final_df = pd.concat(all_dfs, ignore_index=True)
                    final_df = final_df.drop_duplicates(subset=selected_columns, keep="first")
                    final_df = final_df.sort_values(by=key_col).reset_index(drop=True)

                    st.success(f"✅ Merge complete! **{final_df.shape[0]} rows × {final_df.shape[1]} columns**")

                    # Preview
                    st.subheader("Preview of Merged Data")
                    st.dataframe(final_df, use_container_width=True)

                    # Download button
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
        st.error(f"Error reading Primary file: {e}")

else:
    st.info("👆 Please upload a Primary Excel file to begin.")