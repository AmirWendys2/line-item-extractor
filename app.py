import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import os
import re

# Path to master Excel file (adjust to your local path)
# Optional: Replace this with your actual SharePoint sync folder path
MASTER_FILE_PATH = os.path.join(os.getcwd(), "LineItemMaster.xlsx")

def extract_quotation_date(lines):
    for line in lines:
        match = re.search(r"Date:\s+(\d{4}-\d{2}-\d{2})", line)
        if match:
            return match.group(1)
    return ""

def parse_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    lines = []
    for page in doc:
        lines.extend(page.get_text().split("\n"))

    quotation_date = extract_quotation_date(lines)

    items = []
    i = 0
    while i < len(lines) - 5:
        line = lines[i].strip()
        next_lines = [lines[i + j].strip() for j in range(1, 6)]

        try:
            quantity = int(next_lines[2].replace(",", ""))
            net_price = float(next_lines[3].replace(",", ""))
            amount = float(next_lines[4].replace(",", ""))

            item = {
                "Item Number": line,
                "Item Description": next_lines[0],
                "Category": next_lines[1],
                "Quantity": quantity,
                "Net Price": net_price,
                "Amount": amount,
                "Quotation Date": quotation_date
            }
            items.append(item)
            i += 6
        except:
            i += 1

    return pd.DataFrame(items)

st.title("📄 Multi-PDF Line Item Extractor")

# Session state for uploaded files and clearing
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = []

if st.button("🧹 Clear All PDFs"):
    st.session_state.uploaded_files = []
    st.session_state.clear_output = True
    st.session_state.append_to_master = False
    st.session_state.reset_uploader = not st.session_state.get("reset_uploader", False)
    st.success("PDF uploads and results cleared. Please re-upload your files if needed.")

uploaded_files = st.file_uploader(
    "Upload one or more PDF quotes",
    type="pdf",
    accept_multiple_files=True,
    key="file_uploader" if not st.session_state.get("reset_uploader") else "file_uploader_reset"
)

# Replace session state with current upload if new ones provided
if uploaded_files:
    st.session_state.uploaded_files = uploaded_files

append_to_master = st.checkbox("✅ Append to master file on desktop", value=st.session_state.get("append_to_master", False))
st.session_state.append_to_master = append_to_master

# Button to download current master file
if st.button("📂 Download Current Master File"):
    if os.path.exists(MASTER_FILE_PATH):
        master_df = pd.read_excel(MASTER_FILE_PATH)
        output = BytesIO()
        master_df.to_excel(output, index=False)
        st.download_button("⬇️ Click to Download Master File", output.getvalue(), file_name="LineItemMaster.xlsx")
    else:
        st.warning("⚠️ Master file not found at the expected location.")

if st.session_state.uploaded_files:
    session_dfs = []

    for uploaded_file in st.session_state.uploaded_files:
        st.write(f"Processing: {uploaded_file.name}")
        try:
            df = parse_pdf(uploaded_file)
            df["Source File"] = uploaded_file.name
            session_dfs.append(df)
            st.success(f"Extracted {len(df)} line items from {uploaded_file.name}")
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {e}")

    if session_dfs:
        final_df = pd.concat(session_dfs, ignore_index=True)
        if not st.session_state.get("clear_output"):
            st.dataframe(final_df)
        st.session_state.clear_output = False

        # Download session report
        output = BytesIO()
        final_df.to_excel(output, index=False)
        st.download_button("📥 Download This Session Report", output.getvalue(), file_name="session_line_items.xlsx")

        # Append to master file if selected
        if append_to_master:
            try:
                if os.path.exists(MASTER_FILE_PATH):
                    master_df = pd.read_excel(MASTER_FILE_PATH)
                    before_count = len(master_df)
                    combined_df = pd.concat([master_df, final_df], ignore_index=True)
                    combined_df.drop_duplicates(
                        subset=["Item Number", "Item Description", "Category", "Quantity", "Net Price", "Amount", "Source File"],
                        inplace=True
                    )
                    after_count = len(combined_df)
                else:
                    combined_df = final_df
                    before_count = 0
                    after_count = len(combined_df)

                combined_df.to_excel(MASTER_FILE_PATH, index=False)
                new_records = after_count - before_count
                st.success(f"✅ Master file updated at: {MASTER_FILE_PATH} — Added {new_records} new records.")
            except Exception as e:
                st.error(f"❌ Failed to update master file: {e}")
    else:
        st.warning("⚠️ No line items extracted from uploaded files.")
