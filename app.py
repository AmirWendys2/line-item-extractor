import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import os
import re
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\\Users\\ARASUC2\\Downloads\\tesseract.exe"

# Path to master Excel file (adjust to your local path)
MASTER_FILE_PATH = "C:\\Users\\ARASUC2\\OneDrive - Wendy’s Portal\\Restaurant and Digital Technology - Amir's Acrelec Master File\\LineItemMaster.xlsx"


# --------------------- Parsers ---------------------
def extract_quotation_date(lines):
    for line in lines:
        match = re.search(r"Date:\s+(\d{4}-\d{2}-\d{2})", line)
        if match:
            return match.group(1)
    return ""
def parse_acrelec_pdf(file):
    import fitz
    import pandas as pd
    import re

    def extract_quotation_date(lines):
        for line in lines:
            match = re.search(r"Date:\s+(\d{4}-\d{2}-\d{2})", line)
            if match:
                return match.group(1)
        return ""

    file.seek(0)
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

def parse_pdg_pdf(file, mode="Strict"):
    import fitz
    import pandas as pd
    import re

    def is_float(val):
        try:
            float(val.replace(",", ""))
            return True
        except:
            return False

    def clean_float(val):
        return float(val.replace(",", ""))

    file.seek(0)
    doc = fitz.open(stream=file.read(), filetype="pdf")
    lines = []
    for page in doc:
        lines.extend(page.get_text().split("\n"))

    # Get quotation date
    quote_date = ""
    for line in lines:
        match = re.search(r"\d{1,2}/\d{1,2}/\d{4}", line)
        if match:
            quote_date = match.group(0)
            break

    items = []
    i = 0
    while i < len(lines) - 6:
        block = lines[i:i + 10]
        try:
            qty_line = block[0].strip()
            price_line = block[1].strip()
            um_line = block[2].strip()
            desc_line = block[3].strip()
            ext_line = block[4].strip()
            item_line = block[5].strip()
            part_line = re.sub(r"[ \t]+[ITE]?$", "", block[6].strip())

            # STRICT MODE
            if mode == "Strict":
                if not (is_float(qty_line) and is_float(price_line) and is_float(ext_line)):
                    i += 1
                    continue
                if um_line not in ["EA", "SET", "FT", "HR"]:
                    i += 1
                    continue
                if not re.fullmatch(r"\d{3}", item_line):
                    i += 1
                    continue
                if not re.match(r"^[A-Z0-9\-]{4,}$", part_line):
                    i += 1
                    continue

            # Try to add the line item
            items.append({
                "Quantity": clean_float(qty_line) if is_float(qty_line) else None,
                "Unit Price": clean_float(price_line) if is_float(price_line) else None,
                "UM": um_line if um_line in ["EA", "SET", "FT", "HR"] else "",
                "Description": desc_line,
                "Extension": clean_float(ext_line) if is_float(ext_line) else None,
                "Item No": item_line if re.fullmatch(r"\d{3}", item_line) else "",
                "Part No": part_line,
                "Quotation Date": quote_date
            })
            i += 8
        except:
            i += 1

    if not items:
        st.warning(f"⚠️ No line items found in {mode} mode.")
    else:
        st.success(f"✅ {len(items)} line items extracted using {mode} mode.")
    return pd.DataFrame(items)

# --------------------- Streamlit App ---------------------
st.title("📄 Multi-PDF Line Item Extractor")

parser_type = st.radio("Select PDF Type", ["Acrelec", "PDG"], horizontal=True)
st.sidebar.info("💬 Any questions or concerns contact Amir Rasul")

# Session state management
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

if uploaded_files:
    st.session_state.uploaded_files = uploaded_files

append_to_master = st.checkbox("✅ Append to master file", value=st.session_state.get("append_to_master", False))
st.session_state.append_to_master = append_to_master


if st.session_state.uploaded_files:
    session_dfs = []

    for uploaded_file in st.session_state.uploaded_files:
        st.write(f"Processing: {uploaded_file.name}")
        try:
            df = parse_acrelec_pdf(uploaded_file) if parser_type == "Acrelec" else parse_pdg_pdf(uploaded_file)
            df["Source File"] = uploaded_file.name
            session_dfs.append(df)
            st.success(f"Extracted {len(df)} line items from {uploaded_file.name}")
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {e}")

    if session_dfs:
        final_df = pd.concat(session_dfs, ignore_index=True)
        st.session_state.final_df = final_df

        # Display top-level download buttons
        col1, col2 = st.columns(2)
        with col1:
            if os.path.exists(MASTER_FILE_PATH):
                master_df = pd.read_excel(MASTER_FILE_PATH, sheet_name=None)
                output_master = BytesIO()
                with pd.ExcelWriter(output_master, engine="openpyxl") as writer:
                    for sheet, data in master_df.items():
                        data.to_excel(writer, sheet_name=sheet, index=False)
                st.download_button("⬇️ Download Master File", output_master.getvalue(), file_name="LineItemMaster.xlsx", key="master_top")
            else:
                st.warning("⚠️ Master file not found.")
        with col2:
            output_session = BytesIO()
            final_df.to_excel(output_session, index=False)
            st.download_button("📥 Download This Session Report", output_session.getvalue(), file_name="session_line_items.xlsx", key="session_top")

        # Show summary
        summary = final_df.groupby("Source File").size().reset_index(name="Line Items Extracted")
        st.markdown("### 📊 Summary by PDF File")
        st.dataframe(summary)

        if append_to_master:
            try:
                os.makedirs(os.path.dirname(MASTER_FILE_PATH), exist_ok=True)
                if parser_type == "PDG":
                    with pd.ExcelWriter(MASTER_FILE_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                        final_df.to_excel(writer, sheet_name="PDG", index=False)
                    st.success(f"✅ PDG data written to 'PDG' sheet in: {MASTER_FILE_PATH}")
                else:  # Acrelec
                    if os.path.exists(MASTER_FILE_PATH):
                        existing = pd.read_excel(MASTER_FILE_PATH, sheet_name="Acrelec")
                        before_count = len(existing)
                        combined = pd.concat([existing, final_df], ignore_index=True)
                        combined.drop_duplicates(
                            subset=["Item Number", "Item Description", "Category", "Quantity", "Net Price", "Amount", "Source File"],
                            inplace=True
                        )
                        after_count = len(combined)
                    else:
                        combined = final_df
                        before_count = 0
                        after_count = len(combined)

                    with pd.ExcelWriter(MASTER_FILE_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                        combined.to_excel(writer, sheet_name="Acrelec", index=False)

                    st.success(f"✅ Acrelec sheet updated — {after_count - before_count} new records added.")
            except Exception as e:
                st.error(f"❌ Failed to update master file: {e}")
