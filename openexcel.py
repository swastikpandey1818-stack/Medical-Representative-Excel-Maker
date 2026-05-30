import streamlit as st
import pandas as pd
from google import genai
import pypdf
import docx
import os
import json

EXCEL_FILE = "mr_consolidated_data.xlsx"
client = genai.Client(api_key="GEMINI")

st.set_page_config(page_title="MR Multi-File AI Data Merger", layout="wide")
st.title("📑 MR Multi-File AI Data Merger")
st.write("Upload files and add custom AI instructions to merge records into one master sheet.")

# 1. Sidebar for Text Input and Custom Recommendation
st.sidebar.header("🔧 Inputs & Settings")
whatsapp_text = st.sidebar.text_area("Paste raw text notes here:", height=150)

# The custom recommendation box you requested!
custom_recommendation = st.sidebar.text_input(
    "💡 Custom AI Instruction / Recommendation:", 
    placeholder="e.g., If quantity is missing, assume it is 5."
)

# 2. Main Page for File Uploads
st.header("📁 Upload Source Documents")
uploaded_files = st.file_uploader(
    "Drag and drop or browse up to 3 files simultaneously", 
    type=["pdf", "docx", "xlsx", "xls"], 
    accept_multiple_files=True
)

if len(uploaded_files) > 3:
    st.error("⚠️ Maximum 3 files allowed. Please remove the extra files.")
    uploaded_files = uploaded_files[:3]

# Helper function to read raw text
def extract_text_from_file(file):
    text = ""
    file_type = file.name.split(".")[-1].lower()
    if file_type == "pdf":
        pdf_reader = pypdf.PdfReader(file)
        for page in pdf_reader.pages:
            text += str(page.extract_text()) + "\n"
    elif file_type == "docx":
        doc = docx.Document(file)  # Updated to call docx.Document()
        for para in doc.paragraphs:
            text += para.text + "\n"

# 3. Process and Merge System
if st.button("🚀 Process and Group All Data ⚡"):
    all_extracted_records = []
    direct_excel_dfs = []

    combined_raw_text = whatsapp_text + "\n"
    for file in uploaded_files:
        file_type = file.name.split(".")[-1].lower()
        if file_type in ["pdf", "docx"]:
            combined_raw_text += f"\n--- Content from {file.name} ---\n"
            combined_raw_text += extract_text_from_file(file)
        elif file_type in ["xlsx", "xls"]:
            try:
                df = pd.read_excel(file)
                df.columns = [col.strip().title() for col in df.columns]
                direct_excel_dfs.append(df)
            except Exception as e:
                st.error(f"Could not read spreadsheet {file.name}: {e}")

    # Send data + custom instruction to Gemini
    if combined_raw_text.strip():
        with st.spinner("AI is analyzing files..."):
            
            # Here is the f-string prompt using your custom recommendation box input!
            prompt = f"""
            You are a data entry automation script for a medical representative.
            Extract all instances of Doctor visits, products pitched, and sample quantities into JSON format.
            The response must be a valid JSON array of objects with exactly these keys:
            "Doctor Name", "Product Pitched", "Quantity".
            Ensure "Quantity" is strictly an integer.
            
            CRITICAL EXTRA INSTRUCTION FROM THE USER:
            {custom_recommendation}
            
            Text Data to process:
            {combined_raw_text}
            
            Return ONLY raw valid JSON text code. No markdown boxes, no talk.
            """
            
            try:
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt,
                )
                clean_json = response.text.strip().replace("```json", "").replace("```", "")
                ai_data = json.loads(clean_json)
                all_extracted_records.extend(ai_data)
            except Exception as e:
                st.error("AI extraction failed on file text. Check formatting.")
                st.text(response.text)

    # Compile everything into a primary DataFrame
    final_new_df = pd.DataFrame(all_extracted_records)

    # Append direct spreadsheets data if present
    for df in direct_excel_dfs:
        rename_map = {}
        for col in df.columns:
            if "doc" in col.lower(): rename_map[col] = "Doctor Name"
            elif "prod" in col.lower() or "med" in col.lower(): rename_map[col] = "Product Pitched"
            elif "qty" in col.lower() or "quant" in col.lower() or "sam" in col.lower(): rename_map[col] = "Quantity"
        df = df.rename(columns=rename_map)
        valid_cols = [c for c in ["Doctor Name", "Product Pitched", "Quantity"] if c in df.columns]
        df = df[valid_cols]
        final_new_df = pd.concat([final_new_df, df], ignore_index=True)

    # Clean, standardise, and SUM quantities
    if not final_new_df.empty:
        final_new_df["Doctor Name"] = final_new_df["Doctor Name"].fillna("Unknown").astype(str).str.strip().str.title()
        final_new_df["Product Pitched"] = final_new_df["Product Pitched"].fillna("Unknown").astype(str).str.strip().str.title()
        final_new_df["Quantity"] = pd.to_numeric(final_new_df["Quantity"], errors='coerce').fillna(0).astype(int)

        consolidated_df = final_new_df.groupby(["Doctor Name", "Product Pitched"], as_index=False)["Quantity"].sum()

        if os.path.exists(EXCEL_FILE):
            existing_df = pd.read_excel(EXCEL_FILE)
            combined_master = pd.concat([existing_df, consolidated_df], ignore_index=True)
            combined_master = combined_master.groupby(["Doctor Name", "Product Pitched"], as_index=False)["Quantity"].sum()
        else:
            combined_master = consolidated_df

        combined_master.to_excel(EXCEL_FILE, index=False)
        st.success("Successfully processed all files and custom instructions!")
        st.dataframe(consolidated_df)
    else:
        st.warning("No data found to process.")

# 4. Download Section
if os.path.exists(EXCEL_FILE):
    st.markdown("---")
    st.subheader("📊 Master Spreadsheet")
    master_df = pd.read_excel(EXCEL_FILE)
    st.dataframe(master_df)
    
    with open(EXCEL_FILE, "rb") as file:
        st.download_button(
            label="📥 Download Consolidated Data for LibreOffice",
            data=file,
            file_name="mr_master_consolidated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )