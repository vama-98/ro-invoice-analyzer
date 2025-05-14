import streamlit as st
import pandas as pd
import os
import fitz  # PyMuPDF
import re
import tempfile
import zipfile
import openpyxl

st.set_page_config(page_title="RO & Invoice Analyzer", layout="wide")
st.title("📦 Return Order & Invoice Analyzer")

st.markdown("""
This dashboard lets you:
1. Upload your RO summary PDF (with invoice mappings)
2. Upload a ZIP file containing all RO Excel files
3. Run the analysis and download enriched results
""")

# === Upload PDF ===
pdf_file = st.file_uploader("Upload HEVN PDF (RO Summary)", type="pdf")

# === Upload RO Excel ZIP ===
zip_file = st.file_uploader("Upload ZIP containing RO Excel files", type="zip")

run_button = st.button("🚀 Run Analysis")

if run_button:
    if not pdf_file or not zip_file:
        st.error("Please upload both the PDF and the ZIP file.")
    else:
        with st.spinner("Processing... Please wait."):
            temp_dir = tempfile.mkdtemp()

            # Save and read PDF
            pdf_path = os.path.join(temp_dir, "HEVN_temp.pdf")
            with open(pdf_path, "wb") as f:
                f.write(pdf_file.read())

            # Extract ZIP contents
            zip_path = os.path.join(temp_dir, "ro_excels.zip")
            with open(zip_path, "wb") as f:
                f.write(zip_file.read())

            extracted_folder = os.path.join(temp_dir, "extracted")
            os.makedirs(extracted_folder, exist_ok=True)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extracted_folder)

            output_csv_path = os.path.join(temp_dir, "ro_invoice_mapping_detailed.csv")

            # === STEP 1: Build RO -> Invoice + Debit Note Mapping ===
            doc = fitz.open(pdf_path)
            lines = []
            for page in doc:
                lines.extend(page.get_text().split('\n'))

            invoice_map = {}
            for i, line in enumerate(lines):
                if "invoice number" in line.lower() and i + 1 < len(lines):
                    invoice_number = lines[i + 1].strip()
                    debit_note = None

                    # Look up to 5 lines above for debit note reference
                    for j in range(i - 1, max(i - 6, -1), -1):
                        if "debit note reference" in lines[j].lower():
                            match = re.search(r"(\d{6,})", lines[j])
                            if match:
                                debit_note = match.group(1)
                                break
                            if j + 1 < len(lines):
                                match_next = re.search(r"(\d{6,})", lines[j + 1])
                                if match_next:
                                    debit_note = match_next.group(1)
                                    break

                    if invoice_number:
                        invoice_map[i] = {
                            "invoice_number": invoice_number,
                            "debit_note": debit_note
                        }

            entries = []
            ro_pattern = re.compile(r"RO\d{10,}", re.IGNORECASE)

            for i, line in enumerate(lines):
                if "remarks" in line.lower() and "RO" in line:
                    ro_match = ro_pattern.search(line)
                    if ro_match:
                        ro_number = ro_match.group().strip()
                        closest = max([key for key in invoice_map if key < i], default=None)
                        if closest is not None:
                            invoice_data = invoice_map[closest]
                            entries.append({
                                "RO Number": ro_number,
                                "Invoice Number": invoice_data['invoice_number'],
                                "Debit Note Reference": invoice_data['debit_note']
                            })

            ro_df = pd.DataFrame(entries)

            # === STEP 2: Enrich with SKU grouping and amount ===
            final_records = []

            all_files = [f for f in os.listdir(extracted_folder) if f.endswith('.xlsx')]

            for _, row in ro_df.iterrows():
                ro_number = row['RO Number']
                invoice_number = row['Invoice Number']
                debit_note = row['Debit Note Reference']

                matching_file = next((f for f in all_files if ro_number in f), None)
                if not matching_file:
                    st.warning(f"⚠️ File not found for {ro_number}")
                    continue

                filepath = os.path.join(extracted_folder, matching_file)
                df = pd.read_excel(filepath, sheet_name='RO')

                cols = df.columns.tolist()
                if "Reject Reason" in cols:
                    misnamed_index = cols.index("Reject Reason")
                    true_reason_index = misnamed_index + 1 if misnamed_index + 1 < len(cols) else None

                    cols[misnamed_index] = "Item Value"
                    if true_reason_index:
                        cols[true_reason_index] = "Reject Reason"

                    df.columns = cols

                df['Item Value'] = pd.to_numeric(df['Item Value'], errors='coerce')

                group_cols = ["PO Code", "Vendor Article Name", "Vendor Article Number", "GTIN",
                              "Size", "Colour", "Brand", "Reject Reason"]

                grouped = df.groupby(group_cols).agg(
                    Quantity=('Item Value', 'count'),
                    Total_Amount_With_Tax=('Item Value', 'sum')
                ).reset_index()

                grouped['Base_Value'] = grouped['Total_Amount_With_Tax'] / 1.18
                grouped['Tax_Value'] = grouped['Total_Amount_With_Tax'] - grouped['Base_Value']

                for _, data_row in grouped.iterrows():
                    enriched_row = {
                        "RO Number": ro_number,
                        "Invoice Number": invoice_number,
                        "Debit Note Reference": debit_note
                    }
                    enriched_row.update(data_row.to_dict())
                    final_records.append(enriched_row)

            final_df = pd.DataFrame(final_records)

            # Save CSV
            final_df.to_csv(output_csv_path, index=False)
            st.success("✅ Analysis complete!")

            # Display and offer download
            st.dataframe(final_df)
            csv = final_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                "📥 Download Enriched CSV",
                data=csv,
                file_name="ro_invoice_mapping_detailed.csv",
                mime="text/csv"
            )
