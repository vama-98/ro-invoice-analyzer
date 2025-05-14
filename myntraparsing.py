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

            # === STEP 1: Build RO -> Invoice Mapping ===
            doc = fitz.open(pdf_path)
            lines = []
            for page in doc:
                lines.extend(page.get_text().split('\n'))

            invoice_map = {}  # line index -> invoice number
            for i, line in enumerate(lines):
                if line.strip().lower() == "invoice number":
                    if i + 1 < len(lines):
                        inv_value = lines[i + 1].strip()
                        if re.match(r"BHR2-\d{4}-[A-Z0-9:/\-]+", inv_value):
                            invoice_map[i] = inv_value

            entries = []
            ro_pattern = re.compile(r"RO\d{10,}", re.IGNORECASE)
            for i, line in enumerate(lines):
                if "Remarks" in line and "RO" in line:
                    ro_match = ro_pattern.search(line)
                    if ro_match:
                        ro_number = ro_match.group().strip()
                        invoice_number = None
                        for inv_line in sorted(invoice_map.keys(), reverse=True):
                            if inv_line < i:
                                invoice_number = invoice_map[inv_line]
                                break
                        entries.append({"RO Number": ro_number, "Invoice Number": invoice_number})

            ro_df = pd.DataFrame(entries)

            # === STEP 2: Enrich with SKU grouping and amount ===
            final_records = []

            for _, row in ro_df.iterrows():
                ro_number = row['RO Number']
                invoice_number = row['Invoice Number']
                filename = f"RO-{ro_number}.xlsx"
                filepath = os.path.join(extracted_folder, filename)

                if not os.path.exists(filepath):
                    st.warning(f"⚠️ File not found for {ro_number}: {filename}")
                    continue

                df = pd.read_excel(filepath, sheet_name='RO')

                # Fix misnamed columns: 'Reject Reason' actually holds Item Value, and the next unnamed column is real Reject Reason
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
                        "Invoice Number": invoice_number
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
