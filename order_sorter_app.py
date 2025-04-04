import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
import tempfile
import os

st.set_page_config(page_title="Order PDF Sorter", layout="centered")

st.title("📦 Order PDF Sorter")

st.markdown("Upload your **Excel file** (with `Order ID` column) and **PDF file** (invoices), and get a sorted PDF in return.")

# --- File Uploads ---
excel_file = st.file_uploader("📄 Upload Excel File", type=["xlsx"])
pdf_file = st.file_uploader("📑 Upload Merged PDF", type=["pdf"])

if excel_file and pdf_file:
    with tempfile.TemporaryDirectory() as tmpdir:
        # Save uploaded files
        excel_path = os.path.join(tmpdir, "uploaded.xlsx")
        pdf_path = os.path.join(tmpdir, "uploaded.pdf")

        with open(excel_path, "wb") as f:
            f.write(excel_file.read())
        with open(pdf_path, "wb") as f:
            f.write(pdf_file.read())

        # --- Step 1: Read Excel ---
        df = pd.read_excel(excel_path)
        df.columns = df.columns.str.strip().str.lower()  # Normalize column names

        if 'order id' not in df.columns:
            st.error("❌ Excel file must contain a column named 'Order ID'")
            st.stop()

        order_ids = df['order id'].dropna().drop_duplicates().astype(str).tolist()

        # --- Step 2: Search in PDF ---
        doc = fitz.open(pdf_path)
        order_to_page = {}
        for page_num in range(len(doc)):
            text = doc[page_num].get_text()
            for order_id in order_ids:
                if order_id in text and order_id not in order_to_page:
                    order_to_page[order_id] = page_num

        # --- Step 3: Sort & Create Output PDF ---
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        missing_ids = []

        for order_id in order_ids:
            if order_id in order_to_page:
                page_num = order_to_page[order_id]
                writer.add_page(reader.pages[page_num])
            else:
                missing_ids.append(order_id)

        output_path = os.path.join(tmpdir, "sorted_by_order_id.pdf")
        with open(output_path, "wb") as f:
            writer.write(f)

        # --- Step 4: Download + Summary ---
        st.success("✅ Sorted PDF created!")

        with open(output_path, "rb") as f:
            st.download_button("⬇️ Download Sorted PDF", f, file_name="sorted_by_order_id.pdf")

        if missing_ids:
            st.warning(f"⚠️ {len(missing_ids)} Order IDs not found in the PDF:")
            st.write(missing_ids)
        else:
            st.balloons()
            st.info("🎉 All Order IDs were successfully matched and sorted!")
