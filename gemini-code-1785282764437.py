import os
import pandas as pd
import streamlit as st
from io import BytesIO
from reportlab.lib.pagesizes import landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# 1. Konfigurasi Halaman Streamlit
st.set_page_config(page_title="Sistem Isian Data & PDF F4", layout="wide")
st.title("📋 Aplikasi Isian Data & Laporan")

# Nama file Excel bawaan di folder
FILE_NAME = "BERKAS_CATIN_F4.xlsb"

# 2. Cek Keberadaan File
if os.path.exists(FILE_NAME):
    # Baca aliran byte asli agar tidak merusak rumus Excel saat didownload
    with open(FILE_NAME, "rb") as f:
        original_bytes = f.read()

    try:
        # Buka file Excel dengan engine pyxlsb
        excel_file = pd.ExcelFile(FILE_NAME, engine='pyxlsb')
        sheet_names = excel_file.sheet_names
        
        # Pencarian nama sheet secara presisi (Abaikan huruf besar/kecil)
        target_sheet = None
        for sheet in sheet_names:
            if sheet.strip().upper() == "ISIAN DATA":
                target_sheet = sheet
                break
        
        # Jika tidak ditemukan, gunakan sheet pertama sebagai acuan
        if not target_sheet:
            target_sheet = sheet_names[0]

        # Baca data dari sheet target tanpa mengurangi/mengubah kolom apapun
        df = pd.read_excel(FILE_NAME, sheet_name=target_sheet, engine='pyxlsb')

        st.subheader(f"1. Isian Data (Sheet: {target_sheet})")
        st.caption("Semua kolom ditampilkan lengkap. Anda bisa langsung mengisi/mengedit data di bawah ini:")

        # Editor Tabel Interaktif Streamlit
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        st.markdown("---")
        st.subheader("2. Panel Download File")

        col1, col2 = st.columns(2)

        # --- TOMBOL 1: DOWNLOAD EXCEL UTUH BERISI RUMUS ---
        with col1:
            st.download_button(
                label="📥 Download File Excel Utuh (.xlsb)",
                data=original_bytes,
                file_name=FILE_NAME,
                mime="application/vnd.ms-excel.sheet.binary.macroenabled.12",
                use_container_width=True
            )

        # --- TOMBOL 2: DOWNLOAD RINGKASAN PDF F4 1 LEMBAR ---
        def generate_pdf(dataframe):
            buffer = BytesIO()
            # Ukuran Kertas F4 Landscape (215.9 mm x 330.2 mm)
            F4_SIZE = landscape((612, 936))
            doc = SimpleDocTemplate(
                buffer, 
                pagesize=F4_SIZE, 
                rightMargin=15, 
                leftMargin=15, 
                topMargin=15, 
                bottomMargin=15
            )
            
            elements = []
            styles = getSampleStyleSheet()
            
            title_style = ParagraphStyle(
                'TitleStyle', 
                parent=styles['Heading1'], 
                fontSize=12, 
                alignment=1, 
                spaceAfter=8
            )
            elements.append(Paragraph("<b>RINGKASAN ISIAN DATA</b>", title_style))
            
            # Susun data tabel untuk PDF
            table_data = [list(dataframe.columns)]
            for _, row in dataframe.iterrows():
                table_data.append([str(val) if pd.notna(val) else "" for val in row])
            
            # Format Style Tabel PDF agar rapi
            pdf_table = Table(table_data, repeatRows=1)
            pdf_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTSIZE', (0, 0), (-1, -1), 7),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ]))
            
            elements.append(pdf_table)
            doc.build(elements)
            buffer.seek(0)
            return buffer

        with col2:
            pdf_bytes = generate_pdf(edited_df)
            st.download_button(
                label="📄 Download Ringkasan PDF (Ukuran F4)",
                data=pdf_bytes,
                file_name="Ringkasan_Isian_Data_F4.pdf",
                mime="application/pdf",
                use_container_width=True
            )

    except Exception as err:
        st.error(f"Gagal memproses file Excel: {err}")

else:
    st.error(f"File '{FILE_NAME}' tidak ditemukan di folder proyek. Pastikan file disimpan dalam satu folder dengan app.py.")
