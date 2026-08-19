import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ---------------------------------------------------------
# 1. KONFIGURASI HALAMAN STREAMLIT
# ---------------------------------------------------------
st.set_page_config(
    page_title="Sistem Input Data & Ringkasan F4",
    layout="wide"
)

st.title("📋 Form Isian Data & Export Laporan")
st.markdown("---")

# Ukuran Kertas F4 dalam Point (215.9 mm x 330.2 mm)
F4_SIZE = (612, 936) 
F4_LANDSCAPE = landscape(F4_SIZE)

# ---------------------------------------------------------
# 2. UPLOAD FILE EXCEL SUMBER (ASLI)
# ---------------------------------------------------------
uploaded_file = st.sidebar.file_uploader(
    "Unggah File Excel Utama (.xlsx / .xlsb)", 
    type=["xlsx", "xlsb"]
)

if uploaded_file is not None:
    # Simpan byte file asli agar file tidak berubah/rusak sedikitpun saat didownload ulang
    original_bytes = uploaded_file.getvalue()

    st.subheader("1. Form Isian Data")
    st.info("💡 Isi data pada form di bawah ini. Semua kolom diambil lengkap dari sheet utama tanpa ada yang dikurangi.")

    # ---------------------------------------------------------
    # 3. BACA SHEET "ISIAN DATA" DENGAN PANDAS
    # ---------------------------------------------------------
    try:
        # Menampilkan sheet 'isian data' (atau sheet pertama jika nama sheet bervariasi)
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_target = "isian data" if "isian data" in [s.lower() for s in excel_file.sheet_names] else excel_file.sheet_names[0]
        
        df_source = pd.read_excel(uploaded_file, sheet_name=sheet_target)

        # Menampilkan Dataframe yang dapat diisi/diedit pengguna di web
        edited_df = st.data_editor(
            df_source,
            num_rows="dynamic",
            use_container_width=True,
            key="data_editor"
        )

        st.markdown("---")
        st.subheader("2. Panel Fitur & Download")

        col1, col2 = st.columns(2)

        # ---------------------------------------------------------
        # FITUR A: DOWNLOAD FILE EXCEL UTUH (ORIGINAL)
        # ---------------------------------------------------------
        with col1:
            st.markdown("### 📥 Download Excel")
            st.caption("Mendownload file Excel asli beserta seluruh sheet, struktur, dan rumus di dalamnya.")
            st.download_button(
                label="Download File Excel Utuh (.xlsx)",
                data=original_bytes,
                file_name="File_Utuh_Sistem.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        # ---------------------------------------------------------
        # FITUR B: GENERATE & DOWNLOAD PDF 1 LEMBAR F4
        # ---------------------------------------------------------
        def generate_pdf_f4(df):
            buffer = BytesIO()
            # Set margin pas agar muat tepat 1 lembar F4
            doc = SimpleDocTemplate(
                buffer, 
                pagesize=F4_LANDSCAPE,
                rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20
            )
            elements = []
            styles = getSampleStyleSheet()

            # Judul Laporan
            title_style = ParagraphStyle(
                'TitleStyle',
                parent=styles['Heading1'],
                fontSize=14,
                leading=16,
                alignment=1, # Center
                spaceAfter=10
            )
            elements.append(Paragraph("<b>RINGKASAN ISIAN DATA LAPORAN</b>", title_style))
            elements.append(Spacer(1, 5))

            # Konversi seluruh isi DataFrame ke format tabel ReportLab
            table_data = [list(df.columns)] # Header
            for idx, row in df.iterrows():
                row_data = [str(val) if pd.notna(val) else "" for val in row]
                table_data.append(row_data)

            # Style Tabel PDF
            pdf_table = Table(table_data, repeatRows=1)
            pdf_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ]))

            elements.append(pdf_table)
            doc.build(elements)
            buffer.seek(0)
            return buffer

        with col2:
            st.markdown("### 📄 Download PDF Ringkasan F4")
            st.caption("Mencetak seluruh kolom 'Isian Data' ke dalam PDF 1 Lembar Ukuran F4.")
            
            pdf_bytes = generate_pdf_f4(edited_df)
            
            st.download_button(
                label="Download Ringkasan PDF (F4)",
                data=pdf_bytes,
                file_name="Ringkasan_Isian_Data_F4.pdf",
                mime="application/pdf",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file Excel: {e}")

else:
    st.warning("👈 Silakan unggah file Excel Anda melalui menu di sebelah kiri untuk memulai.")
