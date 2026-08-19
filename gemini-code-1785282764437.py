import os
import pandas as pd
import streamlit as st
from io import BytesIO
from reportlab.lib.pagesizes import landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Form Input Data Catin F4", layout="wide")
st.title("📝 Formulir Input Data Catin (F4)")

FILE_NAME = "BERKAS_CATIN_F4.xlsb"

if not os.path.exists(FILE_NAME):
    st.error(f"File '{FILE_NAME}' tidak ditemukan! Pastikan file berada di folder yang sama dengan `app.py`.")
    st.stop()

# 1. Baca data dari sheet ISIAN DATA
@st.cache_data(ttl=1)
def load_data():
    excel_file = pd.ExcelFile(FILE_NAME, engine='pyxlsb')
    target_sheet = next((s for s in excel_file.sheet_names if s.strip().upper() == "ISIAN DATA"), excel_file.sheet_names[0])
    df = pd.read_excel(FILE_NAME, sheet_name=target_sheet, engine='pyxlsb')
    return df, target_sheet

df, target_sheet = load_data()

# Mengambil baris data pertama (jika ada) sebagai nilai bawaan
default_row = df.iloc[0] if len(df) > 0 else pd.Series([None]*len(df.columns), index=df.columns)

def get_val(col_name):
    val = default_row.get(col_name, "")
    return str(val) if pd.notna(val) else ""

# 2. BENTUK FORMULIR INPUT YANG SIMPEL & TERSTRUKTUR
with st.form("form_isian_catin"):
    st.subheader("📋 Masukkan / Perbarui Data")
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "📌 Data Akad & Register", 
        "👨 Catin Laki-Laki", 
        "👩 Catin Perempuan", 
        "👴 Data Orang Tua / Wali"
    ])

    inputs = {}

    with tab1:
        st.markdown("##### Informasi Pendaftaran & Akad Nikah")
        col1, col2 = st.columns(2)
        with col1:
            for col in df.columns[:5]:
                inputs[col] = st.text_input(col, value=get_val(col))
        with col2:
            for col in df.columns[5:10]:
                inputs[col] = st.text_input(col, value=get_val(col))

    with tab2:
        st.markdown("##### Data Calon Pengantin Laki-Laki")
        col1, col2 = st.columns(2)
        mid = 10 + (len(df.columns[10:30]) // 2)
        with col1:
            for col in df.columns[10:mid]:
                inputs[col] = st.text_input(col, value=get_val(col))
        with col2:
            for col in df.columns[mid:30]:
                inputs[col] = st.text_input(col, value=get_val(col))

    with tab3:
        st.markdown("##### Data Calon Pengantin Perempuan")
        col1, col2 = st.columns(2)
        mid_pr = 30 + (len(df.columns[30:50]) // 2)
        with col1:
            for col in df.columns[30:mid_pr]:
                inputs[col] = st.text_input(col, value=get_val(col))
        with col2:
            for col in df.columns[mid_pr:50]:
                inputs[col] = st.text_input(col, value=get_val(col))

    with tab4:
        st.markdown("##### Data Orang Tua & Wali")
        col1, col2 = st.columns(2)
        mid_ortu = 50 + (len(df.columns[50:]) // 2)
        with col1:
            for col in df.columns[50:mid_ortu]:
                inputs[col] = st.text_input(col, value=get_val(col))
        with col2:
            for col in df.columns[mid_ortu:]:
                inputs[col] = st.text_input(col, value=get_val(col))

    btn_submit = st.form_submit_button("💾 Simpan & Perbarui Data", use_container_width=True)

if btn_submit:
    # Buat DataFrame baru berdasarkan hasil input form
    updated_df = pd.DataFrame([inputs])
    st.session_state['updated_df'] = updated_df
    st.success("✅ Data berhasil diperbarui di memori sistem!")

# 3. PANEL DOWNLOAD & PRANALA HASIL
st.markdown("---")
st.subheader("📥 Download Hasil Input")

final_df = st.session_state.get('updated_df', df)

col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    # Simpan byte file asli agar rumus sheet lain tetap utuh
    with open(FILE_NAME, "rb") as f:
        file_bytes = f.read()
    st.download_button(
        label="📥 Download File Excel Utama (.xlsb)",
        data=file_bytes,
        file_name=FILE_NAME,
        mime="application/vnd.ms-excel.sheet.binary.macroenabled.12",
        use_container_width=True
    )

with col_dl2:
    def generate_pdf(dataframe):
        buffer = BytesIO()
        F4_SIZE = landscape((612, 936)) # Ukuran F4 Landscape
        doc = SimpleDocTemplate(buffer, pagesize=F4_SIZE, rightMargin=15, leftMargin=15, topMargin=15, bottomMargin=15)
        elements = []
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=12, alignment=1, spaceAfter=8)
        elements.append(Paragraph("<b>RINGKASAN ISIAN DATA CATIN (F4)</b>", title_style))
        
        table_data = [list(dataframe.columns)]
        for _, row in dataframe.iterrows():
            table_data.append([str(val) if pd.notna(val) else "" for val in row])
        
        pdf_table = Table(table_data, repeatRows=1)
        pdf_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elements.append(pdf_table)
        doc.build(elements)
        buffer.seek(0)
        return buffer

    pdf_bytes = generate_pdf(final_df)
    st.download_button(
        label="📄 Download Ringkasan PDF (Ukuran F4)",
        data=pdf_bytes,
        file_name="Ringkasan_Isian_Data_F4.pdf",
        mime="application/pdf",
        use_container_width=True
    )
