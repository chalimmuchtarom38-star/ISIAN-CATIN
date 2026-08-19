import os
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from io import BytesIO
from reportlab.lib.pagesizes import landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Form Input Catin F4", layout="wide")
st.title("📝 Formulir Input Isian Data Catin (F4)")

FILE_NAME = "BERKAS_CATIN_F4.xlsb"

if not os.path.exists(FILE_NAME):
    st.error(f"File '{FILE_NAME}' tidak ditemukan di folder aplikasi! Pastikan file berada di folder yang sama dengan app.py.")
    st.stop()

# Fungsi Konversi Otomatis Angka Serial Excel -> Format Tanggal Teks (DD/MM/YYYY)
def convert_excel_date(val):
    if isinstance(val, (int, float)):
        try:
            date_obj = datetime(1899, 12, 30) + timedelta(days=int(val))
            return date_obj.strftime("%d/%m/%Y")
        except:
            return str(val)
    elif isinstance(val, str) and val.strip().isdigit():
        try:
            date_obj = datetime(1899, 12, 30) + timedelta(days=int(val.strip()))
            return date_obj.strftime("%d/%m/%Y")
        except:
            return val
    return str(val) if pd.notna(val) else ""

# Fungsi Membaca Data dari Sheet ISIAN DATA (Kolom B = Label, Kolom G/F = Isian Data)
@st.cache_data(ttl=1)
def load_excel_data():
    try:
        excel_file = pd.ExcelFile(FILE_NAME, engine='pyxlsb')
        sheet_name = next((s for s in excel_file.sheet_names if s.strip().upper() == "ISIAN DATA"), excel_file.sheet_names[0])
        df = pd.read_excel(FILE_NAME, sheet_name=sheet_name, engine='pyxlsb', header=None)
        
        mapping_rows = [
            ("Baris 2", "Nomor Register", 1),
            ("Baris 3", "Tanggal Surat", 2),
            ("Baris 4", "Tanggal Pelaksanaan", 3),
            ("Baris 5", "Jam Pelaksanaan", 4),
            ("Baris 6", "Tempat Akad Nikah", 5),
            
            # Catin Laki-Laki
            ("Baris 8", "Nama Catin Laki-Laki", 7),
            ("Baris 9", "Bin (Ayah Laki-Laki)", 8),
            ("Baris 10", "TTL Catin Laki-Laki", 9),
            ("Baris 11", "NIK Catin Laki-Laki", 10),
            ("Baris 12", "Pekerjaan Laki-Laki", 11),
            ("Baris 13", "Status Laki-Laki", 12),
            ("Baris 14", "Jenis Kelamin Laki-Laki", 13),
            ("Baris 15", "Nama Istri Terdahulu", 14),
            ("Baris 16", "Alamat Catin Laki-Laki", 15),
            ("Baris 17", "Pendidikan Laki-Laki", 16),
            
            # Ortu Laki-Laki
            ("Baris 19", "Nama Ayah Laki-Laki", 18),
            ("Baris 20", "NIK Ayah Laki-Laki", 19),
            ("Baris 21", "TTL Ayah Laki-Laki", 20),
            ("Baris 22", "Pekerjaan Ayah Laki-Laki", 21),
            ("Baris 23", "Alamat Ayah Laki-Laki", 22),
            ("Baris 26", "Nama Ibu Laki-Laki", 25),
            ("Baris 27", "NIK Ibu Laki-Laki", 26),
            ("Baris 28", "TTL Ibu Laki-Laki", 27),
            ("Baris 29", "Pekerjaan Ibu Laki-Laki", 28),
            ("Baris 30", "Alamat Ibu Laki-Laki", 29),
            
            # Catin Perempuan
            ("Baris 34", "Nama Catin Perempuan", 33),
            ("Baris 35", "Binti (Ayah Perempuan)", 34),
            ("Baris 36", "TTL Catin Perempuan", 35),
            ("Baris 37", "NIK Catin Perempuan", 36),
            ("Baris 38", "Pekerjaan Perempuan", 37),
            ("Baris 39", "Status Perempuan", 38),
            ("Baris 40", "Jenis Kelamin Perempuan", 39),
            ("Baris 41", "Alamat Catin Perempuan", 40),
            ("Baris 42", "Nama Suami Terdahulu", 41),
            ("Baris 43", "Pendidikan Perempuan", 42),
            
            # Ortu Perempuan
            ("Baris 45", "Nama Ayah Perempuan", 44),
            ("Baris 46", "NIK Ayah Perempuan", 45),
            ("Baris 47", "TTL Ayah Perempuan", 46),
            ("Baris 48", "Pekerjaan Ayah Perempuan", 47),
            ("Baris 49", "Alamat Ayah Perempuan", 48),
            ("Baris 52", "Nama Ibu Perempuan", 51),
            ("Baris 53", "NIK Ibu Perempuan", 52),
            ("Baris 54", "TTL Ibu Perempuan", 53),
            ("Baris 55", "Pekerjaan Ibu Perempuan", 54),
            ("Baris 56", "Alamat Ibu Perempuan", 55),
            
            # Wali & Saksi
            ("Baris 58", "Nama Wali", 57),
            ("Baris 59", "Bin Wali", 58),
            ("Baris 60", "NIK Wali", 59),
            ("Baris 61", "TTL Wali", 60),
            ("Baris 62", "Pekerjaan Wali", 61),
            ("Baris 63", "Alamat Wali", 62),
            ("Baris 64", "Hubungan Wali", 63),
            ("Baris 65", "Mahar / Maskawin", 64),
            ("Baris 70", "Nama Saksi 1", 69),
            ("Baris 71", "TTL Saksi 1", 70),
            ("Baris 72", "NIK Saksi 1", 71),
            ("Baris 73", "Pekerjaan Saksi 1", 72),
            ("Baris 74", "Alamat Saksi 1", 73),
            ("Baris 76", "Nama Saksi 2", 75),
            ("Baris 77", "TTL Saksi 2", 76),
            ("Baris 78", "NIK Saksi 2", 77),
            ("Baris 79", "Pekerjaan Saksi 2", 78),
            ("Baris 80", "Alamat Saksi 2", 79)
        ]
        
        extracted_data = []
        for ref, label, r_idx in mapping_rows:
            val = ""
            if r_idx < len(df):
                raw_val = df.iloc[r_idx, 6] if pd.notna(df.iloc[r_idx, 6]) else df.iloc[r_idx, 5]
                val = convert_excel_date(raw_val)
                if val.strip() == ":":
                    val = ""
            extracted_data.append((ref, label, val))
            
        return extracted_data
    except Exception as e:
        st.error(f"Gagal membaca Excel: {e}")
        return []

data_list = load_excel_data()

# FORM INPUT INTERAKTIF STREAMLIT
with st.form("form_catin"):
    st.subheader("Isi / Sesuaikan Data Catin Secara Praktis")
    
    t1, t2, t3, t4 = st.tabs([
        "📄 Register & Surat", 
        "👨 Catin Laki-Laki", 
        "👩 Catin Perempuan", 
        "🤝 Wali & Saksi"
    ])
    
    user_inputs = {}

    with t1:
        st.markdown("##### Data Surat & Pelaksanaan Akad")
        for ref, label, val in data_list[:5]:
            user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val)

    with t2:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### -- Data Catin Laki-Laki --")
            for ref, label, val in data_list[5:15]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val)
        with col2:
            st.markdown("##### -- Data Orang Tua Laki-Laki --")
            for ref, label, val in data_list[15:25]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val)

    with t3:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### -- Data Catin Perempuan --")
            for ref, label, val in data_list[25:35]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val)
        with col2:
            st.markdown("##### -- Data Orang Tua Perempuan --")
            for ref, label, val in data_list[35:45]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val)

    with t4:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### -- Data Wali & Mahar --")
            for ref, label, val in data_list[45:53]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val)
        with col2:
            st.markdown("##### -- Data Saksi 1 & Saksi 2 --")
            for ref, label, val in data_list[53:]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val)

    btn_simpan = st.form_submit_button("💾 Simpan & Perbarui Data Input", use_container_width=True)

if btn_simpan:
    st.session_state['input_data'] = user_inputs
    st.success("✅ Data berhasil diperbarui!")

# PANEL DOWNLOAD RESULT
st.markdown("---")
st.subheader("📥 Download Hasil")

col_d1, col_d2 = st.columns(2)

with col_d1:
    with open(FILE_NAME, "rb") as f:
        file_bytes = f.read()
    st.download_button(
        label="📥 Download File Excel Utama (.xlsb)",
        data=file_bytes,
        file_name=FILE_NAME,
        mime="application/vnd.ms-excel.sheet.binary.macroenabled.12",
        use_container_width=True
    )

with col_d2:
    def generate_pdf(data_dict):
        buffer = BytesIO()
        F4_SIZE = landscape((612, 936))
        doc = SimpleDocTemplate(buffer, pagesize=F4_SIZE, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        elements = []
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=12, alignment=1, spaceAfter=10)
        elements.append(Paragraph("<b>RINGKASAN ISIAN DATA CATIN (F4)</b>", title_style))
        
        table_rows = [["FIELD / BARIS EXCEL", "ISIAN DATA"]]
        current_data = data_dict if data_dict else {item[1]: item[2] for item in data_list}
        for k, v in current_data.items():
            table_rows.append([str(k), str(v)])
            
        pdf_table = Table(table_rows, colWidths=[250, 600])
        pdf_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elements.append(pdf_table)
        doc.build(elements)
        buffer.seek(0)
        return buffer

    pdf_bytes = generate_pdf(st.session_state.get('input_data'))
    st.download_button(
        label="📄 Download Laporan PDF (Ukuran F4)",
        data=pdf_bytes,
        file_name="Isian_Data_Catin_F4.pdf",
        mime="application/pdf",
        use_container_width=True
    )
