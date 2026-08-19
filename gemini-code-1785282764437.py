import os
import re
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from io import BytesIO
import openpyxl
from reportlab.lib.pagesizes import landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Form Input Catin F4", layout="wide")
st.title("📝 Formulir Input Isian Data Catin (F4)")

FILE_NAME = "BERKAS_CATIN_F4.xlsb"

if not os.path.exists(FILE_NAME):
    st.error(f"File '{FILE_NAME}' tidak ditemukan di folder aplikasi! Pastikan file berada di folder yang sama dengan app.py.")
    st.stop()

# Daftar Nama Bulan Bahasa Indonesia
NAMA_BULAN = [
    "", "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
    "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"
]

def format_tanggal_indonesia(val):
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, (int, float)):
        try:
            date_obj = datetime(1899, 12, 30) + timedelta(days=int(val))
            return f"{date_obj.day} {NAMA_BULAN[date_obj.month]} {date_obj.year}"
        except:
            return str(val)
    val_str = str(val).strip()
    if val_str.isdigit() and len(val_str) >= 5:
        try:
            date_obj = datetime(1899, 12, 30) + timedelta(days=int(val_str))
            return f"{date_obj.day} {NAMA_BULAN[date_obj.month]} {date_obj.year}"
        except:
            return val_str
    return val_str

# Peta Baris Excel (Excel Row Index 1-based, Kolom G / Index 7)
MAPPING_ROWS = [
    ("Baris 2", "Nomor Register", 2),
    ("Baris 3", "Tanggal Surat", 3),
    ("Baris 4", "Tanggal Pelaksanaan", 4),
    ("Baris 5", "Jam Pelaksanaan", 5),
    ("Baris 6", "Tempat Akad Nikah", 6),
    
    # Catin Laki-Laki
    ("Baris 8", "Nama Catin Laki-Laki", 8),
    ("Baris 9", "Bin (Ayah Laki-Laki)", 9),
    ("Baris 10", "TTL Catin Laki-Laki", 10),
    ("Baris 11", "NIK Catin Laki-Laki", 11),
    ("Baris 12", "Pekerjaan Laki-Laki", 12),
    ("Baris 13", "Status Laki-Laki", 13),
    ("Baris 14", "Jenis Kelamin Laki-Laki", 14),
    ("Baris 15", "Nama Istri Terdahulu", 15),
    ("Baris 16", "Alamat Catin Laki-Laki", 16),
    ("Baris 17", "Pendidikan Laki-Laki", 17),
    
    # Ortu Laki-Laki
    ("Baris 19", "Nama Ayah Laki-Laki", 19),
    ("Baris 20", "NIK Ayah Laki-Laki", 20),
    ("Baris 21", "TTL Ayah Laki-Laki", 21),
    ("Baris 22", "Pekerjaan Ayah Laki-Laki", 22),
    ("Baris 23", "Alamat Ayah Laki-Laki", 23),
    ("Baris 26", "Nama Ibu Laki-Laki", 26),
    ("Baris 27", "NIK Ibu Laki-Laki", 27),
    ("Baris 28", "TTL Ibu Laki-Laki", 28),
    ("Baris 29", "Pekerjaan Ibu Laki-Laki", 29),
    ("Baris 30", "Alamat Ibu Laki-Laki", 30),
    
    # Catin Perempuan
    ("Baris 34", "Nama Catin Perempuan", 34),
    ("Baris 35", "Binti (Ayah Perempuan)", 35),
    ("Baris 36", "TTL Catin Perempuan", 36),
    ("Baris 37", "NIK Catin Perempuan", 37),
    ("Baris 38", "Pekerjaan Perempuan", 38),
    ("Baris 39", "Status Perempuan", 39),
    ("Baris 40", "Jenis Kelamin Perempuan", 40),
    ("Baris 41", "Alamat Catin Perempuan", 41),
    ("Baris 42", "Nama Suami Terdahulu", 42),
    ("Baris 43", "Pendidikan Perempuan", 43),
    
    # Ortu Perempuan
    ("Baris 45", "Nama Ayah Perempuan", 45),
    ("Baris 46", "NIK Ayah Perempuan", 46),
    ("Baris 47", "TTL Ayah Perempuan", 47),
    ("Baris 48", "Pekerjaan Ayah Perempuan", 48),
    ("Baris 49", "Alamat Ayah Perempuan", 49),
    ("Baris 52", "Nama Ibu Perempuan", 52),
    ("Baris 53", "NIK Ibu Perempuan", 53),
    ("Baris 54", "TTL Ibu Perempuan", 54),
    ("Baris 55", "Pekerjaan Ibu Perempuan", 55),
    ("Baris 56", "Alamat Ibu Perempuan", 56),
    
    # Wali & Saksi
    ("Baris 58", "Nama Wali", 58),
    ("Baris 59", "Bin Wali", 59),
    ("Baris 60", "NIK Wali", 60),
    ("Baris 61", "TTL Wali", 61),
    ("Baris 62", "Pekerjaan Wali", 62),
    ("Baris 63", "Alamat Wali", 63),
    ("Baris 64", "Hubungan Wali", 64),
    ("Baris 65", "Mahar / Maskawin", 65),
    ("Baris 70", "Nama Saksi 1", 70),
    ("Baris 71", "TTL Saksi 1", 71),
    ("Baris 72", "NIK Saksi 1", 72),
    ("Baris 73", "Pekerjaan Saksi 1", 73),
    ("Baris 74", "Alamat Saksi 1", 74),
    ("Baris 76", "Nama Saksi 2", 76),
    ("Baris 77", "TTL Saksi 2", 77),
    ("Baris 78", "NIK Saksi 2", 78),
    ("Baris 79", "Pekerjaan Saksi 2", 79),
    ("Baris 80", "Alamat Saksi 2", 80)
]

@st.cache_data(ttl=1)
def load_excel_data():
    try:
        excel_file = pd.ExcelFile(FILE_NAME, engine='pyxlsb')
        sheet_name = next((s for s in excel_file.sheet_names if s.strip().upper() == "ISIAN DATA"), excel_file.sheet_names[0])
        df = pd.read_excel(FILE_NAME, sheet_name=sheet_name, engine='pyxlsb', header=None)
        
        extracted_data = []
        for ref, label, r_idx in MAPPING_ROWS:
            val = ""
            row_0based = r_idx - 1
            if row_0based < len(df):
                raw_val = df.iloc[row_0based, 6] if pd.notna(df.iloc[row_0based, 6]) else df.iloc[row_0based, 5]
                val = format_tanggal_indonesia(raw_val)
                if str(val).strip() == ":":
                    val = ""
            extracted_data.append((ref, label, val, r_idx))
            
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
        for ref, label, val, row_num in data_list[:5]:
            user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val, key=f"inp_{row_num}")

    with t2:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### -- Data Catin Laki-Laki --")
            for ref, label, val, row_num in data_list[5:15]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val, key=f"inp_{row_num}")
        with col2:
            st.markdown("##### -- Data Orang Tua Laki-Laki --")
            for ref, label, val, row_num in data_list[15:25]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val, key=f"inp_{row_num}")

    with t3:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### -- Data Catin Perempuan --")
            for ref, label, val, row_num in data_list[25:35]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val, key=f"inp_{row_num}")
        with col2:
            st.markdown("##### -- Data Orang Tua Perempuan --")
            for ref, label, val, row_num in data_list[35:45]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val, key=f"inp_{row_num}")

    with t4:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### -- Data Wali & Mahar --")
            for ref, label, val, row_num in data_list[45:53]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val, key=f"inp_{row_num}")
        with col2:
            st.markdown("##### -- Data Saksi 1 & Saksi 2 --")
            for ref, label, val, row_num in data_list[53:]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val, key=f"inp_{row_num}")

    btn_simpan = st.form_submit_button("💾 Simpan & Perbarui Data Input", use_container_width=True)

if btn_simpan:
    st.session_state['input_data'] = user_inputs
    st.success("✅ Data berhasil diperbarui di memori sistem!")

# PANEL DOWNLOAD RESULT
st.markdown("---")
st.subheader("📥 Download Hasil Isian")

col_d1, col_d2 = st.columns(2)

# FUNGSI MEMPERBARUI SEL DALAM EXCEL
def generate_updated_excel(data_dict):
    """
    Membaca data master, menyusun ulang dataframe/sheet,
    lalu menuliskan isian baru ke Kolom G (Kolom ke-7).
    """
    output = BytesIO()
    excel_file = pd.ExcelFile(FILE_NAME, engine='pyxlsb')
    sheet_name = next((s for s in excel_file.sheet_names if s.strip().upper() == "ISIAN DATA"), excel_file.sheet_names[0])
    
    df = pd.read_excel(FILE_NAME, sheet_name=sheet_name, engine='pyxlsb', header=None)
    
    current_inputs = data_dict if data_dict else {item[1]: item[2] for item in data_list}
    
    # Update data pada Kolom G (indeks kolom 6 di Pandas 0-based)
    label_to_row = {item[1]: item[2] for item in MAPPING_ROWS}
    for label, val in current_inputs.items():
        if label in label_to_row:
            r_idx = label_to_row[label] - 1  # Konversi ke 0-based index
            if r_idx < len(df):
                df.iloc[r_idx, 6] = val

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, header=False, sheet_name="ISIAN DATA")
        
    output.seek(0)
    return output

with col_d1:
    excel_bytes = generate_updated_excel(st.session_state.get('input_data'))
    
    curr_data = st.session_state.get('input_data', {item[1]: item[2] for item in data_list})
    pria = curr_data.get("Nama Catin Laki-Laki", "").strip().upper()
    wanita = curr_data.get("Nama Catin Perempuan", "").strip().upper()
    file_excel_name = f"BERKAS_CATIN_{pria}_&_{wanita}.xlsx" if (pria or wanita) else "BERKAS_CATIN_TERISI.xlsx"

    st.download_button(
        label="📊 Download File Excel Terisi (.xlsx)",
        data=excel_bytes,
        file_name=file_excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with col_d2:
    def generate_pdf_formal(data_dict):
        buffer = BytesIO()
        F4_LANDSCAPE = landscape((612, 936))
        
        doc = SimpleDocTemplate(
            buffer, 
            pagesize=F4_LANDSCAPE, 
            rightMargin=15, 
            leftMargin=15, 
            topMargin=15, 
            bottomMargin=15
        )
        elements = []
        styles = getSampleStyleSheet()
        
        header_style = ParagraphStyle(
            'HeaderStyle',
            parent=styles['Heading1'],
            fontSize=11,
            leading=13,
            alignment=1,
            fontName='Helvetica-Bold',
            textColor=colors.HexColor("#0F2C59")
        )
        
        sub_header_style = ParagraphStyle(
            'SubHeaderStyle',
            parent=styles['Normal'],
            fontSize=8,
            leading=10,
            alignment=1,
            fontName='Helvetica-Oblique',
            textColor=colors.HexColor("#333333")
        )
        
        cell_bold = ParagraphStyle('CB', fontSize=6.5, leading=7.5, fontName='Helvetica-Bold')
        cell_norm = ParagraphStyle('CN', fontSize=6.5, leading=7.5, fontName='Helvetica')
        
        elements.append(Paragraph("RINGKASAN ISIAN DATA BERKAS CATIN (F4)", header_style))
        elements.append(Paragraph("Daftar Pemeriksaan & Verifikasi Data Pernikahan", sub_header_style))
        elements.append(Spacer(1, 6))
        
        current_data = data_dict if data_dict else {item[1]: item[2] for item in data_list}
        items = list(current_data.items())
        
        table_rows = [
            [
                Paragraph("<b>FIELD / PARAMETER (A)</b>", cell_bold),
                Paragraph("<b>ISIAN DATA (A)</b>", cell_bold),
                Paragraph("<b>FIELD / PARAMETER (B)</b>", cell_bold),
                Paragraph("<b>ISIAN DATA (B)</b>", cell_bold)
            ]
        ]
        
        half = (len(items) + 1) // 2
        for i in range(half):
            k1, v1 = items[i]
            k2, v2 = items[i + half] if (i + half) < len(items) else ("", "")
            
            table_rows.append([
                Paragraph(f"<b>{k1}</b>", cell_bold),
                Paragraph(str(v1), cell_norm),
                Paragraph(f"<b>{k2}</b>" if k2 else "", cell_bold),
                Paragraph(str(v2) if v2 else "", cell_norm)
            ])
            
        col_widths = [160, 293, 160, 293]
        
        pdf_table = Table(table_rows, colWidths=col_widths, repeatRows=1)
        pdf_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F2C59")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 1.8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1.8),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F9F9F9")]),
        ]))
        
        elements.append(pdf_table)
        doc.build(elements)
        buffer.seek(0)
        return buffer

    pdf_bytes = generate_pdf_formal(st.session_state.get('input_data'))
    st.download_button(
        label="📄 Download Laporan PDF Formal (1 Lembar F4)",
        data=pdf_bytes,
        file_name="Isian_Data_Catin_F4_Formal.pdf",
        mime="application/pdf",
        use_container_width=True
    )
