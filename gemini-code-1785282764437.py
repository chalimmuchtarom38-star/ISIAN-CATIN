import os
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from io import BytesIO
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

@st.cache_data(ttl=1)
def load_excel_data():
    try:
        excel_file = pd.ExcelFile(FILE_NAME, engine='pyxlsb')
        sheet_name = next((s for s in excel_file.sheet_names if s.strip().upper() == "ISIAN DATA"), excel_file.sheet_names[0])
        df = pd.read_excel(FILE_NAME, sheet_name=sheet_name, engine='pyxlsb', header=None)
        
        mapping_rows = [
            # Group: SURAT & AKAD
            ("SURAT", "Baris 2", "Nomor Register", 1),
            ("SURAT", "Baris 3", "Tanggal Surat", 2),
            ("SURAT", "Baris 4", "Tanggal Pelaksanaan", 3),
            ("SURAT", "Baris 5", "Jam Pelaksanaan", 4),
            ("SURAT", "Baris 6", "Tempat Akad Nikah", 5),
            
            # Group: CATIN PRIA
            ("CATIN_L", "Baris 8", "Nama Catin Laki-Laki", 7),
            ("CATIN_L", "Baris 9", "Bin (Ayah Laki-Laki)", 8),
            ("CATIN_L", "Baris 10", "TTL Catin Laki-Laki", 9),
            ("CATIN_L", "Baris 11", "NIK Catin Laki-Laki", 10),
            ("CATIN_L", "Baris 12", "Pekerjaan Laki-Laki", 11),
            ("CATIN_L", "Baris 13", "Status Laki-Laki", 12),
            ("CATIN_L", "Baris 14", "Jenis Kelamin Laki-Laki", 13),
            ("CATIN_L", "Baris 15", "Nama Istri Terdahulu", 14),
            ("CATIN_L", "Baris 16", "Alamat Catin Laki-Laki", 15),
            ("CATIN_L", "Baris 17", "Pendidikan Laki-Laki", 16),
            
            # Group: ORTU PRIA
            ("ORTU_L", "Baris 19", "Nama Ayah Laki-Laki", 18),
            ("ORTU_L", "Baris 20", "NIK Ayah Laki-Laki", 19),
            ("ORTU_L", "Baris 21", "TTL Ayah Laki-Laki", 20),
            ("ORTU_L", "Baris 22", "Pekerjaan Ayah Laki-Laki", 21),
            ("ORTU_L", "Baris 23", "Alamat Ayah Laki-Laki", 22),
            ("ORTU_L", "Baris 26", "Nama Ibu Laki-Laki", 25),
            ("ORTU_L", "Baris 27", "NIK Ibu Laki-Laki", 26),
            ("ORTU_L", "Baris 28", "TTL Ibu Laki-Laki", 27),
            ("ORTU_L", "Baris 29", "Pekerjaan Ibu Laki-Laki", 28),
            ("ORTU_L", "Baris 30", "Alamat Ibu Laki-Laki", 29),
            
            # Group: CATIN WANITA
            ("CATIN_P", "Baris 34", "Nama Catin Perempuan", 33),
            ("CATIN_P", "Baris 35", "Binti (Ayah Perempuan)", 34),
            ("CATIN_P", "Baris 36", "TTL Catin Perempuan", 35),
            ("CATIN_P", "Baris 37", "NIK Catin Perempuan", 36),
            ("CATIN_P", "Baris 38", "Pekerjaan Perempuan", 37),
            ("CATIN_P", "Baris 39", "Status Perempuan", 38),
            ("CATIN_P", "Baris 40", "Jenis Kelamin Perempuan", 39),
            ("CATIN_P", "Baris 41", "Alamat Catin Perempuan", 40),
            ("CATIN_P", "Baris 42", "Nama Suami Terdahulu", 41),
            ("CATIN_P", "Baris 43", "Pendidikan Perempuan", 42),
            
            # Group: ORTU WANITA
            ("ORTU_P", "Baris 45", "Nama Ayah Perempuan", 44),
            ("ORTU_P", "Baris 46", "NIK Ayah Perempuan", 45),
            ("ORTU_P", "Baris 47", "TTL Ayah Perempuan", 46),
            ("ORTU_P", "Baris 48", "Pekerjaan Ayah Perempuan", 47),
            ("ORTU_P", "Baris 49", "Alamat Ayah Perempuan", 48),
            ("ORTU_P", "Baris 52", "Nama Ibu Perempuan", 51),
            ("ORTU_P", "Baris 53", "NIK Ibu Perempuan", 52),
            ("ORTU_P", "Baris 54", "TTL Ibu Perempuan", 53),
            ("ORTU_P", "Baris 55", "Pekerjaan Ibu Perempuan", 54),
            ("ORTU_P", "Baris 56", "Alamat Ibu Perempuan", 55),
            
            # Group: WALI & SAKSI
            ("WALI_SAKSI", "Baris 58", "Nama Wali", 57),
            ("WALI_SAKSI", "Baris 59", "Bin Wali", 58),
            ("WALI_SAKSI", "Baris 60", "NIK Wali", 59),
            ("WALI_SAKSI", "Baris 61", "TTL Wali", 60),
            ("WALI_SAKSI", "Baris 62", "Pekerjaan Wali", 61),
            ("WALI_SAKSI", "Baris 63", "Alamat Wali", 62),
            ("WALI_SAKSI", "Baris 64", "Hubungan Wali", 63),
            ("WALI_SAKSI", "Baris 65", "Mahar / Maskawin", 64),
            ("WALI_SAKSI", "Baris 70", "Nama Saksi 1", 69),
            ("WALI_SAKSI", "Baris 71", "TTL Saksi 1", 70),
            ("WALI_SAKSI", "Baris 72", "NIK Saksi 1", 71),
            ("WALI_SAKSI", "Baris 73", "Pekerjaan Saksi 1", 72),
            ("WALI_SAKSI", "Baris 74", "Alamat Saksi 1", 73),
            ("WALI_SAKSI", "Baris 76", "Nama Saksi 2", 75),
            ("WALI_SAKSI", "Baris 77", "TTL Saksi 2", 76),
            ("WALI_SAKSI", "Baris 78", "NIK Saksi 2", 77),
            ("WALI_SAKSI", "Baris 79", "Pekerjaan Saksi 2", 78),
            ("WALI_SAKSI", "Baris 80", "Alamat Saksi 2", 79)
        ]
        
        extracted_data = []
        for group, ref, label, r_idx in mapping_rows:
            val = ""
            if r_idx < len(df):
                raw_val = df.iloc[r_idx, 6] if pd.notna(df.iloc[r_idx, 6]) else df.iloc[r_idx, 5]
                val = format_tanggal_indonesia(raw_val)
                if val.strip() == ":":
                    val = ""
            extracted_data.append((group, ref, label, val))
            
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
        for group, ref, label, val in [item for item in data_list if item[0] == "SURAT"]:
            user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val)

    with t2:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### -- Data Catin Laki-Laki --")
            for group, ref, label, val in [item for item in data_list if item[0] == "CATIN_L"]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val)
        with col2:
            st.markdown("##### -- Data Orang Tua Laki-Laki --")
            for group, ref, label, val in [item for item in data_list if item[0] == "ORTU_L"]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val)

    with t3:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### -- Data Catin Perempuan --")
            for group, ref, label, val in [item for item in data_list if item[0] == "CATIN_P"]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val)
        with col2:
            st.markdown("##### -- Data Orang Tua Perempuan --")
            for group, ref, label, val in [item for item in data_list if item[0] == "ORTU_P"]:
                user_inputs[label] = st.text_input(f"[{ref}] {label}", value=val)

    with t4:
        st.markdown("##### -- Data Wali, Mahar & Saksi --")
        for group, ref, label, val in [item for item in data_list if item[0] == "WALI_SAKSI"]:
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
    # FUNGSI PEMBUATAN PDF KELOMPOK TERSTRUKTUR (1 LEMBAR F4 LANDSCAPE)
    def generate_pdf_grouped(data_dict):
        buffer = BytesIO()
        F4_LANDSCAPE = landscape((612, 936)) # Ukuran F4 Landscape
        
        doc = SimpleDocTemplate(
            buffer, 
            pagesize=F4_LANDSCAPE, 
            rightMargin=15, 
            leftMargin=15, 
            topMargin=12, 
            bottomMargin=12
        )
        elements = []
        styles = getSampleStyleSheet()
        
        # Styles
        title_style = ParagraphStyle('T', fontSize=10, leading=11, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor("#0F2C59"))
        sub_title = ParagraphStyle('ST', fontSize=7.5, leading=8.5, alignment=1, fontName='Helvetica-Oblique', textColor=colors.HexColor("#444444"))
        
        head_cat = ParagraphStyle('HC', fontSize=6.5, leading=7.5, fontName='Helvetica-Bold', textColor=colors.whitesmoke, alignment=1)
        lbl_style = ParagraphStyle('L', fontSize=6, leading=7, fontName='Helvetica-Bold')
        val_style = ParagraphStyle('V', fontSize=6, leading=7, fontName='Helvetica')
        
        # Header Dokumen
        elements.append(Paragraph("RINGKASAN ISIAN DATA BERKAS CATIN (F4)", title_style))
        elements.append(Paragraph("Dokumen Verifikasi & Data Isian Pernikahan", sub_title))
        elements.append(Spacer(1, 4))
        
        # Gabungkan Data Terbaru
        current_data = data_dict if data_dict else {item[2]: item[3] for item in data_list}
        
        def get_v(lbl):
            return str(current_data.get(lbl, ""))

        # SUSUNAN BLOK KELOMPOK DALAM DUO KOLOM
        # Kolom Kiri: Catin Laki-Laki & Ortu Laki-Laki
        # Kolom Kanan: Catin Perempuan & Ortu Perempuan
        
        # 1. Trik Tabel untuk Header Akad & Register (Melintang Atas)
        row_akad = [
            Paragraph("<b>No. Register:</b> " + get_v("Nomor Register"), val_style),
            Paragraph("<b>Tgl Surat:</b> " + get_v("Tanggal Surat"), val_style),
            Paragraph("<b>Tgl Pelaksanaan:</b> " + get_v("Tanggal Pelaksanaan"), val_style),
            Paragraph("<b>Jam:</b> " + get_v("Jam Pelaksanaan"), val_style),
            Paragraph("<b>Tempat Akad:</b> " + get_v("Tempat Akad Nikah"), val_style)
        ]
        t_akad = Table([row_akad], colWidths=[180, 180, 180, 120, 246])
        t_akad.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#EAECEE")),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDC3C7")),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('LEFTPADDING', (0,0), (-1,-1), 4),
        ]))
        elements.append(t_akad)
        elements.append(Spacer(1, 4))
        
        # 2. Tabel Utama 2 Kolom Besar (Sisi Pria vs Sisi Wanita)
        col_pria = [
            [Paragraph("DATA CATIN LAKI-LAKI", head_cat), ""],
            [Paragraph("Nama Lengkap", lbl_style), Paragraph(get_v("Nama Catin Laki-Laki") + " Bin " + get_v("Bin (Ayah Laki-Laki)"), val_style)],
            [Paragraph("NIK / TTL", lbl_style), Paragraph(get_v("NIK Catin Laki-Laki") + " / " + get_v("TTL Catin Laki-Laki"), val_style)],
            [Paragraph("Status / Pekerjaan", lbl_style), Paragraph(get_v("Status Laki-Laki") + " / " + get_v("Pekerjaan Laki-Laki"), val_style)],
            [Paragraph("Alamat", lbl_style), Paragraph(get_v("Alamat Catin Laki-Laki"), val_style)],
            [Paragraph("DATA ORANG TUA (AYAH & IBU) LAKI-LAKI", head_cat), ""],
            [Paragraph("Ayah / NIK", lbl_style), Paragraph(get_v("Nama Ayah Laki-Laki") + " / " + get_v("NIK Ayah Laki-Laki"), val_style)],
            [Paragraph("TTL / Pekerjaan", lbl_style), Paragraph(get_v("TTL Ayah Laki-Laki") + " / " + get_v("Pekerjaan Ayah Laki-Laki"), val_style)],
            [Paragraph("Ibu / NIK", lbl_style), Paragraph(get_v("Nama Ibu Laki-Laki") + " / " + get_v("NIK Ibu Laki-Laki"), val_style)],
            [Paragraph("TTL / Pekerjaan", lbl_style), Paragraph(get_v("TTL Ibu Laki-Laki") + " / " + get_v("Pekerjaan Ibu Laki-Laki"), val_style)],
            [Paragraph("Alamat Ortu", lbl_style), Paragraph(get_v("Alamat Ayah Laki-Laki"), val_style)]
        ]
        
        col_wanita = [
            [Paragraph("DATA CATIN PEREMPUAN", head_cat), ""],
            [Paragraph("Nama Lengkap", lbl_style), Paragraph(get_v("Nama Catin Perempuan") + " Binti " + get_v("Binti (Ayah Perempuan)"), val_style)],
            [Paragraph("NIK / TTL", lbl_style), Paragraph(get_v("NIK Catin Perempuan") + " / " + get_v("TTL Catin Perempuan"), val_style)],
            [Paragraph("Status / Pekerjaan", lbl_style), Paragraph(get_v("Status Perempuan") + " / " + get_v("Pekerjaan Perempuan"), val_style)],
            [Paragraph("Alamat", lbl_style), Paragraph(get_v("Alamat Catin Perempuan"), val_style)],
            [Paragraph("DATA ORANG TUA (AYAH & IBU) PEREMPUAN", head_cat), ""],
            [Paragraph("Ayah / NIK", lbl_style), Paragraph(get_v("Nama Ayah Perempuan") + " / " + get_v("NIK Ayah Perempuan"), val_style)],
            [Paragraph("TTL / Pekerjaan", lbl_style), Paragraph(get_v("TTL Ayah Perempuan") + " / " + get_v("Pekerjaan Ayah Perempuan"), val_style)],
            [Paragraph("Ibu / NIK", lbl_style), Paragraph(get_v("Nama Ibu Perempuan") + " / " + get_v("NIK Ibu Perempuan"), val_style)],
            [Paragraph("TTL / Pekerjaan", lbl_style), Paragraph(get_v("TTL Ibu Perempuan") + " / " + get_v("Pekerjaan Ibu Perempuan"), val_style)],
            [Paragraph("Alamat Ortu", lbl_style), Paragraph(get_v("Alamat Ayah Perempuan"), val_style)]
        ]
        
        t_pria = Table(col_pria, colWidths=[100, 340])
        t_pria.setStyle(TableStyle([
            ('SPAN', (0,0), (1,0)),
            ('BACKGROUND', (0,0), (1,0), colors.HexColor("#1F4E78")),
            ('SPAN', (0,5), (1,5)),
            ('BACKGROUND', (0,5), (1,5), colors.HexColor("#1F4E78")),
            ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#D5D8DC")),
            ('TOPPADDING', (0,0), (-1,-1), 1.8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 1.8),
        ]))

        t_wanita = Table(col_wanita, colWidths=[100, 340])
        t_wanita.setStyle(TableStyle([
            ('SPAN', (0,0), (1,0)),
            ('BACKGROUND', (0,0), (1,0), colors.HexColor("#2E75B6")),
            ('SPAN', (0,5), (1,5)),
            ('BACKGROUND', (0,5), (1,5), colors.HexColor("#2E75B6")),
            ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#D5D8DC")),
            ('TOPPADDING', (0,0), (-1,-1), 1.8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 1.8),
        ]))

        main_grid = Table([[t_pria, t_wanita]], colWidths=[448, 448])
        main_grid.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ]))
        elements.append(main_grid)
        elements.append(Spacer(1, 4))
        
        # 3. Kelompok Wali, Mahar & Saksi (Bawah Melintang)
        col_wali_saksi = [
            [Paragraph("DATA WALI, MAHAR & SAKSI-SAKSI NIKAH", head_cat), "", "", ""],
            [
                Paragraph("<b>Wali Nikah:</b> " + get_v("Nama Wali") + " Bin " + get_v("Bin Wali") + " (" + get_v("Hubungan Wali") + ")", val_style),
                Paragraph("<b>NIK / TTL Wali:</b> " + get_v("NIK Wali") + " / " + get_v("TTL Wali"), val_style),
                Paragraph("<b>Pekerjaan / Alamat:</b> " + get_v("Pekerjaan Wali") + " / " + get_v("Alamat Wali"), val_style),
                Paragraph("<b>Mahar / Maskawin:</b> " + get_v("Mahar / Maskawin"), val_style)
            ],
            [
                Paragraph("<b>Saksi 1:</b> " + get_v("Nama Saksi 1"), val_style),
                Paragraph("<b>NIK / TTL Saksi 1:</b> " + get_v("NIK Saksi 1") + " / " + get_v("TTL Saksi 1"), val_style),
                Paragraph("<b>Pekerjaan / Alamat:</b> " + get_v("Pekerjaan Saksi 1") + " / " + get_v("Alamat Saksi 1"), val_style),
                ""
            ],
            [
                Paragraph("<b>Saksi 2:</b> " + get_v("Nama Saksi 2"), val_style),
                Paragraph("<b>NIK / TTL Saksi 2:</b> " + get_v("NIK Saksi 2") + " / " + get_v("TTL Saksi 2"), val_style),
                Paragraph("<b>Pekerjaan / Alamat:</b> " + get_v("Pekerjaan Saksi 2") + " / " + get_v("Alamat Saksi 2"), val_style),
                ""
            ]
        ]
        
        t_ws = Table(col_wali_saksi, colWidths=[240, 220, 280, 166])
        t_ws.setStyle(TableStyle([
            ('SPAN', (0,0), (3,0)),
            ('BACKGROUND', (0,0), (3,0), colors.HexColor("#34495E")),
            ('SPAN', (2,2), (3,2)),
            ('SPAN', (2,3), (3,3)),
            ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#D5D8DC")),
            ('TOPPADDING', (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ('LEFTPADDING', (0,0), (-1,-1), 4),
        ]))
        elements.append(t_ws)

        doc.build(elements)
        buffer.seek(0)
        return buffer

    pdf_bytes = generate_pdf_grouped(st.session_state.get('input_data'))
    st.download_button(
        label="📄 Download Laporan PDF Formal (Ter-Kelompok 1 Lembar F4)",
        data=pdf_bytes,
        file_name="Isian_Data_Catin_F4_Rapi.pdf",
        mime="application/pdf",
        use_container_width=True
    )
