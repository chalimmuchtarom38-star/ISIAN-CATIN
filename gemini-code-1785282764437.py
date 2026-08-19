import os
import re
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from io import BytesIO
from reportlab.lib.pagesizes import portrait
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

# FUNGSI MENEVALUASI RUMUS UMUR SECARA AKURAT BERDASARKAN TAHUN BERJALAN
def hitung_umur_otomatis(ttl_val, raw_excel_val=""):
    tahun_sekarang = datetime.now().year
    
    # 1. Ekstrak tahun dari TTL (misal: "Pemalang, 15-08-1989" -> 1989)
    if ttl_val:
        years = re.findall(r'\b(19\d\d|20\d\d)\b', str(ttl_val))
        if years:
            tahun_lahir = int(years[-1])
            umur_kalkulasi = tahun_sekarang - tahun_lahir
            if 10 <= umur_kalkulasi <= 100:
                return f"{umur_kalkulasi} TAHUN"

    # 2. Jika ada nilai mentah di Excel dan valid (>10 tahun), gunakan nilai tersebut
    if raw_excel_val:
        digits = re.sub(r'[^0-9]', '', str(raw_excel_val))
        if digits.isdigit():
            u_num = int(digits)
            # Koreksi jika angka simpanan lama (cached) di xlsb masih 27 tapi lahir 1989
            if u_num >= 10:
                return f"{u_num} TAHUN"

    return ""

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
        ttl_l, ttl_p = "", ""
        raw_u_l, raw_u_p = "", ""

        # Membaca kolom samping umur jika ada di Excel
        try:
            for col_idx in range(6, df.shape[1]):
                val_l = df.iloc[9, col_idx] if pd.notna(df.iloc[9, col_idx]) else ""
                val_p = df.iloc[35, col_idx] if pd.notna(df.iloc[35, col_idx]) else ""
                if not raw_u_l and ("TH" in str(val_l).upper() or str(val_l).isdigit()):
                    raw_u_l = str(val_l)
                if not raw_u_p and ("TH" in str(val_p).upper() or str(val_p).isdigit()):
                    raw_u_p = str(val_p)
        except:
            pass

        for group, ref, label, r_idx in mapping_rows:
            val = ""
            if r_idx < len(df):
                raw_val = df.iloc[r_idx, 6] if pd.notna(df.iloc[r_idx, 6]) else df.iloc[r_idx, 5]
                val = format_tanggal_indonesia(raw_val)
                if val.strip() == ":":
                    val = ""
            
            if label == "TTL Catin Laki-Laki":
                ttl_l = val
            elif label == "TTL Catin Perempuan":
                ttl_p = val
                
            extracted_data.append((group, ref, label, val))
            
        # Sisipkan Umur Otomatis Hasil Kalkulasi Rumus Presisi
        umur_l_final = hitung_umur_otomatis(ttl_l, raw_u_l)
        umur_p_final = hitung_umur_otomatis(ttl_p, raw_u_p)

        extracted_data.append(("CATIN_L", "Baris 10", "Umur Catin Laki-Laki", umur_l_final))
        extracted_data.append(("CATIN_P", "Baris 36", "Umur Catin Perempuan", umur_p_final))

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
    def generate_pdf_portrait_f4(data_dict):
        buffer = BytesIO()
        F4_PORTRAIT = portrait((612, 936))
        
        doc = SimpleDocTemplate(
            buffer, 
            pagesize=F4_PORTRAIT, 
            rightMargin=25, 
            leftMargin=25, 
            topMargin=15, 
            bottomMargin=15
        )
        elements = []
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle('T', fontSize=12, leading=14, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor("#0F2C59"))
        sub_title = ParagraphStyle('ST', fontSize=8.5, leading=10, alignment=1, fontName='Helvetica-Oblique', textColor=colors.HexColor("#333333"))
        
        head_sec = ParagraphStyle('HS', fontSize=8, leading=9, fontName='Helvetica-Bold', textColor=colors.whitesmoke)
        lbl_style = ParagraphStyle('L', fontSize=7.5, leading=8.5, fontName='Helvetica-Bold')
        val_style = ParagraphStyle('V', fontSize=7.5, leading=8.5, fontName='Helvetica')
        center_style = ParagraphStyle('CS', fontSize=7.5, leading=9.5, alignment=1, fontName='Helvetica')
        center_bold = ParagraphStyle('CB', fontSize=8, leading=10, alignment=1, fontName='Helvetica-Bold')

        # Header Dokumen
        elements.append(Paragraph("PEMERINTAH KABUPATEN PEMALANG", title_style))
        elements.append(Paragraph("BERKAS VERIFIKASI ISIAN DATA CATIN (FORM F4) - DESA TAMBI", sub_title))
        elements.append(Spacer(1, 6))
        
        current_data = data_dict if data_dict else {item[2]: item[3] for item in data_list}
        def get_v(lbl):
            return str(current_data.get(lbl, "")).strip()

        # Format Umur Dinamis & Presisi
        u_l = get_v("Umur Catin Laki-Laki") or hitung_umur_otomatis(get_v("TTL Catin Laki-Laki"))
        u_p = get_v("Umur Catin Perempuan") or hitung_umur_otomatis(get_v("TTL Catin Perempuan"))
        str_u_l = f" / {u_l}" if u_l else ""
        str_u_p = f" / {u_p}" if u_p else ""

        # 1. TABLE REGIST & AKAD
        data_akad = [
            [Paragraph("<b>No. Register:</b> " + get_v("Nomor Register"), val_style), Paragraph("<b>Tgl Surat:</b> " + get_v("Tanggal Surat"), val_style)],
            [Paragraph("<b>Tgl Pelaksanaan:</b> " + get_v("Tanggal Pelaksanaan"), val_style), Paragraph("<b>Jam / Tempat Akad:</b> " + get_v("Jam Pelaksanaan") + " / " + get_v("Tempat Akad Nikah"), val_style)]
        ]
        t_akad = Table(data_akad, colWidths=[281, 281])
        t_akad.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#EAEDED")),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDC3C7")),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('LEFTPADDING', (0,0), (-1,-1), 5),
        ]))
        elements.append(t_akad)
        elements.append(Spacer(1, 5))

        def make_section(title, rows_data):
            content = [[Paragraph(title, head_sec), ""]]
            for r in rows_data:
                content.append([Paragraph(r[0], lbl_style), Paragraph(r[1], val_style)])
            t = Table(content, colWidths=[140, 422])
            t.setStyle(TableStyle([
                ('SPAN', (0,0), (1,0)),
                ('BACKGROUND', (0,0), (1,0), colors.HexColor("#1F4E78")),
                ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#D5D8DC")),
                ('TOPPADDING', (0,0), (-1,-1), 2.5),
                ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
                ('LEFTPADDING', (0,0), (-1,-1), 5),
            ]))
            return t

        # 2. CATIN LAKI-LAKI
        t_pria = make_section("I. DATA CATIN LAKI-LAKI", [
            ("Nama Lengkap / Bin", get_v("Nama Catin Laki-Laki") + " Bin " + get_v("Bin (Ayah Laki-Laki)")),
            ("NIK / TTL / Umur", get_v("NIK Catin Laki-Laki") + " / " + get_v("TTL Catin Laki-Laki") + str_u_l),
            ("Status / Pekerjaan", get_v("Status Laki-Laki") + " / " + get_v("Pekerjaan Laki-Laki")),
            ("Alamat", get_v("Alamat Catin Laki-Laki")),
            ("Ayah / NIK / TTL / Pekerjaan", get_v("Nama Ayah Laki-Laki") + " / " + get_v("NIK Ayah Laki-Laki") + " / " + get_v("TTL Ayah Laki-Laki") + " / " + get_v("Pekerjaan Ayah Laki-Laki")),
            ("Ibu / NIK / TTL / Pekerjaan", get_v("Nama Ibu Laki-Laki") + " / " + get_v("NIK Ibu Laki-Laki") + " / " + get_v("TTL Ibu Laki-Laki") + " / " + get_v("Pekerjaan Ibu Laki-Laki"))
        ])
        elements.append(t_pria)
        elements.append(Spacer(1, 5))

        # 3. CATIN PEREMPUAN
        t_wanita = make_section("II. DATA CATIN PEREMPUAN", [
            ("Nama Lengkap / Binti", get_v("Nama Catin Perempuan") + " Binti " + get_v("Binti (Ayah Perempuan)")),
            ("NIK / TTL / Umur", get_v("NIK Catin Perempuan") + " / " + get_v("TTL Catin Perempuan") + str_u_p),
            ("Status / Pekerjaan", get_v("Status Perempuan") + " / " + get_v("Pekerjaan Perempuan")),
            ("Alamat", get_v("Alamat Catin Perempuan")),
            ("Ayah / NIK / TTL / Pekerjaan", get_v("Nama Ayah Perempuan") + " / " + get_v("NIK Ayah Perempuan") + " / " + get_v("TTL Ayah Perempuan") + " / " + get_v("Pekerjaan Ayah Perempuan")),
            ("Ibu / NIK / TTL / Pekerjaan", get_v("Nama Ibu Perempuan") + " / " + get_v("NIK Ibu Perempuan") + " / " + get_v("TTL Ibu Perempuan") + " / " + get_v("Pekerjaan Ibu Perempuan"))
        ])
        elements.append(t_wanita)
        elements.append(Spacer(1, 5))

        # 4. DATA WALI & SAKSI
        t_wali = make_section("III. DATA WALI NIKAH, MAHAR & SAKSI-SAKSI", [
            ("Nama Wali / Bin", get_v("Nama Wali") + " Bin " + get_v("Bin Wali") + f" ({get_v('Hubungan Wali')})"),
            ("NIK / TTL Wali", get_v("NIK Wali") + " / " + get_v("TTL Wali")),
            ("Pekerjaan & Alamat Wali", get_v("Pekerjaan Wali") + " / " + get_v("Alamat Wali")),
            ("Mahar / Maskawin", get_v("Mahar / Maskawin")),
            ("Saksi I", get_v("Nama Saksi 1") + " / NIK: " + get_v("NIK Saksi 1") + " / Pekerjaan: " + get_v("Pekerjaan Saksi 1")),
            ("Saksi II", get_v("Nama Saksi 2") + " / NIK: " + get_v("NIK Saksi 2") + " / Pekerjaan: " + get_v("Pekerjaan Saksi 2"))
        ])
        elements.append(t_wali)
        elements.append(Spacer(1, 12))

        # 5. KOLOM TANDA TANGAN (KASI PELAYANAN DI KANAN, KADES DI KIRI)
        ttd_data = [
            [
                Paragraph("Mengetahui,<br/><b>KEPALA DESA TAMBI</b>", center_style),
                Paragraph("Desa Tambi, " + format_tanggal_indonesia(datetime.now().strftime("%Y-%m-%d")) + "<br/><b>KASI PELAYANAN DESA TAMBI</b>", center_style)
            ],
            ["", ""],
            [
                Paragraph("<b><u>JURI</u></b>", center_bold),
                Paragraph("<b><u>CHALIM MUCHTAROM, S.Pd.I</u></b>", center_bold)
            ]
        ]
        
        t_ttd = Table(ttd_data, colWidths=[281, 281])
        t_ttd.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,1), (-1,1), 30),
        ]))
        elements.append(t_ttd)

        doc.build(elements)
        buffer.seek(0)
        return buffer

    pdf_bytes = generate_pdf_portrait_f4(st.session_state.get('input_data'))
    st.download_button(
        label="📄 Download Laporan PDF Portrait (1 Lembar F4 Presisi)",
        data=pdf_bytes,
        file_name="Isian_Data_Catin_Desa_Tambi_F4.pdf",
        mime="application/pdf",
        use_container_width=True
    )
