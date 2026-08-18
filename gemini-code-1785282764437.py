import streamlit as st
import openpyxl
import re
import json
import os
from io import BytesIO
from datetime import datetime, date

# Import modul untuk generate PDF F4 1 Lembar
from reportlab.lib.pagesizes import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(
    page_title="Aplikasi Berkas Catin - Desa Tambi",
    page_icon="📜",
    layout="wide"
)

EXCEL_FILE = "BERKAS CATIN .xlsx"
DRAFT_FILE = "draf_terakhir.json"

# --- FUNGSI HARI INDONESIA ---
HARI_INDONESIA = {
    'Monday': 'Senin',
    'Tuesday': 'Selasa',
    'Wednesday': 'Rabu',
    'Thursday': 'Kamis',
    'Friday': 'Jumat',
    'Saturday': 'Sabtu',
    'Sunday': 'Minggu'
}

def get_hari_tgl(tgl_obj):
    if isinstance(tgl_obj, str):
        try:
            tgl_obj = date.fromisoformat(tgl_obj)
        except Exception:
            return tgl_obj
    if isinstance(tgl_obj, (date, datetime)):
        nama_hari = HARI_INDONESIA.get(tgl_obj.strftime('%A'), '')
        return f"{nama_hari}, {tgl_obj.strftime('%d-%m-%Y')}"
    return str(tgl_obj)

# --- FUNGSI RUMUS HITUNG UMUR ---
def hitung_umur(ttl_str):
    if not ttl_str:
        return 0
    match = re.search(r'\b(19\d{2}|20\d{2})\b', str(ttl_str))
    if match:
        tahun_lahir = int(match.group(1))
        tahun_sekarang = datetime.now().year
        return max(0, tahun_sekarang - tahun_lahir)
    return 0

# --- FUNGSI LOAD & SAVE DRAF ---
def load_draft():
    if os.path.exists(DRAFT_FILE):
        try:
            with open(DRAFT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_draft_file():
    draft_data = {}
    # Kunci-kunci internal Streamlit atau data binary yang harus diabaikan saat serialisasi JSON
    ignored_keys = {'excel_bytes', 'pdf_bytes', 'is_processed'}
    
    for key in st.session_state:
        if key in ignored_keys:
            continue
        val = st.session_state[key]
        if isinstance(val, (date, datetime)):
            draft_data[key] = str(val)
        elif isinstance(val, (str, int, float, bool, list, dict)) or val is None:
            draft_data[key] = val
            
    with open(DRAFT_FILE, "w", encoding="utf-8") as f:
        json.dump(draft_data, f, ensure_ascii=False, indent=2)

# --- FUNGSI UPDATE DATA EXCEL ---
def update_excel_data(excel_path, data):
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active

    cell_mapping = {
        'no_register': 'C4',
        'tgl_surat': 'C5',
        'tgl_pelaksanaan': 'C6',
        'jam_akad': 'C7',
        'tempat_akad': 'C8',
        'email_catin': 'C9',
        'mahar': 'C10',

        'nama_lk': 'C13',
        'bin_lk': 'C14',
        'ttl_lk': 'C15',
        'nik_lk': 'C16',
        'pekerjaan_lk': 'C17',
        'status_lk': 'C18',
        'jk_lk': 'C19',
        'istri_terdahulu': 'C20',
        'alamat_lk': 'C21',
        'pendidikan_lk': 'C22',

        'nama_ayah_lk': 'C25',
        'bin_ayah_lk': 'C26',
        'nik_ayah_lk': 'C27',
        'ttl_ayah_lk': 'C28',
        'pekerjaan_ayah_lk': 'C29',
        'alamat_ayah_lk': 'C30',

        'nama_ibu_lk': 'C33',
        'bin_ibu_lk': 'C34',
        'nik_ibu_lk': 'C35',
        'ttl_ibu_lk': 'C36',
        'pekerjaan_ibu_lk': 'C37',
        'alamat_ibu_lk': 'C38',

        'nama_pr': 'F13',
        'binti_pr': 'F14',
        'ttl_pr': 'F15',
        'nik_pr': 'F16',
        'pekerjaan_pr': 'F17',
        'status_pr': 'F18',
        'jk_pr': 'F19',
        'suami_terdahulu': 'F20',
        'alamat_pr': 'F21',
        'pendidikan_pr': 'F22',

        'nama_ayah_pr': 'F25',
        'bin_ayah_pr': 'F26',
        'nik_ayah_pr': 'F27',
        'ttl_ayah_pr': 'F28',
        'pekerjaan_ayah_pr': 'F29',
        'alamat_ayah_pr': 'F30',

        'nama_ibu_pr': 'F33',
        'bin_ibu_pr': 'F34',
        'nik_ibu_pr': 'F35',
        'ttl_ibu_pr': 'F36',
        'pekerjaan_ibu_pr': 'F37',
        'alamat_ibu_pr': 'F38',

        'nama_wali': 'C41',
        'bin_wali': 'C42',
        'nik_wali': 'C43',
        'ttl_wali': 'C44',
        'pekerjaan_wali': 'C45',
        'alamat_wali': 'C46',
        'hubungan_wali': 'C47',
        'nama_wali_lengkap': 'C48',

        'saksi1_nama': 'F41',
        'saksi1_ttl': 'F42',
        'saksi1_nik': 'F43',
        'saksi1_pekerjaan': 'F44',
        'saksi1_alamat': 'F45',

        'saksi2_nama': 'F47',
        'saksi2_ttl': 'F48',
        'saksi2_nik': 'F49',
        'saksi2_pekerjaan': 'F50',
        'saksi2_alamat': 'F51',
    }

    for key, cell in cell_mapping.items():
        if key in data:
            val = data[key]
            if isinstance(val, (date, datetime)):
                val = str(val)
            sheet[cell] = val

    output_stream = BytesIO()
    wb.save(output_stream)
    wb.close()
    return output_stream.getvalue()

# --- FUNGSI GENERATE PDF F4 LENGKAP (1 LEMBAR) ---
def generate_pdf_f4(data):
    buffer = BytesIO()
    f4_size = (215 * mm, 330 * mm)
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=f4_size,
        leftMargin=7 * mm,
        rightMargin=7 * mm,
        topMargin=7 * mm,
        bottomMargin=7 * mm
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=10.5,
        leading=12,
        alignment=1,
        fontName='Helvetica-Bold'
    )
    
    sec_title_style = ParagraphStyle(
        'SecTitleStyle',
        fontSize=8,
        leading=9.5,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#003366')
    )
    
    lbl_style = ParagraphStyle('LblStyle', fontSize=7, leading=8.5, fontName='Helvetica-Bold')
    val_style = ParagraphStyle('ValStyle', fontSize=7, leading=8.5, fontName='Helvetica')
    
    elements = []
    
    elements.append(Paragraph("PEMERINTAH KABUPATEN PEMALANG - KECAMATAN WATUKUMPUL", ParagraphStyle('Kop1', fontSize=7.5, alignment=1, fontName='Helvetica-Bold')))
    elements.append(Paragraph("RINGKASAN LEMBAR VERIFIKASI BERKAS CALON PENGANTIN DESA TAMBI", title_style))
    elements.append(Paragraph(f"No. Register: <b>{data.get('no_register', '-')}</b> | Tanggal Surat: <b>{data.get('tgl_surat', '-')}</b>", ParagraphStyle('SubTitle', fontSize=7.5, alignment=1, leading=9)))
    elements.append(Spacer(1, 3))
    
    def row1(lbl, val):
        return [Paragraph(lbl, lbl_style), Paragraph(":", lbl_style), Paragraph(str(val or '-'), val_style)]

    def row2(lbl1, val1, lbl2, val2):
        return [
            Paragraph(lbl1, lbl_style), Paragraph(":", lbl_style), Paragraph(str(val1 or '-'), val_style),
            Paragraph(lbl2, lbl_style), Paragraph(":", lbl_style), Paragraph(str(val2 or '-'), val_style)
        ]

    hari_tgl_akad = f"{get_hari_tgl(data.get('tgl_pelaksanaan', '-'))} (Jam: {data.get('jam_akad', '-')})"
    tabel_akad_data = [
        [Paragraph("I. PELAKSANAAN AKAD NIKAH", sec_title_style), "", ""],
        row1("Hari & Tgl / Jam Akad", hari_tgl_akad),
        row1("Tempat Akad Nikah", data.get('tempat_akad', '-')),
        row1("Maskawin / Mahar", data.get('mahar', '-')),
        row1("Email Catin", data.get('email_catin', '-')),
    ]

    tabel_catin_data = [
        [Paragraph("II. CALON PENGANTIN LAKI-LAKI", sec_title_style), "", "", Paragraph("III. CALON PENGANTIN PEREMPUAN", sec_title_style), "", ""],
        row2("Nama Lengkap", data.get('nama_lk','-'), "Nama Lengkap", data.get('nama_pr','-')),
        row2("Bin / Binti", data.get('bin_lk','-'), "Bin / Binti", data.get('binti_pr','-')),
        row2("NIK", data.get('nik_lk','-'), "NIK", data.get('nik_pr','-')),
        row2("Tempat, Tgl Lahir", data.get('ttl_lk','-'), "Tempat, Tgl Lahir", data.get('ttl_pr','-')),
        row2("Umur Catin", f"{hitung_umur(data.get('ttl_lk'))} Tahun", "Umur Catin", f"{hitung_umur(data.get('ttl_pr'))} Tahun"),
        row2("Status / Gender", f"{data.get('status_lk','-')} / {data.get('jk_lk','-')}", "Status / Gender", f"{data.get('status_pr','-')} / {data.get('jk_pr','-')}"),
        row2("Pekerjaan", data.get('pekerjaan_lk','-'), "Pekerjaan", data.get('pekerjaan_pr','-')),
        row2("Pendidikan", data.get('pendidikan_lk','-'), "Pendidikan", data.get('pendidikan_pr','-')),
        row2("Ex Pasangan", data.get('istri_terdahulu','-'), "Ex Pasangan", data.get('suami_terdahulu','-')),
        row2("Alamat Lengkap", data.get('alamat_lk','-'), "Alamat Lengkap", data.get('alamat_pr','-')),
    ]

    tabel_ortu_data = [
        [Paragraph("IV. ORANG TUA LAKI-LAKI", sec_title_style), "", "", Paragraph("V. ORANG TUA PEREMPUAN", sec_title_style), "", ""],
        row2("Ayah / Bin", f"{data.get('nama_ayah_lk','-')} bin {data.get('bin_ayah_lk','-')}", "Ayah / Bin", f"{data.get('nama_ayah_pr','-')} bin {data.get('bin_ayah_pr','-')}"),
        row2("NIK / TTL Ayah", f"{data.get('nik_ayah_lk','-')} / {data.get('ttl_ayah_lk','-')}", "NIK / TTL Ayah", f"{data.get('nik_ayah_pr','-')} / {data.get('ttl_ayah_pr','-')}"),
        row2("Pekerjaan Ayah", data.get('pekerjaan_ayah_lk','-'), "Pekerjaan Ayah", data.get('pekerjaan_ayah_pr','-')),
        row2("Alamat Ayah", data.get('alamat_ayah_lk','-'), "Alamat Ayah", data.get('alamat_ayah_pr','-')),
        row2("Ibu / Binti", f"{data.get('nama_ibu_lk','-')} bin {data.get('bin_ibu_lk','-')}", "Ibu / Binti", f"{data.get('nama_ibu_pr','-')} bin {data.get('bin_ibu_pr','-')}"),
        row2("NIK / TTL Ibu", f"{data.get('nik_ibu_lk','-')} / {data.get('ttl_ibu_lk','-')}", "NIK / TTL Ibu", f"{data.get('nik_ibu_pr','-')} / {data.get('ttl_ibu_pr','-')}"),
        row2("Pekerjaan Ibu", data.get('pekerjaan_ibu_lk','-'), "Pekerjaan Ibu", data.get('pekerjaan_ibu_pr','-')),
        row2("Alamat Ibu", data.get('alamat_ibu_lk','-'), "Alamat Ibu", data.get('alamat_ibu_pr','-')),
    ]

    tabel_wali_saksi = [
        [Paragraph("VI. DATA WALI NIKAH", sec_title_style), "", "", Paragraph("VII. DATA SAKSI-SAKSI AKAD", sec_title_style), "", ""],
        row2("Nama Wali", f"{data.get('nama_wali','-')} bin {data.get('bin_wali','-')}", "Saksi 1", data.get('saksi1_nama','-')),
        row2("NIK / TTL Wali", f"{data.get('nik_wali','-')} / {data.get('ttl_wali','-')}", "NIK / TTL Saksi 1", f"{data.get('saksi1_nik','-')} / {data.get('saksi1_ttl','-')}"),
        row2("Hubungan Wali", data.get('hubungan_wali','-'), "Pekerjaan Saksi 1", data.get('saksi1_pekerjaan','-')),
        row2("Pekerjaan Wali", data.get('pekerjaan_wali','-'), "Alamat Saksi 1", data.get('saksi1_alamat','-')),
        row2("Alamat Wali", data.get('alamat_wali','-'), "Saksi 2", data.get('saksi2_nama','-')),
        row2("Wali Lengkap", data.get('nama_wali_lengkap','-'), "NIK / TTL Saksi 2", f"{data.get('saksi2_nik','-')} / {data.get('saksi2_ttl','-')}"),
        row2("", "", "Pekerjaan Saksi 2", data.get('saksi2_pekerjaan','-')),
        row2("", "", "Alamat Saksi 2", data.get('saksi2_alamat','-')),
    ]

    style_table = TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0.3),
        ('TOPPADDING', (0,0), (-1,-1), 0.3),
        ('LEFTPADDING', (0,0), (-1,-1), 1),
        ('RIGHTPADDING', (0,0), (-1,-1), 1),
        ('LINEBELOW', (0,0), (-1,0), 0.5, colors.HexColor('#CCCCCC')),
    ])

    t1 = Table(tabel_akad_data, colWidths=[34*mm, 3*mm, 164*mm])
    t1.setStyle(style_table)
    
    t2 = Table(tabel_catin_data, colWidths=[26*mm, 3*mm, 71*mm, 26*mm, 3*mm, 71*mm])
    t2.setStyle(style_table)

    t3 = Table(tabel_ortu_data, colWidths=[26*mm, 3*mm, 71*mm, 26*mm, 3*mm, 71*mm])
    t3.setStyle(style_table)

    t4 = Table(tabel_wali_saksi, colWidths=[26*mm, 3*mm, 71*mm, 26*mm, 3*mm, 71*mm])
    t4.setStyle(style_table)

    elements.extend([
        t1, Spacer(1, 2), 
        t2, Spacer(1, 2), 
        t3, Spacer(1, 2), 
        t4, Spacer(1, 6)
    ])
    
    ttd_catin = [
        [Paragraph("Catin Laki-Laki", lbl_style), Paragraph("Catin Perempuan", lbl_style), Paragraph("Wali Nikah", lbl_style)],
        [Spacer(1, 14), Spacer(1, 14), Spacer(1, 14)],
        [
            Paragraph(f"( <b>{data.get('nama_lk','...')}</b> )", val_style), 
            Paragraph(f"( <b>{data.get('nama_pr','...')}</b> )", val_style), 
            Paragraph(f"( <b>{data.get('nama_wali','...')}</b> )", val_style)
        ]
    ]
    t_ttd1 = Table(ttd_catin, colWidths=[67*mm, 67*mm, 67*mm])
    t_ttd1.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    
    tgl_surat_str = data.get('tgl_surat', 'Tambi, ................. 2026')
    ttd_pemdes = [
        [
            Paragraph("Mengetahui,<br/><b>KEPALA DESA TAMBI</b>", lbl_style), 
            Paragraph("", lbl_style), 
            Paragraph(f"{tgl_surat_str}<br/><b>KASI PELAYANAN DESA TAMBI</b>", lbl_style)
        ],
        [Spacer(1, 16), Spacer(1, 16), Spacer(1, 16)],
        [
            Paragraph("<b><u>JURI</u></b>", val_style), 
            Paragraph("", val_style), 
            Paragraph("<b><u>CHALIM MUCHTAROM, S.Pd.I</u></b>", val_style)
        ]
    ]
    t_ttd2 = Table(ttd_pemdes, colWidths=[80*mm, 41*mm, 80*mm])
    t_ttd2.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))

    elements.extend([t_ttd1, Spacer(1, 5), t_ttd2])
    
    doc.build(elements)
    pdf_data = buffer.getvalue()
    buffer.close()
    return pdf_data


draft = load_draft()

st.title("📜 Form Input Berkas Catin Desa Tambi")

# --- TOMBOL RESET ---
col_top1, col_top2 = st.columns([1, 4])
with col_top1:
    if st.button("🔄 Reset Form / Hapus Draf"):
        if os.path.exists(DRAFT_FILE):
            os.remove(DRAFT_FILE)
        st.session_state.clear()
        st.rerun()

# --- FORM INPUT ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📝 Register & Akad",
    "👨 Catin Laki-Laki & Ortu",
    "👩 Catin Perempuan & Ortu",
    "🤝 Data Wali",
    "👥 Data Saksi 1 & 2"
])

with tab1:
    st.subheader("Surat & Pelaksanaan Akad Nikah")
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Nomor Register", value=draft.get("no_register", "400.12.3.2/010/ VIII/ 2026"), key="no_register")
        st.text_input("Tanggal Surat", value=draft.get("tgl_surat", "TAMBI, 11 AGUSTUS 2026"), key="tgl_surat")
        
        tgl_default = date(2026, 9, 7)
        if "tgl_pelaksanaan" in draft:
            try:
                tgl_default = date.fromisoformat(draft["tgl_pelaksanaan"])
            except Exception:
                pass
        st.date_input("Tanggal Pelaksanaan Akad", value=tgl_default, key="tgl_pelaksanaan")
        st.text_input("Jam Akad", value=draft.get("jam_akad", "JAM. 08.00"), key="jam_akad")
    with col2:
        st.text_input("Tempat Akad Nikah", value=draft.get("tempat_akad", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="tempat_akad")
        st.text_input("Email Catin", value=draft.get("email_catin", ""), key="email_catin")
        st.text_input("Maskawin / Mahar", value=draft.get("mahar", "Seperangkat Alat Sholat"), key="mahar")

with tab2:
    st.subheader("Data Calon Pengantin Laki-Laki")
    col_lk1, col_lk2 = st.columns(2)
    with col_lk1:
        st.text_input("Nama Catin Laki-Laki", value=draft.get("nama_lk", "Miftahul Anam"), key="nama_lk")
        st.text_input("Bin (Ayah Laki-Laki)", value=draft.get("bin_lk", "Nur Karim"), key="bin_lk")
        
        ttl_lk = st.text_input("TTL Laki-Laki", value=draft.get("ttl_lk", "Pemalang, 18 Februari 1999"), key="ttl_lk")
        st.info(f"💡 Umur Catin LK: **{hitung_umur(ttl_lk)} Tahun**")
        
        st.text_input("NIK Laki-Laki", value=draft.get("nik_lk", "3327031802990004"), key="nik_lk")
    with col_lk2:
        st.text_input("Pekerjaan Laki-Laki", value=draft.get("pekerjaan_lk", "Swasta"), key="pekerjaan_lk")
        st.text_input("Status Laki-Laki", value=draft.get("status_lk", "BELUM KAWIN"), key="status_lk")
        st.text_input("Jenis Kelamin Laki-Laki", value=draft.get("jk_lk", "Laki-Laki"), key="jk_lk")
        st.text_input("Nama Istri Terdahulu", value=draft.get("istri_terdahulu", ""), key="istri_terdahulu")
        st.text_area("Alamat Laki-Laki", value=draft.get("alamat_lk", "RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang"), key="alamat_lk")
        st.text_input("Pendidikan Laki-Laki", value=draft.get("pendidikan_lk", "SLTA"), key="pendidikan_lk")

    st.divider()
    st.subheader("Data Ayah & Ibu Laki-Laki")
    col_alk, col_ilk = st.columns(2)
    with col_alk:
        st.markdown("**Ayah Laki-Laki**")
        st.text_input("Nama Ayah Laki-Laki", value=draft.get("nama_ayah_lk", "Nur Karim"), key="nama_ayah_lk")
        st.text_input("bin Ayah LK", value=draft.get("bin_ayah_lk", "Kasturi"), key="bin_ayah_lk")
        st.text_input("NIK Ayah LK", value=draft.get("nik_ayah_lk", "3327030608680006"), key="nik_ayah_lk")
        
        ttl_ayah_lk = st.text_input("TTL Ayah LK", value=draft.get("ttl_ayah_lk", "Pemalang, 06 Agustus 1968"), key="ttl_ayah_lk")
        st.caption(f"Umur Ayah LK: {hitung_umur(ttl_ayah_lk)} Tahun")
        
        st.text_input("Pekerjaan Ayah LK", value=draft.get("pekerjaan_ayah_lk", "PETANI/ PEKEBUN"), key="pekerjaan_ayah_lk")
        st.text_area("Alamat Ayah LK", value=draft.get("alamat_ayah_lk", "RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang"), key="alamat_ayah_lk")

    with col_ilk:
        st.markdown("**Ibu Laki-Laki**")
        st.text_input("Nama Ibu Laki-Laki", value=draft.get("nama_ibu_lk", "Samijah"), key="nama_ibu_lk")
        st.text_input("bin Ibu LK", value=draft.get("bin_ibu_lk", "Taryad"), key="bin_ibu_lk")
        st.text_input("NIK Ibu LK", value=draft.get("nik_ibu_lk", "3327035405740004"), key="nik_ibu_lk")
        
        ttl_ibu_lk = st.text_input("TTL Ibu LK", value=draft.get("ttl_ibu_lk", "Pemalang, 14 Mei 1974"), key="ttl_ibu_lk")
        st.caption(f"Umur Ibu LK: {hitung_umur(ttl_ibu_lk)} Tahun")
        
        st.text_input("Pekerjaan Ibu LK", value=draft.get("pekerjaan_ibu_lk", "Mengurus Rumah Tangga"), key="pekerjaan_ibu_lk")
        st.text_area("Alamat Ibu LK", value=draft.get("alamat_ibu_lk", "RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang"), key="alamat_ibu_lk")

with tab3:
    st.subheader("Data Calon Pengantin Perempuan")
    col_pr1, col_pr2 = st.columns(2)
    with col_pr1:
        st.text_input("Nama Catin Perempuan", value=draft.get("nama_pr", "Diyan Solehatin"), key="nama_pr")
        st.text_input("Binti (Ayah Perempuan)", value=draft.get("binti_pr", "Disun"), key="binti_pr")
        
        ttl_pr = st.text_input("TTL Perempuan", value=draft.get("ttl_pr", "Pemalang, 29 Juni 2007"), key="ttl_pr")
        st.info(f"💡 Umur Catin PR: **{hitung_umur(ttl_pr)} Tahun**")
        
        st.text_input("NIK Perempuan", value=draft.get("nik_pr", "3327046906070010"), key="nik_pr")
    with col_pr2:
        st.text_input("Pekerjaan Perempuan", value=draft.get("pekerjaan_pr", "BELUM/ TIDAK BEKERJA"), key="pekerjaan_pr")
        st.text_input("Status Perempuan", value=draft.get("status_pr", "BELUM KAWIN"), key="status_pr")
        st.text_input("Jenis Kelamin Perempuan", value=draft.get("jk_pr", "PEREMPUAN"), key="jk_pr")
        st.text_input("Nama Suami Terdahulu", value=draft.get("suami_terdahulu", ""), key="suami_terdahulu")
        st.text_area("Alamat Perempuan", value=draft.get("alamat_pr", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="alamat_pr")
        st.text_input("Pendidikan Perempuan", value=draft.get("pendidikan_pr", "SLTP"), key="pendidikan_pr")

    st.divider()
    st.subheader("Data Ayah & Ibu Perempuan")
    col_apr, col_ipr = st.columns(2)
    with col_apr:
        st.markdown("**Ayah Perempuan**")
        st.text_input("Nama Ayah PR", value=draft.get("nama_ayah_pr", "Disun"), key="nama_ayah_pr")
        st.text_input("bin Ayah PR", value=draft.get("bin_ayah_pr", "Tawiroji"), key="bin_ayah_pr")
        st.text_input("NIK Ayah PR", value=draft.get("nik_ayah_pr", "3327042504840003"), key="nik_ayah_pr")
        
        ttl_ayah_pr = st.text_input("TTL Ayah PR", value=draft.get("ttl_ayah_pr", "Pemalang, 21 April 1984"), key="ttl_ayah_pr")
        st.caption(f"Umur Ayah PR: {hitung_umur(ttl_ayah_pr)} Tahun")
        
        st.text_input("Pekerjaan Ayah PR", value=draft.get("pekerjaan_ayah_pr", "PETANI/ PEKEBUN"), key="pekerjaan_ayah_pr")
        st.text_area("Alamat Ayah PR", value=draft.get("alamat_ayah_pr", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="alamat_ayah_pr")

    with col_ipr:
        st.markdown("**Ibu Perempuan**")
        st.text_input("Nama Ibu PR", value=draft.get("nama_ibu_pr", "Mutirah"), key="nama_ibu_pr")
        st.text_input("bin Ibu PR", value=draft.get("bin_ibu_pr", "Tamiarjo"), key="bin_ibu_pr")
        st.text_input("NIK Ibu PR", value=draft.get("nik_ibu_pr", "3327044411840003"), key="nik_ibu_pr")
        
        ttl_ibu_pr = st.text_input("TTL Ibu PR", value=draft.get("ttl_ibu_pr", "Pemalang, 04 November 1984"), key="ttl_ibu_pr")
        st.caption(f"Umur Ibu PR: {hitung_umur(ttl_ibu_pr)} Tahun")
        
        st.text_input("Pekerjaan Ibu PR", value=draft.get("pekerjaan_ibu_pr", "Mengurus Rumah Tangga"), key="pekerjaan_ibu_pr")
        st.text_area("Alamat Ibu PR", value=draft.get("alamat_ibu_pr", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="alamat_ibu_pr")

with tab4:
    st.subheader("Data Wali Nikah")
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        st.text_input("Nama Wali", value=draft.get("nama_wali", "Disun"), key="nama_wali")
        st.text_input("Bin Wali", value=draft.get("bin_wali", "Tawiroji"), key="bin_wali")
        st.text_input("NIK Wali", value=draft.get("nik_wali", "3327042504840003"), key="nik_wali")
        
        ttl_wali = st.text_input("TTL Wali", value=draft.get("ttl_wali", "Pemalang, 21 April 1984"), key="ttl_wali")
        st.caption(f"Umur Wali: {hitung_umur(ttl_wali)} Tahun")
    with col_w2:
        st.text_input("Pekerjaan Wali", value=draft.get("pekerjaan_wali", "PETANI/ PEKEBUN"), key="pekerjaan_wali")
        st.text_area("Alamat Wali", value=draft.get("alamat_wali", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="alamat_wali")
        st.text_input("Hubungan Wali", value=draft.get("hubungan_wali", "AYAH KANDUNG"), key="hubungan_wali")
        st.text_input("Nama Wali Lengkap", value=draft.get("nama_wali_lengkap", "Disun Bin Tawiroji"), key="nama_wali_lengkap")

with tab5:
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        st.subheader("Data Saksi 1")
        st.text_input("Nama Saksi 1", value=draft.get("saksi1_nama", "Chalim Muchtarom"), key="saksi1_nama")
        st.text_input("TTL Saksi 1", value=draft.get("saksi1_ttl", "Pemalang, 21 Oktober 1989"), key="saksi1_ttl")
        st.text_input("NIK Saksi 1", value=draft.get("saksi1_nik", "3327042110890004"), key="saksi1_nik")
        st.text_input("Pekerjaan Saksi 1", value=draft.get("saksi1_pekerjaan", "Perangkat Desa"), key="saksi1_pekerjaan")
        st.text_area("Alamat Saksi 1", value=draft.get("saksi1_alamat", "RT 002 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="saksi1_alamat")

    with col_s2:
        st.subheader("Data Saksi 2")
        st.text_input("Nama Saksi 2", value=draft.get("saksi2_nama", "Sidin"), key="saksi2_nama")
        st.text_input("TTL Saksi 2", value=draft.get("saksi2_ttl", "Pemalang, 15 Mei 1980"), key="saksi2_ttl")
        st.text_input("NIK Saksi 2", value=draft.get("saksi2_nik", "3327041505800002"), key="saksi2_nik")
        st.text_input("Pekerjaan Saksi 2", value=draft.get("saksi2_pekerjaan", "Petani/Pekebun"), key="saksi2_pekerjaan")
        st.text_area("Alamat Saksi 2", value=draft.get("saksi2_alamat", "RT 003 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="saksi2_alamat")

# --- TOMBOL AKSI ---
st.divider()
col_act1, col_act2 = st.columns(2)

with col_act1:
    if st.button("💾 SIMPAN DRAF SEMENTARA", use_container_width=True):
        save_draft_file()
        st.toast("✅ Draf berhasil disimpan!", icon="💾")

with col_act2:
    submit = st.button("🚀 PROSES KE EXCEL & GENERATE BERKAS (PDF/EXCEL)", type="primary", use_container_width=True)

# --- PROSES GENERATE FILE ---
if submit:
    save_draft_file()
    
    with st.spinner("⏳ Memproses Excel & Membuat PDF F4 Lengkap dengan Pengesahan Kades & Kasi..."):
        data_dict = {key: st.session_state[key] for key in st.session_state}
        
        # 1. GENERATE EXCEL DENGAN MENULISKAN DATA KE CELL
        try:
            if os.path.exists(EXCEL_FILE):
                excel_bytes = update_excel_data(EXCEL_FILE, data_dict)
                st.session_state['excel_bytes'] = excel_bytes
            else:
                st.warning(f"File master Excel '{EXCEL_FILE}' tidak ditemukan di folder. Menyiapkan PDF saja.")
        except Exception as e:
            st.error(f"Error memproses Excel: {e}")
            
        # 2. GENERATE PDF F4
        try:
            pdf_bytes = generate_pdf_f4(data_dict)
            st.session_state['pdf_bytes'] = pdf_bytes
        except Exception as e:
            st.error(f"Error memproses PDF F4: {e}")
            
        st.session_state['is_processed'] = True

# --- MENAMPILKAN TOMBOL DOWNLOAD SETELAH PROSES ---
if st.session_state.get('is_processed', False):
    st.divider()
    st.success("🎉 Berkas Berhasil Diproses! Silakan unduh file di bawah ini:")
    
    nama_lk_file = st.session_state.get('nama_lk', 'Catin').replace(" ", "_")
    nama_pr_file = st.session_state.get('nama_pr', 'Catin').replace(" ", "_")
    
    col_dl1, col_dl2 = st.columns(2)
    
    # Tombol Download Excel
    with col_dl1:
        if 'excel_bytes' in st.session_state:
            filename_excel = f"BERKAS_CATIN_{nama_lk_file}_dan_{nama_pr_file}.xlsx"
            st.download_button(
                label="📊 UNDUH FILE EXCEL TERBARU",
                data=st.session_state['excel_bytes'],
                file_name=filename_excel,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
    # Tombol Download PDF F4
    with col_dl2:
        if 'pdf_bytes' in st.session_state:
            filename_pdf = f"RINGKASAN_CATIN_F4_{nama_lk_file}_dan_{nama_pr_file}.pdf"
            st.download_button(
                label="📄 UNDUH PDF F4 (1 LEMBAR UTUH)",
                data=st.session_state['pdf_bytes'],
                file_name=filename_pdf,
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
