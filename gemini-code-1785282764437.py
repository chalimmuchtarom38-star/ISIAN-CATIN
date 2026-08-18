import os
import re
from datetime import datetime, date
import streamlit as st
import openpyxl

# Import ReportLab untuk Export PDF
from reportlab.lib.pagesizes import mm
from reportlab.lib.units import mm as mm_unit
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY

# ==========================================
# KONFIGURASI HALAMAN STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Sistem Berkas Catin Desa Tambi",
    page_icon="💍",
    layout="wide"
)

EXCEL_FILE = "BERKAS CATIN .xlsx"

# ==========================================
# FUNGSI BANTUAN (UTILITIES)
# ==========================================
HARI_INDONESIA = {
    'Monday': 'Senin',
    'Tuesday': 'Selasa',
    'Wednesday': 'Rabu',
    'Thursday': 'Kamis',
    'Friday': 'Jumat',
    'Saturday': 'Sabtu',
    'Sunday': 'Minggu'
}

def get_hari_tgl(val):
    """Mengubah nilai date/datetime/string ISO menjadi format: Hari, DD-MM-YYYY"""
    if not val:
        return ""
    if isinstance(val, str):
        try:
            val = datetime.strptime(val, "%Y-%m-%d").date()
        except ValueError:
            return val
    if isinstance(val, (date, datetime)):
        hari_en = val.strftime('%A')
        hari_id = HARI_INDONESIA.get(hari_en, hari_en)
        return f"{hari_id}, {val.strftime('%d-%m-%Y')}"
    return str(val)

def hitung_umur(ttl_str):
    """Mengekstrak tahun dari string TTL dan menghitung umur relatif terhadap tahun berjalan"""
    if not ttl_str:
        return ""
    matches = re.findall(r'\b(19\d{2}|20\d{2})\b', ttl_str)
    if matches:
        tahun_lahir = int(matches[-1])
        tahun_sekarang = datetime.now().year
        return str(tahun_sekarang - tahun_lahir)
    return ""

# ==========================================
# FUNGSI EXPORT EXCEL (openpyxl)
# ==========================================
def update_excel_data(data, master_path):
    """Memasukkan data dari form ke dalam template Excel 'BERKAS CATIN .xlsx'"""
    wb = openpyxl.load_workbook(master_path)
    ws = wb.active

    # Pemetaan Field ke Sel Excel
    mapping = {
        # CATIN PRIA
        'catin_pria_nama': 'C13',
        'catin_pria_nik': 'C14',
        'catin_pria_bin': 'C15',
        'catin_pria_ttl': 'C16',
        'catin_pria_kewarganegaraan': 'C17',
        'catin_pria_agama': 'C18',
        'catin_pria_pekerjaan': 'C19',
        'catin_pria_alamat': 'C20',

        # CATIN WANITA
        'catin_wanita_nama': 'F13',
        'catin_wanita_nik': 'F14',
        'catin_wanita_binti': 'F15',
        'catin_wanita_ttl': 'F16',
        'catin_wanita_kewarganegaraan': 'F17',
        'catin_wanita_agama': 'F18',
        'catin_wanita_pekerjaan': 'F19',
        'catin_wanita_alamat': 'F20',

        # WALI NIKAH
        'wali_nama': 'C26',
        'wali_nik': 'C27',
        'wali_bin': 'C28',
        'wali_ttl': 'C29',
        'wali_kewarganegaraan': 'C30',
        'wali_agama': 'C31',
        'wali_pekerjaan': 'C32',
        'wali_alamat': 'C33',
        'wali_hubungan': 'C34',

        # AYAH CATIN PRIA
        'ayah_pria_nama': 'C40',
        'ayah_pria_nik': 'C41',
        'ayah_pria_bin': 'C42',
        'ayah_pria_ttl': 'C43',
        'ayah_pria_kewarganegaraan': 'C44',
        'ayah_pria_agama': 'C45',
        'ayah_pria_pekerjaan': 'C46',
        'ayah_pria_alamat': 'C47',

        # IBU CATIN PRIA
        'ibu_pria_nama': 'F40',
        'ibu_pria_nik': 'F41',
        'ibu_pria_bin': 'F42',
        'ibu_pria_ttl': 'F43',
        'ibu_pria_kewarganegaraan': 'F44',
        'ibu_pria_agama': 'F45',
        'ibu_pria_pekerjaan': 'F46',
        'ibu_pria_alamat': 'F47',

        # AYAH CATIN WANITA
        'ayah_wanita_nama': 'C53',
        'ayah_wanita_nik': 'C54',
        'ayah_wanita_bin': 'C55',
        'ayah_wanita_ttl': 'C56',
        'ayah_wanita_kewarganegaraan': 'C57',
        'ayah_wanita_agama': 'C58',
        'ayah_wanita_pekerjaan': 'C59',
        'ayah_wanita_alamat': 'C60',

        # IBU CATIN WANITA
        'ibu_wanita_nama': 'F53',
        'ibu_wanita_nik': 'F54',
        'ibu_wanita_bin': 'F55',
        'ibu_wanita_ttl': 'F56',
        'ibu_wanita_kewarganegaraan': 'F57',
        'ibu_wanita_agama': 'F58',
        'ibu_wanita_pekerjaan': 'F59',
        'ibu_wanita_alamat': 'F60',

        # AKAD NIKAH & LAINNYA
        'hari_tgl_akad': 'C66',
        'waktu_akad': 'C67',
        'mas_kawin': 'C68',
        'tempat_akad': 'C69',
        'tgl_surat': 'C70',
        'status_pria': 'C71',
        'status_wanita': 'C72',
    }

    for key, cell in mapping.items():
        ws[cell] = data.get(key, "")

    # Hitung Umur Otomatis
    ws['C21'] = hitung_umur(data.get('catin_pria_ttl', ''))
    ws['F21'] = hitung_umur(data.get('catin_wanita_ttl', ''))

    output_path = "BERKAS_CATIN_TERISI.xlsx"
    wb.save(output_path)
    return output_path

# ==========================================
# FUNGSI EXPORT PDF F4 (ReportLab)
# ==========================================
def generate_pdf_f4(data, output_pdf_path):
    """Mencetak Lembar Verifikasi Data Catin Ukuran Folio/F4 (1 Halaman Presisi)"""
    # Ukuran F4 / Folio = 215mm x 330mm
    F4_SIZE = (215 * mm_unit, 330 * mm_unit)
    doc = SimpleDocTemplate(
        output_pdf_path,
        pagesize=F4_SIZE,
        leftMargin=15 * mm_unit,
        rightMargin=15 * mm_unit,
        topMargin=12 * mm_unit,
        bottomMargin=12 * mm_unit
    )

    story = []
    styles = getSampleStyleSheet()

    # Style Kustom
    title_style = ParagraphStyle('TitleStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, leading=15, alignment=TA_CENTER)
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=13, alignment=TA_CENTER)
    section_style = ParagraphStyle('SectionStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, leading=11, textColor=colors.HexColor('#1E3A8A'))
    cell_bold = ParagraphStyle('CellBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, leading=9.5)
    cell_text = ParagraphStyle('CellText', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=9.5)

    # 1. KOP / JUDUL
    story.append(Paragraph("PEMERINTAH KABUPATEN PEMALANG", title_style))
    story.append(Paragraph("KECAMATAN WATUKUMPUL - DESA TAMBI", title_style))
    story.append(Paragraph("LEMBAR VERIFIKASI BERKAS CALON PENGANTIN (CATIN)", subtitle_style))
    story.append(Spacer(1, 4 * mm_unit))

    # Helper Tabel 2 Kolom
    def make_section_table(title, items_left, items_right):
        rows = []
        rows.append([Paragraph(f"<b>{title}</b>", section_style), ""])
        max_rows = max(len(items_left), len(items_right))
        for i in range(max_rows):
            left_lbl, left_val = items_left[i] if i < len(items_left) else ("", "")
            right_lbl, right_val = items_right[i] if i < len(items_right) else ("", "")

            left_p = Paragraph(f"<b>{left_lbl}:</b> {left_val}" if left_lbl else "", cell_text)
            right_p = Paragraph(f"<b>{right_lbl}:</b> {right_val}" if right_lbl else "", cell_text)
            rows.append([left_p, right_p])

        t = Table(rows, colWidths=[92 * mm_unit, 92 * mm_unit])
        t.setStyle(TableStyle([
            ('SPAN', (0, 0), (1, 0)),
            ('BACKGROUND', (0, 0), (1, 0), colors.HexColor('#F1F5F9')),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1.5),
            ('TOPPADDING', (0, 0), (-1, -1), 1.5),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LINEBELOW', (0, 0), (1, 0), 0.5, colors.HexColor('#CBD5E1')),
        ]))
        return t

    # DATA CATIN
    c_pria = [
        ("Nama", data.get('catin_pria_nama', '')),
        ("NIK", data.get('catin_pria_nik', '')),
        ("Bin", data.get('catin_pria_bin', '')),
        ("TTL / Umur", f"{data.get('catin_pria_ttl', '')} ({hitung_umur(data.get('catin_pria_ttl', ''))} th)"),
        ("Agama / Kwn", f"{data.get('catin_pria_agama', '')} / {data.get('catin_pria_kewarganegaraan', '')}"),
        ("Pekerjaan", data.get('catin_pria_pekerjaan', '')),
        ("Status", data.get('status_pria', '')),
        ("Alamat", data.get('catin_pria_alamat', '')),
    ]
    c_wanita = [
        ("Nama", data.get('catin_wanita_nama', '')),
        ("NIK", data.get('catin_wanita_nik', '')),
        ("Binti", data.get('catin_wanita_binti', '')),
        ("TTL / Umur", f"{data.get('catin_wanita_ttl', '')} ({hitung_umur(data.get('catin_wanita_ttl', ''))} th)"),
        ("Agama / Kwn", f"{data.get('catin_wanita_agama', '')} / {data.get('catin_wanita_kewarganegaraan', '')}"),
        ("Pekerjaan", data.get('catin_wanita_pekerjaan', '')),
        ("Status", data.get('status_wanita', '')),
        ("Alamat", data.get('catin_wanita_alamat', '')),
    ]
    story.append(make_section_table("I. DATA CALON PENGANTIN (PRIA & WANITA)", c_pria, c_wanita))
    story.append(Spacer(1, 3 * mm_unit))

    # DATA ORANG TUA
    ortu_pria = [
        ("Ayah", data.get('ayah_pria_nama', '')),
        ("NIK / Bin", f"{data.get('ayah_pria_nik', '')} / {data.get('ayah_pria_bin', '')}"),
        ("TTL / Agama", f"{data.get('ayah_pria_ttl', '')} / {data.get('ayah_pria_agama', '')}"),
        ("Ibu", data.get('ibu_pria_nama', '')),
        ("NIK / Bin", f"{data.get('ibu_pria_nik', '')} / {data.get('ibu_pria_bin', '')}"),
        ("TTL / Agama", f"{data.get('ibu_pria_ttl', '')} / {data.get('ibu_pria_agama', '')}"),
    ]
    ortu_wanita = [
        ("Ayah", data.get('ayah_wanita_nama', '')),
        ("NIK / Bin", f"{data.get('ayah_wanita_nik', '')} / {data.get('ayah_wanita_bin', '')}"),
        ("TTL / Agama", f"{data.get('ayah_wanita_ttl', '')} / {data.get('ayah_wanita_agama', '')}"),
        ("Ibu", data.get('ibu_wanita_nama', '')),
        ("NIK / Bin", f"{data.get('ibu_wanita_nik', '')} / {data.get('ibu_wanita_bin', '')}"),
        ("TTL / Agama", f"{data.get('ibu_wanita_ttl', '')} / {data.get('ibu_wanita_agama', '')}"),
    ]
    story.append(make_section_table("II. DATA ORANG TUA (AYAH & IBU)", ortu_pria, ortu_wanita))
    story.append(Spacer(1, 3 * mm_unit))

    # WALI & AKAD NIKAH
    wali_info = [
        ("Nama Wali", data.get('wali_nama', '')),
        ("NIK / Bin", f"{data.get('wali_nik', '')} / {data.get('wali_bin', '')}"),
        ("TTL / Hubungan", f"{data.get('wali_ttl', '')} / {data.get('wali_hubungan', '')}"),
        ("Pekerjaan / Alamat", f"{data.get('wali_pekerjaan', '')} - {data.get('wali_alamat', '')}"),
    ]
    akad_info = [
        ("Hari / Tgl Akad", data.get('hari_tgl_akad', '')),
        ("Waktu Akad", data.get('waktu_akad', '')),
        ("Tempat Akad", data.get('tempat_akad', '')),
        ("Mas Kawin", data.get('mas_kawin', '')),
        ("Tanggal Surat", data.get('tgl_surat', '')),
    ]
    story.append(make_section_table("III. DATA WALI NIKAH & RENCANA AKAD", wali_info, akad_info))
    story.append(Spacer(1, 5 * mm_unit))

    # TANDA TANGAN (4 Blok)
    p_center = ParagraphStyle('PCenter', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=9.5, alignment=TA_CENTER)
    p_center_bold = ParagraphStyle('PCenterBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, leading=9.5, alignment=TA_CENTER)

    tgl_srt = data.get('tgl_surat', '....................')
    ttd_data = [
        [
            Paragraph(f"Tambi, {tgl_srt}<br/>Calon Pengantin Pria", p_center),
            Paragraph(f"Tambi, {tgl_srt}<br/>Calon Pengantin Wanita", p_center),
            Paragraph(f"Tambi, {tgl_srt}<br/>Wali Nikah", p_center),
        ],
        [Paragraph("", p_center), Paragraph("", p_center), Paragraph("", p_center)],
        [
            Paragraph(f"<u><b>{data.get('catin_pria_nama', '....................')}</b></u>", p_center_bold),
            Paragraph(f"<u><b>{data.get('catin_wanita_nama', '....................')}</b></u>", p_center_bold),
            Paragraph(f"<u><b>{data.get('wali_nama', '....................')}</b></u>", p_center_bold),
        ],
        [Paragraph("<br/>Mengetahui,<br/>Kasi Pelayanan Desa Tambi", p_center), Paragraph("", p_center), Paragraph("<br/>Mengetahui,<br/>Kepala Desa Tambi", p_center)],
        [Paragraph("", p_center), Paragraph("", p_center), Paragraph("", p_center)],
        [
            Paragraph("<u><b>....................................</b></u>", p_center_bold),
            Paragraph("", p_center),
            Paragraph("<u><b>KASNO</b></u>", p_center_bold),
        ]
    ]

    t_ttd = Table(ttd_data, colWidths=[61 * mm_unit, 61 * mm_unit, 62 * mm_unit])
    t_ttd.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('HEIGHT', (0, 1), (-1, 1), 12 * mm_unit),
        ('HEIGHT', (0, 4), (-1, 4), 12 * mm_unit),
    ]))
    story.append(t_ttd)

    doc.build(story)

# ==========================================
# INTERFACE UTAMA (STREAMLIT UI)
# ==========================================
st.title("💍 Pengelolaan Berkas Catin Desa Tambi")
st.markdown("Isi formulir di bawah ini untuk memperbarui berkas master Excel dan mencetak lembar verifikasi PDF F4.")

with st.form("catin_form"):
    # TAB 1: CATIN
    tab1, tab2, tab3, tab4 = st.tabs(["🤵👰 Catin Pria & Wanita", "👨‍👩‍👦 Orang Tua", "🕌 Wali & Akad", "📌 Status & Dokumen"])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Data Calon Pengantin Pria")
            catin_pria_nama = st.text_input("Nama Lengkap Pria")
            catin_pria_nik = st.text_input("NIK Pria")
            catin_pria_bin = st.text_input("Bin (Nama Ayah Pria)")
            catin_pria_ttl = st.text_input("TTL Pria", placeholder="Pemalang, 12 Mei 1998")
            catin_pria_kewarganegaraan = st.text_input("Kewarganegaraan Pria", value="WNI")
            catin_pria_agama = st.text_input("Agama Pria", value="Islam")
            catin_pria_pekerjaan = st.text_input("Pekerjaan Pria")
            catin_pria_alamat = st.text_area("Alamat Pria", height=80)

        with col2:
            st.subheader("Data Calon Pengantin Wanita")
            catin_wanita_nama = st.text_input("Nama Lengkap Wanita")
            catin_wanita_nik = st.text_input("NIK Wanita")
            catin_wanita_binti = st.text_input("Binti (Nama Ayah Wanita)")
            catin_wanita_ttl = st.text_input("TTL Wanita", placeholder="Pemalang, 20 Agustus 2001")
            catin_wanita_kewarganegaraan = st.text_input("Kewarganegaraan Wanita", value="WNI")
            catin_wanita_agama = st.text_input("Agama Wanita", value="Islam")
            catin_wanita_pekerjaan = st.text_input("Pekerjaan Wanita")
            catin_wanita_alamat = st.text_area("Alamat Wanita", height=80)

    # TAB 2: ORANG TUA
    with tab2:
        col_o1, col_o2 = st.columns(2)
        with col_o1:
            st.subheader("Orang Tua Catin Pria")
            st.markdown("**Ayah Pria:**")
            ayah_pria_nama = st.text_input("Nama Ayah Pria")
            ayah_pria_nik = st.text_input("NIK Ayah Pria")
            ayah_pria_bin = st.text_input("Bin Ayah Pria")
            ayah_pria_ttl = st.text_input("TTL Ayah Pria")
            ayah_pria_kewarganegaraan = st.text_input("Kewarganegaraan Ayah Pria", value="WNI")
            ayah_pria_agama = st.text_input("Agama Ayah Pria", value="Islam")
            ayah_pria_pekerjaan = st.text_input("Pekerjaan Ayah Pria")
            ayah_pria_alamat = st.text_area("Alamat Ayah Pria", height=60)

            st.markdown("**Ibu Pria:**")
            ibu_pria_nama = st.text_input("Nama Ibu Pria")
            ibu_pria_nik = st.text_input("NIK Ibu Pria")
            ibu_pria_bin = st.text_input("Binti Ibu Pria")
            ibu_pria_ttl = st.text_input("TTL Ibu Pria")
            ibu_pria_kewarganegaraan = st.text_input("Kewarganegaraan Ibu Pria", value="WNI")
            ibu_pria_agama = st.text_input("Agama Ibu Pria", value="Islam")
            ibu_pria_pekerjaan = st.text_input("Pekerjaan Ibu Pria")
            ibu_pria_alamat = st.text_area("Alamat Ibu Pria", height=60)

        with col_o2:
            st.subheader("Orang Tua Catin Wanita")
            st.markdown("**Ayah Wanita:**")
            ayah_wanita_nama = st.text_input("Nama Ayah Wanita")
            ayah_wanita_nik = st.text_input("NIK Ayah Wanita")
            ayah_wanita_bin = st.text_input("Bin Ayah Wanita")
            ayah_wanita_ttl = st.text_input("TTL Ayah Wanita")
            ayah_wanita_kewarganegaraan = st.text_input("Kewarganegaraan Ayah Wanita", value="WNI")
            ayah_wanita_agama = st.text_input("Agama Ayah Wanita", value="Islam")
            ayah_wanita_pekerjaan = st.text_input("Pekerjaan Ayah Wanita")
            ayah_wanita_alamat = st.text_area("Alamat Ayah Wanita", height=60)

            st.markdown("**Ibu Wanita:**")
            ibu_wanita_nama = st.text_input("Nama Ibu Wanita")
            ibu_wanita_nik = st.text_input("NIK Ibu Wanita")
            ibu_wanita_bin = st.text_input("Binti Ibu Wanita")
            ibu_wanita_ttl = st.text_input("TTL Ibu Wanita")
            ibu_wanita_kewarganegaraan = st.text_input("Kewarganegaraan Ibu Wanita", value="WNI")
            ibu_wanita_agama = st.text_input("Agama Ibu Wanita", value="Islam")
            ibu_wanita_pekerjaan = st.text_input("Pekerjaan Ibu Wanita")
            ibu_wanita_alamat = st.text_area("Alamat Ibu Wanita", height=60)

    # TAB 3: WALI & AKAD
    with tab3:
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            st.subheader("Data Wali Nikah")
            wali_nama = st.text_input("Nama Wali")
            wali_nik = st.text_input("NIK Wali")
            wali_bin = st.text_input("Bin Wali")
            wali_ttl = st.text_input("TTL Wali")
            wali_kewarganegaraan = st.text_input("Kewarganegaraan Wali", value="WNI")
            wali_agama = st.text_input("Agama Wali", value="Islam")
            wali_pekerjaan = st.text_input("Pekerjaan Wali")
            wali_alamat = st.text_area("Alamat Wali", height=60)
            wali_hubungan = st.text_input("Hubungan Wali", placeholder="Ayah Kandung / Paman / Kakak")

        with col_w2:
            st.subheader("Rencana Akad Nikah")
            hari_tgl_akad_input = st.date_input("Tanggal Akad Nikah", value=date.today())
            waktu_akad = st.text_input("Waktu Akad Nikah", value="09:00 WIB")
            mas_kawin = st.text_input("Mas Kawin / Mahar", placeholder="Uang Tunai Rp 500.000,- dibayar tunai")
            tempat_akad = st.text_input("Tempat Akad Nikah", value="KUA Kec. Watukumpul / Rumah Catin Wanita")

    # TAB 4: STATUS & DOKUMEN
    with tab4:
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            st.subheader("Status & Tanggal Surat")
            status_pria = st.selectbox("Status Catin Pria", ["Perjaka", "Duda Mati", "Duda Cerai"])
            status_wanita = st.selectbox("Status Catin Wanita", ["Perawan", "Janda Mati", "Janda Cerai"])
            tgl_surat_input = st.date_input("Tanggal Penerbitan Surat", value=date.today())

    st.markdown("---")
    submit = st.form_submit_button("🚀 Proses & Generate Berkas (Excel & PDF)")

# ==========================================
# EKSEKUSI SETELAH SUBMIT
# ==========================================
if submit:
    hari_tgl_akad_str = get_hari_tgl(hari_tgl_akad_input)
    tgl_surat_str = get_hari_tgl(tgl_surat_input)

    form_data = {
        'catin_pria_nama': catin_pria_nama,
        'catin_pria_nik': catin_pria_nik,
        'catin_pria_bin': catin_pria_bin,
        'catin_pria_ttl': catin_pria_ttl,
        'catin_pria_kewarganegaraan': catin_pria_kewarganegaraan,
        'catin_pria_agama': catin_pria_agama,
        'catin_pria_pekerjaan': catin_pria_pekerjaan,
        'catin_pria_alamat': catin_pria_alamat,

        'catin_wanita_nama': catin_wanita_nama,
        'catin_wanita_nik': catin_wanita_nik,
        'catin_wanita_binti': catin_wanita_binti,
        'catin_wanita_ttl': catin_wanita_ttl,
        'catin_wanita_kewarganegaraan': catin_wanita_kewarganegaraan,
        'catin_wanita_agama': catin_wanita_agama,
        'catin_wanita_pekerjaan': catin_wanita_pekerjaan,
        'catin_wanita_alamat': catin_wanita_alamat,

        'wali_nama': wali_nama,
        'wali_nik': wali_nik,
        'wali_bin': wali_bin,
        'wali_ttl': wali_ttl,
        'wali_kewarganegaraan': wali_kewarganegaraan,
        'wali_agama': wali_agama,
        'wali_pekerjaan': wali_pekerjaan,
        'wali_alamat': wali_alamat,
        'wali_hubungan': wali_hubungan,

        'ayah_pria_nama': ayah_pria_nama,
        'ayah_pria_nik': ayah_pria_nik,
        'ayah_pria_bin': ayah_pria_bin,
        'ayah_pria_ttl': ayah_pria_ttl,
        'ayah_pria_kewarganegaraan': ayah_pria_kewarganegaraan,
        'ayah_pria_agama': ayah_pria_agama,
        'ayah_pria_pekerjaan': ayah_pria_pekerjaan,
        'ayah_pria_alamat': ayah_pria_alamat,

        'ibu_pria_nama': ibu_pria_nama,
        'ibu_pria_nik': ibu_pria_nik,
        'ibu_pria_bin': ibu_pria_bin,
        'ibu_pria_ttl': ibu_pria_ttl,
        'ibu_pria_kewarganegaraan': ibu_pria_kewarganegaraan,
        'ibu_pria_agama': ibu_pria_agama,
        'ibu_pria_pekerjaan': ibu_pria_pekerjaan,
        'ibu_pria_alamat': ibu_pria_alamat,

        'ayah_wanita_nama': ayah_wanita_nama,
        'ayah_wanita_nik': ayah_wanita_nik,
        'ayah_wanita_bin': ayah_wanita_bin,
        'ayah_wanita_ttl': ayah_wanita_ttl,
        'ayah_wanita_kewarganegaraan': ayah_wanita_kewarganegaraan,
        'ayah_wanita_agama': ayah_wanita_agama,
        'ayah_wanita_pekerjaan': ayah_wanita_pekerjaan,
        'ayah_wanita_alamat': ayah_wanita_alamat,

        'ibu_wanita_nama': ibu_wanita_nama,
        'ibu_wanita_nik': ibu_wanita_nik,
        'ibu_wanita_bin': ibu_wanita_bin,
        'ibu_wanita_ttl': ibu_wanita_ttl,
        'ibu_wanita_kewarganegaraan': ibu_wanita_kewarganegaraan,
        'ibu_wanita_agama': ibu_wanita_agama,
        'ibu_wanita_pekerjaan': ibu_wanita_pekerjaan,
        'ibu_wanita_alamat': ibu_wanita_alamat,

        'hari_tgl_akad': hari_tgl_akad_str,
        'waktu_akad': waktu_akad,
        'mas_kawin': mas_kawin,
        'tempat_akad': tempat_akad,
        'tgl_surat': tgl_surat_str,
        'status_pria': status_pria,
        'status_wanita': status_wanita,
    }

    st.success("✅ Data berhasil diproses!")

    # 1. GENERATE EXCEL
    if os.path.exists(EXCEL_FILE):
        try:
            excel_out = update_excel_data(form_data, EXCEL_FILE)
            with open(excel_out, "rb") as f:
                st.download_button(
                    label="📊 Unduh File Excel Terisi (.xlsx)",
                    data=f,
                    file_name=f"BERKAS_CATIN_{catin_pria_nama.replace(' ', '_')}_{catin_wanita_nama.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Gagal mengupdate file Excel: {e}")
    else:
        st.warning(f"File master '{EXCEL_FILE}' tidak ditemukan di direktori root.")

    # 2. GENERATE PDF F4
    try:
        pdf_out = "BERKAS_VERIFIKASI_CATIN_F4.pdf"
        generate_pdf_f4(form_data, pdf_out)
        with open(pdf_out, "rb") as f_pdf:
            st.download_button(
                label="📄 Unduh Lembar Verifikasi PDF F4 (.pdf)",
                data=f_pdf,
                file_name=f"VERIFIKASI_CATIN_{catin_pria_nama.replace(' ', '_')}_{catin_wanita_nama.replace(' ', '_')}.pdf",
                mime="application/pdf"
            )
    except Exception as e:
        st.error(f"Gagal membuat file PDF: {e}")
