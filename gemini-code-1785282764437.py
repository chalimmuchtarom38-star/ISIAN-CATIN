import streamlit as st
import openpyxl
from io import BytesIO
from datetime import date

# Library ReportLab untuk Pembuatan PDF Ukuran F4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Ukuran Kertas F4 / Folio dalam Points (215.9 mm x 330 mm)
F4_WIDTH = 215.9 * 2.83465
F4_HEIGHT = 330.0 * 2.83465
F4_SIZE = (F4_WIDTH, F4_HEIGHT)

st.set_page_config(
    page_title="Aplikasi Berkas Catin - Desa Tambi",
    page_icon="📜",
    layout="wide"
)

st.title("📜 Form Input & Cetak PDF ISIAN DATA Catin")
st.caption("Memproses dan mengunduh seluruh isi ringkasan 'ISIAN DATA' dalam format PDF ukuran F4/Folio tanpa ada yang terbuang.")

EXCEL_FILE = "BERKAS CATIN .xlsx"

# --------------------------------------------------
# FUNGSI MEMBUAT PDF SELURUH 'ISIAN DATA' UKURAN F4
# --------------------------------------------------
def generate_pdf_isian_data(d):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=F4_SIZE,
        leftMargin=28,
        rightMargin=28,
        topMargin=25,
        bottomMargin=25
    )
    
    styles = getSampleStyleSheet()
    
    style_kop = ParagraphStyle('KopHead', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, alignment=1, leading=13)
    style_title = ParagraphStyle('TitlePDF', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, alignment=1, leading=14)
    style_sec = ParagraphStyle('SecHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9.5, leading=11, textColor=colors.HexColor('#1A365D'))
    style_lbl = ParagraphStyle('Lbl', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=10.5)
    style_bold = ParagraphStyle('LblB', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=10.5)
    style_val = ParagraphStyle('Val', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=10.5)

    elements = []

    # KOP SURAT & JUDUL
    elements.append(Paragraph("PEMERINTAH KABUPATEN PEMALANG - KECAMATAN WATUKUMPUL", style_kop))
    elements.append(Paragraph("PEMERINTAH DESA TAMBI", style_kop))
    elements.append(Spacer(1, 4))
    elements.append(Paragraph("<u>RINGKASAN ISIAN DATA BERKAS CATIN</u>", style_title))
    elements.append(Spacer(1, 6))

    # HELPER TABLE BUILDER
    def build_table(data_rows, col_widths=[20, 150, 10, 350]):
        table_content = []
        for row in data_rows:
            no_str = f"{row[0]}." if row[0] else ""
            table_content.append([
                Paragraph(no_str, style_lbl),
                Paragraph(row[1], style_lbl),
                Paragraph(":", style_lbl),
                Paragraph(str(row[2]) if row[2] else "-", style_bold if len(row) > 3 and row[3] else style_val)
            ])
        t = Table(table_content, colWidths=col_widths)
        t.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('TOPPADDING', (0,0), (-1,-1), 1),
            ('BOTTOMPADDING', (0,0), (-1,-1), 1),
            ('LEFTPADDING', (0,0), (-1,-1), 2),
            ('RIGHTPADDING', (0,0), (-1,-1), 2),
        ]))
        return t

    # SECTION 1: REGISTER & PELAKSANAAN
    elements.append(Paragraph("A. REGISTER & PELAKSANAAN AKAD NIKAH", style_sec))
    reg_rows = [
        ["", "Nomor Register", d['no_register']],
        ["", "Tanggal Surat", d['tgl_surat']],
        ["", "Tanggal Pelaksanaan", f"{d['tgl_pelaksanaan']} (Jam: {d['jam_akad']})"],
        ["", "Tempat Akad Nikah", d['tempat_akad']],
        ["", "Email Catin", d['email_catin']],
        ["", "Maskawin / Mahar", d['mahar'], True]
    ]
    elements.append(build_table(reg_rows))
    elements.append(Spacer(1, 4))

    # SECTION 2: CATIN LAKI-LAKI
    elements.append(Paragraph("B. DATA CALON PENGANTIN LAKI-LAKI", style_sec))
    lk_rows = [
        [1, "Nama Lengkap", d['nama_lk'], True],
        [2, "Bin", d['bin_lk']],
        [3, "Tempat, Tanggal Lahir (Umur)", f"{d['ttl_lk']} ({d['umur_lk']} Tahun)"],
        [4, "NIK", d['nik_lk']],
        [5, "Pekerjaan", d['pekerjaan_lk']],
        [6, "Status Pernikahan", d['status_lk']],
        [7, "Jenis Kelamin", d['jk_lk']],
        [8, "Nama Istri Terdahulu", d['istri_terdahulu']],
        [9, "Alamat Tempat Tinggal", d['alamat_lk']],
        [10, "Pendidikan Terakhir", d['pendidikan_lk']]
    ]
    elements.append(build_table(lk_rows))
    elements.append(Spacer(1, 4))

    # AYAH LAKI-LAKI
    elements.append(Paragraph("C. DATA AYAH CATIN LAKI-LAKI", style_sec))
    ayah_lk_rows = [
        [1, "Nama Ayah Laki-Laki", f"{d['nama_ayah_lk']} bin {d['bin_ayah_lk']}"],
        [2, "NIK", d['nik_ayah_lk']],
        [3, "Tempat, Tanggal Lahir (Umur)", f"{d['ttl_ayah_lk']} ({d['umur_ayah_lk']} Tahun)"],
        [4, "Pekerjaan", d['pekerjaan_ayah_lk']],
        [5, "Alamat Tempat Tinggal", d['alamat_ayah_lk']]
    ]
    elements.append(build_table(ayah_lk_rows))
    elements.append(Spacer(1, 4))

    # IBU LAKI-LAKI
    elements.append(Paragraph("D. DATA IBU CATIN LAKI-LAKI", style_sec))
    ibu_lk_rows = [
        [1, "Nama Ibu Laki-Laki", f"{d['nama_ibu_lk']} bin {d['bin_ibu_lk']}"],
        [2, "NIK", d['nik_ibu_lk']],
        [3, "Tempat, Tanggal Lahir (Umur)", f"{d['ttl_ibu_lk']} ({d['umur_ibu_lk']} Tahun)"],
        [4, "Pekerjaan", d['pekerjaan_ibu_lk']],
        [5, "Alamat Tempat Tinggal", d['alamat_ibu_lk']]
    ]
    elements.append(build_table(ibu_lk_rows))
    elements.append(Spacer(1, 4))

    # SECTION 3: CATIN PEREMPUAN
    elements.append(Paragraph("E. DATA CALON PENGANTIN PEREMPUAN", style_sec))
    pr_rows = [
        [1, "Nama Lengkap", d['nama_pr'], True],
        [2, "Binti", d['binti_pr']],
        [3, "Tempat, Tanggal Lahir (Umur)", f"{d['ttl_pr']} ({d['umur_pr']} Tahun)"],
        [4, "NIK", d['nik_pr']],
        [5, "Pekerjaan", d['pekerjaan_pr']],
        [6, "Status Pernikahan", d['status_pr']],
        [7, "Jenis Kelamin", d['jk_pr']],
        [8, "Alamat Tempat Tinggal", d['alamat_pr']],
        [9, "Nama Suami Terdahulu", d['suami_terdahulu']],
        [10, "Pendidikan Terakhir", d['pendidikan_pr']]
    ]
    elements.append(build_table(pr_rows))
    elements.append(Spacer(1, 4))

    # AYAH PEREMPUAN
    elements.append(Paragraph("F. DATA AYAH CATIN PEREMPUAN", style_sec))
    ayah_pr_rows = [
        [1, "Nama Ayah Perempuan", f"{d['nama_ayah_pr']} bin {d['bin_ayah_pr']}"],
        [2, "NIK", d['nik_ayah_pr']],
        [3, "Tempat, Tanggal Lahir (Umur)", f"{d['ttl_ayah_pr']} ({d['umur_ayah_pr']} Tahun)"],
        [4, "Pekerjaan", d['pekerjaan_ayah_pr']],
        [5, "Alamat Tempat Tinggal", d['alamat_ayah_pr']]
    ]
    elements.append(build_table(ayah_pr_rows))
    elements.append(Spacer(1, 4))

    # IBU PEREMPUAN
    elements.append(Paragraph("G. DATA IBU CATIN PEREMPUAN", style_sec))
    ibu_pr_rows = [
        [1, "Nama Ibu Perempuan", f"{d['nama_ibu_pr']} bin {d['bin_ibu_pr']}"],
        [2, "NIK", d['nik_ibu_pr']],
        [3, "Tempat, Tanggal Lahir (Umur)", f"{d['ttl_ibu_pr']} ({d['umur_ibu_pr']} Tahun)"],
        [4, "Pekerjaan", d['pekerjaan_ibu_pr']],
        [5, "Alamat Tempat Tinggal", d['alamat_ibu_pr']]
    ]
    elements.append(build_table(ibu_pr_rows))
    elements.append(Spacer(1, 4))

    # SECTION 4: DATA WALI
    elements.append(Paragraph("H. DATA WALI NIKAH", style_sec))
    wali_rows = [
        [1, "Nama Wali", d['nama_wali']],
        [2, "Bin Wali", d['bin_wali']],
        [3, "NIK Wali", d['nik_wali']],
        [4, "Tempat, Tanggal Lahir (Umur)", f"{d['ttl_wali']} ({d['umur_wali']} Tahun)"],
        [5, "Pekerjaan Wali", d['pekerjaan_wali']],
        [6, "Alamat Tempat Tinggal", d['alamat_wali']],
        [7, "Hubungan Wali", d['hubungan_wali']],
        [8, "Nama Wali Lengkap", d['nama_wali_lengkap'], True]
    ]
    elements.append(build_table(wali_rows))
    elements.append(Spacer(1, 4))

    # SECTION 5: DATA SAKSI 1 & SAKSI 2
    elements.append(Paragraph("I. DATA SAKSI NIKAH (SAKSI 1 & SAKSI 2)", style_sec))
    saksi_rows = [
        [1, "Nama Saksi 1", d['saksi1_nama']],
        [2, "TTL / Umur Saksi 1", f"{d['saksi1_ttl']} ({d['saksi1_umur']} Tahun)"],
        [3, "NIK Saksi 1", d['saksi1_nik']],
        [4, "Pekerjaan / Alamat Saksi 1", f"{d['saksi1_pekerjaan']} / {d['saksi1_alamat']}"],
        [5, "Nama Saksi 2", d['saksi2_nama']],
        [6, "TTL / Umur Saksi 2", f"{d['saksi2_ttl']} ({d['saksi2_umur']} Tahun)"],
        [7, "NIK Saksi 2", d['saksi2_nik']],
        [8, "Pekerjaan / Alamat Saksi 2", f"{d['saksi2_pekerjaan']} / {d['saksi2_alamat']}"]
    ]
    elements.append(build_table(saksi_rows))
    elements.append(Spacer(1, 10))

    # TANDA TANGAN
    ttd_table = Table([
        ["", f"{d['tgl_surat']}"],
        ["", "Kepala Desa Tambi / Kasi Pelayanan"],
        ["", "\n\n"],
        ["", "<u><b>CHALIM MUCHTAROM, S.Pd.I</b></u>"]
    ], colWidths=[300, 230])
    ttd_table.setStyle(TableStyle([
        ('ALIGN', (1,0), (1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    elements.append(ttd_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer


# --------------------------------------------------
# FORMULIR INPUT DATA
# --------------------------------------------------
with st.form("form_catin"):
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
            no_register = st.text_input("Nomor Register", value="400.12.3.2/010/ VIII/ 2026")
            tgl_surat = st.text_input("Tanggal Surat", value="TAMBI, 11 AGUSTUS 2026")
            tgl_pelaksanaan = st.date_input("Tanggal Pelaksanaan Akad", value=date(2026, 9, 7))
            jam_akad = st.text_input("Jam Akad", value="JAM. 08.00")
        with col2:
            tempat_akad = st.text_input("Tempat Akad Nikah", value="RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang")
            email_catin = st.text_input("Email Catin", value="")
            mahar = st.text_input("Maskawin / Mahar", value="Seperangkat Alat Sholat")

    with tab2:
        st.subheader("Data Calon Pengantin Laki-Laki")
        col_lk1, col_lk2 = st.columns(2)
        with col_lk1:
            nama_lk = st.text_input("Nama Calon Pengantin Laki-Laki", value="MIfahul Anam")
            bin_lk = st.text_input("Bin (Ayah Laki-Laki)", value="Nur Karim")
            ttl_lk = st.text_input("Tempat, Tanggal Lahir Laki-Laki", value="Pemalang, 18 Februari 1999")
            umur_lk = st.number_input("Umur Laki-Laki", value=27)
            nik_lk = st.text_input("NIK Laki-Laki", value="3327031802990004")
        with col_lk2:
            pekerjaan_lk = st.text_input("Pekerjaan Laki-Laki", value="Swasta")
            status_lk = st.text_input("Status Laki-Laki", value="BELUM KAWIN")
            jk_lk = st.text_input("Jenis Kelamin Laki-Laki", value="Laki-Laki")
            istri_terdahulu = st.text_input("Nama Istri Terdahulu (Jika ada)", value="")
            alamat_lk = st.text_area("Alamat Laki-Laki", value="RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang")
            pendidikan_lk = st.text_input("Pendidikan Laki-Laki", value="SLTA")

        st.divider()
        st.subheader("Data Ayah & Ibu Laki-Laki")
        col_alk, col_ilk = st.columns(2)
        with col_alk:
            st.markdown("**Ayah Laki-Laki**")
            nama_ayah_lk = st.text_input("Nama Ayah Laki-Laki", value="Nur Karim")
            bin_ayah_lk = st.text_input("bin (Kakek Laki-Laki)", value="Kasturi")
            nik_ayah_lk = st.text_input("NIK Ayah Laki-Laki", value="3327030608680006")
            ttl_ayah_lk = st.text_input("TTL Ayah Laki-Laki", value="Pemalang, 06 Agustus 1968")
            umur_ayah_lk = st.number_input("Umur Ayah Laki-Laki", value=58)
            pekerjaan_ayah_lk = st.text_input("Pekerjaan Ayah Laki-Laki", value="PETANI/ PEKEBUN")
            alamat_ayah_lk = st.text_area("Alamat Ayah Laki-Laki", value="RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang")

        with col_ilk:
            st.markdown("**Ibu Laki-Laki**")
            nama_ibu_lk = st.text_input("Nama Ibu Laki-Laki", value="Samijah")
            bin_ibu_lk = st.text_input("bin (Kakek Ibu Laki-Laki)", value="Taryad")
            nik_ibu_lk = st.text_input("NIK Ibu Laki-Laki", value="3327035405740004")
            ttl_ibu_lk = st.text_input("TTL Ibu Laki-Laki", value="Pemalang, 14 Mei 1974")
            umur_ibu_lk = st.number_input("Umur Ibu Laki-Laki", value=52)
            pekerjaan_ibu_lk = st.text_input("Pekerjaan Ibu Laki-Laki", value="Mengurus Rumah Tangga")
            alamat_ibu_lk = st.text_area("Alamat Ibu Laki-Laki", value="RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang")

    with tab3:
        st.subheader("Data Calon Pengantin Perempuan")
        col_pr1, col_pr2 = st.columns(2)
        with col_pr1:
            nama_pr = st.text_input("Nama Calon Pengantin Perempuan", value="Diyan Solehatin")
            binti_pr = st.text_input("Binti (Ayah Perempuan)", value="Disun")
            ttl_pr = st.text_input("Tempat, Tanggal Lahir Perempuan", value="Pemalang, 29 Juni 2007")
            umur_pr = st.number_input("Umur Perempuan", value=19)
            nik_pr = st.text_input("NIK Perempuan", value="3327046906070010")
        with col_pr2:
            pekerjaan_pr = st.text_input("Pekerjaan Perempuan", value="BELUM/ TIDAK BEKERJA")
            status_pr = st.text_input("Status Perempuan", value="BELUM KAWIN")
            jk_pr = st.text_input("Jenis Kelamin Perempuan", value="PEREMPUAN")
            suami_terdahulu = st.text_input("Nama Suami Terdahulu (Jika ada)", value="")
            alamat_pr = st.text_area("Alamat Perempuan", value="RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang")
            pendidikan_pr = st.text_input("Pendidikan Perempuan", value="SLTP")

        st.divider()
        st.subheader("Data Ayah & Ibu Perempuan")
        col_apr, col_ipr = st.columns(2)
        with col_apr:
            st.markdown("**Ayah Perempuan**")
            nama_ayah_pr = st.text_input("Nama Ayah Perempuan", value="Disun")
            bin_ayah_pr = st.text_input("bin (Kakek Perempuan)", value="Tawiroji")
            nik_ayah_pr = st.text_input("NIK Ayah Perempuan", value="3327042504840003")
            ttl_ayah_pr = st.text_input("TTL Ayah Perempuan", value="Pemalang, 21 April 1984")
            umur_ayah_pr = st.number_input("Umur Ayah Perempuan", value=42)
            pekerjaan_ayah_pr = st.text_input("Pekerjaan Ayah Perempuan", value="PETANI/ PEKEBUN")
            alamat_ayah_pr = st.text_area("Alamat Ayah Perempuan", value="RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang")

        with col_ipr:
            st.markdown("**Ibu Perempuan**")
            nama_ibu_pr = st.text_input("Nama Ibu Perempuan", value="Mutirah")
            bin_ibu_pr = st.text_input("bin (Kakek Ibu Perempuan)", value="Tamiarjo")
            nik_ibu_pr = st.text_input("NIK Ibu Perempuan", value="3327044411840003")
            ttl_ibu_pr = st.text_input("TTL Ibu Perempuan", value="Pemalang, 04 November 1984")
            umur_ibu_pr = st.number_input("Umur Ibu Perempuan", value=42)
            pekerjaan_ibu_pr = st.text_input("Pekerjaan Ibu Perempuan", value="Mengurus Rumah Tangga")
            alamat_ibu_pr = st.text_area("Alamat Ibu Perempuan", value="RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang")

    with tab4:
        st.subheader("Data Wali Nikah")
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            nama_wali = st.text_input("Nama Wali", value="Disun")
            bin_wali = st.text_input("Bin Wali", value="Tawiroji")
            nik_wali = st.text_input("NIK Wali", value="3327042504840003")
            ttl_wali = st.text_input("TTL Wali", value="PEMALANG, 21 April 1984")
            umur_wali = st.number_input("Umur Wali", value=42)
        with col_w2:
            pekerjaan_wali = st.text_input("Pekerjaan Wali", value="PETANI/ PEKEBUN")
            alamat_wali = st.text_area("Alamat Wali", value="RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang")
            hubungan_wali = st.text_input("Hubungan Wali", value="AYAH KANDUNG")
            nama_wali_lengkap = st.text_input("Nama Wali Lengkap", value="Disun Bin Tawiroji")

    with tab5:
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            st.subheader("Data Saksi 1")
            saksi1_nama = st.text_input("Nama Saksi 1", value="Chalim Muchtarom")
            saksi1_ttl = st.text_input("TTL Saksi 1", value="Pemalang, 21 Oktober 1989")
            saksi1_umur = st.number_input("Umur Saksi 1", value=37)
            saksi1_nik = st.text_input("NIK Saksi 1", value="3327042110890004")
            saksi1_pekerjaan = st.text_input("Pekerjaan Saksi 1", value="Perangkat Desa")
            saksi1_alamat = st.text_area("Alamat Saksi 1", value="RT 002 RW 001 Desa Tambi Kecamatan Watukumpu Kabupaten Pemalang")

        with col_s2:
            st.subheader("Data Saksi 2")
            saksi2_nama = st.text_input("Nama Saksi 2", value="Sidin")
            saksi2_ttl = st.text_input("TTL Saksi 2", value="Pemalang, ")
            saksi2_umur = st.number_input("Umur Saksi 2", value=0)
            saksi2_nik = st.text_input("NIK Saksi 2", value="0000000000000000")
            saksi2_pekerjaan = st.text_input("Pekerjaan Saksi 2", value="")
            saksi2_alamat = st.text_area("Alamat Saksi 2", value="")

    submit = st.form_submit_button("💾 PROSES DATA & GENERATE BERKAS")


# --------------------------------------------------
# PROSES UPDATE EXCEL & GENERATE PDF F4
# --------------------------------------------------
if submit:
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=False)
        sheet = wb['ISIAN DATA']

        cell_updates = {
            'G2': no_register,
            'H3': tgl_surat,
            'G4': tgl_pelaksanaan.strftime('%Y-%m-%d'),
            'G5': jam_akad,
            'G6': tempat_akad,
            'K4': email_catin,
            'G8': nama_lk,
            'G9': bin_lk,
            'G10': ttl_lk,
            'K10': umur_lk,
            'G11': nik_lk,
            'G12': pekerjaan_lk,
            'G13': status_lk,
            'G14': jk_lk,
            'G15': istri_terdahulu,
            'G16': alamat_lk,
            'G17': pendidikan_lk,
            'G19': nama_ayah_lk,
            'J19': bin_ayah_lk,
            'G20': nik_ayah_lk,
            'G21': ttl_ayah_lk,
            'K21': umur_ayah_lk,
            'G22': pekerjaan_ayah_lk,
            'G23': alamat_ayah_lk,
            'G26': nama_ibu_lk,
            'I26': bin_ibu_lk,
            'G27': nik_ibu_lk,
            'G28': ttl_ibu_lk,
            'K28': umur_ibu_lk,
            'G29': pekerjaan_ibu_lk,
            'G30': alamat_ibu_lk,
            'G34': nama_pr,
            'G35': binti_pr,
            'G36': ttl_pr,
            'K36': umur_pr,
            'G37': nik_pr,
            'G38': pekerjaan_pr,
            'G39': status_pr,
            'G40': jk_pr,
            'G41': alamat_pr,
            'G42': suami_terdahulu,
            'G43': pendidikan_pr,
            'G45': nama_ayah_pr,
            'I45': bin_ayah_pr,
            'G46': nik_ayah_pr,
            'G47': ttl_ayah_pr,
            'K47': umur_ayah_pr,
            'G48': pekerjaan_ayah_pr,
            'G49': alamat_ayah_pr,
            'G52': nama_ibu_pr,
            'I52': bin_ibu_pr,
            'G53': nik_ibu_pr,
            'G54': ttl_ibu_pr,
            'K54': umur_ibu_pr,
            'G55': pekerjaan_ibu_pr,
            'G56': alamat_ibu_pr,
            'G58': nama_wali,
            'G59': bin_wali,
            'G60': nik_wali,
            'G61': ttl_wali,
            'K61': umur_wali,
            'G62': pekerjaan_wali,
            'G63': alamat_wali,
            'G64': hubungan_wali,
            'G65': mahar,
            'G68': nama_wali_lengkap,
            'G70': saksi1_nama,
            'G71': saksi1_ttl,
            'K71': saksi1_umur,
            'G72': saksi1_nik,
            'G73': saksi1_pekerjaan,
            'G74': saksi1_alamat,
            'G76': saksi2_nama,
            'G77': saksi2_ttl,
            'K77': saksi2_umur,
            'G78': saksi2_nik,
            'G79': saksi2_pekerjaan,
            'G80': saksi2_alamat,
        }

        for cell_ref, val in cell_updates.items():
            sheet[cell_ref] = val

        output_excel = BytesIO()
        wb.save(output_excel)
        output_excel.seek(0)

        data_dict = {
            'no_register': no_register, 'tgl_surat': tgl_surat, 'tgl_pelaksanaan': tgl_pelaksanaan.strftime('%d-%m-%Y'),
            'jam_akad': jam_akad, 'tempat_akad': tempat_akad, 'email_catin': email_catin, 'mahar': mahar,
            'nama_lk': nama_lk, 'bin_lk': bin_lk, 'ttl_lk': ttl_lk, 'umur_lk': umur_lk, 'nik_lk': nik_lk,
            'pekerjaan_lk': pekerjaan_lk, 'status_lk': status_lk, 'jk_lk': jk_lk, 'istri_terdahulu': istri_terdahulu,
            'alamat_lk': alamat_lk, 'pendidikan_lk': pendidikan_lk, 'nama_ayah_lk': nama_ayah_lk, 'bin_ayah_lk': bin_ayah_lk,
            'nik_ayah_lk': nik_ayah_lk, 'ttl_ayah_lk': ttl_ayah_lk, 'umur_ayah_lk': umur_ayah_lk, 'pekerjaan_ayah_lk': pekerjaan_ayah_lk,
            'alamat_ayah_lk': alamat_ayah_lk, 'nama_ibu_lk': nama_ibu_lk, 'bin_ibu_lk': bin_ibu_lk, 'nik_ibu_lk': nik_ibu_lk,
            'ttl_ibu_lk': ttl_ibu_lk, 'umur_ibu_lk': umur_ibu_lk, 'pekerjaan_ibu_lk': pekerjaan_ibu_lk, 'alamat_ibu_lk': alamat_ibu_lk,
            'nama_pr': nama_pr, 'binti_pr': binti_pr, 'ttl_pr': ttl_pr, 'umur_pr': umur_pr, 'nik_pr': nik_pr,
            'pekerjaan_pr': pekerjaan_pr, 'status_pr': status_pr, 'jk_pr': jk_pr, 'alamat_pr': alamat_pr,
            'suami_terdahulu': suami_terdahulu, 'pendidikan_pr': pendidikan_pr, 'nama_ayah_pr': nama_ayah_pr, 'bin_ayah_pr': bin_ayah_pr,
            'nik_ayah_pr': nik_ayah_pr, 'ttl_ayah_pr': ttl_ayah_pr, 'umur_ayah_pr': umur_ayah_pr, 'pekerjaan_ayah_pr': pekerjaan_ayah_pr,
            'alamat_ayah_pr': alamat_ayah_pr, 'nama_ibu_pr': nama_ibu_pr, 'bin_ibu_pr': bin_ibu_pr, 'nik_ibu_pr': nik_ibu_pr,
            'ttl_ibu_pr': ttl_ibu_pr, 'umur_ibu_pr': umur_ibu_pr, 'pekerjaan_ibu_pr': pekerjaan_ibu_pr, 'alamat_ibu_pr': alamat_ibu_pr,
            'nama_wali': nama_wali, 'bin_wali': bin_wali, 'nik_wali': nik_wali, 'ttl_wali': ttl_wali, 'umur_wali': umur_wali,
            'pekerjaan_wali': pekerjaan_wali, 'alamat_wali': alamat_wali, 'hubungan_wali': hubungan_wali, 'nama_wali_lengkap': nama_wali_lengkap,
            'saksi1_nama': saksi1_nama, 'saksi1_ttl': saksi1_ttl, 'saksi1_umur': saksi1_umur, 'saksi1_nik': saksi1_nik,
            'saksi1_pekerjaan': saksi1_pekerjaan, 'saksi1_alamat': saksi1_alamat, 'saksi2_nama': saksi2_nama, 'saksi2_ttl': saksi2_ttl,
            'saksi2_umur': saksi2_umur, 'saksi2_nik': saksi2_nik, 'saksi2_pekerjaan': saksi2_pekerjaan, 'saksi2_alamat': saksi2_alamat
        }

        output_pdf = generate_pdf_isian_data(data_dict)

        st.success("✅ Data berhasil diproses!")

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                label="📥 DOWNLOAD EXCEL UTUH (.XLSX)",
                data=output_excel,
                file_name=f"BERKAS_CATIN_{nama_lk}_{nama_pr}.xlsx".replace(" ", "_"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_d2:
            st.download_button(
                label="📄 DOWNLOAD PDF ISIAN DATA (UKURAN F4)",
                data=output_pdf,
                file_name=f"ISIAN_DATA_F4_{nama_lk}.pdf".replace(" ", "_"),
                mime="application/pdf",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Gagal memproses file: {e}")
