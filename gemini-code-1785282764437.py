import streamlit as st
import openpyxl
from io import BytesIO
from datetime import datetime, date

# Library untuk Pembuatan PDF Ukuran F4
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Ukuran Kertas F4 dalam Points (215.9 mm x 330 mm)
F4_WIDTH = 215.9 * 2.83465
F4_HEIGHT = 330.0 * 2.83465
F4_SIZE = (F4_WIDTH, F4_HEIGHT)

st.set_page_config(
    page_title="Aplikasi Berkas Catin - Desa Tambi",
    page_icon="📜",
    layout="wide"
)

st.title("📜 Form Input Berkas Catin Desa Tambi")
st.caption("Mengisi sheet 'ISIAN DATA' tanpa merusak rumus. Menyediakan unduhan Excel utuh & PDF Surat Pengantar (F4).")

EXCEL_FILE = "BERKAS CATIN .xlsx"

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
    
    # --- TAB 1: REGISTER & AKAD ---
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

    # --- TAB 2: CATIN LAKI-LAKI & ORTU ---
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

    # --- TAB 3: CATIN PEREMPUAN & ORTU ---
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

    # --- TAB 4: DATA WALI ---
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

    # --- TAB 5: DATA SAKSI 1 & 2 ---
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
            saksi2_nik = st.text_input("NIK Saksi 2", value="0000000000000000")
            saksi2_pekerjaan = st.text_input("Pekerjaan Saksi 2", value="")
            saksi2_alamat = st.text_area("Alamat Saksi 2", value="")

    submit = st.form_submit_button("💾 PROSES DATA & GENERATE BERKAS")


# --------------------------------------------------
# FUNGSI MEMBUAT PDF SURAT PENGANTAR HALAMAN DEPAN (UKURAN F4)
# --------------------------------------------------
def create_f4_pdf(data_dict):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=F4_SIZE,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    
    styles = getSampleStyleSheet()
    
    style_kop_head = ParagraphStyle('KopHead', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, alignment=1, leading=14)
    style_kop_sub = ParagraphStyle('KopSub', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=14, alignment=1, leading=16)
    style_title = ParagraphStyle('TitlePDF', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, alignment=1, leading=14)
    style_body = ParagraphStyle('BodyPDF', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=13)
    style_bold = ParagraphStyle('BoldPDF', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=13)

    elements = []

    # KOP SURAT
    elements.append(Paragraph("PEMERINTAH KABUPATEN PEMALANG", style_kop_head))
    elements.append(Paragraph("KECAMATAN WATUKUMPUL", style_kop_head))
    elements.append(Paragraph("DESA TAMBI", style_kop_sub))
    elements.append(Spacer(1, 10))

    # JUDUL SURAT
    elements.append(Paragraph("<u>PENGANTAR NIKAH</u>", style_title))
    elements.append(Paragraph(f"Nomor: {data_dict['no_register']}", style_title))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("Yang bertanda tangan di bawah ini menjelaskan dengan sesungguhnya bahwa:", style_body))
    elements.append(Spacer(1, 8))

    # TABLE DATA CATIN LAKI-LAKI
    table_data = [
        [Paragraph("1.", style_body), Paragraph("Nama Lengkap", style_body), Paragraph(":", style_body), Paragraph(data_dict['nama_lk'], style_bold)],
        [Paragraph("2.", style_body), Paragraph("NIK", style_body), Paragraph(":", style_body), Paragraph(data_dict['nik_lk'], style_body)],
        [Paragraph("3.", style_body), Paragraph("Jenis Kelamin", style_body), Paragraph(":", style_body), Paragraph(data_dict['jk_lk'], style_body)],
        [Paragraph("4.", style_body), Paragraph("Tempat, Tgl Lahir", style_body), Paragraph(":", style_body), Paragraph(data_dict['ttl_lk'], style_body)],
        [Paragraph("5.", style_body), Paragraph("Pekerjaan", style_body), Paragraph(":", style_body), Paragraph(data_dict['pekerjaan_lk'], style_body)],
        [Paragraph("6.", style_body), Paragraph("Status Pernikahan", style_body), Paragraph(":", style_body), Paragraph(data_dict['status_lk'], style_body)],
        [Paragraph("7.", style_body), Paragraph("Alamat", style_body), Paragraph(":", style_body), Paragraph(data_dict['alamat_lk'], style_body)],
    ]

    t1 = Table(table_data, colWidths=[20, 130, 15, 380])
    t1.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('BOTTOMPADDING', (0,0), (-1,-1), 2)]))
    elements.append(t1)
    elements.append(Spacer(1, 10))

    elements.append(Paragraph("Adalah benar-benar anak dari pernikahan seorang pria:", style_body))
    elements.append(Spacer(1, 6))

    # TABLE AYAH
    table_ayah = [
        [Paragraph("1.", style_body), Paragraph("Nama Ayah", style_body), Paragraph(":", style_body), Paragraph(f"{data_dict['nama_ayah_lk']} bin {data_dict['bin_ayah_lk']}", style_body)],
        [Paragraph("2.", style_body), Paragraph("NIK Ayah", style_body), Paragraph(":", style_body), Paragraph(data_dict['nik_ayah_lk'], style_body)],
        [Paragraph("3.", style_body), Paragraph("Tempat, Tgl Lahir", style_body), Paragraph(":", style_body), Paragraph(data_dict['ttl_ayah_lk'], style_body)],
        [Paragraph("4.", style_body), Paragraph("Pekerjaan / Alamat", style_body), Paragraph(":", style_body), Paragraph(f"{data_dict['pekerjaan_ayah_lk']} / {data_dict['alamat_ayah_lk']}", style_body)],
    ]
    t2 = Table(table_ayah, colWidths=[20, 130, 15, 380])
    t2.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('BOTTOMPADDING', (0,0), (-1,-1), 2)]))
    elements.append(t2)
    elements.append(Spacer(1, 10))

    # CALON PASANGAN
    elements.append(Paragraph("Dan hendak menikah dengan calon pasangan:", style_body))
    elements.append(Spacer(1, 6))

    table_pasangan = [
        [Paragraph("1.", style_body), Paragraph("Nama Calon Istri", style_body), Paragraph(":", style_body), Paragraph(f"{data_dict['nama_pr']} binti {data_dict['binti_pr']}", style_bold)],
        [Paragraph("2.", style_body), Paragraph("NIK", style_body), Paragraph(":", style_body), Paragraph(data_dict['nik_pr'], style_body)],
        [Paragraph("3.", style_body), Paragraph("Tempat, Tgl Lahir", style_body), Paragraph(":", style_body), Paragraph(data_dict['ttl_pr'], style_body)],
        [Paragraph("4.", style_body), Paragraph("Alamat", style_body), Paragraph(":", style_body), Paragraph(data_dict['alamat_pr'], style_body)],
    ]
    t3 = Table(table_pasangan, colWidths=[20, 130, 15, 380])
    t3.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('BOTTOMPADDING', (0,0), (-1,-1), 2)]))
    elements.append(t3)
    elements.append(Spacer(1, 15))

    # TANDA TANGAN
    data_ttd = [
        ["", f"{data_dict['tgl_surat']}"],
        ["", "Kepala Desa / Kasi Pelayanan"],
        ["", "\n\n\n"],
        ["", "<u><b>CHALIM MUCHTAROM, S.Pd.I</b></u>"]
    ]
    t_ttd = Table(data_ttd, colWidths=[300, 245])
    t_ttd.setStyle(TableStyle([
        ('ALIGN', (1,0), (1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    elements.append(t_ttd)

    doc.build(elements)
    buffer.seek(0)
    return buffer


# --------------------------------------------------
# PROSES EKSEKUSI PENYIMPANAN EXCEL & GENERATE PDF
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
            'G78': saksi2_nik,
            'G79': saksi2_pekerjaan,
            'G80': saksi2_alamat,
        }

        for cell_ref, val in cell_updates.items():
            sheet[cell_ref] = val

        # Save Excel to memory
        output_excel = BytesIO()
        wb.save(output_excel)
        output_excel.seek(0)

        # Build PDF Data Dict
        data_pdf = {
            'no_register': no_register,
            'tgl_surat': tgl_surat,
            'nama_lk': nama_lk,
            'bin_lk': bin_lk,
            'nik_lk': nik_lk,
            'jk_lk': jk_lk,
            'ttl_lk': ttl_lk,
            'pekerjaan_lk': pekerjaan_lk,
            'status_lk': status_lk,
            'alamat_lk': alamat_lk,
            'nama_ayah_lk': nama_ayah_lk,
            'bin_ayah_lk': bin_ayah_lk,
            'nik_ayah_lk': nik_ayah_lk,
            'ttl_ayah_lk': ttl_ayah_lk,
            'pekerjaan_ayah_lk': pekerjaan_ayah_lk,
            'alamat_ayah_lk': alamat_ayah_lk,
            'nama_pr': nama_pr,
            'binti_pr': binti_pr,
            'nik_pr': nik_pr,
            'ttl_pr': ttl_pr,
            'alamat_pr': alamat_pr
        }

        pdf_bytes = create_f4_pdf(data_pdf)

        st.success("✅ Data berhasil diproses! Silakan pilih berkas yang ingin diunduh di bawah ini:")

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
                label="📄 DOWNLOAD HALAMAN DEPAN PDF (F4)",
                data=pdf_bytes,
                file_name=f"PENGANTAR_NIKAH_F4_{nama_lk}.pdf".replace(" ", "_"),
                mime="application/pdf",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Gagal memproses file: {e}")
