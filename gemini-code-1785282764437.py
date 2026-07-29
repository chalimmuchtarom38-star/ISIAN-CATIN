import streamlit as st
import openpyxl
import io
import re
import gc
from datetime import datetime

# Impor reportlab secara aman
HAS_REPORTLAB = True
try:
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
except ImportError:
    HAS_REPORTLAB = False

st.set_page_config(page_title="Form Input Data Catin", page_icon="📝", layout="centered")

st.title("📝 Form Pengisian Data Catin")
st.write("Isi formulir di bawah ini. Anda dapat mendownload file **Excel Utuh (Master 100%)** atau **PDF Isian Data (Full 1 Halaman F4)**.")

EXCEL_MASTER = "N SUSILAH RT 010 RW 002.xlsx"

@st.cache_resource(show_spinner=False)
def load_master_bytes():
    with open(EXCEL_MASTER, "rb") as f:
        return f.read()

def hitung_umur(ttl_text):
    if not ttl_text:
        return ""
    match = re.search(r'\b(19\d\d|20\d\d)\b', str(ttl_text))
    if match:
        tahun_lahir = int(match.group(1))
        tahun_sekarang = datetime.now().year
        umur = tahun_sekarang - tahun_lahir
        return umur if umur >= 0 else ""
    return ""

def validasi_nik(nik, label):
    nik_str = str(nik).strip()
    if not nik_str.isdigit():
        return f"❌ {label} harus berupa angka saja!"
    if len(nik_str) != 16:
        return f"❌ {label} harus tepat 16 digit! (Saat ini: {len(nik_str)} digit)"
    return None

def generate_pdf_isian_data(data_dict):
    if not HAS_REPORTLAB:
        return None
    
    buffer = io.BytesIO()
    
    # Ukuran Kertas F4 dalam Point (215.9mm x 330mm = 612pt x 936pt)
    f4_size = (612, 936)
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=f4_size,
        rightMargin=22,
        leftMargin=22,
        topMargin=22,
        bottomMargin=22
    )
    
    styles = getSampleStyleSheet()
    
    # Judul Dokumen
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=12,
        alignment=1, # Center
        spaceAfter=8,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor("#1A365D")
    )
    
    # Judul Sub-Seksi
    section_style = ParagraphStyle(
        'SectionStyle',
        parent=styles['Heading2'],
        fontSize=8.5,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor("#1A365D"),
        spaceBefore=0,
        spaceAfter=0
    )
    
    # Text dalam Sel (Ditingkatkan ukurannya agar pas mengisi halaman)
    cell_label_style = ParagraphStyle('CellLabel', fontSize=8, fontName='Helvetica-Bold', leading=10, textColor=colors.HexColor("#2D3748"))
    cell_val_style = ParagraphStyle('CellVal', fontSize=8, fontName='Helvetica', leading=10, textColor=colors.HexColor("#1A202C"))

    story = []
    story.append(Paragraph("<b>LEMBAR ISIAN DATA CATIN & PELAKSANAAN AKAD</b>", title_style))

    table_data = []
    t_style = [
        ('BACKGROUND', (0,0), (-1,-1), colors.white),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.5), # Padding ditingkatkan agar tinggi tabel full 1 halaman
        ('TOPPADDING', (0,0), (-1,-1), 2.5),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
    ]

    row_idx = 0
    for section, fields in data_dict.items():
        # Header Seksi
        sec_p = Paragraph(f"<b>{section.upper()}</b>", section_style)
        table_data.append([sec_p, ""])
        t_style.append(('SPAN', (0, row_idx), (1, row_idx)))
        t_style.append(('BACKGROUND', (0, row_idx), (1, row_idx), colors.HexColor("#E2E8F0")))
        row_idx += 1
        
        # Baris Data
        for k, v in fields.items():
            p_k = Paragraph(k, cell_label_style)
            p_v = Paragraph(str(v) if v else "-", cell_val_style)
            table_data.append([p_k, p_v])
            row_idx += 1

    t = Table(table_data, colWidths=[155, 413])
    t.setStyle(TableStyle(t_style))
    story.append(t)
    
    doc.build(story)
    buffer.seek(0)
    return buffer

try:
    master_bytes = load_master_bytes()
except Exception as e:
    st.error(f"⚠️ Gagal membaca file master Excel '{EXCEL_MASTER}'. Error: {e}")
    st.stop()

with st.form("catin_form"):
    st.header("📌 1. Data Surat & Pelaksanaan")
    no_register = st.text_input("Nomor Register", "400.12.3.2/007/ VII/ 2026")
    tgl_surat = st.text_input("Tanggal Surat", "TAMBI, 29 JULI 2026")
    tgl_pelaksanaan = st.date_input("Tanggal Pelaksanaan Akad")
    jam_pelaksanaan = st.text_input("Jam Pelaksanaan", "JAM. 08.00")
    tempat_akad = st.text_area("Tempat Akad Nikah", "DI RUMAH MEMPELAI PUTRI RT 010 RW 002 DESA TAMBI WATUKUMPUL PEMALANG")

    st.markdown("---")
    st.header("👨 2. Calon Pengantin Laki-Laki")
    pria_nama = st.text_input("Nama Laki-Laki", "SAEFULOH")
    pria_bin = st.text_input("BIN (Ayah Laki-Laki)", "KUNENI")
    pria_ttl = st.text_input("Tempat Tanggal Lahir (Pria)", "PURBALINGGA , 07 APRIL 2000")
    pria_nik = st.text_input("NIK (Pria)", "3303170704000001", max_chars=16)
    pria_pekerjaan = st.text_input("Pekerjaan (Pria)", "KARYAWAN SWASTA")
    pria_status = st.selectbox("Status (Pria)", ["BELUM KAWIN", "DUDA"], index=0)
    pria_pendidikan = st.text_input("Pendidikan (Pria)", "SLTA")
    pria_alamat = st.text_area("Alamat (Pria)", "RT 004 RW 001 DESA JINGKANG KECAMATAN KARANGJAMBU KABUPATEN PURBALINGGA")

    st.subheader("👴 Data Ayah Laki-Laki")
    ayah_pria_nama = st.text_input("Nama Ayah Laki-Laki", "KUNENI")
    ayah_pria_bin = st.text_input("BIN Ayah Laki-Laki", "SANBISRI")
    ayah_pria_nik = st.text_input("NIK Ayah Laki-Laki", "3303170406740001", max_chars=16)
    ayah_pria_ttl = st.text_input("TTL Ayah Laki-Laki", "PURBALINGGA, 04 JUNI 1974")
    ayah_pria_pekerjaan = st.text_input("Pekerjaan Ayah Laki-Laki", "PETANI/ PEKEBUN")
    ayah_pria_alamat = st.text_area("Alamat Ayah Laki-Laki", "RT 004 RW 001 DESA JINGKANG KECAMATAN KARANGJAMBU KABUPATEN PURBALINGGA")

    st.subheader("👵 Data Ibu Laki-Laki")
    ibu_pria_nama = st.text_input("Nama Ibu Laki-Laki", "DARYATI")
    ibu_pria_bin = st.text_input("BIN Ibu Laki-Laki", "SUMIARTO")
    ibu_pria_nik = st.text_input("NIK Ibu Laki-Laki", "3303174705800003", max_chars=16)
    ibu_pria_ttl = st.text_input("TTL Ibu Laki-Laki", "PURBALINGGA, 07 JUNI 1980")
    ibu_pria_pekerjaan = st.text_input("Pekerjaan Ibu Laki-Laki", "Mengurus Rumah Tangga")
    ibu_pria_alamat = st.text_area("Alamat Ibu Laki-Laki", "RT 004 RW 001 DESA JINGKANG KECAMATAN KARANGJAMBU KABUPATEN PURBALINGGA")

    st.markdown("---")
    st.header("👩 3. Calon Pengantin Perempuan")
    wanita_nama = st.text_input("Nama Perempuan", "SUSILAH")
    wanita_binti = st.text_input("BINTI (Ayah Perempuan)", "RUSMAN")
    wanita_ttl = st.text_input("Tempat Tanggal Lahir (Perempuan)", "PEMALANG, 11 SEPTEMBER 1999")
    wanita_nik = st.text_input("NIK (Perempuan)", "3327045109990007", max_chars=16)
    wanita_pekerjaan = st.text_input("Pekerjaan (Perempuan)", "BELUM/ TIDAK BEKERJA")
    wanita_status = st.selectbox("Status (Perempuan)", ["BELUM KAWIN", "JANDA"], index=0)
    wanita_pendidikan = st.text_input("Pendidikan (Perempuan)", "SLTA")
    wanita_alamat = st.text_area("Alamat (Perempuan)", "RT 002 RW 004 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    st.subheader("👴 Data Ayah Perempuan")
    ayah_wanita_nama = st.text_input("Nama Ayah Perempuan", "RUSMAN")
    ayah_wanita_bin = st.text_input("BIN Ayah Perempuan", "MARTA")
    ayah_wanita_nik = st.text_input("NIK Ayah Perempuan", "3327040107740063", max_chars=16)
    ayah_wanita_ttl = st.text_input("TTL Ayah Perempuan", "PEMALANG, 01 JULI 1974")
    ayah_wanita_pekerjaan = st.text_input("Pekerjaan Ayah Perempuan", "PETANI/ PEKEBUN")
    ayah_wanita_alamat = st.text_area("Alamat Ayah Perempuan", "RT 002 RW 004 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    st.subheader("👵 Data Ibu Perempuan")
    ibu_wanita_nama = st.text_input("Nama Ibu Perempuan", "KHOSINGAH")
    ibu_wanita_bin = st.text_input("BIN Ibu Perempuan", "NASIR")
    ibu_wanita_nik = st.text_input("NIK Ibu Perempuan", "3327044107810112", max_chars=16)
    ibu_wanita_ttl = st.text_input("TTL Ibu Perempuan", "PEMALANG 01 JULI 1981")
    ibu_wanita_pekerjaan = st.text_input("Pekerjaan Ibu Perempuan", "Mengurus Rumah Tangga")
    ibu_wanita_alamat = st.text_area("Alamat Ibu Perempuan", "RT 002 RW 004 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    st.markdown("---")
    st.header("🤝 4. Data Wali & Mahar")
    wali_nama = st.text_input("Nama Wali", "RUSMAN")
    wali_bin = st.text_input("BIN Wali", "MARTA")
    wali_nik = st.text_input("NIK Wali", "3327040107740063", max_chars=16)
    wali_ttl = st.text_input("TTL Wali", "PEMALANG, 01 JULI 1974")
    wali_pekerjaan = st.text_input("Pekerjaan Wali", "PETANI/ PEKEBUN")
    wali_alamat = st.text_area("Alamat Wali", "RT 002 RW 004 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")
    wali_hubungan = st.text_input("Hubungan Wali", "AYAH KANDUNG")
    mahar = st.text_input("Mahar / Maskawin", "Seperangkat Alat Sholat")
    nama_lengkap_wali_b68 = st.text_input("Nama Wali Lengkap beserta BIN", "RUSMAN BIN MARTA")

    st.markdown("---")
    st.header("📜 5. Data Saksi-Saksi")
    st.subheader("👤 Saksi 1")
    saksi1_nama = st.text_input("Nama Saksi 1", "JURI")
    saksi1_ttl = st.text_input("TTL Saksi 1", "PEMALANG, 04 MEI 1966")
    saksi1_nik = st.text_input("NIK Saksi 1", "3327040405660002", max_chars=16)
    saksi1_pekerjaan = st.text_input("Pekerjaan Saksi 1", "KEPALA DESA")
    saksi1_alamat = st.text_area("Alamat Saksi 1", "RT 010 RW 002 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    st.subheader("👤 Saksi 2")
    saksi2_nama = st.text_input("Nama Saksi 2", "GHOFUR")
    saksi2_ttl = st.text_input("TTL Saksi 2", "PURBALINGGA, 10 FEBRUARI 1988")
    saksi2_nik = st.text_input("NIK Saksi 2", "3303171002880003", max_chars=16)
    saksi2_pekerjaan = st.text_input("Pekerjaan Saksi 2", "WIRASWASTA")
    saksi2_alamat = st.text_area("Alamat Saksi 2", "RT 001 RW 001 DESA JINGKANG KECAMATAN KARANGJAMBU KABUPATEN PURBALINGGA")

    submitted = st.form_submit_button("💾 Proses Data & Siapkan Download")

if submitted:
    daftar_nik = [
        (pria_nik, "NIK Catin Laki-Laki"),
        (ayah_pria_nik, "NIK Ayah Laki-Laki"),
        (ibu_pria_nik, "NIK Ibu Laki-Laki"),
        (wanita_nik, "NIK Catin Perempuan"),
        (ayah_wanita_nik, "NIK Ayah Perempuan"),
        (ibu_wanita_nik, "NIK Ibu Perempuan"),
        (wali_nik, "NIK Wali"),
        (saksi1_nik, "NIK Saksi 1"),
        (saksi2_nik, "NIK Saksi 2")
    ]

    errors = []
    for nik_val, label in daftar_nik:
        err = validasi_nik(nik_val, label)
        if err:
            errors.append(err)

    if errors:
        for err in errors:
            st.error(err)
    else:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(master_bytes))
            ws = wb["ISIAN DATA"]

            ws["G2"] = no_register
            ws["H3"] = f", {tgl_surat}"
            ws["G4"] = str(tgl_pelaksanaan)
            ws["G5"] = jam_pelaksanaan
            ws["G6"] = tempat_akad

            ws["G8"] = pria_nama
            ws["G9"] = pria_bin
            ws["G10"] = pria_ttl
            ws["K10"] = hitung_umur(pria_ttl)
            ws["G11"] = str(pria_nik)
            ws["G12"] = pria_pekerjaan
            ws["G13"] = pria_status
            ws["G16"] = pria_alamat
            ws["G17"] = pria_pendidikan

            ws["G19"] = ayah_pria_nama
            ws["J19"] = ayah_pria_bin
            ws["G20"] = str(ayah_pria_nik)
            ws["G21"] = ayah_pria_ttl
            ws["K21"] = hitung_umur(ayah_pria_ttl)
            ws["G22"] = ayah_pria_pekerjaan
            ws["G23"] = ayah_pria_alamat

            ws["G26"] = ibu_pria_nama
            ws["I26"] = ibu_pria_bin
            ws["G27"] = str(ibu_pria_nik)
            ws["G28"] = ibu_pria_ttl
            ws["K28"] = hitung_umur(ibu_pria_ttl)
            ws["G29"] = ibu_pria_pekerjaan
            ws["G30"] = ibu_pria_alamat

            ws["G34"] = wanita_nama
            ws["G35"] = wanita_binti
            ws["G36"] = wanita_ttl
            ws["K36"] = hitung_umur(wanita_ttl)
            ws["G37"] = str(wanita_nik)
            ws["G38"] = wanita_pekerjaan
            ws["G39"] = wanita_status
            ws["G41"] = wanita_alamat
            ws["G43"] = wanita_pendidikan

            ws["G45"] = ayah_wanita_nama
            ws["I45"] = ayah_wanita_bin
            ws["G46"] = str(ayah_wanita_nik)
            ws["G47"] = ayah_wanita_ttl
            ws["K47"] = hitung_umur(ayah_wanita_ttl)
            ws["G48"] = ayah_wanita_pekerjaan
            ws["G49"] = ayah_wanita_alamat

            ws["G52"] = ibu_wanita_nama
            ws["I52"] = ibu_wanita_bin
            ws["G53"] = str(ibu_wanita_nik)
            ws["G54"] = ibu_wanita_ttl
            ws["K54"] = hitung_umur(ibu_wanita_ttl)
            ws["G55"] = ibu_wanita_pekerjaan
            ws["G56"] = ibu_wanita_alamat

            ws["G58"] = wali_nama
            ws["G59"] = wali_bin
            ws["G60"] = str(wali_nik)
            ws["G61"] = wali_ttl
            ws["K61"] = hitung_umur(wali_ttl)
            ws["G62"] = wali_pekerjaan
            ws["G63"] = wali_alamat
            ws["G64"] = wali_hubungan
            ws["G65"] = mahar
            ws["G68"] = nama_lengkap_wali_b68

            ws["G70"] = saksi1_nama
            ws["G71"] = saksi1_ttl
            ws["K71"] = hitung_umur(saksi1_ttl)
            ws["G72"] = str(saksi1_nik)
            ws["G73"] = saksi1_pekerjaan
            ws["G74"] = saksi1_alamat

            ws["G76"] = saksi2_nama
            ws["G77"] = saksi2_ttl
            ws["K77"] = hitung_umur(saksi2_ttl)
            ws["G78"] = str(saksi2_nik)
            ws["G79"] = saksi2_pekerjaan
            ws["G80"] = saksi2_alamat

            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)

            st.success("✅ Data berhasil diproses!")

            col1, col2 = st.columns(2)
            
            with col1:
                st.download_button(
                    label="📊 Download Excel Utuh (100% Asli Master)",
                    data=excel_buffer,
                    file_name=f"BERKAS_CATIN_{wanita_nama}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            with col2:
                if HAS_REPORTLAB:
                    pdf_data = {
                        "1. Surat & Pelaksanaan Akad": {
                            "No Register": no_register,
                            "Tanggal Surat": tgl_surat,
                            "Tanggal Pelaksanaan Akad": str(tgl_pelaksanaan),
                            "Jam Pelaksanaan": jam_pelaksanaan,
                            "Tempat Akad Nikah": tempat_akad
                        },
                        "2. Calon Pengantin Laki-Laki": {
                            "Nama Laki-Laki": pria_nama,
                            "BIN (Ayah Laki-Laki)": pria_bin,
                            "TTL (Pria)": pria_ttl,
                            "NIK (Pria)": pria_nik,
                            "Pekerjaan (Pria)": pria_pekerjaan,
                            "Status / Pendidikan": f"{pria_status} / {pria_pendidikan}",
                            "Alamat (Pria)": pria_alamat
                        },
                        "Orang Tua Laki-Laki": {
                            "Ayah Laki-Laki": f"{ayah_pria_nama} bin {ayah_pria_bin} (NIK: {ayah_pria_nik} | TTL: {ayah_pria_ttl})",
                            "Pekerjaan / Alamat Ayah": f"{ayah_pria_pekerjaan} - {ayah_pria_alamat}",
                            "Ibu Laki-Laki": f"{ibu_pria_nama} binti {ibu_pria_bin} (NIK: {ibu_pria_nik} | TTL: {ibu_pria_ttl})",
                            "Pekerjaan / Alamat Ibu": f"{ibu_pria_pekerjaan} - {ibu_pria_alamat}"
                        },
                        "3. Calon Pengantin Perempuan": {
                            "Nama Perempuan": wanita_nama,
                            "BINTI (Ayah Perempuan)": wanita_binti,
                            "TTL (Perempuan)": wanita_ttl,
                            "NIK (Perempuan)": wanita_nik,
                            "Pekerjaan (Perempuan)": wanita_pekerjaan,
                            "Status / Pendidikan": f"{wanita_status} / {wanita_pendidikan}",
                            "Alamat (Perempuan)": wanita_alamat
                        },
                        "Orang Tua Perempuan": {
                            "Ayah Perempuan": f"{ayah_wanita_nama} bin {ayah_wanita_bin} (NIK: {ayah_wanita_nik} | TTL: {ayah_wanita_ttl})",
                            "Pekerjaan / Alamat Ayah": f"{ayah_wanita_pekerjaan} - {ayah_wanita_alamat}",
                            "Ibu Perempuan": f"{ibu_wanita_nama} binti {ibu_wanita_bin} (NIK: {ibu_wanita_nik} | TTL: {ibu_wanita_ttl})",
                            "Pekerjaan / Alamat Ibu": f"{ibu_wanita_pekerjaan} - {ibu_wanita_alamat}"
                        },
                        "4. Wali & Mahar": {
                            "Nama Wali / BIN": f"{wali_nama} bin {wali_bin} (Lengkap: {nama_lengkap_wali_b68})",
                            "NIK / TTL Wali": f"{wali_nik} / {wali_ttl}",
                            "Pekerjaan / Hubungan": f"{wali_pekerjaan} / {wali_hubungan}",
                            "Alamat Wali": wali_alamat,
                            "Mahar": mahar
                        },
                        "5. Saksi-Saksi Nikah": {
                            "Saksi 1 (Nama / NIK)": f"{saksi1_nama} (NIK: {saksi1_nik})",
                            "TTL / Pekerjaan Saksi 1": f"{saksi1_ttl} / {saksi1_pekerjaan}",
                            "Alamat Saksi 1": saksi1_alamat,
                            "Saksi 2 (Nama / NIK)": f"{saksi2_nama} (NIK: {saksi2_nik})",
                            "TTL / Pekerjaan Saksi 2": f"{saksi2_ttl} / {saksi2_pekerjaan}",
                            "Alamat Saksi 2": saksi2_alamat
                        }
                    }
                    pdf_buffer = generate_pdf_isian_data(pdf_data)
                    st.download_button(
                        label="📄 Download PDF (1 Halaman Full F4)",
                        data=pdf_buffer,
                        file_name=f"ISIAN_DATA_{wanita_nama}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                else:
                    st.warning("⚠️ Pustaka `reportlab` belum tersedia. Pastikan sudah dimasukkan ke requirements.txt.")

            wb.close()
            gc.collect()

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat memproses data: {e}")
