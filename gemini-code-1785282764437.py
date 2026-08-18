import re
from datetime import date, datetime
from io import BytesIO

import openpyxl
from reportlab.lib import colors

# Import ReportLab untuk generate PDF 1 Halaman
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st

st.set_page_config(
    page_title="Aplikasi Berkas Catin - Desa Tambi",
    page_icon="📜",
    layout="wide",
)

st.title("📜 Form Input Berkas Catin")
st.caption(
    "Aplikasi ini mengisi nilai pada sheet 'ISIAN DATA'. Umur dan hari akad"
    " dihitung otomatis di latar belakang."
)

# Membuka file Excel template utama
EXCEL_FILE = "BERKAS CATIN .xlsx"

# Dictionary konversi hari ke Bahasa Indonesia
NAMA_HARI = {
    "Monday": "Senin",
    "Tuesday": "Selasa",
    "Wednesday": "Rabu",
    "Thursday": "Kamis",
    "Friday": "Jumat",
    "Saturday": "Sabtu",
    "Sunday": "Minggu",
}


# FUNGSI HELPER HITUNG UMUR OTOMATIS
def hitung_umur(ttl_str):
  """Menghitung umur otomatis dengan mengekstrak 4 digit tahun lahir dari teks

  Tempat Tanggal Lahir (contoh: 'Pemalang, 18 Februari 1999').
  """
  if not ttl_str:
    return 0
  match = re.search(r"\b(19\d{2}|20\d{2})\b", str(ttl_str))
  if match:
    tahun_lahir = int(match.group(1))
    tahun_sekarang = datetime.now().year
    return max(0, tahun_sekarang - tahun_lahir)
  return 0


# FUNGSI MEMBUAT PDF REKAP ISIAN DATA
def generate_pdf_rekap(data_dict):
  buffer = BytesIO()

  doc = SimpleDocTemplate(
      buffer,
      pagesize=A4,
      leftMargin=10,
      rightMargin=10,
      topMargin=10,
      bottomMargin=10,
  )

  styles = getSampleStyleSheet()

  style_title = ParagraphStyle(
      "TitleStyle",
      parent=styles["Heading1"],
      fontSize=12,
      leading=14,
      alignment=1,
      fontName="Helvetica-Bold",
      textColor=colors.HexColor("#0f172a"),
      spaceAfter=2,
  )

  style_sub = ParagraphStyle(
      "SubTitleStyle",
      parent=styles["Normal"],
      fontSize=9,
      leading=11,
      alignment=1,
      fontName="Helvetica-Bold",
      textColor=colors.HexColor("#334155"),
  )

  style_section = ParagraphStyle(
      "SecStyle",
      parent=styles["Normal"],
      fontSize=8.5,
      leading=10,
      fontName="Helvetica-Bold",
      textColor=colors.HexColor("#1e3a8a"),
  )

  style_cell_label = ParagraphStyle(
      "LabelStyle",
      parent=styles["Normal"],
      fontSize=8,
      leading=9.5,
      fontName="Helvetica-Bold",
      textColor=colors.HexColor("#1e293b"),
  )
  style_cell_val = ParagraphStyle(
      "ValStyle",
      parent=styles["Normal"],
      fontSize=8,
      leading=9.5,
      fontName="Helvetica",
      textColor=colors.HexColor("#0f172a"),
  )

  style_ttd_title = ParagraphStyle(
      "TTDTitle",
      parent=styles["Normal"],
      fontSize=8.5,
      leading=10,
      fontName="Helvetica-Bold",
      alignment=2,
  )
  style_ttd_name = ParagraphStyle(
      "TTDName",
      parent=styles["Normal"],
      fontSize=8.5,
      leading=10,
      fontName="Helvetica-Bold",
      alignment=2,
  )

  elements = []

  elements.append(
      Paragraph("REKAP ISIAN DATA BERKAS CATIN DESA TAMBI", style_title)
  )
  elements.append(
      Paragraph(
          f"<b>No. Register:</b> {data_dict.get('G2','')} &nbsp;|&nbsp; <b>Tanggal"
          f" Surat:</b> {data_dict.get('H3','')}",
          style_sub,
      )
  )
  elements.append(Spacer(1, 4))

  def make_p(txt, is_label=False):
    st_use = style_cell_label if is_label else style_cell_val
    return Paragraph(str(txt) if txt else "-", st_use)

  table_data = [
      # SECTION 1
      [
          Paragraph("1. PELAKSANAAN AKAD NIKAH", style_section),
          "",
          Paragraph("2. CATIN LAKI-LAKI", style_section),
          "",
      ],
      [
          make_p("Hari / Tgl Akad", True),
          make_p(f"{data_dict.get('I4','')} / {data_dict.get('G4','')}"),
          make_p("Nama / Bin", True),
          make_p(f"{data_dict.get('G8','')} bin {data_dict.get('G9','')}"),
      ],
      [
          make_p("Jam / Tempat", True),
          make_p(f"{data_dict.get('G5','')} @ {data_dict.get('G6','')}"),
          make_p("NIK / TTL", True),
          make_p(
              f"{data_dict.get('G11','')} / {data_dict.get('G10','')}"
              f" ({data_dict.get('K10','')} th)"
          ),
      ],
      [
          make_p("Mahar", True),
          make_p(data_dict.get("G65", "")),
          make_p("Pekerjaan / Status", True),
          make_p(f"{data_dict.get('G12','')} / {data_dict.get('G13','')}"),
      ],
      [
          make_p("Pendidikan LK", True),
          make_p(data_dict.get("G17", "")),
          make_p("Alamat LK", True),
          make_p(data_dict.get("G16", "")),
      ],
      # SECTION 2
      [
          Paragraph("3. AYAH & IBU CATIN LAKI-LAKI", style_section),
          "",
          Paragraph("4. CATIN PEREMPUAN", style_section),
          "",
      ],
      [
          make_p("Ayah LK", True),
          make_p(
              f"{data_dict.get('G19','')} bin {data_dict.get('J19','')}"
              f" ({data_dict.get('G20','')})"
          ),
          make_p("Nama / Binti", True),
          make_p(f"{data_dict.get('G34','')} binti {data_dict.get('G35','')}"),
      ],
      [
          make_p("Pekerjaan/Alamat", True),
          make_p(f"{data_dict.get('G22','')} / {data_dict.get('G23','')}"),
          make_p("NIK / TTL", True),
          make_p(
              f"{data_dict.get('G37','')} / {data_dict.get('G36','')}"
              f" ({data_dict.get('K36','')} th)"
          ),
      ],
      [
          make_p("Ibu LK", True),
          make_p(
              f"{data_dict.get('G26','')} bin {data_dict.get('I26','')}"
              f" ({data_dict.get('G27','')})"
          ),
          make_p("Pekerjaan / Status", True),
          make_p(f"{data_dict.get('G38','')} / {data_dict.get('G39','')}"),
      ],
      [
          make_p("Pekerjaan/Alamat", True),
          make_p(f"{data_dict.get('G29','')} / {data_dict.get('G30','')}"),
          make_p("Alamat PR", True),
          make_p(data_dict.get("G41", "")),
      ],
      # SECTION 3
      [
          Paragraph("5. AYAH & IBU CATIN PEREMPUAN", style_section),
          "",
          Paragraph("6. DATA WALI NIKAH", style_section),
          "",
      ],
      [
          make_p("Ayah PR", True),
          make_p(
              f"{data_dict.get('G45','')} bin {data_dict.get('I45','')}"
              f" ({data_dict.get('G46','')})"
          ),
          make_p("Nama Wali", True),
          make_p(f"{data_dict.get('G68','')} ({data_dict.get('G64','')})"),
      ],
      [
          make_p("Pekerjaan/Alamat", True),
          make_p(f"{data_dict.get('G48','')} / {data_dict.get('G49','')}"),
          make_p("NIK / TTL / Umur", True),
          make_p(
              f"{data_dict.get('G60','')} / {data_dict.get('G61','')}"
              f" ({data_dict.get('K61','')} th)"
          ),
      ],
      [
          make_p("Ibu PR", True),
          make_p(
              f"{data_dict.get('G52','')} bin {data_dict.get('I52','')}"
              f" ({data_dict.get('G53','')})"
          ),
          make_p("Pekerjaan / Alamat", True),
          make_p(f"{data_dict.get('G62','')} / {data_dict.get('G63','')}"),
      ],
      [
          make_p("Pekerjaan/Alamat", True),
          make_p(f"{data_dict.get('G55','')} / {data_dict.get('G56','')}"),
          make_p("Mahar", True),
          make_p(data_dict.get("G65", "")),
      ],
      # SECTION 4
      [Paragraph("7. SAKSI 1 & SAKSI 2", style_section), "", "", ""],
      [
          make_p("Saksi 1", True),
          make_p(
              f"{data_dict.get('G70','')} | NIK: {data_dict.get('G72','')} |"
              f" TTL: {data_dict.get('G71','')} ({data_dict.get('K71','')} th) |"
              f" Pekerjaan: {data_dict.get('G73','')} | Alamat:"
              f" {data_dict.get('G74','')}"
          ),
      ],
      [
          make_p("Saksi 2", True),
          make_p(
              f"{data_dict.get('G76','')} | NIK: {data_dict.get('G78','')} |"
              f" TTL: {data_dict.get('G77','')} ({data_dict.get('K77','')} th) |"
              f" Pekerjaan: {data_dict.get('G79','')} | Alamat:"
              f" {data_dict.get('G80','')}"
          ),
      ],
  ]

  t = Table(table_data, colWidths=[98, 189, 98, 190])
  t.setStyle(
      TableStyle([
          ("SPAN", (0, 0), (1, 0)),
          ("SPAN", (2, 0), (3, 0)),
          ("SPAN", (0, 5), (1, 5)),
          ("SPAN", (2, 5), (3, 5)),
          ("SPAN", (0, 10), (1, 10)),
          ("SPAN", (2, 10), (3, 10)),
          ("SPAN", (0, 15), (3, 15)),
          ("SPAN", (1, 16), (3, 16)),
          ("SPAN", (1, 17), (3, 17)),
          ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8fafc")),
          ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#94a3b8")),
          ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
          ("TOPPADDING", (0, 0), (-1, -1), 2.5),
          ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
          ("LEFTPADDING", (0, 0), (-1, -1), 4),
          ("RIGHTPADDING", (0, 0), (-1, -1), 4),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#e2e8f0")),
          ("BACKGROUND", (2, 0), (3, 0), colors.HexColor("#e2e8f0")),
          ("BACKGROUND", (0, 5), (1, 5), colors.HexColor("#e2e8f0")),
          ("BACKGROUND", (2, 5), (3, 5), colors.HexColor("#e2e8f0")),
          ("BACKGROUND", (0, 10), (1, 10), colors.HexColor("#e2e8f0")),
          ("BACKGROUND", (2, 10), (3, 10), colors.HexColor("#e2e8f0")),
          ("BACKGROUND", (0, 15), (3, 15), colors.HexColor("#e2e8f0")),
      ])
  )

  elements.append(t)
  elements.append(Spacer(1, 8))

  ttd_data = [
      [
          "",
          Paragraph(
              "Tambi, " + data_dict.get("H3", "").replace("TAMBI, ", ""),
              style_ttd_title,
          ),
      ],
      [
          "",
          Paragraph(
              "Yang Mengantar,<br/><b>Kasi Pelayanan</b>", style_ttd_title
          ),
      ],
      ["", Spacer(1, 22)],
      ["", Paragraph("<u><b>CHALIM MUCHTAROM, S.Pd.I</b></u>", style_ttd_name)],
  ]

  table_ttd = Table(ttd_data, colWidths=[370, 205])
  table_ttd.setStyle(
      TableStyle([
          ("VALIGN", (0, 0), (-1, -1), "TOP"),
          ("ALIGN", (1, 0), (1, -1), "RIGHT"),
          ("TOPPADDING", (0, 0), (-1, -1), 1),
          ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
      ])
  )

  elements.append(table_ttd)
  doc.build(elements)
  buffer.seek(0)
  return buffer


# FORMULIR INPUT DATA
with st.form("form_catin"):
  tab1, tab2, tab3, tab4, tab5 = st.tabs([
      "📝 Register & Akad",
      "👨 Catin Laki-Laki & Ortu",
      "👩 Catin Perempuan & Ortu",
      "🤝 Data Wali",
      "👥 Data Saksi 1 & 2",
  ])

  with tab1:
    st.subheader("Surat & Pelaksanaan Akad Nikah")
    col1, col2 = st.columns(2)
    with col1:
      no_register = st.text_input(
          "Nomor Register", value="400.12.3.2/010/ VIII/ 2026"
      )
      tgl_surat = st.text_input(
          "Tanggal Surat (Format: TAMBI, DD AGUSTUS YYYY)",
          value="TAMBI, 11 AGUSTUS 2026",
      )
      tgl_pelaksanaan = st.date_input(
          "Tanggal Pelaksanaan Akad", value=date(2026, 9, 7)
      )
      jam_akad = st.text_input("Jam Akad", value="JAM. 08.00")
    with col2:
      tempat_akad = st.text_input(
          "Tempat Akad Nikah",
          value=(
              "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"
          ),
      )
      email_catin = st.text_input("Email Catin", value="")
      mahar = st.text_input("Maskawin / Mahar", value="Seperangkat Alat Sholat")

  with tab2:
    st.subheader("Data Calon Pengantin Laki-Laki")
    col_lk1, col_lk2 = st.columns(2)
    with col_lk1:
      nama_lk = st.text_input(
          "Nama Calon Pengantin Laki-Laki", value="Miftahul Anam"
      )
      bin_lk = st.text_input("Bin (Ayah Laki-Laki)", value="Nur Karim")
      ttl_lk = st.text_input(
          "Tempat, Tanggal Lahir Laki-Laki", value="Pemalang, 18 Februari 1999"
      )
      nik_lk = st.text_input("NIK Laki-Laki", value="3327031802990004")
    with col_lk2:
      pekerjaan_lk = st.text_input("Pekerjaan Laki-Laki", value="Swasta")
      status_lk = st.text_input("Status Laki-Laki", value="BELUM KAWIN")
      jk_lk = st.text_input("Jenis Kelamin Laki-Laki", value="Laki-Laki")
      istri_terdahulu = st.text_input(
          "Nama Istri Terdahulu (Jika ada)", value=""
      )
      alamat_lk = st.text_area(
          "Alamat Laki-Laki",
          value=(
              "RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang"
          ),
      )
      pendidikan_lk = st.text_input("Pendidikan Laki-Laki", value="SLTA")

    st.divider()
    st.subheader("Data Ayah & Ibu Laki-Laki")
    col_alk, col_ilk = st.columns(2)
    with col_alk:
      st.markdown("**Ayah Laki-Laki**")
      nama_ayah_lk = st.text_input("Nama Ayah Laki-Laki", value="Nur Karim")
      bin_ayah_lk = st.text_input("bin (Kakek Laki-Laki)", value="Kasturi")
      nik_ayah_lk = st.text_input(
          "NIK Ayah Laki-Laki", value="3327030608680006"
      )
      ttl_ayah_lk = st.text_input(
          "TTL Ayah Laki-Laki", value="Pemalang, 06 Agustus 1968"
      )
      pekerjaan_ayah_lk = st.text_input(
          "Pekerjaan Ayah Laki-Laki", value="PETANI/ PEKEBUN"
      )
      alamat_ayah_lk = st.text_area(
          "Alamat Ayah Laki-Laki",
          value=(
              "RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang"
          ),
      )

    with col_ilk:
      st.markdown("**Ibu Laki-Laki**")
      nama_ibu_lk = st.text_input("Nama Ibu Laki-Laki", value="Samijah")
      bin_ibu_lk = st.text_input(
          "bin (Kakek dari Ibu Laki-Laki)", value="Taryad"
      )
      nik_ibu_lk = st.text_input("NIK Ibu Laki-Laki", value="3327035405740004")
      ttl_ibu_lk = st.text_input(
          "TTL Ibu Laki-Laki", value="Pemalang, 14 Mei 1974"
      )
      pekerjaan_ibu_lk = st.text_input(
          "Pekerjaan Ibu Laki-Laki", value="Mengurus Rumah Tangga"
      )
      alamat_ibu_lk = st.text_area(
          "Alamat Ibu Laki-Laki",
          value=(
              "RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang"
          ),
      )

  with tab3:
    st.subheader("Data Calon Pengantin Perempuan")
    col_pr1, col_pr2 = st.columns(2)
    with col_pr1:
      nama_pr = st.text_input(
          "Nama Calon Pengantin Perempuan", value="Diyan Solehatin"
      )
      binti_pr = st.text_input("Binti (Ayah Perempuan)", value="Disun")
      ttl_pr = st.text_input(
          "Tempat, Tanggal Lahir Perempuan", value="Pemalang, 29 Juni 2007"
      )
      nik_pr = st.text_input("NIK Perempuan", value="3327046906070010")
    with col_pr2:
      pekerjaan_pr = st.text_input(
          "Pekerjaan Perempuan", value="BELUM/ TIDAK BEKERJA"
      )
      status_pr = st.text_input("Status Perempuan", value="BELUM KAWIN")
      jk_pr = st.text_input("Jenis Kelamin Perempuan", value="PEREMPUAN")
      suami_terdahulu = st.text_input(
          "Nama Suami Terdahulu (Jika ada)", value=""
      )
      alamat_pr = st.text_area(
          "Alamat Perempuan",
          value=(
              "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"
          ),
      )
      pendidikan_pr = st.text_input("Pendidikan Perempuan", value="SLTP")

    st.divider()
    st.subheader("Data Ayah & Ibu Perempuan")
    col_apr, col_ipr = st.columns(2)
    with col_apr:
      st.markdown("**Ayah Perempuan**")
      nama_ayah_pr = st.text_input("Nama Ayah Perempuan", value="Disun")
      bin_ayah_pr = st.text_input("bin (Kakek Perempuan)", value="Tawiroji")
      nik_ayah_pr = st.text_input(
          "NIK Ayah Perempuan", value="3327042504840003"
      )
      ttl_ayah_pr = st.text_input(
          "TTL Ayah Perempuan", value="Pemalang, 21 April 1984"
      )
      pekerjaan_ayah_pr = st.text_input(
          "Pekerjaan Ayah Perempuan", value="PETANI/ PEKEBUN"
      )
      alamat_ayah_pr = st.text_area(
          "Alamat Ayah Perempuan",
          value=(
              "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"
          ),
      )

    with col_ipr:
      st.markdown("**Ibu Perempuan**")
      nama_ibu_pr = st.text_input("Nama Ibu Perempuan", value="Mutirah")
      bin_ibu_pr = st.text_input(
          "bin (Kakek dari Ibu Perempuan)", value="Tamiarjo"
      )
      nik_ibu_pr = st.text_input("NIK Ibu Perempuan", value="3327044411840003")
      ttl_ibu_pr = st.text_input(
          "TTL Ibu Perempuan", value="Pemalang, 04 November 1984"
      )
      pekerjaan_ibu_pr = st.text_input(
          "Pekerjaan Ibu Perempuan", value="Mengurus Rumah Tangga"
      )
      alamat_ibu_pr = st.text_area(
          "Alamat Ibu Perempuan",
          value=(
              "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"
          ),
      )

  with tab4:
    st.subheader("Data Wali Nikah")
    col_w1, col_w2 = st.columns(2)
    with col_w1:
      nama_wali = st.text_input("Nama Wali", value="Disun")
      bin_wali = st.text_input("Bin Wali", value="Tawiroji")
      nik_wali = st.text_input("NIK Wali", value="3327042504840003")
      ttl_wali = st.text_input("TTL Wali", value="Pemalang, 21 April 1984")
    with col_w2:
      pekerjaan_wali = st.text_input(
          "Pekerjaan Wali", value="PETANI/ PEKEBUN"
      )
      alamat_wali = st.text_area(
          "Alamat Wali",
          value=(
              "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"
          ),
      )
      hubungan_wali = st.text_input("Hubungan Wali", value="AYAH KANDUNG")
      nama_wali_lengkap = st.text_input(
          "Nama Wali Lengkap (Nama Bin)", value="Disun Bin Tawiroji"
      )

  with tab5:
    col_s1, col_s2 = st.columns(2)
    with col_s1:
      st.subheader("Data Saksi 1")
      saksi1_nama = st.text_input("Nama Saksi 1", value="Chalim Muchtarom")
      saksi1_ttl = st.text_input(
          "TTL Saksi 1", value="Pemalang, 21 Oktober 1989"
      )
      saksi1_nik = st.text_input("NIK Saksi 1", value="3327042110890004")
      saksi1_pekerjaan = st.text_input(
          "Pekerjaan Saksi 1", value="Perangkat Desa"
      )
      saksi1_alamat = st.text_area(
          "Alamat Saksi 1",
          value=(
              "RT 002 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"
          ),
      )

    with col_s2:
      st.subheader("Data Saksi 2")
      saksi2_nama = st.text_input("Nama Saksi 2", value="Sidin")
      saksi2_ttl = st.text_input("TTL Saksi 2", value="Pemalang, 15 Mei 1980")
      saksi2_nik = st.text_input("NIK Saksi 2", value="3327041505800002")
      saksi2_pekerjaan = st.text_input(
          "Pekerjaan Saksi 2", value="Petani/Pekebun"
      )
      saksi2_alamat = st.text_area(
          "Alamat Saksi 2",
          value=(
              "RT 003 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"
          ),
      )

  submit = st.form_submit_button("💾 ISIKAN KE EXCEL & GENERATE BERKAS")

if submit:
  try:
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=False)

    # BARIS BARU: Mengubah ukuran kertas seluruh sheet di workbook menjadi F4 (paperSize = 14)
    for ws in wb.worksheets:
      ws.page_setup.paperSize = 14

    sheet = wb["ISIAN DATA"]

    hari_akad = NAMA_HARI.get(tgl_pelaksanaan.strftime("%A"), "")

    # HITUNG UMUR OTOMATIS SAAT SUBMIT
    umur_lk = hitung_umur(ttl_lk)
    umur_ayah_lk = hitung_umur(ttl_ayah_lk)
    umur_ibu_lk = hitung_umur(ttl_ibu_lk)
    umur_pr = hitung_umur(ttl_pr)
    umur_ayah_pr = hitung_umur(ttl_ayah_pr)
    umur_ibu_pr = hitung_umur(ttl_ibu_pr)
    umur_wali = hitung_umur(ttl_wali)
    saksi1_umur = hitung_umur(saksi1_ttl)
    saksi2_umur = hitung_umur(saksi2_ttl)

    cell_updates = {
        "G2": no_register,
        "H3": tgl_surat,
        "G4": tgl_pelaksanaan.strftime("%Y-%m-%d"),
        "I4": hari_akad,
        "G5": jam_akad,
        "G6": tempat_akad,
        "G8": nama_lk,
        "G9": bin_lk,
        "G10": ttl_lk,
        "K10": umur_lk,
        "G11": nik_lk,
        "G12": pekerjaan_lk,
        "G13": status_lk,
        "G14": jk_lk,
        "G15": istri_terdahulu,
        "G16": alamat_lk,
        "G17": pendidikan_lk,
        "G19": nama_ayah_lk,
        "J19": bin_ayah_lk,
        "G20": nik_ayah_lk,
        "G21": ttl_ayah_lk,
        "K21": umur_ayah_lk,
        "G22": pekerjaan_ayah_lk,
        "G23": alamat_ayah_lk,
        "G26": nama_ibu_lk,
        "I26": bin_ibu_lk,
        "G27": nik_ibu_lk,
        "G28": ttl_ibu_lk,
        "K28": umur_ibu_lk,
        "G29": pekerjaan_ibu_lk,
        "G30": alamat_ibu_lk,
        "G34": nama_pr,
        "G35": binti_pr,
        "G36": ttl_pr,
        "K36": umur_pr,
        "G37": nik_pr,
        "G38": pekerjaan_pr,
        "G39": status_pr,
        "G40": jk_pr,
        "G41": alamat_pr,
        "G42": suami_terdahulu,
        "G43": pendidikan_pr,
        "G45": nama_ayah_pr,
        "I45": bin_ayah_pr,
        "G46": nik_ayah_pr,
        "G47": ttl_ayah_pr,
        "K47": umur_ayah_pr,
        "G48": pekerjaan_ayah_pr,
        "G49": alamat_ayah_pr,
        "G52": nama_ibu_pr,
        "I52": bin_ibu_pr,
        "G53": nik_ibu_pr,
        "G54": ttl_ibu_pr,
        "K54": umur_ibu_pr,
        "G55": pekerjaan_ibu_pr,
        "G56": alamat_ibu_pr,
        "G58": nama_wali,
        "G59": bin_wali,
        "G60": nik_wali,
        "G61": ttl_wali,
        "K61": umur_wali,
        "G62": pekerjaan_wali,
        "G63": alamat_wali,
        "G64": hubungan_wali,
        "G65": mahar,
        "G68": nama_wali_lengkap,
        "G70": saksi1_nama,
        "G71": saksi1_ttl,
        "K71": saksi1_umur,
        "G72": saksi1_nik,
        "G73": saksi1_pekerjaan,
        "G74": saksi1_alamat,
        "G76": saksi2_nama,
        "G77": saksi2_ttl,
        "K77": saksi2_umur,
        "G78": saksi2_nik,
        "G79": saksi2_pekerjaan,
        "G80": saksi2_alamat,
    }

    for cell_ref, val in cell_updates.items():
      sheet[cell_ref] = val

    output_excel = BytesIO()
    wb.save(output_excel)
    output_excel.seek(0)

    output_pdf = generate_pdf_rekap(cell_updates)

    st.success(
        "✅ Success! Data berhasil diisikan ke sheet ISIAN DATA. Seluruh sheet"
        " otomatis tersetel ke ukuran kertas F4."
    )

    filename_base = f"BERKAS_CATIN_{nama_lk}_{nama_pr}".replace(" ", "_")

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
      st.download_button(
          label="📥 Download File Excel Berkas Catin",
          data=output_excel,
          file_name=f"{filename_base}.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    with col_dl2:
      st.download_button(
          label="📄 Download Rekap Isian Data (PDF 1 Lembar Full)",
          data=output_pdf,
          file_name=f"REKAP_ISIAN_DATA_{filename_base}.pdf",
          mime="application/pdf",
      )

  except Exception as e:
    st.error(f"Gagal memproses berkas: {e}")
