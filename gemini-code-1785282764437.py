import streamlit as st
import openpyxl
import re
from io import BytesIO
from datetime import datetime, date

from reportlab.lib.pagesizes import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- CONFIG PAGE ---
st.set_page_config(page_title="Verifikasi Catin Desa Tambi", page_icon="📄", layout="centered")

HARI_INDONESIA = {
    'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu',
    'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu', 'Sunday': 'Minggu'
}

def get_hari_tgl(tgl_obj):
    if not tgl_obj or tgl_obj == '-':
        return '-'
    if isinstance(tgl_obj, str):
        try:
            tgl_obj = date.fromisoformat(tgl_obj)
        except:
            return str(tgl_obj)
    if isinstance(tgl_obj, (date, datetime)):
        nama_hari = HARI_INDONESIA.get(tgl_obj.strftime('%A'), '')
        return f"{nama_hari}, {tgl_obj.strftime('%d-%m-%Y')}"
    return str(tgl_obj)

def hitung_umur(ttl_str):
    if not ttl_str or ttl_str == '-':
        return 0
    match = re.search(r'\b(19\d{2}|20\d{2})\b', str(ttl_str))
    if match:
        tahun_lahir = int(match.group(1))
        tahun_sekarang = datetime.now().year
        return max(0, tahun_sekarang - tahun_lahir)
    return 0

def clean_val(val):
    if val is None:
        return "-"
    val_str = str(val).strip()
    return val_str if val_str != "" else "-"

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
        'TitleStyle', parent=styles['Heading1'], fontSize=10.5, leading=12, alignment=1, fontName='Helvetica-Bold'
    )
    sec_title_style = ParagraphStyle(
        'SecTitleStyle', fontSize=8, leading=9.5, fontName='Helvetica-Bold', textColor=colors.HexColor('#003366')
    )
    lbl_style = ParagraphStyle('LblStyle', fontSize=7, leading=8.5, fontName='Helvetica-Bold')
    val_style = ParagraphStyle('ValStyle', fontSize=7, leading=8.5, fontName='Helvetica')
    
    elements = []
    
    elements.append(Paragraph("PEMERINTAH KABUPATEN PEMALANG - KECAMATAN WATUKUMPUL", ParagraphStyle('Kop1', fontSize=7.5, alignment=1, fontName='Helvetica-Bold')))
    elements.append(Paragraph("RINGKASAN LEMBAR VERIFIKASI BERKAS CALON PENGANTIN DESA TAMBI", title_style))
    elements.append(Paragraph(f"No. Register: <b>{clean_val(data.get('no_register'))}</b> | Tanggal Surat: <b>{clean_val(data.get('tgl_surat'))}</b>", ParagraphStyle('SubTitle', fontSize=7.5, alignment=1, leading=9)))
    elements.append(Spacer(1, 3))
    
    def row1(lbl, val):
        return [Paragraph(lbl, lbl_style), Paragraph(":", lbl_style), Paragraph(clean_val(val), val_style)]

    def row2(lbl1, val1, lbl2, val2):
        return [
            Paragraph(lbl1, lbl_style), Paragraph(":", lbl_style), Paragraph(clean_val(val1), val_style),
            Paragraph(lbl2, lbl_style), Paragraph(":", lbl_style), Paragraph(clean_val(val2), val_style)
        ]

    hari_tgl_akad = f"{get_hari_tgl(data.get('tgl_pelaksanaan'))} (Jam: {clean_val(data.get('jam_akad'))})"
    tabel_akad_data = [
        [Paragraph("I. PELAKSANAAN AKAD NIKAH", sec_title_style), "", ""],
        row1("Hari & Tgl / Jam Akad", hari_tgl_akad),
        row1("Tempat Akad Nikah", data.get('tempat_akad')),
        row1("Maskawin / Mahar", data.get('mahar')),
        row1("Email Catin", data.get('email_catin')),
    ]

    tabel_catin_data = [
        [Paragraph("II. CALON PENGANTIN LAKI-LAKI", sec_title_style), "", "", Paragraph("III. CALON PENGANTIN PEREMPUAN", sec_title_style), "", ""],
        row2("Nama Lengkap", data.get('nama_lk'), "Nama Lengkap", data.get('nama_pr')),
        row2("Bin / Binti", data.get('bin_lk'), "Bin / Binti", data.get('binti_pr')),
        row2("NIK", data.get('nik_lk'), "NIK", data.get('nik_pr')),
        row2("Tempat, Tgl Lahir", data.get('ttl_lk'), "Tempat, Tgl Lahir", data.get('ttl_pr')),
        row2("Umur Catin", f"{hitung_umur(data.get('ttl_lk'))} Tahun", "Umur Catin", f"{hitung_umur(data.get('ttl_pr'))} Tahun"),
        row2("Status / Gender", f"{clean_val(data.get('status_lk'))} / {clean_val(data.get('jk_lk'))}", "Status / Gender", f"{clean_val(data.get('status_pr'))} / {clean_val(data.get('jk_pr'))}"),
        row2("Pekerjaan", data.get('pekerjaan_lk'), "Pekerjaan", data.get('pekerjaan_pr')),
        row2("Pendidikan", data.get('pendidikan_lk'), "Pendidikan", data.get('pendidikan_pr')),
        row2("Ex Pasangan", data.get('istri_terdahulu'), "Ex Pasangan", data.get('suami_terdahulu')),
        row2("Alamat Lengkap", data.get('alamat_lk'), "Alamat Lengkap", data.get('alamat_pr')),
    ]

    tabel_ortu_data = [
        [Paragraph("IV. ORANG TUA LAKI-LAKI", sec_title_style), "", "", Paragraph("V. ORANG TUA PEREMPUAN", sec_title_style), "", ""],
        row2("Ayah / Bin", f"{clean_val(data.get('nama_ayah_lk'))} bin {clean_val(data.get('bin_ayah_lk'))}", "Ayah / Bin", f"{clean_val(data.get('nama_ayah_pr'))} bin {clean_val(data.get('bin_ayah_pr'))}"),
        row2("NIK / TTL Ayah", f"{clean_val(data.get('nik_ayah_lk'))} / {clean_val(data.get('ttl_ayah_lk'))}", "NIK / TTL Ayah", f"{clean_val(data.get('nik_ayah_pr'))} / {clean_val(data.get('ttl_ayah_pr'))}"),
        row2("Pekerjaan Ayah", data.get('pekerjaan_ayah_lk'), "Pekerjaan Ayah", data.get('pekerjaan_ayah_pr')),
        row2("Alamat Ayah", data.get('alamat_ayah_lk'), "Alamat Ayah", data.get('alamat_ayah_pr')),
        row2("Ibu / Binti", f"{clean_val(data.get('nama_ibu_lk'))} bin {clean_val(data.get('bin_ibu_lk'))}", "Ibu / Binti", f"{clean_val(data.get('nama_ibu_pr'))} bin {clean_val(data.get('bin_ibu_pr'))}"),
        row2("NIK / TTL Ibu", f"{clean_val(data.get('nik_ibu_lk'))} / {clean_val(data.get('ttl_ibu_lk'))}", "NIK / TTL Ibu", f"{clean_val(data.get('nik_ibu_pr'))} / {clean_val(data.get('ttl_ibu_pr'))}"),
        row2("Pekerjaan Ibu", data.get('pekerjaan_ibu_lk'), "Pekerjaan Ibu", data.get('pekerjaan_ibu_pr')),
        row2("Alamat Ibu", data.get('alamat_ibu_lk'), "Alamat Ibu", data.get('alamat_ibu_pr')),
    ]

    tabel_wali_saksi = [
        [Paragraph("VI. DATA WALI NIKAH", sec_title_style), "", "", Paragraph("VII. DATA SAKSI-SAKSI AKAD", sec_title_style), "", ""],
        row2("Nama Wali", f"{clean_val(data.get('nama_wali'))} bin {clean_val(data.get('bin_wali'))}", "Saksi 1", data.get('saksi1_nama')),
        row2("NIK / TTL Wali", f"{clean_val(data.get('nik_wali'))} / {clean_val(data.get('ttl_wali'))}", "NIK / TTL Saksi 1", f"{clean_val(data.get('saksi1_nik'))} / {clean_val(data.get('saksi1_ttl'))}"),
        row2("Hubungan Wali", data.get('hubungan_wali'), "Pekerjaan Saksi 1", data.get('saksi1_pekerjaan')),
        row2("Pekerjaan Wali", data.get('pekerjaan_wali'), "Alamat Saksi 1", data.get('saksi1_alamat')),
        row2("Alamat Wali", data.get('alamat_wali'), "Saksi 2", data.get('saksi2_nama')),
        row2("Wali Lengkap", data.get('nama_wali_lengkap'), "NIK / TTL Saksi 2", f"{clean_val(data.get('saksi2_nik'))} / {clean_val(data.get('saksi2_ttl'))}"),
        row2("", "", "Pekerjaan Saksi 2", data.get('saksi2_pekerjaan')),
        row2("", "", "Alamat Saksi 2", data.get('saksi2_alamat')),
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
        t4, Spacer(1, 5)
    ])
    
    ttd_catin = [
        [Paragraph("Catin Laki-Laki", lbl_style), Paragraph("Catin Perempuan", lbl_style), Paragraph("Wali Nikah", lbl_style)],
        [Spacer(1, 18), Spacer(1, 18), Spacer(1, 18)],
        [
            Paragraph(f"( <b>{clean_val(data.get('nama_lk'))}</b> )", val_style), 
            Paragraph(f"( <b>{clean_val(data.get('nama_pr'))}</b> )", val_style), 
            Paragraph(f"( <b>{clean_val(data.get('nama_wali'))}</b> )", val_style)
        ]
    ]
    t_ttd1 = Table(ttd_catin, colWidths=[67*mm, 67*mm, 67*mm])
    t_ttd1.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    
    tgl_surat_str = clean_val(data.get('tgl_surat'))
    if tgl_surat_str == '-':
        tgl_surat_str = 'Tambi, ................. 2026'
    else:
        tgl_surat_str = f"Tambi, {tgl_surat_str}"

    ttd_pemdes = [
        [
            Paragraph("Mengetahui,<br/><b>KASI PELAYANAN DESA TAMBI</b>", lbl_style), 
            Paragraph(f"{tgl_surat_str}<br/><b>KEPALA DESA TAMBI</b>", lbl_style)
        ],
        [Spacer(1, 22), Spacer(1, 22)],
        [
            Paragraph("<b><u>CHALIM MUCHTAROM, S.Pd.I</u></b>", val_style), 
            Paragraph("<b><u>JURI</u></b>", val_style)
        ]
    ]
    t_ttd2 = Table(ttd_pemdes, colWidths=[100*mm, 101*mm])
    t_ttd2.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))

    elements.extend([t_ttd1, Spacer(1, 6), t_ttd2])
    
    doc.build(elements)
    pdf_data = buffer.getvalue()
    buffer.close()
    return pdf_data

# --- APP LAYOUT ---
st.title("📄 Verifikasi Berkas Catin Desa Tambi")
st.write("Upload file Excel data Catin untuk membuat ringkasan verifikasi PDF.")

if 'data_catin' not in st.session_state:
    st.session_state.data_catin = None

uploaded_file = st.file_uploader("Upload File Excel (.xlsx)", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        ws = wb.active
        
        # Penarikan data dari cell Excel (Sesuaikan koordinat cell jika ada perubahan)
        st.session_state.data_catin = {
            'no_register': ws['C4'].value,
            'tgl_surat': ws['C5'].value,
            'tgl_pelaksanaan': ws['C8'].value,
            'jam_akad': ws['C9'].value,
            'tempat_akad': ws['C10'].value,
            'mahar': ws['C11'].value,
            'email_catin': ws['C12'].value,
            
            'nama_lk': ws['C15'].value,
            'bin_lk': ws['C16'].value,
            'nik_lk': ws['C17'].value,
            'ttl_lk': ws['C18'].value,
            'status_lk': ws['C19'].value,
            'jk_lk': ws['C20'].value,
            'pekerjaan_lk': ws['C21'].value,
            'pendidikan_lk': ws['C22'].value,
            'istri_terdahulu': ws['C23'].value,
            'alamat_lk': ws['C24'].value,

            'nama_pr': ws['F15'].value,
            'binti_pr': ws['F16'].value,
            'nik_pr': ws['F17'].value,
            'ttl_pr': ws['F18'].value,
            'status_pr': ws['F19'].value,
            'jk_pr': ws['F20'].value,
            'pekerjaan_pr': ws['F21'].value,
            'pendidikan_pr': ws['F22'].value,
            'suami_terdahulu': ws['F23'].value,
            'alamat_pr': ws['F24'].value,

            'nama_ayah_lk': ws['C27'].value,
            'bin_ayah_lk': ws['C28'].value,
            'nik_ayah_lk': ws['C29'].value,
            'ttl_ayah_lk': ws['C30'].value,
            'pekerjaan_ayah_lk': ws['C31'].value,
            'alamat_ayah_lk': ws['C32'].value,

            'nama_ibu_lk': ws['C34'].value,
            'bin_ibu_lk': ws['C35'].value,
            'nik_ibu_lk': ws['C36'].value,
            'ttl_ibu_lk': ws['C37'].value,
            'pekerjaan_ibu_lk': ws['C38'].value,
            'alamat_ibu_lk': ws['C39'].value,

            'nama_ayah_pr': ws['F27'].value,
            'bin_ayah_pr': ws['F28'].value,
            'nik_ayah_pr': ws['F29'].value,
            'ttl_ayah_pr': ws['F30'].value,
            'pekerjaan_ayah_pr': ws['F31'].value,
            'alamat_ayah_pr': ws['F32'].value,

            'nama_ibu_pr': ws['F34'].value,
            'bin_ibu_pr': ws['F35'].value,
            'nik_ibu_pr': ws['F36'].value,
            'ttl_ibu_pr': ws['F37'].value,
            'pekerjaan_ibu_pr': ws['F38'].value,
            'alamat_ibu_pr': ws['F39'].value,

            'nama_wali': ws['C42'].value,
            'bin_wali': ws['C43'].value,
            'nik_wali': ws['C44'].value,
            'ttl_wali': ws['C45'].value,
            'hubungan_wali': ws['C46'].value,
            'pekerjaan_wali': ws['C47'].value,
            'alamat_wali': ws['C48'].value,
            'nama_wali_lengkap': ws['C49'].value,

            'saksi1_nama': ws['F42'].value,
            'saksi1_nik': ws['F43'].value,
            'saksi1_ttl': ws['F44'].value,
            'saksi1_pekerjaan': ws['F45'].value,
            'saksi1_alamat': ws['F46'].value,

            'saksi2_nama': ws['F48'].value,
            'saksi2_nik': ws['F49'].value,
            'saksi2_ttl': ws['F50'].value,
            'saksi2_pekerjaan': ws['F51'].value,
            'saksi2_alamat': ws['F52'].value,
        }
        st.success("File Excel berhasil diproses!")
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")

if st.session_state.data_catin is not None:
    try:
        pdf_bytes = generate_pdf_f4(st.session_state.data_catin)
        
        nama_lk = clean_val(st.session_state.data_catin.get('nama_lk'))
        nama_pr = clean_val(st.session_state.data_catin.get('nama_pr'))
        file_name_pdf = f"Ringkasan_Verifikasi_Catin_{nama_lk}_dan_{nama_pr}.pdf"
        
        st.download_button(
            label="📥 Download PDF Ringkasan Verifikasi (F4)",
            data=pdf_bytes,
            file_name=file_name_pdf,
            mime="application/pdf"
        )
    except Exception as e:
        st.error(f"Gagal membuat PDF: {e}")
