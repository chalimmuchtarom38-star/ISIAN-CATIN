import os
import pandas as pd
import streamlit as st
from io import BytesIO
from reportlab.lib.pagesizes import landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Form Input Catin F4", layout="wide")
st.title("📝 Formulir Input Isian Data Catin (F4)")

FILE_NAME = "BERKAS_CATIN_F4.xlsb"

if not os.path.exists(FILE_NAME):
    st.error(f"File '{FILE_NAME}' tidak ditemukan di folder aplikasi! Pastikan nama file sudah pas.")
    st.stop()

# Daftar Field Seleksi Khusus dari Kolom B (Baris 2 s/d 80)
FIELD_MAPPING = [
    # --- DATA SURAT & AKAD NIKAH ---
    ("B2", "Nomor Register", "400.12.3.2/010/ VIII/ 2026"),
    ("B3", "Tanggal Surat", "TAMBI, 11 AGUSTUS 2026"),
    ("B4", "Tanggal Pelaksanaan", ""),
    ("B5", "Jam Pelaksanaan", "JAM. 08.00"),
    ("B6", "Tempat Akad Nikah", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"),
    
    # --- CATIN LAKI-LAKI ---
    ("B8", "Nama Catin Laki-Laki", "MIfahul Anam"),
    ("B9", "Bin (Ayah Laki-Laki)", "Nur Karim"),
    ("B10", "Tempat, Tanggal Lahir (Laki-Laki)", "Pemalang, 18 Februari 1999"),
    ("B11", "NIK (Laki-Laki)", "3327031802990004"),
    ("B12", "Pekerjaan (Laki-Laki)", "Swasta"),
    ("B13", "Status (Laki-Laki)", "BELUM KAWIN"),
    ("B14", "Jenis Kelamin (Laki-Laki)", "Laki-Laki"),
    ("B15", "Nama Istri Terdahulu", ""),
    ("B16", "Alamat (Laki-Laki)", "RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang"),
    ("B17", "Pendidikan (Laki-Laki)", "SLTA"),
    
    # --- ORANG TUA LAKI-LAKI ---
    ("B19", "Nama Ayah Laki-Laki", "Nur Karim"),
    ("B20", "NIK Ayah Laki-Laki", "3327030608680006"),
    ("B21", "TTL Ayah Laki-Laki", "Pemalang, 06 Agustus 1968"),
    ("B22", "Pekerjaan Ayah Laki-Laki", "PETANI/ PEKEBUN"),
    ("B23", "Alamat Ayah Laki-Laki", "RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang"),
    ("B26", "Nama Ibu Laki-Laki", "Samijah"),
    ("B27", "NIK Ibu Laki-Laki", "3327035405740004"),
    ("B28", "TTL Ibu Laki-Laki", "Pemalang, 14 Mei 1974"),
    ("B29", "Pekerjaan Ibu Laki-Laki", "Mengurus Rumah Tangga"),
    ("B30", "Alamat Ibu Laki-Laki", "RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang"),
    
    # --- CATIN PEREMPUAN ---
    ("B34", "Nama Catin Perempuan", "Diyan Solehatin"),
    ("B35", "Binti (Ayah Perempuan)", "Disun"),
    ("B36", "Tempat, Tanggal Lahir (Perempuan)", "Pemalang, 29 Juni 2007"),
    ("B37", "NIK (Perempuan)", "3327046906070010"),
    ("B38", "Pekerjaan (Perempuan)", "BELUM/ TIDAK BEKERJA"),
    ("B39", "Status (Perempuan)", "BELUM KAWIN"),
    ("B40", "Jenis Kelamin (Perempuan)", "PEREMPUAN"),
    ("B41", "Alamat (Perempuan)", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"),
    ("B42", "Nama Suami Terdahulu", ""),
    ("B43", "Pendidikan (Perempuan)", "SLTP"),
    
    # --- ORANG TUA PEREMPUAN ---
    ("B45", "Nama Ayah Perempuan", "Disun"),
    ("B46", "NIK Ayah Perempuan", "3327042504840003"),
    ("B47", "TTL Ayah Perempuan", "Pemalang, 21 April 1989"),
    ("B48", "Pekerjaan Ayah Perempuan", "PETANI/ PEKEBUN"),
    ("B49", "Alamat Ayah Perempuan", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"),
    ("B52", "Nama Ibu Perempuan", "Mutirah"),
    ("B53", "NIK Ibu Perempuan", "3327044411840003"),
    ("B54", "TTL Ibu Perempuan", "Pemalang, 04 November 1984"),
    ("B55", "Pekerjaan Ibu Perempuan", "Mengurus Rumah Tangga"),
    ("B56", "Alamat Ibu Perempuan", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"),
    
    # --- WALI & SAKSI ---
    ("B58", "Nama Wali", "Disun"),
    ("B59", "Bin Wali", "Tawiroji"),
    ("B60", "NIK Wali", "3327042504840003"),
    ("B61", "TTL Wali", "PEMALANG, 21 April 1984"),
    ("B62", "Pekerjaan Wali", "PETANI/ PEKEBUN"),
    ("B63", "Alamat Wali", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"),
    ("B64", "Hubungan Wali", "AYAH KANDUNG"),
    ("B65", "Mahar / Maskawin", "Seperangkat Alat Sholat"),
    ("B70", "Nama Saksi 1", "Chalim Muchtarom"),
    ("B71", "TTL Saksi 1", "Pemalang, 21 Oktober 1989"),
    ("B72", "NIK Saksi 1", "3327042110890004"),
    ("B73", "Pekerjaan Saksi 1", "Perangkat Desa"),
    ("B74", "Alamat Saksi 1", "RT 002 RW 001 Desa Tambi Kecamatan Watukumpu Kabupaten Pemalang"),
    ("B76", "Nama Saksi 2", "Sidin"),
    ("B77", "TTL Saksi 2", "Pemalang,"),
    ("B78", "NIK Saksi 2", "0000000000000000"),
    ("B79", "Pekerjaan Saksi 2", ""),
    ("B80", "Alamat Saksi 2", "")
]

# FORMULIR TERBAGI RAPI
with st.form("form_isian_catin"):
    st.subheader("Isikan Data Sesuai Urutan Kolom B Excel")
    
    tab_surat, tab_pria, tab_wanita, tab_wali_saksi = st.tabs([
        "📄 Register & Surat", 
        "👨 Catin Laki-Laki & Ortu", 
        "👩 Catin Perempuan & Ortu", 
        "🤝 Wali & Saksi"
    ])

    user_inputs = {}

    with tab_surat:
        st.markdown("##### 1. Data Surat & Pelaksanaan Akad")
        for cell_ref, label, def_val in FIELD_MAPPING[:5]:
            user_inputs[label] = st.text_input(f"[{cell_ref}] {label}", value=def_val)

    with tab_pria:
        st.markdown("##### 2. Data Calon Pengantin Laki-Laki & Orang Tua")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**-- Data Catin Laki-Laki --**")
            for cell_ref, label, def_val in FIELD_MAPPING[5:15]:
                user_inputs[label] = st.text_input(f"[{cell_ref}] {label}", value=def_val)
        with col2:
            st.markdown("**-- Data Orang Tua Laki-Laki --**")
            for cell_ref, label, def_val in FIELD_MAPPING[15:25]:
                user_inputs[label] = st.text_input(f"[{cell_ref}] {label}", value=def_val)

    with tab_wanita:
        st.markdown("##### 3. Data Calon Pengantin Perempuan & Orang Tua")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**-- Data Catin Perempuan --**")
            for cell_ref, label, def_val in FIELD_MAPPING[25:35]:
                user_inputs[label] = st.text_input(f"[{cell_ref}] {label}", value=def_val)
        with col2:
            st.markdown("**-- Data Orang Tua Perempuan --**")
            for cell_ref, label, def_val in FIELD_MAPPING[35:45]:
                user_inputs[label] = st.text_input(f"[{cell_ref}] {label}", value=def_val)

    with tab_wali_saksi:
        st.markdown("##### 4. Data Wali, Saksi & Mahar")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**-- Data Wali & Mahar --**")
            for cell_ref, label, def_val in FIELD_MAPPING[45:53]:
                user_inputs[label] = st.text_input(f"[{cell_ref}] {label}", value=def_val)
        with col2:
            st.markdown("**-- Data Saksi 1 & Saksi 2 --**")
            for cell_ref, label, def_val in FIELD_MAPPING[53:]:
                user_inputs[label] = st.text_input(f"[{cell_ref}] {label}", value=def_val)

    btn_simpan = st.form_submit_button("💾 Simpan & Perbarui Data Input", use_container_width=True)

if btn_simpan:
    st.session_state['input_data'] = user_inputs
    st.success("✅ Data berhasil disimpan! Anda bisa langsung mendownload hasilnya di bawah.")

# PANEL DOWNLOAD
st.markdown("---")
st.subheader("📥 Download Berkas")

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
    def generate_pdf(data_dict):
        buffer = BytesIO()
        F4_SIZE = landscape((612, 936))
        doc = SimpleDocTemplate(buffer, pagesize=F4_SIZE, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        elements = []
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=12, alignment=1, spaceAfter=10)
        elements.append(Paragraph("<b>RINGKASAN ISIAN DATA CATIN (F4)</b>", title_style))
        
        # Susun tabel 2 kolom (Field, Value)
        table_rows = [["FIELD (SELEKSI B2-B80)", "ISIAN DATA"]]
        current_data = data_dict if data_dict else {item[1]: item[2] for item in FIELD_MAPPING}
        for k, v in current_data.items():
            table_rows.append([str(k), str(v)])
            
        pdf_table = Table(table_rows, colWidths=[250, 600])
        pdf_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elements.append(pdf_table)
        doc.build(elements)
        buffer.seek(0)
        return buffer

    pdf_bytes = generate_pdf(st.session_state.get('input_data'))
    st.download_button(
        label="📄 Download Laporan PDF (Ukuran F4)",
        data=pdf_bytes,
        file_name="Isian_Data_Catin_F4.pdf",
        mime="application/pdf",
        use_container_width=True
    )
