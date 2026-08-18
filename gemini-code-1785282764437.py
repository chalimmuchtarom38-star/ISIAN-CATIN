import streamlit as st
import openpyxl
import re
import json
import os
from io import BytesIO
from datetime import datetime, date

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(
    page_title="Aplikasi Berkas Catin - Desa Tambi",
    page_icon="📜",
    layout="wide"
)

st.title("📜 Form Input Berkas Catin")
st.caption("Aplikasi dilengkapi fitur **Simpan Draf** agar data tidak hilang saat jaringan lambat/error.")

EXCEL_FILE = "BERKAS CATIN .xlsx"
DRAFT_FILE = "draf_terakhir.json"

NAMA_HARI = {
    "Monday": "Senin", "Tuesday": "Selasa", "Wednesday": "Rabu",
    "Thursday": "Kamis", "Friday": "Jumat", "Saturday": "Sabtu", "Sunday": "Minggu"
}

def hitung_umur(ttl_str):
    if not ttl_str:
        return 0
    match = re.search(r'\b(19\d{2}|20\d{2})\b', str(ttl_str))
    if match:
        tahun_lahir = int(match.group(1))
        tahun_sekarang = datetime.now().year
        return max(0, tahun_sekarang - tahun_lahir)
    return 0

# --- INIALISASI / MUAT DRAF JIKA ADA ---
def load_draft():
    if os.path.exists(DRAFT_FILE):
        try:
            with open(DRAFT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_draft(data):
    with open(DRAFT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

draft_data = load_draft()

# Tombol aksi atas untuk Draf
col_btn1, col_btn2 = st.columns([1, 4])
with col_btn1:
    if st.button("🔄 Reset / Clear Draf"):
        if os.path.exists(DRAFT_FILE):
            os.remove(DRAFT_FILE)
        st.session_state.clear()
        st.rerun()

# --- FORM INPUT DATA (Menggunakan key & nilai bawaan dari draf) ---
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
        no_register = st.text_input("Nomor Register", value=draft_data.get("no_register", "400.12.3.2/010/ VIII/ 2026"), key="no_register")
        tgl_surat = st.text_input("Tanggal Surat", value=draft_data.get("tgl_surat", "TAMBI, 11 AGUSTUS 2026"), key="tgl_surat")
        tgl_pelaksanaan = st.date_input("Tanggal Pelaksanaan Akad", value=date.fromisoformat(draft_data.get("tgl_pelaksanaan", "2026-09-07")), key="tgl_pelaksanaan")
        jam_akad = st.text_input("Jam Akad", value=draft_data.get("jam_akad", "JAM. 08.00"), key="jam_akad")
    with col2:
        tempat_akad = st.text_input("Tempat Akad Nikah", value=draft_data.get("tempat_akad", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="tempat_akad")
        email_catin = st.text_input("Email Catin", value=draft_data.get("email_catin", ""), key="email_catin")
        mahar = st.text_input("Maskawin / Mahar", value=draft_data.get("mahar", "Seperangkat Alat Sholat"), key="mahar")

with tab2:
    st.subheader("Data Calon Pengantin Laki-Laki")
    col_lk1, col_lk2 = st.columns(2)
    with col_lk1:
        nama_lk = st.text_input("Nama Catin Laki-Laki", value=draft_data.get("nama_lk", "Miftahul Anam"), key="nama_lk")
        bin_lk = st.text_input("Bin (Ayah Laki-Laki)", value=draft_data.get("bin_lk", "Nur Karim"), key="bin_lk")
        ttl_lk = st.text_input("TTL Laki-Laki", value=draft_data.get("ttl_lk", "Pemalang, 18 Februari 1999"), key="ttl_lk")
        nik_lk = st.text_input("NIK Laki-Laki", value=draft_data.get("nik_lk", "3327031802990004"), key="nik_lk")
    with col_lk2:
        pekerjaan_lk = st.text_input("Pekerjaan Laki-Laki", value=draft_data.get("pekerjaan_lk", "Swasta"), key="pekerjaan_lk")
        status_lk = st.text_input("Status Laki-Laki", value=draft_data.get("status_lk", "BELUM KAWIN"), key="status_lk")
        jk_lk = st.text_input("Jenis Kelamin Laki-Laki", value=draft_data.get("jk_lk", "Laki-Laki"), key="jk_lk")
        istri_terdahulu = st.text_input("Nama Istri Terdahulu", value=draft_data.get("istri_terdahulu", ""), key="istri_terdahulu")
        alamat_lk = st.text_area("Alamat Laki-Laki", value=draft_data.get("alamat_lk", "RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang"), key="alamat_lk")
        pendidikan_lk = st.text_input("Pendidikan Laki-Laki", value=draft_data.get("pendidikan_lk", "SLTA"), key="pendidikan_lk")

    st.divider()
    st.subheader("Data Ayah & Ibu Laki-Laki")
    col_alk, col_ilk = st.columns(2)
    with col_alk:
        st.markdown("**Ayah Laki-Laki**")
        nama_ayah_lk = st.text_input("Nama Ayah Laki-Laki", value=draft_data.get("nama_ayah_lk", "Nur Karim"), key="nama_ayah_lk")
        bin_ayah_lk = st.text_input("bin Ayah LK", value=draft_data.get("bin_ayah_lk", "Kasturi"), key="bin_ayah_lk")
        nik_ayah_lk = st.text_input("NIK Ayah LK", value=draft_data.get("nik_ayah_lk", "3327030608680006"), key="nik_ayah_lk")
        ttl_ayah_lk = st.text_input("TTL Ayah LK", value=draft_data.get("ttl_ayah_lk", "Pemalang, 06 Agustus 1968"), key="ttl_ayah_lk")
        pekerjaan_ayah_lk = st.text_input("Pekerjaan Ayah LK", value=draft_data.get("pekerjaan_ayah_lk", "PETANI/ PEKEBUN"), key="pekerjaan_ayah_lk")
        alamat_ayah_lk = st.text_area("Alamat Ayah LK", value=draft_data.get("alamat_ayah_lk", "RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang"), key="alamat_ayah_lk")

    with col_ilk:
        st.markdown("**Ibu Laki-Laki**")
        nama_ibu_lk = st.text_input("Nama Ibu Laki-Laki", value=draft_data.get("nama_ibu_lk", "Samijah"), key="nama_ibu_lk")
        bin_ibu_lk = st.text_input("bin Ibu LK", value=draft_data.get("bin_ibu_lk", "Taryad"), key="bin_ibu_lk")
        nik_ibu_lk = st.text_input("NIK Ibu LK", value=draft_data.get("nik_ibu_lk", "3327035405740004"), key="nik_ibu_lk")
        ttl_ibu_lk = st.text_input("TTL Ibu LK", value=draft_data.get("ttl_ibu_lk", "Pemalang, 14 Mei 1974"), key="ttl_ibu_lk")
        pekerjaan_ibu_lk = st.text_input("Pekerjaan Ibu LK", value=draft_data.get("pekerjaan_ibu_lk", "Mengurus Rumah Tangga"), key="pekerjaan_ibu_lk")
        alamat_ibu_lk = st.text_area("Alamat Ibu LK", value=draft_data.get("alamat_ibu_lk", "RT 002 RW 001 Desa Badak Kecamatan Belik Kabupaten Pemalang"), key="alamat_ibu_lk")

with tab3:
    st.subheader("Data Calon Pengantin Perempuan")
    col_pr1, col_pr2 = st.columns(2)
    with col_pr1:
        nama_pr = st.text_input("Nama Catin Perempuan", value=draft_data.get("nama_pr", "Diyan Solehatin"), key="nama_pr")
        binti_pr = st.text_input("Binti (Ayah Perempuan)", value=draft_data.get("binti_pr", "Disun"), key="binti_pr")
        ttl_pr = st.text_input("TTL Perempuan", value=draft_data.get("ttl_pr", "Pemalang, 29 Juni 2007"), key="ttl_pr")
        nik_pr = st.text_input("NIK Perempuan", value=draft_data.get("nik_pr", "3327046906070010"), key="nik_pr")
    with col_pr2:
        pekerjaan_pr = st.text_input("Pekerjaan Perempuan", value=draft_data.get("pekerjaan_pr", "BELUM/ TIDAK BEKERJA"), key="pekerjaan_pr")
        status_pr = st.text_input("Status Perempuan", value=draft_data.get("status_pr", "BELUM KAWIN"), key="status_pr")
        jk_pr = st.text_input("Jenis Kelamin Perempuan", value=draft_data.get("jk_pr", "PEREMPUAN"), key="jk_pr")
        suami_terdahulu = st.text_input("Nama Suami Terdahulu", value=draft_data.get("suami_terdahulu", ""), key="suami_terdahulu")
        alamat_pr = st.text_area("Alamat Perempuan", value=draft_data.get("alamat_pr", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="alamat_pr")
        pendidikan_pr = st.text_input("Pendidikan Perempuan", value=draft_data.get("pendidikan_pr", "SLTP"), key="pendidikan_pr")

    st.divider()
    st.subheader("Data Ayah & Ibu Perempuan")
    col_apr, col_ipr = st.columns(2)
    with col_apr:
        st.markdown("**Ayah Perempuan**")
        nama_ayah_pr = st.text_input("Nama Ayah PR", value=draft_data.get("nama_ayah_pr", "Disun"), key="nama_ayah_pr")
        bin_ayah_pr = st.text_input("bin Ayah PR", value=draft_data.get("bin_ayah_pr", "Tawiroji"), key="bin_ayah_pr")
        nik_ayah_pr = st.text_input("NIK Ayah PR", value=draft_data.get("nik_ayah_pr", "3327042504840003"), key="nik_ayah_pr")
        ttl_ayah_pr = st.text_input("TTL Ayah PR", value=draft_data.get("ttl_ayah_pr", "Pemalang, 21 April 1984"), key="ttl_ayah_pr")
        pekerjaan_ayah_pr = st.text_input("Pekerjaan Ayah PR", value=draft_data.get("pekerjaan_ayah_pr", "PETANI/ PEKEBUN"), key="pekerjaan_ayah_pr")
        alamat_ayah_pr = st.text_area("Alamat Ayah PR", value=draft_data.get("alamat_ayah_pr", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="alamat_ayah_pr")

    with col_ipr:
        st.markdown("**Ibu Perempuan**")
        nama_ibu_pr = st.text_input("Nama Ibu PR", value=draft_data.get("nama_ibu_pr", "Mutirah"), key="nama_ibu_pr")
        bin_ibu_pr = st.text_input("bin Ibu PR", value=draft_data.get("bin_ibu_pr", "Tamiarjo"), key="bin_ibu_pr")
        nik_ibu_pr = st.text_input("NIK Ibu PR", value=draft_data.get("nik_ibu_pr", "3327044411840003"), key="nik_ibu_pr")
        ttl_ibu_pr = st.text_input("TTL Ibu PR", value=draft_data.get("ttl_ibu_pr", "Pemalang, 04 November 1984"), key="ttl_ibu_pr")
        pekerjaan_ibu_pr = st.text_input("Pekerjaan Ibu PR", value=draft_data.get("pekerjaan_ibu_pr", "Mengurus Rumah Tangga"), key="pekerjaan_ibu_pr")
        alamat_ibu_pr = st.text_area("Alamat Ibu PR", value=draft_data.get("alamat_ibu_pr", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="alamat_ibu_pr")

with tab4:
    st.subheader("Data Wali Nikah")
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        nama_wali = st.text_input("Nama Wali", value=draft_data.get("nama_wali", "Disun"), key="nama_wali")
        bin_wali = st.text_input("Bin Wali", value=draft_data.get("bin_wali", "Tawiroji"), key="bin_wali")
        nik_wali = st.text_input("NIK Wali", value=draft_data.get("nik_wali", "3327042504840003"), key="nik_wali")
        ttl_wali = st.text_input("TTL Wali", value=draft_data.get("ttl_wali", "Pemalang, 21 April 1984"), key="ttl_wali")
    with col_w2:
        pekerjaan_wali = st.text_input("Pekerjaan Wali", value=draft_data.get("pekerjaan_wali", "PETANI/ PEKEBUN"), key="pekerjaan_wali")
        alamat_wali = st.text_area("Alamat Wali", value=draft_data.get("alamat_wali", "RT 004 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="alamat_wali")
        hubungan_wali = st.text_input("Hubungan Wali", value=draft_data.get("hubungan_wali", "AYAH KANDUNG"), key="hubungan_wali")
        nama_wali_lengkap = st.text_input("Nama Wali Lengkap", value=draft_data.get("nama_wali_lengkap", "Disun Bin Tawiroji"), key="nama_wali_lengkap")

with tab5:
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        st.subheader("Data Saksi 1")
        saksi1_nama = st.text_input("Nama Saksi 1", value=draft_data.get("saksi1_nama", "Chalim Muchtarom"), key="saksi1_nama")
        saksi1_ttl = st.text_input("TTL Saksi 1", value=draft_data.get("saksi1_ttl", "Pemalang, 21 Oktober 1989"), key="saksi1_ttl")
        saksi1_nik = st.text_input("NIK Saksi 1", value=draft_data.get("saksi1_nik", "3327042110890004"), key="saksi1_nik")
        saksi1_pekerjaan = st.text_input("Pekerjaan Saksi 1", value=draft_data.get("saksi1_pekerjaan", "Perangkat Desa"), key="saksi1_pekerjaan")
        saksi1_alamat = st.text_area("Alamat Saksi 1", value=draft_data.get("saksi1_alamat", "RT 002 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="saksi1_alamat")

    with col_s2:
        st.subheader("Data Saksi 2")
        saksi2_nama = st.text_input("Nama Saksi 2", value=draft_data.get("saksi2_nama", "Sidin"), key="saksi2_nama")
        saksi2_ttl = st.text_input("TTL Saksi 2", value=draft_data.get("saksi2_ttl", "Pemalang, 15 Mei 1980"), key="saksi2_ttl")
        saksi2_nik = st.text_input("NIK Saksi 2", value=draft_data.get("saksi2_nik", "3327041505800002"), key="saksi2_nik")
        saksi2_pekerjaan = st.text_input("Pekerjaan Saksi 2", value=draft_data.get("saksi2_pekerjaan", "Petani/Pekebun"), key="saksi2_pekerjaan")
        saksi2_alamat = st.text_area("Alamat Saksi 2", value=draft_data.get("saksi2_alamat", "RT 003 RW 001 Desa Tambi Kecamatan Watukumpul Kabupaten Pemalang"), key="saksi2_alamat")

# --- TOMBOL AKSI ---
st.divider()
col_act1, col_act2 = st.columns(2)

with col_act1:
    if st.button("💾 SIMPAN DRAF SEMENTARA (Antisipasi Jaringan Putus)", use_container_width=True):
        current_data = {
            "no_register": no_register, "tgl_surat": tgl_surat, "tgl_pelaksanaan": str(tgl_pelaksanaan),
            "jam_akad": jam_akad, "tempat_akad": tempat_akad, "email_catin": email_catin, "mahar": mahar,
            "nama_lk": nama_lk, "bin_lk": bin_lk, "ttl_lk": ttl_lk, "nik_lk": nik_lk,
            "pekerjaan_lk": pekerjaan_lk, "status_lk": status_lk, "jk_lk": jk_lk, "istri_terdahulu": istri_terdahulu,
            "alamat_lk": alamat_lk, "pendidikan_lk": pendidikan_lk, "nama_ayah_lk": nama_ayah_lk,
            "bin_ayah_lk": bin_ayah_lk, "nik_ayah_lk": nik_ayah_lk, "ttl_ayah_lk": ttl_ayah_lk,
            "pekerjaan_ayah_lk": pekerjaan_ayah_lk, "alamat_ayah_lk": alamat_ayah_lk,
            "nama_ibu_lk": nama_ibu_lk, "bin_ibu_lk": bin_ibu_lk, "nik_ibu_lk": nik_ibu_lk,
            "ttl_ibu_lk": ttl_ibu_lk, "pekerjaan_ibu_lk": pekerjaan_ibu_lk, "alamat_ibu_lk": alamat_ibu_lk,
            "nama_pr": nama_pr, "binti_pr": binti_pr, "ttl_pr": ttl_pr, "nik_pr": nik_pr,
            "pekerjaan_pr": pekerjaan_pr, "status_pr": status_pr, "jk_pr": jk_pr, "suami_terdahulu": suami_terdahulu,
            "alamat_pr": alamat_pr, "pendidikan_pr": pendidikan_pr, "nama_ayah_pr": nama_ayah_pr,
            "bin_ayah_pr": bin_ayah_pr, "nik_ayah_pr": nik_ayah_pr, "ttl_ayah_pr": ttl_ayah_pr,
            "pekerjaan_ayah_pr": pekerjaan_ayah_pr, "alamat_ayah_pr": alamat_ayah_pr,
            "nama_ibu_pr": nama_ibu_pr, "bin_ibu_pr": bin_ibu_pr, "nik_ibu_pr": nik_ibu_pr,
            "ttl_ibu_pr": ttl_ibu_pr, "pekerjaan_ibu_pr": pekerjaan_ibu_pr, "alamat_ibu_pr": alamat_ibu_pr,
            "nama_wali": nama_wali, "bin_wali": bin_wali, "nik_wali": nik_wali, "ttl_wali": ttl_wali,
            "pekerjaan_wali": pekerjaan_wali, "alamat_wali": alamat_wali, "hubungan_wali": hubungan_wali,
            "nama_wali_lengkap": nama_wali_lengkap, "saksi1_nama": saksi1_nama, "saksi1_ttl": saksi1_ttl,
            "saksi1_nik": saksi1_nik, "saksi1_pekerjaan": saksi1_pekerjaan, "saksi1_alamat": saksi1_alamat,
            "saksi2_nama": saksi2_nama, "saksi2_ttl": saksi2_ttl, "saksi2_nik": saksi2_nik,
            "saksi2_pekerjaan": saksi2_pekerjaan, "saksi2_alamat": saksi2_alamat
        }
        save_draft(current_data)
        st.toast("✅ Draf berhasil disimpan! Aman untuk melanjutkan.", icon="💾")

with col_act2:
    submit = st.button("🚀 PROSES KE EXCEL & GENERATE BERKAS", type="primary", use_container_width=True)

if submit:
    # Proses simpan ke Excel tetap berjalan seperti biasa
    st.info("Memproses berkas...")
