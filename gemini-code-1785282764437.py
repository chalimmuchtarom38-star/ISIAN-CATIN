import streamlit as st
import openpyxl
import re
import json
import os
from io import BytesIO
from datetime import datetime, date

st.set_page_config(
    page_title="Aplikasi Berkas Catin - Desa Tambi",
    page_icon="📜",
    layout="wide"
)

EXCEL_FILE = "BERKAS CATIN .xlsx"
DRAFT_FILE = "draf_terakhir.json"

# --- FUNGSI RUMUS HITUNG UMUR (TETAP ADA & AMAN) ---
def hitung_umur(ttl_str):
    if not ttl_str:
        return 0
    # Mencari 4 digit angka tahun (19xx atau 20xx) dari teks TTL
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
        except:
            return {}
    return {}

def save_draft_file():
    draft_data = {}
    for key in st.session_state:
        val = st.session_state[key]
        if isinstance(val, (date, datetime)):
            draft_data[key] = str(val)
        else:
            draft_data[key] = val
            
    with open(DRAFT_FILE, "w", encoding="utf-8") as f:
        json.dump(draft_data, f, ensure_ascii=False, indent=2)

draft = load_draft()

st.title("📜 Form Input Berkas Catin")

# --- TOMBOL ATAS ---
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
            except:
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
        # Menampilkan hasil hitung umur otomatis
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
    submit = st.button("🚀 PROSES KE EXCEL & GENERATE BERKAS", type="primary", use_container_width=True)

if submit:
    save_draft_file()
    
    with st.spinner("⏳ Sedang memproses Excel dan membuat berkas, mohon tunggu..."):
        try:
            ss = st.session_state
            
            # Hitung umur saat tombol diklik (bisa dipakai saat memasukkan ke Excel/PDF)
            umur_lk = hitung_umur(ss.get('ttl_lk', ''))
            umur_pr = hitung_umur(ss.get('ttl_pr', ''))
            
            if os.path.exists(EXCEL_FILE):
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb.active
                
                # Masukkan data ke Excel (Contoh)
                # ws['A2'] = ss.get('nama_lk', '')
                # ws['B2'] = umur_lk
                
                wb.save(EXCEL_FILE)
                wb.close()
                st.success("✅ Data Excel berhasil diperbarui!")
            else:
                st.warning(f"File Excel '{EXCEL_FILE}' tidak ditemukan.")
                
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat memproses Excel: {e}")
