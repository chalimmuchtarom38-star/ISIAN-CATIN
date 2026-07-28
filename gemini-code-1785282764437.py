import streamlit as st
import openpyxl
import io

st.set_page_config(page_title="Form Input Data Catin", page_icon="📝", layout="centered")

st.title("📝 Form Pengisian Data Catin")
st.write("Silakan isi formulir di bawah ini. File hasil unduhan akan mempertahankan 100% format asli Excel.")

# Form Input
with st.form("catin_form"):
    st.header("📌 Data Surat & Pelaksanaan")
    no_register = st.text_input("Nomor Register", "400.12.3.2/006/ VII/ 2026")
    tgl_surat = st.text_input("Tanggal Surat (Format: TAMBI, 27 JULI 2026)", "TAMBI, 27 JULI 2026")
    tgl_pelaksanaan = st.date_input("Tanggal Pelaksanaan Akad")
    jam_pelaksanaan = st.text_input("Jam", "JAM. 08.00")
    tempat_akad = st.text_area("Tempat Akad Nikah", "DI RUMAH MEMPELAI PUTRI RT 002 RW 004 DESA TAMBI WATUKUMPUL PEMALANG")

    st.markdown("---")
    st.header("👨 Calon Pengantin Laki-Laki")
    pria_nama = st.text_input("Nama Laki-Laki", "UNWAN FALAHI")
    pria_bin = st.text_input("Bin (Ayah Laki-Laki)", "RATAM")
    pria_ttl = st.text_input("Tempat Tanggal Lahir (Pria)", "PEMALANG, 23 JUNI 1997")
    pria_nik = st.text_input("NIK (Pria)", "3327042306970003")
    pria_pekerjaan = st.text_input("Pekerjaan (Pria)", "BURUH HARIAN LEPAS")
    pria_status = st.selectbox("Status (Pria)", ["BELUM KAWIN", "DUDA"], index=0)
    pria_alamat = st.text_area("Alamat (Pria)", "RT 002 RW 001 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    st.markdown("---")
    st.header("👩 Calon Pengantin Perempuan")
    wanita_nama = st.text_input("Nama Perempuan", "IKNA SABELA")
    wanita_binti = st.text_input("Binti (Ayah Perempuan)", "WARYONO")
    wanita_ttl = st.text_input("Tempat Tanggal Lahir (Perempuan)", "PEMALANG, 05 OKTOBER 2008")
    wanita_nik = st.text_input("NIK (Perempuan)", "3327044510080002")
    wanita_pekerjaan = st.text_input("Pekerjaan (Perempuan)", "BELUM/ TIDAK BEKERJA")
    wanita_status = st.selectbox("Status (Perempuan)", ["BELUM KAWIN", "JANDA"], index=0)
    wanita_alamat = st.text_area("Alamat (Perempuan)", "RT 003 RW 001 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    st.markdown("---")
    st.header("👴 Data Ayah & Ibu Laki-Laki")
    ayah_pria_nama = st.text_input("Nama Ayah Laki-Laki", "RATAM")
    ayah_pria_bin = st.text_input("Bin (Kakek Laki-Laki)", "MURTAJA")
    ayah_pria_nik = st.text_input("NIK Ayah Laki-Laki", "3327040308660001")
    
    ibu_pria_nama = st.text_input("Nama Ibu Laki-Laki", "DASRI")
    ibu_pria_bin = st.text_input("Bin Ibu Laki-Laki", "KARTADIWIRYA")
    ibu_pria_nik = st.text_input("NIK Ibu Laki-Laki", "3327045104690001")

    st.markdown("---")
    st.header("🧓 Data Ayah & Ibu Perempuan")
    ayah_wanita_nama = st.text_input("Nama Ayah Perempuan", "WARYONO")
    ayah_wanita_bin = st.text_input("Bin (Kakek Perempuan)", "MURSIDI")
    ayah_wanita_nik = st.text_input("NIK Ayah Perempuan", "3327041209810002")

    ibu_wanita_nama = st.text_input("Nama Ibu Perempuan", "DEWI MURIYAH")
    ibu_wanita_bin = st.text_input("Bin Ibu Perempuan", "SUWARNO")
    ibu_wanita_nik = st.text_input("NIK Ibu Perempuan", "3327044302900010")

    st.markdown("---")
    st.header("🤝 Data Wali & Saksi")
    wali_nama = st.text_input("Nama Wali", "WARYONO")
    wali_bin = st.text_input("Bin Wali", "MURSIDI")
    wali_hubungan = st.text_input("Hubungan Wali", "AYAH KANDUNG")
    
    saksi1_nama = st.text_input("Nama Saksi 1", "WARTIM")
    saksi1_nik = st.text_input("NIK Saksi 1", "3327041605870003")
    
    saksi2_nama = st.text_input("Nama Saksi 2", "CHALIM MUCHTAROM")
    saksi2_nik = st.text_input("NIK Saksi 2", "3327042110890004")

    submitted = st.form_submit_button("💾 Proses Data & Buat File Excel")

if submitted:
    # Buka file template Excel asli
    wb = openpyxl.load_workbook("ISIAN DATA CATIN.xlsx")
    ws = wb["ISIAN DATA"]

    # Injeksi Data ke Koordinat Sel Excel Asli
    ws["G2"] = no_register
    ws["H3"] = f", {tgl_surat}"
    ws["G4"] = str(tgl_pelaksanaan)
    ws["G5"] = jam_pelaksanaan
    ws["G6"] = tempat_akad

    # Data Pria
    ws["G8"] = pria_nama
    ws["G9"] = pria_bin
    ws["G10"] = pria_ttl
    ws["G11"] = pria_nik
    ws["G12"] = pria_pekerjaan
    ws["G13"] = pria_status
    ws["G16"] = pria_alamat

    # Data Ayah & Ibu Pria
    ws["G19"] = ayah_pria_nama
    ws["J19"] = ayah_pria_bin
    ws["G20"] = ayah_pria_nik
    ws["G26"] = ibu_pria_nama
    ws["I26"] = ibu_pria_bin
    ws["G27"] = ibu_pria_nik

    # Data Perempuan
    ws["G34"] = wanita_nama
    ws["G35"] = wanita_binti
    ws["G36"] = wanita_ttl
    ws["G37"] = wanita_nik
    ws["G38"] = wanita_pekerjaan
    ws["G39"] = wanita_status
    ws["G41"] = wanita_alamat

    # Data Ayah & Ibu Perempuan
    ws["G45"] = ayah_wanita_nama
    ws["I45"] = ayah_wanita_bin
    ws["G46"] = ayah_wanita_nik
    ws["G52"] = ibu_wanita_nama
    ws["I52"] = ibu_wanita_bin
    ws["G53"] = ibu_wanita_nik

    # Wali & Saksi
    ws["G58"] = wali_nama
    ws["G59"] = wali_bin
    ws["G64"] = wali_hubungan
    ws["G70"] = saksi1_nama
    ws["G72"] = saksi1_nik
    ws["G76"] = saksi2_nama
    ws["G78"] = saksi2_nik

    # Simpan ke Buffer Memory untuk Unduhan
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.success("✅ Data berhasil diproses ke dalam template Excel!")
    st.download_button(
        label="📥 Download File Excel Terisi",
        data=buffer,
        file_name=f"ISIAN_DATA_CATIN_{pria_nama}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )