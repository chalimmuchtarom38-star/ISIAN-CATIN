import streamlit as st
import openpyxl
import io

st.set_page_config(page_title="Form Input Data Catin", page_icon="📝", layout="centered")

st.title("📝 Form Pengisian Data Catin")
st.write("Silakan isi formulir di bawah ini. Hasil unduhan Excel akan mempertahankan 100% format asli.")

with st.form("catin_form"):
    # -------------------------------------------------------------
    st.header("📌 1. Data Surat & Pelaksanaan")
    no_register = st.text_input("Nomor Register", "400.12.3.2/006/ VII/ 2026")
    tgl_surat = st.text_input("Tanggal Surat (Format: TAMBI, 27 JULI 2026)", "TAMBI, 27 JULI 2026")
    tgl_pelaksanaan = st.date_input("Tanggal Pelaksanaan Akad")
    jam_pelaksanaan = st.text_input("Jam Pelaksanaan", "JAM. 08.00")
    tempat_akad = st.text_area("Tempat Akad Nikah", "DI RUMAH MEMPELAI PUTRI RT 002 RW 004 DESA TAMBI WATUKUMPUL PEMALANG")

    # -------------------------------------------------------------
    st.markdown("---")
    st.header("👨 2. Calon Pengantin Laki-Laki")
    pria_nama = st.text_input("Nama Laki-Laki", "UNWAN FALAHI")
    pria_bin = st.text_input("BIN (Ayah Laki-Laki)", "RATAM")
    pria_ttl = st.text_input("Tempat Tanggal Lahir (Pria)", "PEMALANG, 23 JUNI 1997")
    pria_nik = st.text_input("NIK (Pria)", "3327042306970003")
    pria_pekerjaan = st.text_input("Pekerjaan (Pria)", "BURUH HARIAN LEPAS")
    pria_status = st.selectbox("Status (Pria)", ["BELUM KAWIN", "DUDA"], index=0)
    pria_pendidikan = st.text_input("Pendidikan (Pria)", "SD")
    pria_alamat = st.text_area("Alamat (Pria)", "RT 002 RW 001 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    st.subheader("👴 Data Ayah Laki-Laki")
    ayah_pria_nama = st.text_input("Nama Ayah Laki-Laki", "RATAM")
    ayah_pria_bin = st.text_input("BIN Ayah Laki-Laki (Kakek)", "MURTAJA")
    ayah_pria_nik = st.text_input("NIK Ayah Laki-Laki", "3327040308660001")
    ayah_pria_ttl = st.text_input("TTL Ayah Laki-Laki", "PEMALANG, 03 AGUSTUS 1966")
    ayah_pria_pekerjaan = st.text_input("Pekerjaan Ayah Laki-Laki", "PETANI/ PEKEBUN")
    ayah_pria_alamat = st.text_area("Alamat Ayah Laki-Laki", "RT 002 RW 001 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    st.subheader("👵 Data Ibu Laki-Laki")
    ibu_pria_nama = st.text_input("Nama Ibu Laki-Laki", "DASRI")
    ibu_pria_bin = st.text_input("BIN Ibu Laki-Laki", "KARTADIWIRYA")
    ibu_pria_nik = st.text_input("NIK Ibu Laki-Laki", "3327045104690001")
    ibu_pria_ttl = st.text_input("TTL Ibu Laki-Laki", "PEMALANG, 11 APRIL 1969")
    ibu_pria_pekerjaan = st.text_input("Pekerjaan Ibu Laki-Laki", "Mengurus Rumah Tangga")
    ibu_pria_alamat = st.text_area("Alamat Ibu Laki-Laki", "RT 002 RW 001 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    # -------------------------------------------------------------
    st.markdown("---")
    st.header("👩 3. Calon Pengantin Perempuan")
    wanita_nama = st.text_input("Nama Perempuan", "IKNA SABELA")
    wanita_binti = st.text_input("BINTI (Ayah Perempuan)", "WARYONO")
    wanita_ttl = st.text_input("Tempat Tanggal Lahir (Perempuan)", "PEMALANG, 05 OKTOBER 2008")
    wanita_nik = st.text_input("NIK (Perempuan)", "3327044510080002")
    wanita_pekerjaan = st.text_input("Pekerjaan (Perempuan)", "BELUM/ TIDAK BEKERJA")
    wanita_status = st.selectbox("Status (Perempuan)", ["BELUM KAWIN", "JANDA"], index=0)
    wanita_pendidikan = st.text_input("Pendidikan (Perempuan)", "SD")
    wanita_alamat = st.text_area("Alamat (Perempuan)", "RT 003 RW 001 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    st.subheader("👴 Data Ayah Perempuan")
    ayah_wanita_nama = st.text_input("Nama Ayah Perempuan", "WARYONO")
    ayah_wanita_bin = st.text_input("BIN Ayah Perempuan (Kakek)", "MURSIDI")
    ayah_wanita_nik = st.text_input("NIK Ayah Perempuan", "3327041209810002")
    ayah_wanita_ttl = st.text_input("TTL Ayah Perempuan", "PEMALANG, 12 SEPTEMBER 1981")
    ayah_wanita_pekerjaan = st.text_input("Pekerjaan Ayah Perempuan", "PETANI/ PEKEBUN")
    ayah_wanita_alamat = st.text_area("Alamat Ayah Perempuan", "RT 003 RW 001 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    st.subheader("👵 Data Ibu Perempuan")
    ibu_wanita_nama = st.text_input("Nama Ibu Perempuan", "DEWI MURIYAH")
    ibu_wanita_bin = st.text_input("BIN Ibu Perempuan", "SUWARNO")
    ibu_wanita_nik = st.text_input("NIK Ibu Perempuan", "3327044302900010")
    ibu_wanita_ttl = st.text_input("TTL Ibu Perempuan", "PEMALANG 03 FEBRUARI 1990")
    ibu_wanita_pekerjaan = st.text_input("Pekerjaan Ibu Perempuan", "Mengurus Rumah Tangga")
    ibu_wanita_alamat = st.text_area("Alamat Ibu Perempuan", "RT 003 RW 001 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    # -------------------------------------------------------------
    st.markdown("---")
    st.header("🤝 4. Data Wali & Mahar")
    wali_nama = st.text_input("Nama Wali", "WARYONO")
    wali_bin = st.text_input("BIN Wali", "MURSIDI")
    wali_nik = st.text_input("NIK Wali", "3327040107740063")
    wali_ttl = st.text_input("TTL Wali", "PEMALANG, 12 SEPTEMBER 1999")
    wali_pekerjaan = st.text_input("Pekerjaan Wali", "PETANI/ PEKEBUN")
    wali_alamat = st.text_area("Alamat Wali", "RT 003 RW 001 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")
    wali_hubungan = st.text_input("Hubungan Wali", "AYAH KANDUNG")
    mahar = st.text_input("Mahar / Maskawin", "Seperangkat Alat Sholat")
    
    # Kolom Tambahan Khusus Ringkasan Wali (B68 / G68)
    nama_lengkap_wali_b68 = st.text_input("Nama Wali Lengkap beserta BIN (Untuk Baris 'Nama Wali' ringkasan)", "WARYONO BIN MURSIDI")

    # -------------------------------------------------------------
    st.markdown("---")
    st.header("📜 5. Data Saksi-Saksi")
    
    st.subheader("👤 Saksi 1")
    saksi1_nama = st.text_input("Nama Saksi 1", "WARTIM")
    saksi1_ttl = st.text_input("Tempat, Tanggal Lahir Saksi 1", "PEMALANG 16 MEI 1987")
    saksi1_nik = st.text_input("NIK Saksi 1", "3327041605870003")
    saksi1_pekerjaan = st.text_input("Pekerjaan Saksi 1", "WIRASWASTA")
    saksi1_alamat = st.text_area("Alamat Saksi 1", "RT 003 RW 001 DESA TAMBI KECAMATAN WATUKUMPUL KABUPATEN PEMALANG")

    st.subheader("👤 Saksi 2")
    saksi2_nama = st.text_input("Nama Saksi 2", "CHALIM MUCHTAROM")
    saksi2_ttl = st.text_input("Tempat, Tanggal Lahir Saksi 2", "PEMALANG, 21 OKTOBER 1989")
    saksi2_nik = st.text_input("NIK Saksi 2", "3327042110890004")
    saksi2_pekerjaan = st.text_input("Pekerjaan Saksi 2", "PERANGKAT DESA")
    saksi2_alamat = st.text_area("Alamat Saksi 2", "RT 002 RW 001 DESA TAMBI KEC. WATUKUMPUL KAB. PEMALANG")

    submitted = st.form_submit_button("💾 Proses Data & Buat File Excel")

if submitted:
    wb = openpyxl.load_workbook("ISIAN DATA CATIN.xlsx")
    ws = wb["ISIAN DATA"]

    # 1. Surat & Pelaksanaan
    ws["G2"] = no_register
    ws["H3"] = f", {tgl_surat}"
    ws["G4"] = str(tgl_pelaksanaan)
    ws["G5"] = jam_pelaksanaan
    ws["G6"] = tempat_akad

    # 2. Catin Laki-Laki
    ws["G8"] = pria_nama
    ws["G9"] = pria_bin
    ws["G10"] = pria_ttl
    ws["G11"] = pria_nik
    ws["G12"] = pria_pekerjaan
    ws["G13"] = pria_status
    ws["G16"] = pria_alamat
    ws["G17"] = pria_pendidikan

    # Ayah & Ibu Laki-Laki
    ws["G19"] = ayah_pria_nama
    ws["J19"] = ayah_pria_bin
    ws["G20"] = ayah_pria_nik
    ws["G21"] = ayah_pria_ttl
    ws["G22"] = ayah_pria_pekerjaan
    ws["G23"] = ayah_pria_alamat

    ws["G26"] = ibu_pria_nama
    ws["I26"] = ibu_pria_bin
    ws["G27"] = ibu_pria_nik
    ws["G28"] = ibu_pria_ttl
    ws["G29"] = ibu_pria_pekerjaan
    ws["G30"] = ibu_pria_alamat

    # 3. Catin Perempuan
    ws["G34"] = wanita_nama
    ws["G35"] = wanita_binti
    ws["G36"] = wanita_ttl
    ws["G37"] = wanita_nik
    ws["G38"] = wanita_pekerjaan
    ws["G39"] = wanita_status
    ws["G41"] = wanita_alamat
    ws["G43"] = wanita_pendidikan

    # Ayah & Ibu Perempuan
    ws["G45"] = ayah_wanita_nama
    ws["I45"] = ayah_wanita_bin
    ws["G46"] = ayah_wanita_nik
    ws["G47"] = ayah_wanita_ttl
    ws["G48"] = ayah_wanita_pekerjaan
    ws["G49"] = ayah_wanita_alamat

    ws["G52"] = ibu_wanita_nama
    ws["I52"] = ibu_wanita_bin
    ws["G53"] = ibu_wanita_nik
    ws["G54"] = ibu_wanita_ttl
    ws["G55"] = ibu_wanita_pekerjaan
    ws["G56"] = ibu_wanita_alamat

    # 4. Wali & Mahar
    ws["G58"] = wali_nama
    ws["G59"] = wali_bin
    ws["G60"] = wali_nik
    ws["G61"] = wali_ttl
    ws["G62"] = wali_pekerjaan
    ws["G63"] = wali_alamat
    ws["G64"] = wali_hubungan
    ws["G65"] = mahar
    
    # Memasukkan input khusus nama wali ringkasan di G68
    ws["G68"] = nama_lengkap_wali_b68

    # 5. Saksi 1
    ws["G70"] = saksi1_nama
    ws["G71"] = saksi1_ttl
    ws["G72"] = saksi1_nik
    ws["G73"] = saksi1_pekerjaan
    ws["G74"] = saksi1_alamat

    # 6. Saksi 2
    ws["G76"] = saksi2_nama
    ws["G77"] = saksi2_ttl
    ws["G78"] = saksi2_nik
    ws["G79"] = saksi2_pekerjaan
    ws["G80"] = saksi2_alamat

    # Save ke buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.success("✅ Data berhasil diperbarui!")
    st.download_button(
        label="📥 Download File Excel Terisi",
        data=buffer,
        file_name=f"ISIAN_DATA_CATIN_{pria_nama}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
