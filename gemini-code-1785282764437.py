import streamlit as st
import openpyxl
import datetime
import re
import os

# Set Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="Sistem Pengisian Berkas Catin",
    page_icon="📝",
    layout="wide"
)

# =============================================================
# FUNGSI PEMBANTU (UTILITY FUNCTIONS)
# =============================================================

def hitung_umur(ttl_string):
    """Menghitung perkiraan umur berdasarkan tahun di string TTL"""
    if not ttl_string:
        return ""
    try:
        tahun_match = re.findall(r'\b(19\d\d|20\d\d)\b', str(ttl_string))
        if tahun_match:
            tahun_lahir = int(tahun_match[-1])
            tahun_sekarang = datetime.datetime.now().year
            return f"{tahun_sekarang - tahun_lahir} Tahun"
    except Exception:
        pass
    return ""


def update_excel_data(data, master_path):
    """Memasukkan data dari form ke sheet 'ISIAN DATA' di file Excel master"""
    if not os.path.exists(master_path):
        raise FileNotFoundError(f"File template '{master_path}' tidak ditemukan di folder aplikasi.")

    wb = openpyxl.load_workbook(master_path)
    
    # Targetkan langsung sheet "ISIAN DATA"
    if "ISIAN DATA" in wb.sheetnames:
        ws = wb["ISIAN DATA"]
    else:
        ws = wb.active

    # Penanganan MergedCell (Mencegah AttributeError: read-only)
    def set_cell_value(cell_address, value):
        try:
            ws[cell_address] = value
        except AttributeError:
            for rng in ws.merged_cells.ranges:
                if cell_address in rng:
                    top_left_cell = rng.start_cell.coordinate
                    ws[top_left_cell] = value
                    break

    # Pemetaan Field Form ke Sel Excel (Sheet ISIAN DATA)
    mapping = {
        # CATIN PRIA
        'catin_pria_nama': 'C13',
        'catin_pria_nik': 'C14',
        'catin_pria_bin': 'C15',
        'catin_pria_ttl': 'C16',
        'catin_pria_kewarganegaraan': 'C17',
        'catin_pria_agama': 'C18',
        'catin_pria_pekerjaan': 'C19',
        'catin_pria_alamat': 'C20',

        # CATIN WANITA
        'catin_wanita_nama': 'F13',
        'catin_wanita_nik': 'F14',
        'catin_wanita_binti': 'F15',
        'catin_wanita_ttl': 'F16',
        'catin_wanita_kewarganegaraan': 'F17',
        'catin_wanita_agama': 'F18',
        'catin_wanita_pekerjaan': 'F19',
        'catin_wanita_alamat': 'F20',

        # WALI NIKAH
        'wali_nama': 'C26',
        'wali_nik': 'C27',
        'wali_bin': 'C28',
        'wali_ttl': 'C29',
        'wali_kewarganegaraan': 'C30',
        'wali_agama': 'C31',
        'wali_pekerjaan': 'C32',
        'wali_alamat': 'C33',
        'wali_hubungan': 'C34',

        # AYAH CATIN PRIA
        'ayah_pria_nama': 'C40',
        'ayah_pria_nik': 'C41',
        'ayah_pria_bin': 'C42',
        'ayah_pria_ttl': 'C43',
        'ayah_pria_kewarganegaraan': 'C44',
        'ayah_pria_agama': 'C45',
        'ayah_pria_pekerjaan': 'C46',
        'ayah_pria_alamat': 'C47',

        # IBU CATIN PRIA
        'ibu_pria_nama': 'F40',
        'ibu_pria_nik': 'F41',
        'ibu_pria_bin': 'F42',
        'ibu_pria_ttl': 'F43',
        'ibu_pria_kewarganegaraan': 'F44',
        'ibu_pria_agama': 'F45',
        'ibu_pria_pekerjaan': 'F46',
        'ibu_pria_alamat': 'F47',

        # AYAH CATIN WANITA
        'ayah_wanita_nama': 'C53',
        'ayah_wanita_nik': 'C54',
        'ayah_wanita_bin': 'C55',
        'ayah_wanita_ttl': 'C56',
        'ayah_wanita_kewarganegaraan': 'C57',
        'ayah_wanita_agama': 'C58',
        'ayah_wanita_pekerjaan': 'C59',
        'ayah_wanita_alamat': 'C60',

        # IBU CATIN WANITA
        'ibu_wanita_nama': 'F53',
        'ibu_wanita_nik': 'F54',
        'ibu_wanita_bin': 'F55',
        'ibu_wanita_ttl': 'F56',
        'ibu_wanita_kewarganegaraan': 'F57',
        'ibu_wanita_agama': 'F58',
        'ibu_wanita_pekerjaan': 'F59',
        'ibu_wanita_alamat': 'F60',

        # AKAD NIKAH & LAINNYA
        'hari_tgl_akad': 'C66',
        'waktu_akad': 'C67',
        'mas_kawin': 'C68',
        'tempat_akad': 'C69',
        'tgl_surat': 'C70',
        'status_pria': 'C71',
        'status_wanita': 'C72',
    }

    # Isi seluruh data ke dalam sel Excel
    for key, cell in mapping.items():
        set_cell_value(cell, data.get(key, ""))

    # Hitung Umur Otomatis di Sheet ISIAN DATA
    set_cell_value('C21', hitung_umur(data.get('catin_pria_ttl', '')))
    set_cell_value('F21', hitung_umur(data.get('catin_wanita_ttl', '')))

    # Simpan Hasil
    output_path = "BERKAS_CATIN_TERISI.xlsx"
    wb.save(output_path)
    return output_path


# =============================================================
# TAMPILAN APLIKASI STREAMLIT (UI)
# =============================================================

st.title("📋 Input Data Berkas Calon Pengantin (CATIN)")
st.write("Isi formulir di bawah ini untuk memperbarui sheet **ISIAN DATA** secara otomatis.")

# Penampung Data
data = {}

# TABS UNTUK MENGELOMPOKKAN INPUT FORM
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "👨‍💼 Catin Pria", 
    "👩‍💼 Catin Wanita", 
    "📜 Wali Nikah", 
    "👨‍👩‍👦 Orang Tua Pria", 
    "👨‍👩‍👧 Orang Tua Wanita", 
    "💍 Akad & Pelaksanaan"
])

with tab1:
    st.subheader("Data Calon Pengantin Pria")
    col1, col2 = st.columns(2)
    with col1:
        data['catin_pria_nama'] = st.text_input("Nama Lengkap Pria", key="cp_nama")
        data['catin_pria_nik'] = st.text_input("NIK Pria", key="cp_nik")
        data['catin_pria_bin'] = st.text_input("Bin", key="cp_bin")
        data['catin_pria_ttl'] = st.text_input("Tempat, Tgl Lahir (misal: Pemalang, 12 Mei 1995)", key="cp_ttl")
    with col2:
        data['catin_pria_kewarganegaraan'] = st.text_input("Kewarganegaraan Pria", value="WNI", key="cp_kwg")
        data['catin_pria_agama'] = st.selectbox("Agama Pria", ["Islam", "Kristen", "Katolik", "Hindu", "Buddha", "Khonghucu"], key="cp_agama")
        data['catin_pria_pekerjaan'] = st.text_input("Pekerjaan Pria", key="cp_kerja")
        data['catin_pria_alamat'] = st.text_area("Alamat Pria", key="cp_alamat")

with tab2:
    st.subheader("Data Calon Pengantin Wanita")
    col1, col2 = st.columns(2)
    with col1:
        data['catin_wanita_nama'] = st.text_input("Nama Lengkap Wanita", key="cw_nama")
        data['catin_wanita_nik'] = st.text_input("NIK Wanita", key="cw_nik")
        data['catin_wanita_binti'] = st.text_input("Binti", key="cw_binti")
        data['catin_wanita_ttl'] = st.text_input("Tempat, Tgl Lahir (misal: Pemalang, 20 Agustus 1998)", key="cw_ttl")
    with col2:
        data['catin_wanita_kewarganegaraan'] = st.text_input("Kewarganegaraan Wanita", value="WNI", key="cw_kwg")
        data['catin_wanita_agama'] = st.selectbox("Agama Wanita", ["Islam", "Kristen", "Katolik", "Hindu", "Buddha", "Khonghucu"], key="cw_agama")
        data['catin_wanita_pekerjaan'] = st.text_input("Pekerjaan Wanita", key="cw_kerja")
        data['catin_wanita_alamat'] = st.text_area("Alamat Wanita", key="cw_alamat")

with tab3:
    st.subheader("Data Wali Nikah")
    col1, col2 = st.columns(2)
    with col1:
        data['wali_nama'] = st.text_input("Nama Wali", key="w_nama")
        data['wali_nik'] = st.text_input("NIK Wali", key="w_nik")
        data['wali_bin'] = st.text_input("Bin Wali", key="w_bin")
        data['wali_ttl'] = st.text_input("Tempat, Tgl Lahir Wali", key="w_ttl")
        data['wali_hubungan'] = st.text_input("Hubungan Wali (misal: Ayah Kandung)", key="w_hub")
    with col2:
        data['wali_kewarganegaraan'] = st.text_input("Kewarganegaraan Wali", value="WNI", key="w_kwg")
        data['wali_agama'] = st.selectbox("Agama Wali", ["Islam", "Kristen", "Katolik", "Hindu", "Buddha", "Khonghucu"], key="w_agama")
        data['wali_pekerjaan'] = st.text_input("Pekerjaan Wali", key="w_kerja")
        data['wali_alamat'] = st.text_area("Alamat Wali", key="w_alamat")

with tab4:
    st.subheader("Data Orang Tua Catin Pria")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Data Ayah Pria**")
        data['ayah_pria_nama'] = st.text_input("Nama Ayah Pria", key="ap_nama")
        data['ayah_pria_nik'] = st.text_input("NIK Ayah Pria", key="ap_nik")
        data['ayah_pria_bin'] = st.text_input("Bin Ayah Pria", key="ap_bin")
        data['ayah_pria_ttl'] = st.text_input("TTL Ayah Pria", key="ap_ttl")
        data['ayah_pria_kewarganegaraan'] = st.text_input("Kewarganegaraan Ayah Pria", value="WNI", key="ap_kwg")
        data['ayah_pria_agama'] = st.text_input("Agama Ayah Pria", value="Islam", key="ap_agama")
        data['ayah_pria_pekerjaan'] = st.text_input("Pekerjaan Ayah Pria", key="ap_kerja")
        data['ayah_pria_alamat'] = st.text_area("Alamat Ayah Pria", key="ap_alamat")
    with col2:
        st.markdown("**Data Ibu Pria**")
        data['ibu_pria_nama'] = st.text_input("Nama Ibu Pria", key="ip_nama")
        data['ibu_pria_nik'] = st.text_input("NIK Ibu Pria", key="ip_nik")
        data['ibu_pria_bin'] = st.text_input("Binti Ibu Pria", key="ip_bin")
        data['ibu_pria_ttl'] = st.text_input("TTL Ibu Pria", key="ip_ttl")
        data['ibu_pria_kewarganegaraan'] = st.text_input("Kewarganegaraan Ibu Pria", value="WNI", key="ip_kwg")
        data['ibu_pria_agama'] = st.text_input("Agama Ibu Pria", value="Islam", key="ip_agama")
        data['ibu_pria_pekerjaan'] = st.text_input("Pekerjaan Ibu Pria", key="ip_kerja")
        data['ibu_pria_alamat'] = st.text_area("Alamat Ibu Pria", key="ip_alamat")

with tab5:
    st.subheader("Data Orang Tua Catin Wanita")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Data Ayah Wanita**")
        data['ayah_wanita_nama'] = st.text_input("Nama Ayah Wanita", key="aw_nama")
        data['ayah_wanita_nik'] = st.text_input("NIK Ayah Wanita", key="aw_nik")
        data['ayah_wanita_bin'] = st.text_input("Bin Ayah Wanita", key="aw_bin")
        data['ayah_wanita_ttl'] = st.text_input("TTL Ayah Wanita", key="aw_ttl")
        data['ayah_wanita_kewarganegaraan'] = st.text_input("Kewarganegaraan Ayah Wanita", value="WNI", key="aw_kwg")
        data['ayah_wanita_agama'] = st.text_input("Agama Ayah Wanita", value="Islam", key="aw_agama")
        data['ayah_wanita_pekerjaan'] = st.text_input("Pekerjaan Ayah Wanita", key="aw_kerja")
        data['ayah_wanita_alamat'] = st.text_area("Alamat Ayah Wanita", key="aw_alamat")
    with col2:
        st.markdown("**Data Ibu Wanita**")
        data['ibu_wanita_nama'] = st.text_input("Nama Ibu Wanita", key="iw_nama")
        data['ibu_wanita_nik'] = st.text_input("NIK Ibu Wanita", key="iw_nik")
        data['ibu_wanita_bin'] = st.text_input("Binti Ibu Wanita", key="iw_bin")
        data['ibu_wanita_ttl'] = st.text_input("TTL Ibu Wanita", key="iw_ttl")
        data['ibu_wanita_kewarganegaraan'] = st.text_input("Kewarganegaraan Ibu Wanita", value="WNI", key="iw_kwg")
        data['ibu_wanita_agama'] = st.text_input("Agama Ibu Wanita", value="Islam", key="iw_agama")
        data['ibu_wanita_pekerjaan'] = st.text_input("Pekerjaan Ibu Wanita", key="iw_kerja")
        data['ibu_wanita_alamat'] = st.text_area("Alamat Ibu Wanita", key="iw_alamat")

with tab6:
    st.subheader("Detail Pelaksanaan Akad Nikah")
    col1, col2 = st.columns(2)
    with col1:
        data['hari_tgl_akad'] = st.text_input("Hari & Tanggal Akad (misal: Jumat, 20 Oktober 2026)", key="ak_tgl")
        data['waktu_akad'] = st.text_input("Waktu Akad (misal: 08:00 WIB)", key="ak_waktu")
        data['mas_kawin'] = st.text_input("Mas Kawin / Mahar", key="ak_mahar")
        data['tempat_akad'] = st.text_input("Tempat Akad Nikah", key="ak_tempat")
    with col2:
        data['tgl_surat'] = st.text_input("Tanggal Surat Dokumen", key="ak_surat")
        data['status_pria'] = st.selectbox("Status Pria", ["Perjaka", "Duda"], key="ak_stat_pria")
        data['status_wanita'] = st.selectbox("Status Wanita", ["Perawan", "Janda"], key="ak_stat_wanita")

st.markdown("---")

# PROSES & GENERATE
if st.button("🚀 Proses & Generate Berkas", type="primary", use_container_width=True):
    template_file = "BERKAS CATIN .xlsx"
    
    try:
        with st.spinner("Sedang memproses dan mengisikan data ke file Excel..."):
            output_file = update_excel_data(data, template_file)
        
        st.success("✅ Berhasil! Data telah dimasukkan ke sheet **ISIAN DATA**.")
        
        # Tombol Download
        with open(output_file, "rb") as file:
            st.download_button(
                label="📥 Download File BERKAS CATIN Terisi (.xlsx)",
                data=file,
                file_name="BERKAS_CATIN_TERISI.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    except FileNotFoundError as fnf_err:
        st.error(f"❌ {fnf_err}")
        st.info("Pastikan file **'BERKAS CATIN .xlsx'** berada di folder yang sama dengan file `app.py` ini.")
    except Exception as e:
        st.error(f"❌ Terjadi kesalahan saat memproses data: {e}")
