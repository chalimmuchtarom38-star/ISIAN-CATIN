def update_excel_data(data, master_path):
    """Memasukkan data dari form khusus ke sheet 'ISIAN DATA' pada Excel"""
    wb = openpyxl.load_workbook(master_path)
    
    # PERBAIKAN: Targetkan langsung sheet "ISIAN DATA"
    if "ISIAN DATA" in wb.sheetnames:
        ws = wb["ISIAN DATA"]
    else:
        ws = wb.active  # Fallback jika nama sheet berbeda

    # =============================================================
    # FUNGSI PENANGANAN MERGED CELL (Mencegah Error Read-Only)
    # =============================================================
    def set_cell_value(cell_address, value):
        try:
            # Coba tulis nilai secara langsung ke sel
            ws[cell_address] = value
        except AttributeError:
            # Jika sel berupa MergedCell read-only, cari sel utama (Top-Left)
            for rng in ws.merged_cells.ranges:
                if cell_address in rng:
                    top_left_cell = rng.start_cell.coordinate
                    ws[top_left_cell] = value
                    break

    # =============================================================
    # PEMETAAN FIELD FORM KE SEL EXCEL (Sheet ISIAN DATA)
    # =============================================================
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

    # Mengisi seluruh data menggunakan fungsi aman set_cell_value
    for key, cell in mapping.items():
        set_cell_value(cell, data.get(key, ""))

    # Hitung Umur Otomatis di Sheet ISIAN DATA
    set_cell_value('C21', hitung_umur(data.get('catin_pria_ttl', '')))
    set_cell_value('F21', hitung_umur(data.get('catin_wanita_ttl', '')))

    # Simpan Hasil
    output_path = "BERKAS_CATIN_TERISI.xlsx"
    wb.save(output_path)
    return output_path
