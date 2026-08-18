# --- FUNGSI GENERATE PDF F4 LENGKAP (1 LEMBAR) ---
def generate_pdf_f4(data):
    buffer = BytesIO()
    # Ukuran F4 / Folio: 215mm x 330mm
    f4_size = (215 * mm, 330 * mm)
    
    # Margin 7mm agar seluruh data & tanda tangan muat rapi dalam 1 lembar utuh
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
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=10.5,
        leading=12,
        alignment=1, # Center
        fontName='Helvetica-Bold'
    )
    
    sec_title_style = ParagraphStyle(
        'SecTitleStyle',
        fontSize=8,
        leading=9.5,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor('#003366')
    )
    
    lbl_style = ParagraphStyle('LblStyle', fontSize=7, leading=8.5, fontName='Helvetica-Bold')
    val_style = ParagraphStyle('ValStyle', fontSize=7, leading=8.5, fontName='Helvetica')
    
    elements = []
    
    # KOP / JUDUL DOKUMEN
    elements.append(Paragraph("PEMERINTAH KABUPATEN PEMALANG - KECAMATAN WATUKUMPUL", ParagraphStyle('Kop1', fontSize=7.5, alignment=1, fontName='Helvetica-Bold')))
    elements.append(Paragraph("RINGKASAN LEMBAR VERIFIKASI BERKAS CALON PENGANTIN DESA TAMBI", title_style))
    elements.append(Paragraph(f"No. Register: <b>{data.get('no_register', '-')}</b> | Tanggal Surat: <b>{data.get('tgl_surat', '-')}</b>", ParagraphStyle('SubTitle', fontSize=7.5, alignment=1, leading=9)))
    elements.append(Spacer(1, 3))
    
    # Helper baris tabel tunggal
    def row1(lbl, val):
        return [Paragraph(lbl, lbl_style), Paragraph(":", lbl_style), Paragraph(str(val or '-'), val_style)]

    # Helper baris tabel ganda (LK & PR side by side)
    def row2(lbl1, val1, lbl2, val2):
        return [
            Paragraph(lbl1, lbl_style), Paragraph(":", lbl_style), Paragraph(str(val1 or '-'), val_style),
            Paragraph(lbl2, lbl_style), Paragraph(":", lbl_style), Paragraph(str(val2 or '-'), val_style)
        ]

    # I. PELAKSANAAN AKAD NIKAH
    hari_tgl_akad = f"{get_hari_tgl(data.get('tgl_pelaksanaan', '-'))} (Jam: {data.get('jam_akad', '-')})"
    tabel_akad_data = [
        [Paragraph("I. PELAKSANAAN AKAD NIKAH", sec_title_style), "", ""],
        row1("Hari & Tgl / Jam Akad", hari_tgl_akad),
        row1("Tempat Akad Nikah", data.get('tempat_akad', '-')),
        row1("Maskawin / Mahar", data.get('mahar', '-')),
        row1("Email Catin", data.get('email_catin', '-')),
    ]

    # II & III. CATIN LAKI-LAKI & PEREMPUAN
    tabel_catin_data = [
        [Paragraph("II. CALON PENGANTIN LAKI-LAKI", sec_title_style), "", "", Paragraph("III. CALON PENGANTIN PEREMPUAN", sec_title_style), "", ""],
        row2("Nama Lengkap", data.get('nama_lk','-'), "Nama Lengkap", data.get('nama_pr','-')),
        row2("Bin / Binti", data.get('bin_lk','-'), "Bin / Binti", data.get('binti_pr','-')),
        row2("NIK", data.get('nik_lk','-'), "NIK", data.get('nik_pr','-')),
        row2("Tempat, Tgl Lahir", data.get('ttl_lk','-'), "Tempat, Tgl Lahir", data.get('ttl_pr','-')),
        row2("Umur Catin", f"{hitung_umur(data.get('ttl_lk'))} Tahun", "Umur Catin", f"{hitung_umur(data.get('ttl_pr'))} Tahun"),
        row2("Status / Gender", f"{data.get('status_lk','-')} / {data.get('jk_lk','-')}", "Status / Gender", f"{data.get('status_pr','-')} / {data.get('jk_pr','-')}"),
        row2("Pekerjaan", data.get('pekerjaan_lk','-'), "Pekerjaan", data.get('pekerjaan_pr','-')),
        row2("Pendidikan", data.get('pendidikan_lk','-'), "Pendidikan", data.get('pendidikan_pr','-')),
        row2("Ex Pasangan", data.get('istri_terdahulu','-'), "Ex Pasangan", data.get('suami_terdahulu','-')),
        row2("Alamat Lengkap", data.get('alamat_lk','-'), "Alamat Lengkap", data.get('alamat_pr','-')),
    ]

    # IV & V. ORANG TUA LAKI-LAKI & PEREMPUAN
    tabel_ortu_data = [
        [Paragraph("IV. ORANG TUA LAKI-LAKI", sec_title_style), "", "", Paragraph("V. ORANG TUA PEREMPUAN", sec_title_style), "", ""],
        row2("Ayah / Bin", f"{data.get('nama_ayah_lk','-')} bin {data.get('bin_ayah_lk','-')}", "Ayah / Bin", f"{data.get('nama_ayah_pr','-')} bin {data.get('bin_ayah_pr','-')}"),
        row2("NIK / TTL Ayah", f"{data.get('nik_ayah_lk','-')} / {data.get('ttl_ayah_lk','-')}", "NIK / TTL Ayah", f"{data.get('nik_ayah_pr','-')} / {data.get('ttl_ayah_pr','-')}"),
        row2("Pekerjaan Ayah", data.get('pekerjaan_ayah_lk','-'), "Pekerjaan Ayah", data.get('pekerjaan_ayah_pr','-')),
        row2("Alamat Ayah", data.get('alamat_ayah_lk','-'), "Alamat Ayah", data.get('alamat_ayah_pr','-')),
        row2("Ibu / Binti", f"{data.get('nama_ibu_lk','-')} bin {data.get('bin_ibu_lk','-')}", "Ibu / Binti", f"{data.get('nama_ibu_pr','-')} bin {data.get('bin_ibu_pr','-')}"),
        row2("NIK / TTL Ibu", f"{data.get('nik_ibu_lk','-')} / {data.get('ttl_ibu_lk','-')}", "NIK / TTL Ibu", f"{data.get('nik_ibu_pr','-')} / {data.get('ttl_ibu_pr','-')}"),
        row2("Pekerjaan Ibu", data.get('pekerjaan_ibu_lk','-'), "Pekerjaan Ibu", data.get('pekerjaan_ibu_pr','-')),
        row2("Alamat Ibu", data.get('alamat_ibu_lk','-'), "Alamat Ibu", data.get('alamat_ibu_pr','-')),
    ]

    # VI & VII. WALI NIKAH & SAKSI-SAKSI
    tabel_wali_saksi = [
        [Paragraph("VI. DATA WALI NIKAH", sec_title_style), "", "", Paragraph("VII. DATA SAKSI-SAKSI AKAD", sec_title_style), "", ""],
        row2("Nama Wali", f"{data.get('nama_wali','-')} bin {data.get('bin_wali','-')}", "Saksi 1", data.get('saksi1_nama','-')),
        row2("NIK / TTL Wali", f"{data.get('nik_wali','-')} / {data.get('ttl_wali','-')}", "NIK / TTL Saksi 1", f"{data.get('saksi1_nik','-')} / {data.get('saksi1_ttl','-')}"),
        row2("Hubungan Wali", data.get('hubungan_wali','-'), "Pekerjaan Saksi 1", data.get('saksi1_pekerjaan','-')),
        row2("Pekerjaan Wali", data.get('pekerjaan_wali','-'), "Alamat Saksi 1", data.get('saksi1_alamat','-')),
        row2("Alamat Wali", data.get('alamat_wali','-'), "Saksi 2", data.get('saksi2_nama','-')),
        row2("Wali Lengkap", data.get('nama_wali_lengkap','-'), "NIK / TTL Saksi 2", f"{data.get('saksi2_nik','-')} / {data.get('saksi2_ttl','-')}"),
        row2("", "", "Pekerjaan Saksi 2", data.get('saksi2_pekerjaan','-')),
        row2("", "", "Alamat Saksi 2", data.get('saksi2_alamat','-')),
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
    
    # 1. KOTAK TANDA TANGAN CATIN & WALI
    ttd_catin = [
        [Paragraph("Catin Laki-Laki", lbl_style), Paragraph("Catin Perempuan", lbl_style), Paragraph("Wali Nikah", lbl_style)],
        [Spacer(1, 20), Spacer(1, 20), Spacer(1, 20)],  # Jarak ruang tanda tangan longgar & rapi
        [
            Paragraph(f"( <b>{data.get('nama_lk','...')}</b> )", val_style), 
            Paragraph(f"( <b>{data.get('nama_pr','...')}</b> )", val_style), 
            Paragraph(f"( <b>{data.get('nama_wali','...')}</b> )", val_style)
        ]
    ]
    t_ttd1 = Table(ttd_catin, colWidths=[67*mm, 67*mm, 67*mm])
    t_ttd1.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    
    # 2. KOTAK TANDA TANGAN PEMERINTAH DESA (KASI PELAYANAN SEBELAH KIRI, KEPALA DESA SEBELAH KANAN)
    tgl_surat_str = data.get('tgl_surat', 'Tambi, ................. 2026')
    ttd_pemdes = [
        [
            Paragraph("Mengetahui,<br/><b>KASI PELAYANAN DESA TAMBI</b>", lbl_style), 
            Paragraph(f"{tgl_surat_str}<br/><b>KEPALA DESA TAMBI</b>", lbl_style)
        ],
        [Spacer(1, 24), Spacer(1, 24)],  # Jarak tanda tangan diperlebar (24pt) agar tidak mepet & rapi
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

    elements.extend([t_ttd1, Spacer(1, 8), t_ttd2])
    
    doc.build(elements)
    pdf_data = buffer.getvalue()
    buffer.close()
    return pdf_data
