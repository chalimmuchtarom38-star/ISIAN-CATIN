def generate_pdf_isian_data(data_dict, tgl_surat_val):
    if not HAS_REPORTLAB:
        return None
    
    buffer = io.BytesIO()
    
    # Ukuran Kertas F4 dalam Point (612pt x 936pt)
    f4_size = (612, 936)
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=f4_size,
        rightMargin=22,
        leftMargin=22,
        topMargin=15,
        bottomMargin=15
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=11,
        alignment=1, # Center
        spaceAfter=4,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor("#1A365D")
    )
    
    section_style = ParagraphStyle(
        'SectionStyle',
        parent=styles['Heading2'],
        fontSize=8,
        fontName='Helvetica-Bold',
        textColor=colors.HexColor("#1A365D"),
        spaceBefore=0,
        spaceAfter=0
    )
    
    cell_label_style = ParagraphStyle('CellLabel', fontSize=7.2, fontName='Helvetica-Bold', leading=8.5, textColor=colors.HexColor("#2D3748"))
    cell_val_style = ParagraphStyle('CellVal', fontSize=7.2, fontName='Helvetica', leading=8.5, textColor=colors.HexColor("#1A202C"))
    
    ttd_text_style = ParagraphStyle('TTDText', fontSize=8, fontName='Helvetica', alignment=1, leading=10)
    ttd_nama_style = ParagraphStyle('TTDNama', fontSize=8.5, fontName='Helvetica-Bold', alignment=1, leading=10)

    story = []
    story.append(Paragraph("<b>LEMBAR ISIAN DATA CATIN & PELAKSANAAN AKAD</b>", title_style))

    table_data = []
    t_style = [
        ('BACKGROUND', (0,0), (-1,-1), colors.white),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.0),
        ('TOPPADDING', (0,0), (-1,-1), 1.0),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#CBD5E0")),
    ]

    row_idx = 0
    for section, fields in data_dict.items():
        sec_p = Paragraph(f"<b>{section.upper()}</b>", section_style)
        table_data.append([sec_p, ""])
        t_style.append(('SPAN', (0, row_idx), (1, row_idx)))
        t_style.append(('BACKGROUND', (0, row_idx), (1, row_idx), colors.HexColor("#E2E8F0")))
        row_idx += 1
        
        for k, v in fields.items():
            p_k = Paragraph(k, cell_label_style)
            p_v = Paragraph(str(v) if v else "-", cell_val_style)
            table_data.append([p_k, p_v])
            row_idx += 1

    t = Table(table_data, colWidths=[150, 418])
    t.setStyle(TableStyle(t_style))
    story.append(t)
    
    # Menambahkan jarak pembatas ke kolom TTD agar lebih turun dan rapi
    story.append(Spacer(1, 14))
    
    # Blok Tanda Tangan Pengantar / Kasi Pelayanan
    ttd_data = [
        ["", Paragraph(f"Tambi, {tgl_surat_val.replace('TAMBI,', '').strip()}", ttd_text_style)],
        ["", Paragraph("Pengantar / Kasi Pelayanan", ttd_text_style)],
        ["", Spacer(1, 38)], # Ruang TTD diperluas (lebih turun ke bawah)
        ["", Paragraph("<u><b>CHALIM MUCHTAROM, S.Pd.I</b></u>", ttd_nama_style)]
    ]
    
    ttd_table = Table(ttd_data, colWidths=[338, 230])
    ttd_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (1,0), (1,-1), 'CENTER'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1),
        ('TOPPADDING', (0,0), (-1,-1), 1),
    ]))
    
    story.append(ttd_table)
    
    doc.build(story)
    buffer.seek(0)
    return buffer
