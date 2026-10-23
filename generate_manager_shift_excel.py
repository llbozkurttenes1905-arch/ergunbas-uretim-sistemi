import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_clean_manager_excel():
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # STYLES & FONTS
    # -------------------------------------------------------------
    font_title = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    font_sub = Font(name="Segoe UI", size=9.5, italic=True, color="E2E8F0")
    font_grp = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_head = Font(name="Segoe UI", size=9.5, bold=True, color="FFFFFF")
    font_bold = Font(name="Segoe UI", size=9.5, bold=True, color="0F172A")
    font_reg = Font(name="Segoe UI", size=9, color="1E293B")
    font_blue = Font(name="Segoe UI", size=9, bold=True, color="0369A1")
    font_green = Font(name="Segoe UI", size=9.5, bold=True, color="047857")
    font_amber = Font(name="Segoe UI", size=9.5, bold=True, color="B45309")
    font_red_bold = Font(name="Segoe UI", size=9.5, bold=True, color="FFFFFF")
    
    fill_dark = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    fill_shift = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_prod = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    fill_cap = PatternFill(start_color="B45309", end_color="B45309", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_red = PatternFill(start_color="DC2626", end_color="DC2626", fill_type="solid")
    fill_green_box = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    double_bottom = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='94A3B8'),
        bottom=Side(style='double', color='0F172A')
    )

    align_c = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_l = Alignment(horizontal='left', vertical='center')
    align_r = Alignment(horizontal='right', vertical='center')

    # =============================================================
    # SAYFA 1: MÜDÜR ÖZET HAT MATRİSİ (Fotoğraftaki 1 Satır = 1 Hat)
    # =============================================================
    ws1 = wb.active
    ws1.title = "Müdür Özet Hat Matrisi"
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:K1")
    ws1["A1"] = "ERGUNBAS GROUP — MAKİNE & HAT VARDİYA ÖZETİ VE KAPASİTE MATRİSİ"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_dark
    ws1["A1"].alignment = align_c
    ws1.row_dimensions[1].height = 34

    ws1.merge_cells("A2:K2")
    ws1["A2"] = "1. Vardiya: 08:30 / 20:30 (720 dk Net) | 2. Vardiya: 20:30 / 08:30 (720 dk Net) | 24 Saat Kesintisiz Üretim | Hat 101 - 109 & Hat 201 ve Kalite Kontrol"
    ws1["A2"].font = font_sub
    ws1["A2"].fill = fill_dark
    ws1["A2"].alignment = align_c
    ws1.row_dimensions[2].height = 20

    headers_s1 = [
        ("İSTASYON / HAT", 16, align_l),
        ("MESAİ ARALIĞI", 22, align_c),
        ("MESAİ (DK)", 13, align_r),
        ("NET ÇALIŞMA (DK)", 16, align_r),
        ("A.V. ÇALIŞAN", 14, align_c),
        ("B.V. ÇALIŞAN", 14, align_c),
        ("GÜNLÜK KAPASİTE (24s)", 22, align_c),
        ("BAŞLICA ÜRÜNLER", 34, align_l),
        ("STANDART BOY", 15, align_c),
        ("GÜNLÜK KAPI KARŞILIĞI", 20, align_c),
        ("DURUM / NOT", 20, align_c)
    ]

    for col_idx, (h_title, col_w, h_align) in enumerate(headers_s1, 1):
        c = ws1.cell(row=3, column=col_idx, value=h_title)
        c.font = font_head
        c.fill = fill_dark
        c.alignment = align_c
        c.border = thin_border
        ws1.column_dimensions[get_column_letter(col_idx)].width = col_w
    ws1.row_dimensions[3].height = 28

    summary_hats = [
        ("Hat 101", "08:30/20:30 - 20:30/08:30", 720, 720, 1, 1, "450 Boy", "80x80 & 35x100 Pervazlar", "2.20 m", "~90 Kapılık Pervaz", "Aktif Üretim"),
        ("Hat 102", "08:30/20:30 - 20:30/08:30", 720, 720, 1, 1, "374 Boy", "100 mm Kasa, Cam Çıtası, U Kapama", "2.15 m", "150 Kapılık Kasa", "Aktif Üretim"),
        ("Hat 103", "08:30/20:30 - 20:30/08:30", 720, 720, 1, 1, "756 Boy", "50x80 Pervaz, Masif Seren", "2.20 m", "151 Kapılık Pervaz", "Aktif Üretim"),
        ("Hat 104", "08:30/20:30 - 20:30/08:30", 720, 720, 1, 1, "677 Boy", "Masif Kapı Sereni (38x45 mm)", "2.10 m", "193 Kapılık Seren", "Yüksek Hızlı Seren"),
        ("Hat 105", "08:30/20:30 - 20:30/08:30", 720, 720, 1, 1, "641 Boy", "Masif Kapı Sereni (38x45 mm)", "2.10 m", "183 Kapılık Seren", "Yüksek Hızlı Seren"),
        ("Hat 106", "08:30/20:30 - 20:30/08:30", 720, 720, 1, 1, "336 Boy", "140 mm Kasa, 200 mm Kasa", "2.15 m", "134 Kapılık Kasa", "Aktif Üretim"),
        ("Hat 107", "08:30/20:30 - 20:30/08:30", 720, 720, 1, 1, "670 Boy", "200 Kasa, Seren 39mm, Zemin Deck", "2.10 m", "191 Kapılık Seren", "Geniş Kasa & Deck"),
        ("Hat 108", "08:30/20:30 - 20:30/08:30", 720, 720, 1, 1, "413 Boy", "100 mm Kasa, 140 mm Kasa", "2.15 m", "165 Kapılık Kasa", "Aktif Üretim"),
        ("Hat 109", "08:30/20:30 - 20:30/08:30", 720, 720, 1, 1, "1.169 Boy", "50x80 Pervaz, 80x80 Pervaz", "2.20 m", "234 Kapılık Pervaz", "Yüksek Hızlı Pervaz"),
        ("Hat 201 (Levha)", "08:30/20:30 - 20:30/08:30", 720, 720, 2, 2, "1.000 Plaka", "WPC Kapı Levhaları (Tüm Renkler)", "2.07 m", "500 KAPI / GÜN", "Ana Levha Hattı"),
        ("KALİTE KONTROL", "08:30/20:30 - 20:30/08:30", 720, 720, 3, 3, "4.200 Adet", "Tüm Profil Boyları ve Levha Plakaları", "Tüm Boylar", "Tüm Fabrika Çıkışı", "Son Kontrol & Onay")
    ]

    r_idx = 4
    for item in summary_hats:
        (h_name, mesai, m_dk, net_dk, av, bv, cap, prods, boy, kapi, durum) = item
        row_fill = fill_zebra if r_idx % 2 == 0 else PatternFill(fill_type=None)
        
        vals = [
            (h_name, font_bold, align_l),
            (mesai, font_reg, align_c),
            (m_dk, font_reg, align_r),
            (net_dk, font_bold, align_r),
            (av, font_bold, align_c),
            (bv, font_bold, align_c),
            (cap, font_red_bold if "Plaka" in cap or "1.169" in cap else font_bold, align_c),
            (prods, font_reg, align_l),
            (boy, font_green, align_c),
            (kapi, font_green if "500" in kapi else font_bold, align_c),
            (durum, font_blue, align_c)
        ]
        
        for c_idx, (val, f, al) in enumerate(vals, 1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=val)
            cell.font = f
            cell.alignment = al
            cell.border = thin_border
            if c_idx == 7 and ("Plaka" in cap or "1.169" in cap):
                cell.fill = fill_red
            elif c_idx == 9:
                cell.fill = fill_green_box
            elif row_fill.fill_type:
                cell.fill = row_fill
        ws1.row_dimensions[r_idx].height = 24
        r_idx += 1

    # Total row Sayfa 1
    ws1.cell(row=r_idx, column=1, value="GENEL TOPLAM").font = font_bold
    ws1.cell(row=r_idx, column=5, value="=SUM(E4:E14)").font = font_bold # 13 çalışan
    ws1.cell(row=r_idx, column=5).alignment = align_c
    ws1.cell(row=r_idx, column=6, value="=SUM(F4:F14)").font = font_bold # 13 çalışan
    ws1.cell(row=r_idx, column=6).alignment = align_c
    ws1.cell(row=r_idx, column=7, value="~7.000 Boy + 1.000 Plaka").font = font_bold
    ws1.cell(row=r_idx, column=7).alignment = align_c
    ws1.cell(row=r_idx, column=10, value="500 KAPI / GÜN KAPASİTE").font = font_green
    ws1.cell(row=r_idx, column=10).alignment = align_c

    for c in range(1, 12):
        cell = ws1.cell(row=r_idx, column=c)
        cell.fill = fill_total
        cell.border = double_bottom
    ws1.row_dimensions[r_idx].height = 26
    ws1.freeze_panes = "B4"


    # =============================================================
    # SAYFA 2: ÜRÜN BAZLI DETAYLI KAPASİTE MATRİSİ (SADE & METRE CİNSİNDEN)
    # =============================================================
    ws2 = wb.create_sheet(title="Ürün Detaylı Kapasite Matrisi")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:Q1")
    ws2["A1"] = "ERGUNBAS GROUP — ÜRÜN VE HAT BAZLI DETAYLI VARDİYA KAPASİTE MATRİSİ"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_dark
    ws2["A1"].alignment = align_c
    ws2.row_dimensions[1].height = 34

    ws2.merge_cells("A2:Q2")
    ws2["A2"] = "Tüm Ölçüler Metre Cinsinden (2.07 m Levha, 2.10 m Seren, 2.20 m Pervaz, 2.15 m Kasa) | 720 dk (12s) Vardiya ve 1.440 dk (24s) Günlük Net Kapasite"
    ws2["A2"].font = font_sub
    ws2["A2"].fill = fill_dark
    ws2["A2"].alignment = align_c
    ws2.row_dimensions[2].height = 20

    # Super headers
    ws2.merge_cells("A3:F3")
    ws2["A3"] = "İSTASYON & MESAİ BİLGİSİ (720 DK VARDİYA)"
    ws2["A3"].font = font_grp
    ws2["A3"].fill = fill_shift
    ws2["A3"].alignment = align_c

    ws2.merge_cells("G3:K3")
    ws2["G3"] = "ÜRÜN VE METRE BOYU ÖLÇÜLERİ"
    ws2["G3"].font = font_grp
    ws2["G3"].fill = fill_prod
    ws2["G3"].alignment = align_c

    ws2.merge_cells("L3:Q3")
    ws2["L3"] = "KAPASİTE VE KAPI EŞDEĞERLERİ (12s VARDİYA / 24s GÜNLÜK NET)"
    ws2["L3"].font = font_grp
    ws2["L3"].fill = fill_cap
    ws2["L3"].alignment = align_c
    ws2.row_dimensions[3].height = 25

    headers_s2 = [
        ("İSTASYON / HAT", 16, align_l),
        ("MESAİ ARALIĞI", 18, align_c),
        ("MESAİ (DK)", 13, align_r),
        ("NET ÇALIŞMA (DK)", 15, align_r),
        ("A.V. ÇALIŞAN", 13, align_c),
        ("B.V. ÇALIŞAN", 13, align_c),
        
        # Ürün & Boy Detayları (Tek göz/çift göz yok, sade boy metre cinsinden)
        ("ÜRÜN ADI", 24, align_l),
        ("KATEGORİ", 16, align_l),
        ("BOY (METRE)", 14, align_c), # SADE METRE CİNSİNDEN: 2.07 m, 2.10 m, 2.20 m, 2.15 m
        ("KESİT ÖLÇÜSÜ", 18, align_l),
        ("BİRİM AĞIRLIK", 14, align_r),
        
        # Hız ve Kapasite
        ("SAATLİK HIZ", 15, align_r),
        ("12s VARDİYA KAPASİTESİ", 18, align_r),
        ("24s GÜNLÜK KAPASİTE", 18, align_r),
        ("GÜNLÜK TONAJ (KG)", 16, align_r),
        ("KAPI EŞDEĞERİ (GÜNLÜK)", 18, align_r),
        ("PLANLAMA NOTU", 22, align_l)
    ]

    for col_idx, (h_title, col_w, h_align) in enumerate(headers_s2, 1):
        c = ws2.cell(row=4, column=col_idx, value=h_title)
        c.font = font_head
        c.fill = fill_dark
        c.alignment = align_c
        c.border = thin_border
        ws2.column_dimensions[get_column_letter(col_idx)].width = col_w
    ws2.row_dimensions[4].height = 28

    # Detail items (Boy is ONLY in meters: '2.07 m', '2.10 m', '2.20 m', '2.15 m', '3.00 m')
    detail_data = [
        # HAT 101
        ("Hat 101", "08:30 / 20:30 (2V)", 720, 720, 1, 1, True,
         "80X80 PERVAZ", "Pervaz Profili", "2.20 m", "80 x 80 mm", 2.20, "Boy", 20.3, 5.0, "Kapı", "Ana Hat"),
        ("Hat 101", "20:30 / 08:30", 720, 720, 1, 1, False,
         "35X100 PERVAZ", "Pervaz Profili", "2.20 m", "35 x 100 mm", 2.08, "Boy", 21.8, 5.0, "Kapı", "Ana Hat"),
        ("Hat 101", "—", 720, 720, 1, 1, False,
         "35X150 PERVAZ", "Pervaz Profili", "2.20 m", "35 x 150 mm", 2.97, "Boy", 13.9, 5.0, "Kapı", "Ana Hat"),
        ("Hat 101", "—", 720, 720, 1, 1, False,
         "35X80 PERVAZ", "Pervaz Profili", "2.20 m", "35 x 80 mm", 2.94, "Boy", 15.0, 5.0, "Kapı", "Ana Hat"),

        # HAT 102
        ("Hat 102", "08:30 / 20:30 (2V)", 720, 720, 1, 1, True,
         "100 mm KASA", "Kasa Profili", "2.15 m", "100 x 36 mm", 4.00, "Boy", 15.6, 2.5, "Kapı", "Ana Hat"),
        ("Hat 102", "20:30 / 08:30", 720, 720, 1, 1, False,
         "140 mm KASA", "Kasa Profili", "2.15 m", "140 x 36 mm", 4.95, "Boy", 10.4, 2.5, "Kapı", "Yedek Hat"),
        ("Hat 102", "—", 720, 720, 1, 1, False,
         "CAM ÇITASI", "Cam Çıtası", "2.20 m", "15 x 20 mm", 1.50, "Boy", 33.2, 4.0, "Kapı", "Ana Hat"),
        ("Hat 102", "—", 720, 720, 1, 1, False,
         "U KAPAMA", "U Kapama / Lambiri", "2.20 m", "30 x 15 mm", 0.48, "Boy", 85.3, 0.0, "—", "Ana Hat"),

        # HAT 103
        ("Hat 103", "08:30 / 20:30 (2V)", 720, 720, 1, 1, True,
         "50X80 PERVAZ", "Pervaz Profili", "2.20 m", "50 x 80 mm", 1.55, "Boy", 31.5, 5.0, "Kapı", "Yedek Hat"),
        ("Hat 103", "20:30 / 08:30", 720, 720, 1, 1, False,
         "SEREN", "Seren Çıtası", "2.10 m", "38 x 45 mm", 3.39, "Boy", 37.5, 3.5, "Kapı", "Alternatif Hat"),

        # HAT 104
        ("Hat 104", "08:30 / 20:30 (2V)", 720, 720, 1, 1, True,
         "SEREN", "Seren Çıtası", "2.10 m", "38 x 45 mm", 3.39, "Boy", 28.2, 3.5, "Kapı", "Öncelikli Ana Hat"),

        # HAT 105
        ("Hat 105", "08:30 / 20:30 (2V)", 720, 720, 1, 1, True,
         "SEREN", "Seren Çıtası", "2.10 m", "38 x 45 mm", 3.39, "Boy", 26.7, 3.5, "Kapı", "Öncelikli Ana Hat"),
        ("Hat 105", "20:30 / 08:30", 720, 720, 1, 1, False,
         "140 mm KASA", "Kasa Profili", "2.15 m", "140 x 36 mm", 4.95, "Boy", 17.4, 2.5, "Kapı", "Alternatif Hat"),

        # HAT 106
        ("Hat 106", "08:30 / 20:30 (2V)", 720, 720, 1, 1, True,
         "140 mm KASA", "Kasa Profili", "2.15 m", "140 x 36 mm", 4.95, "Boy", 14.0, 2.5, "Kapı", "Öncelikli Ana Hat"),
        ("Hat 106", "20:30 / 08:30", 720, 720, 1, 1, False,
         "200 mm KASA", "Kasa Profili", "2.15 m", "200 x 36 mm", 6.66, "Boy", 20.5, 2.5, "Kapı", "Yedek Hat"),
        ("Hat 106", "—", 720, 720, 1, 1, False,
         "100 mm KASA", "Kasa Profili", "2.15 m", "100 x 36 mm", 4.00, "Boy", 13.5, 2.5, "Kapı", "Yedek Hat"),

        # HAT 107
        ("Hat 107", "08:30 / 20:30 (2V)", 720, 720, 1, 1, True,
         "200 mm KASA", "Kasa Profili", "2.15 m", "200 x 36 mm", 6.66, "Boy", 12.5, 2.5, "Kapı", "Öncelikli Ana Hat"),
        ("Hat 107", "20:30 / 08:30", 720, 720, 1, 1, False,
         "SEREN(39mm)", "Seren Çıtası", "2.10 m", "39 x 45 mm", 3.81, "Boy", 27.9, 3.5, "Kapı", "Özel Kalın Seren"),
        ("Hat 107", "—", 720, 720, 1, 1, False,
         "KT100-220x24-ANT", "Deck / WPC Zemin", "3.00 m", "220 x 24 mm", 10.33, "Boy", 8.2, 0.0, "—", "Zemin Deck"),
        ("Hat 107", "—", 720, 720, 1, 1, False,
         "KT102-219x26-ANT", "Deck / WPC Zemin", "3.00 m", "219 x 26 mm", 10.52, "Boy", 10.3, 0.0, "—", "Zemin Deck"),
        ("Hat 107", "—", 720, 720, 1, 1, False,
         "SEREN", "Seren Çıtası", "2.10 m", "38 x 45 mm", 3.39, "Boy", 20.6, 3.5, "Kapı", "Alternatif Hat"),

        # HAT 108
        ("Hat 108", "08:30 / 20:30 (2V)", 720, 720, 1, 1, True,
         "100 mm KASA", "Kasa Profili", "2.15 m", "100 x 36 mm", 4.00, "Boy", 17.2, 2.5, "Kapı", "Öncelikli Ana Hat"),
        ("Hat 108", "20:30 / 08:30", 720, 720, 1, 1, False,
         "140 mm KASA", "Kasa Profili", "2.15 m", "140 x 36 mm", 4.95, "Boy", 14.7, 2.5, "Kapı", "Yedek Hat"),

        # HAT 109
        ("Hat 109", "08:30 / 20:30 (2V)", 720, 720, 1, 1, True,
         "50X80 PERVAZ", "Pervaz Profili", "2.20 m", "50 x 80 mm", 1.55, "Boy", 48.7, 5.0, "Kapı", "Yüksek Kapasite"),
        ("Hat 109", "20:30 / 08:30", 720, 720, 1, 1, False,
         "80X80 PERVAZ", "Pervaz Profili", "2.20 m", "80 x 80 mm", 2.20, "Boy", 57.5, 5.0, "Kapı", "Yüksek Kapasite"),

        # HAT 201 (LEVHA MAKİNESİ) - 2.07 m
        ("Hat 201 (Levha)", "08:30 / 20:30 (2V)", 720, 720, 2, 2, True,
         "DIŞBUDAK Levha", "WPC Levha Plaka", "2.07 m", "En: 94 cm (2 mm)", 4.77, "Plaka", 42.7, 2.0, "Kapı", "En Yüksek Hacim"),
        ("Hat 201 (Levha)", "20:30 / 08:30", 720, 720, 2, 2, False,
         "TEAK Levha", "WPC Levha Plaka", "2.07 m", "En: 94 cm (2 mm)", 4.25, "Plaka", 38.6, 2.0, "Kapı", "Seri Üretim"),
        ("Hat 201 (Levha)", "—", 720, 720, 2, 2, False,
         "MIX TEAK Levha", "WPC Levha Plaka", "2.07 m", "En: 94 cm (2 mm)", 4.25, "Plaka", 40.1, 2.0, "Kapı", "Seri Üretim"),
        ("Hat 201 (Levha)", "—", 720, 720, 2, 2, False,
         "BEYAZ TEAK Levha", "WPC Levha Plaka", "2.07 m", "En: 94 cm (2 mm)", 4.41, "Plaka", 40.6, 2.0, "Kapı", "Seri Üretim"),
        ("Hat 201 (Levha)", "—", 720, 720, 2, 2, False,
         "MIX BEYAZ TEAK Levha", "WPC Levha Plaka", "2.07 m", "En: 94 cm (2 mm)", 4.25, "Plaka", 44.1, 2.0, "Kapı", "Yüksek Hızlı"),
        ("Hat 201 (Levha)", "—", 720, 720, 2, 2, False,
         "AY GRİ Levha", "WPC Levha Plaka", "2.07 m", "En: 94 cm (2 mm)", 4.64, "Plaka", 40.0, 2.0, "Kapı", "Seri Üretim"),
        ("Hat 201 (Levha)", "—", 720, 720, 2, 2, False,
         "SONOMO Levha", "WPC Levha Plaka", "2.07 m", "En: 94 cm (2 mm)", 4.33, "Plaka", 38.4, 2.0, "Kapı", "Seri Üretim"),
        ("Hat 201 (Levha)", "—", 720, 720, 2, 2, False,
         "COCO Levha", "WPC Levha Plaka", "2.07 m", "En: 94 cm (2 mm)", 4.79, "Plaka", 41.4, 2.0, "Kapı", "Seri Üretim"),
        ("Hat 201 (Levha)", "—", 720, 720, 2, 2, False,
         "ANTRASİT Levha", "WPC Levha Plaka", "2.07 m", "En: 94 cm (2 mm)", 4.77, "Plaka", 40.0, 2.0, "Kapı", "Seri Üretim"),
        ("Hat 201 (Levha)", "—", 720, 720, 2, 2, False,
         "VIZON Levha", "WPC Levha Plaka", "2.07 m", "En: 94 cm (2 mm)", 4.77, "Plaka", 34.0, 2.0, "Kapı", "Seri Üretim"),
        ("Hat 201 (Levha)", "—", 720, 720, 2, 2, False,
         "KOYU CEVİZ Levha", "WPC Levha Plaka", "2.07 m", "En: 94 cm (2 mm)", 4.79, "Plaka", 28.8, 2.0, "Kapı", "Seri Üretim"),

        # KALİTE KONTROL
        ("KALİTE KONTROL", "08:30 / 20:30 (2V)", 720, 720, 3, 3, True,
         "TÜM PROFİL & LEVHALAR", "Kalite Güvence", "Tüm Boylar", "Tüm Kesitler", 0.0, "Adet", 175.0, 0.0, "4.200 Adet/Gün", "Son Kontrol")
    ]

    cur_row = 5
    for r in detail_data:
        (hat_name, mesai_araligi, mesai_dk, net_dk, av_calisan, bv_calisan, is_line_primary,
         prod_name, cat, boy, kesit, unit_kg, unit_lbl, spd_h, kapi_factor, kapi_unit, plan_notu) = r

        cap_12s = round(spd_h * 12.0)
        cap_24s = round(spd_h * 24.0)
        daily_tonaj = round(cap_24s * unit_kg, 1) if unit_kg > 0 else 0.0
        kapi_24s = round(cap_24s / kapi_factor) if kapi_factor > 0 else 0
        kapi_str = f"{kapi_24s} Kapı/Gün" if kapi_factor > 0 else (kapi_unit if kapi_unit != "—" else "—")

        row_fill = fill_zebra if cur_row % 2 == 0 else PatternFill(fill_type=None)

        vals = [
            (hat_name if is_line_primary else "↳ " + hat_name, font_bold if is_line_primary else font_blue, align_l, "@"),
            (mesai_araligi, font_reg, align_c, "@"),
            (mesai_dk, font_reg, align_r, "#,##0"),
            (net_dk, font_bold, align_r, "#,##0"),
            (av_calisan if is_line_primary else "—", font_bold if is_line_primary else font_reg, align_c, "@" if not is_line_primary else "#,##0"),
            (bv_calisan if is_line_primary else "—", font_bold if is_line_primary else font_reg, align_c, "@" if not is_line_primary else "#,##0"),
            (prod_name, font_bold, align_l, "@"),
            (cat, font_reg, align_l, "@"),
            (boy, font_green, align_c, "@"), # SADECE METRE CİNSİNDEN: 2.07 m
            (kesit, font_reg, align_l, "@"),
            (unit_kg if unit_kg > 0 else "—", font_reg, align_r, '0.00 "kg"' if unit_kg > 0 else "@"),
            (round(spd_h, 1), font_green, align_r, f'#,##0.0 " {unit_lbl}/sa"'),
            (cap_12s, font_reg, align_r, f'#,##0 " {unit_lbl}"'),
            (cap_24s, font_bold, align_r, f'#,##0 " {unit_lbl}"'),
            (daily_tonaj if daily_tonaj > 0 else "—", font_reg, align_r, '#,##0.0 " kg"' if daily_tonaj > 0 else "@"),
            (kapi_str, font_green if kapi_24s > 0 else font_reg, align_r, "@"),
            (plan_notu, font_blue, align_l, "@")
        ]

        for col_idx, (val, c_font, c_align, num_fmt) in enumerate(vals, 1):
            cell = ws2.cell(row=cur_row, column=col_idx, value=val)
            cell.font = c_font
            cell.alignment = c_align
            cell.border = thin_border
            if col_idx == 9: # Boy sütunu
                cell.fill = fill_green_box
            elif row_fill.fill_type:
                cell.fill = row_fill
            if num_fmt and num_fmt != "@":
                cell.number_format = num_fmt

        ws2.row_dimensions[cur_row].height = 20
        cur_row += 1

    last_data_row_s2 = cur_row - 1
    ws2.cell(row=cur_row, column=1, value="GENEL TOPLAM (FİLTREYE DUYARLI)").font = font_bold
    ws2.cell(row=cur_row, column=1).alignment = align_l

    ws2.cell(row=cur_row, column=13, value=f"=SUBTOTAL(109, M5:M{last_data_row_s2})").font = font_bold
    ws2.cell(row=cur_row, column=13).number_format = '#,##0 " Adet/Boy"'

    ws2.cell(row=cur_row, column=14, value=f"=SUBTOTAL(109, N5:N{last_data_row_s2})").font = font_bold
    ws2.cell(row=cur_row, column=14).number_format = '#,##0 " Adet/Boy"'

    for c in range(1, 18):
        cell = ws2.cell(row=cur_row, column=c)
        cell.fill = fill_total
        cell.border = double_bottom
    ws2.row_dimensions[cur_row].height = 24

    ws2.auto_filter.ref = f"A4:Q{last_data_row_s2}"
    ws2.freeze_panes = "B5"

    # Save to desktop and static
    dest_desktop = r"C:\Users\llboz\OneDrive\Desktop\Ergunbas_Ekstruzyon_ve_Kalite_Kontrol_Vardiya_Kapasite_Matrisi.xlsx"
    dest_alt = r"C:\Users\llboz\OneDrive\Desktop\Ergunbas_Ekstruzyon_ve_Kalite_Kontrol_Vardiya_Kapasite_Matrisi (Sade_Metre).xlsx"
    dest_static = r"C:\Users\llboz\OneDrive\Desktop\ergunbas-uretim-sistemi\static\Ergunbas_Ekstruzyon_ve_Kalite_Kontrol_Vardiya_Kapasite_Matrisi.xlsx"

    try:
        wb.save(dest_desktop)
        print(f"Saved: {dest_desktop}")
    except Exception as e:
        print(f"Desktop ana dosya ({e})")

    try:
        wb.save(dest_alt)
        print(f"Saved: {dest_alt}")
    except Exception as e:
        print(f"Alt dosya ({e})")

    try:
        wb.save(dest_static)
        print(f"Saved: {dest_static}")
    except:
        pass

if __name__ == "__main__":
    build_clean_manager_excel()
