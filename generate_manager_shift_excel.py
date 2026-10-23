import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Hat & İstasyon Kapasiteleri"
ws.views.sheetView[0].showGridLines = True

# Fonts
font_title = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
font_sub = Font(name="Segoe UI", size=9.5, italic=True, color="E2E8F0")
font_grp = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
font_head = Font(name="Segoe UI", size=9.5, bold=True, color="FFFFFF")
font_bold = Font(name="Segoe UI", size=9.5, bold=True, color="0F172A")
font_reg = Font(name="Segoe UI", size=9, color="1E293B")
font_blue = Font(name="Segoe UI", size=9.5, bold=True, color="0369A1")
font_green = Font(name="Segoe UI", size=9.5, bold=True, color="047857")
font_amber = Font(name="Segoe UI", size=9.5, bold=True, color="B45309")
font_red_bold = Font(name="Segoe UI", size=9.5, bold=True, color="FFFFFF")
font_cift = Font(name="Segoe UI", size=9, bold=True, color="0284C7")

# Fills
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

# Title Rows
ws.merge_cells("A1:R1")
ws["A1"] = "ERGUNBAS GROUP — MAKİNE & HAT VARDİYA MESAİ VE KAPASİTE MATRİSİ"
ws["A1"].font = font_title
ws["A1"].fill = fill_dark
ws["A1"].alignment = align_c
ws.row_dimensions[1].height = 34

ws.merge_cells("A2:R2")
ws["A2"] = "1. Vardiya: 08:30 / 20:30 (12s / 720 dk Net) | 2. Vardiya: 20:30 / 08:30 (12s / 720 dk Net) | 24 Saat Kesintisiz Çalışma (1.440 dk) | Hat 101 - 109 & Hat 201 ve Kalite Kontrol"
ws["A2"].font = font_sub
ws["A2"].fill = fill_dark
ws["A2"].alignment = align_c
ws.row_dimensions[2].height = 20

# Super Headers
ws.merge_cells("A3:G3")
ws["A3"] = "MÜDÜR İSTASYON & MESAİ PLANI (FOTOĞRAF ŞABLONU)"
ws["A3"].font = font_grp
ws["A3"].fill = fill_shift
ws["A3"].alignment = align_c

ws.merge_cells("H3:M3")
ws["H3"] = "HAT BAZLI ÜRETİLEN ÜRÜNLER VE ÖLÇÜLERİ"
ws["H3"].font = font_grp
ws["H3"].fill = fill_prod
ws["H3"].alignment = align_c

ws.merge_cells("N3:R3")
ws["N3"] = "KAPASİTE VE HIZ DETAYLARI (720 dk VARDİYA / 1.440 dk 24s GÜNLÜK NET)"
ws["N3"].font = font_grp
ws["N3"].fill = fill_cap
ws["N3"].alignment = align_c
ws.row_dimensions[3].height = 25

headers = [
    # A - G: Fotoğraftaki ana sütunlar
    ("İSTASYON / HAT", 17, align_l),
    ("MESAİ ARALIĞI", 20, align_c),
    ("MESAİ SÜRESİ (DK)", 15, align_r),
    ("NET ÇALIŞMA (DK)", 15, align_r),
    ("A.V. ÇALIŞAN", 13, align_c),
    ("B.V. ÇALIŞAN", 13, align_c),
    ("GÜNLÜK NET KAPASİTE", 18, align_r),
    
    # H - R: Sağ tarafındaki ürün & kapasite detayları
    ("ÜRETİLEN ÜRÜN", 24, align_l),
    ("KATEGORİ", 16, align_l),
    ("KALIP ÇIKIŞI", 16, align_c),
    ("STANDART BOY", 18, align_c),
    ("KESİT ÖLÇÜSÜ", 18, align_l),
    ("BİRİM AĞIRLIK", 14, align_r),
    ("SAATLİK HIZ", 15, align_r),
    ("12s VARDİYA KAPASİTESİ", 18, align_r),
    ("24s GÜNLÜK KAPASİTE", 18, align_r),
    ("GÜNLÜK TONAJ (KG)", 16, align_r),
    ("KAPI EŞDEĞERİ (24s / GÜN)", 18, align_r)
]

for col_idx, (h_title, col_w, h_align) in enumerate(headers, 1):
    c = ws.cell(row=4, column=col_idx, value=h_title)
    c.font = font_head
    c.fill = fill_dark
    c.alignment = align_c
    c.border = thin_border
    ws.column_dimensions[get_column_letter(col_idx)].width = col_w
ws.row_dimensions[4].height = 28

# Detailed data list:
# NET ÇALIŞMA SÜRESİ = 720 DK (12 SAAT VARDİYA, 24 SAAT = 1.440 DK KESİNTİSİZ)
rows_data = [
    # HAT 101
    ("Hat 101", "08:30 / 20:30 (2V)", 720, 720, 1, 1, "487 Boy / 97 Kapı", True,
     "80X80 PERVAZ", "Pervaz Profili", "Tek Göz (1 Kafa)", "2.20 Metre", "80 x 80 mm", 2.20, "Boy", 20.3, 5.0, "Kapı"),
    ("Hat 101", "20:30 / 08:30", 720, 720, 1, 1, "523 Boy / 105 Kapı", False,
     "35X100 PERVAZ", "Pervaz Profili", "Tek Göz (1 Kafa)", "2.20 Metre", "35 x 100 mm", 2.08, "Boy", 21.8, 5.0, "Kapı"),
    ("Hat 101", "—", 720, 720, 1, 1, "334 Boy / 67 Kapı", False,
     "35X150 PERVAZ", "Pervaz Profili", "Tek Göz (1 Kafa)", "2.20 Metre", "35 x 150 mm", 2.97, "Boy", 13.9, 5.0, "Kapı"),
    ("Hat 101", "—", 720, 720, 1, 1, "360 Boy / 72 Kapı", False,
     "35X80 PERVAZ", "Pervaz Profili", "Tek Göz (1 Kafa)", "2.20 Metre", "35 x 80 mm", 2.94, "Boy", 15.0, 5.0, "Kapı"),

    # HAT 102
    ("Hat 102", "08:30 / 20:30 (2V)", 720, 720, 1, 1, "374 Boy / 150 Kapı", True,
     "100 mm KASA", "Kasa Profili", "Tek Göz (1 Kafa)", "2.15 Metre", "100 x 36 mm", 4.00, "Boy", 15.6, 2.5, "Kapı"),
    ("Hat 102", "20:30 / 08:30", 720, 720, 1, 1, "250 Boy / 100 Kapı", False,
     "140 mm KASA", "Kasa Profili", "Tek Göz (1 Kafa)", "2.15 Metre", "140 x 36 mm", 4.95, "Boy", 10.4, 2.5, "Kapı"),
    ("Hat 102", "—", 720, 720, 1, 1, "797 Boy", False,
     "CAM ÇITASI", "Cam Çıtası", "Tek Göz (1 Kafa)", "2.20 Metre", "15 x 20 mm", 1.50, "Boy", 33.2, 4.0, "Kapı"),
    ("Hat 102", "—", 720, 720, 1, 1, "2.047 Boy", False,
     "U KAPAMA", "U Kapama / Lambiri", "Tek Göz (1 Kafa)", "2.20 Metre", "30 x 15 mm", 0.48, "Boy", 85.3, 0.0, "—"),

    # HAT 103
    ("Hat 103", "08:30 / 20:30 (2V)", 720, 720, 1, 1, "756 Boy / 151 Kapı", True,
     "50X80 PERVAZ", "Pervaz Profili", "Çift Göz (2 Kafa)", "2.20 Metre", "50 x 80 mm", 1.55, "Boy", 31.5, 5.0, "Kapı"),
    ("Hat 103", "20:30 / 08:30", 720, 720, 1, 1, "900 Boy / 257 Kapı", False,
     "SEREN", "Seren Çıtası", "Çift Göz (2 Kafa)", "2.10 Metre", "38 x 45 mm", 3.39, "Boy", 37.5, 3.5, "Kapı"),

    # HAT 104
    ("Hat 104", "08:30 / 20:30 (2V)", 720, 720, 1, 1, "677 Boy / 193 Kapı", True,
     "SEREN", "Seren Çıtası", "Çift Göz (2 Kafa)", "2.10 Metre", "38 x 45 mm", 3.39, "Boy", 28.2, 3.5, "Kapı"),

    # HAT 105
    ("Hat 105", "08:30 / 20:30 (2V)", 720, 720, 1, 1, "641 Boy / 183 Kapı", True,
     "SEREN", "Seren Çıtası", "Çift Göz (2 Kafa)", "2.10 Metre", "38 x 45 mm", 3.39, "Boy", 26.7, 3.5, "Kapı"),
    ("Hat 105", "20:30 / 08:30", 720, 720, 1, 1, "418 Boy / 167 Kapı", False,
     "140 mm KASA", "Kasa Profili", "Tek Göz (1 Kafa)", "2.15 Metre", "140 x 36 mm", 4.95, "Boy", 17.4, 2.5, "Kapı"),

    # HAT 106
    ("Hat 106", "08:30 / 20:30 (2V)", 720, 720, 1, 1, "336 Boy / 134 Kapı", True,
     "140 mm KASA", "Kasa Profili", "Tek Göz (1 Kafa)", "2.15 Metre", "140 x 36 mm", 4.95, "Boy", 14.0, 2.5, "Kapı"),
    ("Hat 106", "20:30 / 08:30", 720, 720, 1, 1, "492 Boy / 197 Kapı", False,
     "200 mm KASA", "Kasa Profili", "Tek Göz (1 Kafa)", "2.15 Metre", "200 x 36 mm", 6.66, "Boy", 20.5, 2.5, "Kapı"),
    ("Hat 106", "—", 720, 720, 1, 1, "324 Boy / 130 Kapı", False,
     "100 mm KASA", "Kasa Profili", "Tek Göz (1 Kafa)", "2.15 Metre", "100 x 36 mm", 4.00, "Boy", 13.5, 2.5, "Kapı"),

    # HAT 107
    ("Hat 107", "08:30 / 20:30 (2V)", 720, 720, 1, 1, "300 Boy / 120 Kapı", True,
     "200 mm KASA", "Kasa Profili", "Tek Göz (1 Kafa)", "2.15 Metre", "200 x 36 mm", 6.66, "Boy", 12.5, 2.5, "Kapı"),
    ("Hat 107", "20:30 / 08:30", 720, 720, 1, 1, "670 Boy / 191 Kapı", False,
     "SEREN(39mm)", "Seren Çıtası", "Tek Göz (Hızlı 1.3 m/dk)", "2.10 Metre", "39 x 45 mm", 3.81, "Boy", 27.9, 3.5, "Kapı"),
    ("Hat 107", "—", 720, 720, 1, 1, "197 Boy", False,
     "KT100-220x24-ANT", "Deck / WPC Zemin", "Tek Göz (1 Kafa)", "3.00 Metre", "220 x 24 mm", 10.33, "Boy", 8.2, 0.0, "—"),
    ("Hat 107", "—", 720, 720, 1, 1, "247 Boy", False,
     "KT102-219x26-ANT", "Deck / WPC Zemin", "Tek Göz (1 Kafa)", "3.00 Metre", "219 x 26 mm", 10.52, "Boy", 10.3, 0.0, "—"),
    ("Hat 107", "—", 720, 720, 1, 1, "494 Boy / 141 Kapı", False,
     "SEREN", "Seren Çıtası", "Tek Göz (1 Kafa)", "2.10 Metre", "38 x 45 mm", 3.39, "Boy", 20.6, 3.5, "Kapı"),

    # HAT 108
    ("Hat 108", "08:30 / 20:30 (2V)", 720, 720, 1, 1, "413 Boy / 165 Kapı", True,
     "100 mm KASA", "Kasa Profili", "Tek Göz (1 Kafa)", "2.15 Metre", "100 x 36 mm", 4.00, "Boy", 17.2, 2.5, "Kapı"),
    ("Hat 108", "20:30 / 08:30", 720, 720, 1, 1, "353 Boy / 141 Kapı", False,
     "140 mm KASA", "Kasa Profili", "Tek Göz (1 Kafa)", "2.15 Metre", "140 x 36 mm", 4.95, "Boy", 14.7, 2.5, "Kapı"),

    # HAT 109
    ("Hat 109", "08:30 / 20:30 (2V)", 720, 720, 1, 1, "1.169 Boy / 234 Kapı", True,
     "50X80 PERVAZ", "Pervaz Profili", "Çift Göz (2 Kafa)", "2.20 Metre", "50 x 80 mm", 1.55, "Boy", 48.7, 5.0, "Kapı"),
    ("Hat 109", "20:30 / 08:30", 720, 720, 1, 1, "1.380 Boy / 276 Kapı", False,
     "80X80 PERVAZ", "Pervaz Profili", "Çift Göz (2 Kafa)", "2.20 Metre", "80 x 80 mm", 2.20, "Boy", 57.5, 5.0, "Kapı"),

    # HAT 201 (LEVHA MAKİNESİ) - 720 dk (12s) Net Vardiya, 1.440 dk (24s) Net Gün: ~1.000 Plaka / 500 Kapı!
    ("Hat 201 (Levha)", "08:30 / 20:30 (2V)", 720, 720, 2, 2, "1.025 Plaka / 513 Kapı", True,
     "DIŞBUDAK Levha", "WPC Levha Plaka", "Tek Göz (Levha Kalıbı)", "207 cm (2.07 m)", "En: 94 cm (2 mm)", 4.77, "Plaka", 42.7, 2.0, "Kapı"),
    ("Hat 201 (Levha)", "20:30 / 08:30", 720, 720, 2, 2, "926 Plaka / 463 Kapı", False,
     "TEAK Levha", "WPC Levha Plaka", "Tek Göz (Levha Kalıbı)", "207 cm (2.07 m)", "En: 94 cm (2 mm)", 4.25, "Plaka", 38.6, 2.0, "Kapı"),
    ("Hat 201 (Levha)", "—", 720, 720, 2, 2, "962 Plaka / 481 Kapı", False,
     "MIX TEAK Levha", "WPC Levha Plaka", "Tek Göz (Levha Kalıbı)", "207 cm (2.07 m)", "En: 94 cm (2 mm)", 4.25, "Plaka", 40.1, 2.0, "Kapı"),
    ("Hat 201 (Levha)", "—", 720, 720, 2, 2, "974 Plaka / 487 Kapı", False,
     "BEYAZ TEAK Levha", "WPC Levha Plaka", "Tek Göz (Levha Kalıbı)", "207 cm (2.07 m)", "En: 94 cm (2 mm)", 4.41, "Plaka", 40.6, 2.0, "Kapı"),
    ("Hat 201 (Levha)", "—", 720, 720, 2, 2, "1.058 Plaka / 529 Kapı", False,
     "MIX BEYAZ TEAK Levha", "WPC Levha Plaka", "Tek Göz (Levha Kalıbı)", "207 cm (2.07 m)", "En: 94 cm (2 mm)", 4.25, "Plaka", 44.1, 2.0, "Kapı"),
    ("Hat 201 (Levha)", "—", 720, 720, 2, 2, "960 Plaka / 480 Kapı", False,
     "AY GRİ Levha", "WPC Levha Plaka", "Tek Göz (Levha Kalıbı)", "207 cm (2.07 m)", "En: 94 cm (2 mm)", 4.64, "Plaka", 40.0, 2.0, "Kapı"),
    ("Hat 201 (Levha)", "—", 720, 720, 2, 2, "922 Plaka / 461 Kapı", False,
     "SONOMO Levha", "WPC Levha Plaka", "Tek Göz (Levha Kalıbı)", "207 cm (2.07 m)", "En: 94 cm (2 mm)", 4.33, "Plaka", 38.4, 2.0, "Kapı"),
    ("Hat 201 (Levha)", "—", 720, 720, 2, 2, "994 Plaka / 497 Kapı", False,
     "COCO Levha", "WPC Levha Plaka", "Tek Göz (Levha Kalıbı)", "207 cm (2.07 m)", "En: 94 cm (2 mm)", 4.79, "Plaka", 41.4, 2.0, "Kapı"),
    ("Hat 201 (Levha)", "—", 720, 720, 2, 2, "960 Plaka / 480 Kapı", False,
     "ANTRASİT Levha", "WPC Levha Plaka", "Tek Göz (Levha Kalıbı)", "207 cm (2.07 m)", "En: 94 cm (2 mm)", 4.77, "Plaka", 40.0, 2.0, "Kapı"),
    ("Hat 201 (Levha)", "—", 720, 720, 2, 2, "816 Plaka / 408 Kapı", False,
     "VIZON Levha", "WPC Levha Plaka", "Tek Göz (Levha Kalıbı)", "207 cm (2.07 m)", "En: 94 cm (2 mm)", 4.77, "Plaka", 34.0, 2.0, "Kapı"),
    ("Hat 201 (Levha)", "—", 720, 720, 2, 2, "691 Plaka / 346 Kapı", False,
     "KOYU CEVİZ Levha", "WPC Levha Plaka", "Tek Göz (Levha Kalıbı)", "207 cm (2.07 m)", "En: 94 cm (2 mm)", 4.79, "Plaka", 28.8, 2.0, "Kapı"),

    # KALİTE KONTROL
    ("KALİTE KONTROL", "08:30 / 20:30 (2V)", 720, 720, 3, 3, "4.200 Adet/Gün", True,
     "TÜM PROFİL & LEVHALAR", "Kalite Güvence", "Ölçü / Yüzey Kontrolü", "Tüm Boylar", "Tüm Kesitler", 0.0, "Adet", 175.0, 0.0, "4.200 Adet/Gün")
]

cur_row = 5
for r in rows_data:
    (hat_name, mesai_araligi, mesai_dk, net_dk, av_calisan, bv_calisan, daily_line_cap_str, is_line_primary,
     prod_name, cat, goz, boy, kesit, unit_kg, unit_lbl, spd_h, kapi_factor, kapi_unit) = r

    # Net 12 saatlik vardiya (720 dk) ve Net 24 saatlik gün (1.440 dk)
    cap_12s = round(spd_h * 12.0)
    cap_24s = round(spd_h * 24.0)
    daily_tonaj = round(cap_24s * unit_kg, 1) if unit_kg > 0 else 0.0
    
    kapi_24s = round(cap_24s / kapi_factor) if kapi_factor > 0 else 0
    kapi_str = f"{kapi_24s} Kapı/Gün" if kapi_factor > 0 else (kapi_unit if kapi_unit != "—" else "—")
    
    row_fill = fill_zebra if cur_row % 2 == 0 else PatternFill(fill_type=None)
    goz_font = font_cift if "Çift" in goz else font_reg
    
    vals = [
        (hat_name if is_line_primary else "↳ " + hat_name, font_bold if is_line_primary else font_blue, align_l, "@"),
        (mesai_araligi, font_reg, align_c, "@"),
        (mesai_dk, font_reg, align_r, "#,##0"),
        (net_dk, font_bold, align_r, "#,##0"),
        (av_calisan if is_line_primary else "—", font_bold if is_line_primary else font_reg, align_c, "@" if not is_line_primary else "#,##0"),
        (bv_calisan if is_line_primary else "—", font_bold if is_line_primary else font_reg, align_c, "@" if not is_line_primary else "#,##0"),
        (daily_line_cap_str if is_line_primary else "—", font_red_bold if (is_line_primary and "Kapı" in daily_line_cap_str) else font_bold, align_r if not is_line_primary else align_c, "@"),
        (prod_name, font_bold, align_l, "@"),
        (cat, font_reg, align_l, "@"),
        (goz, goz_font, align_c, "@"),
        (boy, font_green, align_c, "@"),
        (kesit, font_reg, align_l, "@"),
        (unit_kg if unit_kg > 0 else "—", font_reg, align_r, '0.00 "kg"' if unit_kg > 0 else "@"),
        (round(spd_h, 1), font_green, align_r, f'#,##0.0 " {unit_lbl}/sa"'),
        (cap_12s, font_reg, align_r, f'#,##0 " {unit_lbl}"'),
        (cap_24s, font_bold, align_r, f'#,##0 " {unit_lbl}"'),
        (daily_tonaj if daily_tonaj > 0 else "—", font_reg, align_r, '#,##0.0 " kg"' if daily_tonaj > 0 else "@"),
        (kapi_str, font_green if kapi_24s > 0 else font_reg, align_r, "@")
    ]
    
    for col_idx, (val, c_font, c_align, num_fmt) in enumerate(vals, 1):
        cell = ws.cell(row=cur_row, column=col_idx, value=val)
        cell.font = c_font
        cell.alignment = c_align
        cell.border = thin_border
        if col_idx == 7 and is_line_primary and "Kapı" in daily_line_cap_str:
            cell.fill = fill_red
        elif col_idx == 11:
            cell.fill = fill_green_box
        elif row_fill.fill_type:
            cell.fill = row_fill
        if num_fmt and num_fmt != "@":
            cell.number_format = num_fmt

    ws.row_dimensions[cur_row].height = 20
    cur_row += 1

# Total row
last_data_row = cur_row - 1
ws.cell(row=cur_row, column=1, value="GENEL TOPLAM (FİLTREYE DUYARLI)").font = font_bold
ws.cell(row=cur_row, column=1).alignment = align_l

# Subtotals
ws.cell(row=cur_row, column=15, value=f"=SUBTOTAL(109, O5:O{last_data_row})").font = font_bold
ws.cell(row=cur_row, column=15).number_format = '#,##0 " Adet/Boy"'

ws.cell(row=cur_row, column=16, value=f"=SUBTOTAL(109, P5:P{last_data_row})").font = font_bold
ws.cell(row=cur_row, column=16).number_format = '#,##0 " Adet/Boy"'

for c in range(1, 19):
    cell = ws.cell(row=cur_row, column=c)
    cell.fill = fill_total
    cell.border = double_bottom
ws.row_dimensions[cur_row].height = 24

# AutoFilter and Freeze Panes
ws.auto_filter.ref = f"A4:R{last_data_row}"
ws.freeze_panes = "B5"

dest_desktop = r"C:\Users\llboz\OneDrive\Desktop\Ergunbas_Ekstruzyon_ve_Kalite_Kontrol_Vardiya_Kapasite_Matrisi.xlsx"
dest_alt = r"C:\Users\llboz\OneDrive\Desktop\Ergunbas_Ekstruzyon_ve_Kalite_Kontrol_Vardiya_Kapasite_Matrisi (Guncel_720dk).xlsx"
dest_static = r"C:\Users\llboz\OneDrive\Desktop\ergunbas-uretim-sistemi\static\Ergunbas_Ekstruzyon_ve_Kalite_Kontrol_Vardiya_Kapasite_Matrisi.xlsx"

try:
    wb.save(dest_desktop)
    print(f"Excel saved to Desktop: {dest_desktop}")
except Exception as e:
    print(f"Desktop ana dosya açık olabilir ({e}), alternatif isimle kaydediliyor...")

try:
    wb.save(dest_alt)
    print(f"Excel alternatif kopya kaydedildi: {dest_alt}")
except Exception as e:
    print(f"Alt save error: {e}")

try:
    wb.save(dest_static)
    print(f"Excel saved to static: {dest_static}")
except:
    pass
