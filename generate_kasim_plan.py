# -*- coding: utf-8 -*-
"""
Kasım Siparişleri – Termin Tarihine Göre Üretim Planlama Matrisi
Kaynak: sipariş_içeriği_düzenli (2).xlsx
"""
import openpyxl, datetime, math
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────
# 1. KAYNAK VERİLER (tüm sayfa taraması)
# ─────────────────────────────────────────────────────────────────
wb_src = openpyxl.load_workbook(r'C:\Users\llboz\OneDrive\Desktop\siparis_icerik_temp.xlsx', data_only=True)

# İndeks bazlı erişim ile encoding hatalarını önleme
ws_kasim = wb_src.sheetnames[0]
ws_ozet = wb_src.sheetnames[2]
ws_kasa = wb_src['kasa']
ws_pervaz = wb_src['pervaz']

# Sipariş Özeti
ws_ozet_sheet = wb_src[ws_ozet]
siparis_ozet = {}   # no: {kanat, kasa, pervaz}
for r in range(4, ws_ozet_sheet.max_row+1):
    no   = ws_ozet_sheet.cell(r,1).value
    kn   = ws_ozet_sheet.cell(r,2).value or 0
    kas  = ws_ozet_sheet.cell(r,3).value or 0
    per  = ws_ozet_sheet.cell(r,4).value or 0
    if no and str(no).isdigit():
        siparis_ozet[int(no)] = {'kanat': kn, 'kasa': kas, 'pervaz': per}

# Kasa sayfasından termin tarihlerini çek (sütun A=sip_no, D=termin, C=takım)
kasa_termin = {}   # sip_no -> termin
for r in range(2, ws_kasa.max_row+1):
    sip = ws_kasa.cell(r, 1).value
    ter = ws_kasa.cell(r, 4).value
    if sip and isinstance(sip, int) and isinstance(ter, datetime.datetime):
        kasa_termin[sip] = ter

# Pervaz sayfasından termin tarihlerini çek
pervaz_termin = {}
for r in range(2, ws_pervaz.max_row+1):
    sip = ws_pervaz.cell(r, 1).value
    ter = ws_pervaz.cell(r, 4).value
    if sip and isinstance(sip, int) and isinstance(ter, datetime.datetime):
        pervaz_termin[sip] = ter

# Sip no'yu ve termin'i kasa/pervaz'dan birleştir
all_termins = {}
for sno in siparis_ozet:
    t = kasa_termin.get(sno) or pervaz_termin.get(sno)
    all_termins[sno] = t

# ─────────────────────────────────────────────────────────────────
# 2. HAT KAPASİTELERİ (24s fiili ortalama)
# ─────────────────────────────────────────────────────────────────
LINE_CAPS = {
    # Hat: (ürün_grubu, ürün_adı, birim, günlük_çıktı_24s, kapi_faktörü)
    'Hat 102': ('Kasa',   '100 mm Kasa',    'Boy',   376, 2.5),
    'Hat 106': ('Kasa',   '140 mm Kasa',    'Boy',   336, 2.5),
    'Hat 108': ('Kasa',   '140 mm Kasa',    'Boy',   352, 2.5),
    'Hat 103': ('Pervaz', '50×80 Pervaz',   'Boy',   756, 5.0),
    'Hat 109': ('Pervaz', '50×80 Pervaz',   'Boy',  1204, 5.0),
    'Hat 104': ('Seren',  'Seren 38×45 mm', 'Boy',   676, 3.5),
    'Hat 105': ('Seren',  'Seren 38×45 mm', 'Boy',   640, 3.5),
    'Hat 201': ('Levha',  'WPC Levha',      'Plaka', 1000, 2.0),
}

# ─────────────────────────────────────────────────────────────────
# 3. PLAN OLUŞTUR: Her sipariş için gerekli gün / hat ataması
# ─────────────────────────────────────────────────────────────────
def kasa_tipi_sec(kasa_adet):
    """100 ve 140 mm kasa karışımı olduğunu varsay: 106+108+102 kombine."""
    # 3 hat kombine günlük = 336+352+376 = 1064 boy | faktör 2.5
    return 1064

def pervaz_gunluk():
    """Hat 109 + Hat 103 kombine."""
    return 1204 + 756  # 1960 boy/gun

# Sipariş bazlı plan
plan_rows = []
for sno, data in sorted(siparis_ozet.items()):
    kanat  = data['kanat']
    kasa   = data['kasa']
    pervaz = data['pervaz']
    termin = all_termins.get(sno)

    # Kasa: toplam boy = kasa_takım * 5 boy/kapı (kullanıcı reçetesi)
    kasa_boy   = kasa   * 5.0   if kasa   else 0
    pervaz_boy = pervaz * 5.0   if pervaz else 0

    # Kasa günleri: Hat 106 (140mm) + Hat 108 (140mm) + Hat 102 (100mm) → kombine 1064 boy/gün
    kasa_gun   = math.ceil(kasa_boy   / 1064) if kasa_boy   > 0 else 0
    pervaz_gun = math.ceil(pervaz_boy / 1960) if pervaz_boy > 0 else 0

    # Kasa hat ataması
    kasa_hat = 'Hat 106 + 108 + 102' if kasa > 0 else '-'
    pervaz_hat = 'Hat 109 + 103' if pervaz > 0 else '-'

    # Hangi hat daha verimli?
    # 140mm Kasa → Hat 106 önce (Ana Hat); 100mm Kasa → Hat 108 önce (daha hızlı)
    # Pervaz → Hat 109 her zaman önce (2x hız farkı var)
    kasa_verimli   = 'Hat 108 (140mm) > Hat 106 > Hat 102' if kasa > 200 else 'Hat 102 (100mm) veya Hat 106'
    pervaz_verimli = 'Hat 109 (50x80) – 2x daha hızlı'

    plan_rows.append({
        'sip_no': sno,
        'termin': termin,
        'kanat':  kanat,
        'kasa':   kasa,
        'pervaz': pervaz,
        'kasa_boy':    kasa_boy,
        'pervaz_boy':  pervaz_boy,
        'kasa_gun':    kasa_gun,
        'pervaz_gun':  pervaz_gun,
        'kasa_hat':    kasa_hat,
        'pervaz_hat':  pervaz_hat,
        'kasa_verimli':    kasa_verimli,
        'pervaz_verimli':  pervaz_verimli,
    })

# Terminli sıralama (önce yakın termin, terminsi en sona)
plan_rows.sort(key=lambda x: (x['termin'] is None, x['termin'] or datetime.datetime(2099,12,31)))

# ─────────────────────────────────────────────────────────────────
# 4. EXCEL YAZAR
# ─────────────────────────────────────────────────────────────────
def fill(h):    return PatternFill('solid', fgColor=h)
def font(bold=False, size=10, color='000000', italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic, name='Calibri')
def align(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def thin_border():
    s = Side(style='thin', color='AAAAAA')
    return Border(left=s, right=s, top=s, bottom=s)

C_DARK   = '1F3864'
C_MED    = '2E75B6'
C_KASA   = 'DDEEFF'
C_PERVAZ = 'FFF2CC'
C_RED    = 'FFDCE0'
C_GREEN  = 'C6EFCE'
C_HEADER = 'F2F2F2'
C_URGENT = 'FF0000'

wb = openpyxl.Workbook()

# ════════════════════════════════════════════════════════
# SAYFA 1 – SİPARİŞ BAZLI ÜRETİM PLAN (TERMİN SIRALI)
# ════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = 'Termin Bazli Uretim Plani'
ws1.sheet_view.showGridLines = False

# Başlık
ws1.merge_cells('A1:O1')
ws1['A1'] = 'ERGÜNBAŞ – KASIM 2026 SİPARİŞ TERMİN SIRALI ÜRETİM PLANLAMA MATRİSİ'
ws1['A1'].font = Font(bold=True, size=15, color='FFFFFF', name='Calibri')
ws1['A1'].fill = fill(C_DARK)
ws1['A1'].alignment = align()
ws1.row_dimensions[1].height = 32

ws1.merge_cells('A2:O2')
ws1['A2'] = ('Reçete: 1 Kapı = 3.5 Boy Seren | 5 Boy Kasa | 5 Boy Pervaz  |  '
             'Hat Kapasitesi: Hat 106=336 | Hat 108=352 | Hat 102=376 (boy/24s)  |  '
             'Pervaz: Hat 109=1204 | Hat 103=756 (boy/24s)')
ws1['A2'].font = Font(italic=True, size=9, color='FFFFFF', name='Calibri')
ws1['A2'].fill = fill(C_MED)
ws1['A2'].alignment = align()
ws1.row_dimensions[2].height = 16

# Sütun başlıkları
headers = [
    'Sıra', 'Sipariş No', 'Termin Tarihi', 'Aciliyet',
    'Kanat\n(Adet)', 'Kasa\n(Takım)', 'Pervaz\n(Takım)',
    'Kasa\nGerekli Boy', 'Pervaz\nGerekli Boy',
    'Kasa Üretim\nGünü (3 Hat)',
    'Pervaz Üretim\nGünü (2 Hat)',
    'Kasa Atanan\nHatlar',
    'Pervaz Atanan\nHatlar',
    'En Verimli\nKasa Hattı',
    'Notlar'
]
col_widths = [5, 12, 15, 10, 8, 8, 8, 13, 13, 13, 13, 22, 18, 30, 28]

for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
    c = ws1.cell(3, ci, h)
    c.font = Font(bold=True, size=9, color='FFFFFF', name='Calibri')
    c.fill = fill(C_DARK)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = thin_border()
    ws1.column_dimensions[get_column_letter(ci)].width = w
ws1.row_dimensions[3].height = 42

today = datetime.datetime(2026, 10, 24)
ri = 4
for idx, row in enumerate(plan_rows, 1):
    termin_dt = row['termin']
    if termin_dt:
        days_left = (termin_dt - today).days
        if days_left <= 5:
            aciliyet_str = f'⚠ ACİL ({days_left} gün)'
            acil_fill = fill('FFDCE0')
            acil_font_c = 'CC0000'
        elif days_left <= 10:
            aciliyet_str = f'🔶 ÖNCELİKLİ ({days_left} gün)'
            acil_fill = fill('FFEB9C')
            acil_font_c = '9C5700'
        else:
            aciliyet_str = f'🟢 Normal ({days_left} gün)'
            acil_fill = fill(C_GREEN)
            acil_font_c = '375623'
        termin_str = termin_dt.strftime('%d.%m.%Y')
    else:
        aciliyet_str = '– Termin Yok'
        acil_fill = fill('E8E8E8')
        acil_font_c = '666666'
        termin_str = '–'
        days_left = 999

    vals = [
        idx,
        row['sip_no'],
        termin_str,
        aciliyet_str,
        row['kanat'],
        row['kasa'],
        row['pervaz'],
        int(row['kasa_boy'])   if row['kasa_boy']   else '–',
        int(row['pervaz_boy']) if row['pervaz_boy'] else '–',
        row['kasa_gun']    if row['kasa_gun']   else '–',
        row['pervaz_gun']  if row['pervaz_gun'] else '–',
        row['kasa_hat'],
        row['pervaz_hat'],
        row['kasa_verimli'],
        ''
    ]

    zebra = fill('EFF3FA') if ri % 2 == 0 else fill('FFFFFF')
    for ci, val in enumerate(vals, 1):
        c = ws1.cell(ri, ci, val)
        c.border = thin_border()
        c.alignment = Alignment(
            horizontal='center' if ci not in (3,4,12,13,14,15) else 'left',
            vertical='center',
            wrap_text=(ci in (4,12,13,14,15))
        )
        if ci == 4:
            c.fill = acil_fill
            c.font = Font(bold=True, size=9, color=acil_font_c, name='Calibri')
        elif ci in (8,9):
            c.fill = fill(C_KASA) if ci==8 else fill(C_PERVAZ)
            c.font = Font(bold=True, size=10, name='Calibri')
        elif ci in (10,11):
            c.fill = fill(C_GREEN) if isinstance(val, int) and val <= 3 else (
                     fill('FFEB9C') if isinstance(val, int) and val <= 7 else
                     fill(C_RED)    if isinstance(val, int)              else fill('E8E8E8'))
            c.font = Font(bold=True, size=10, name='Calibri')
        elif ci in (12,13):
            c.fill = fill(C_KASA) if ci==12 else fill(C_PERVAZ)
            c.font = Font(size=9, name='Calibri')
        elif ci == 14:
            c.fill = fill('E2EFDA')
            c.font = Font(size=9, italic=True, name='Calibri')
        else:
            c.fill = zebra
            c.font = Font(bold=(ci in (2,)), size=10, name='Calibri')
    ws1.row_dimensions[ri].height = 22
    ri += 1

# Toplam satırı
ws1.merge_cells(f'A{ri}:C{ri}')
ws1.cell(ri, 1, 'KASIM TOPLAMI').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
ws1.cell(ri, 1).fill = fill(C_MED)
ws1.cell(ri, 1).alignment = align()

tot_kasa   = sum(r['kasa']   or 0 for r in plan_rows)
tot_pervaz = sum(r['pervaz'] or 0 for r in plan_rows)
tot_kanat  = sum(r['kanat']  or 0 for r in plan_rows)

for ci, val in [(5, tot_kanat), (6, tot_kasa), (7, tot_pervaz),
                (8, int(tot_kasa*5)), (9, int(tot_pervaz*5))]:
    c = ws1.cell(ri, ci, val)
    c.font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    c.fill = fill(C_DARK)
    c.alignment = align()
    c.border = thin_border()
    c.number_format = '#,##0'
for ci in [1,2,3,4,10,11,12,13,14,15]:
    ws1.cell(ri, ci).fill = fill(C_DARK)
    ws1.cell(ri, ci).border = thin_border()

ws1.row_dimensions[ri].height = 24
ws1.auto_filter.ref = f'A3:O{ri-1}'
ws1.freeze_panes = 'B4'

# ════════════════════════════════════════════════════════
# SAYFA 2 – HAT VERİMLİLİK KARŞILAŞTIRMASI
# ════════════════════════════════════════════════════════
ws2 = wb.create_sheet('Hat Verimlilik Analizi')
ws2.sheet_view.showGridLines = False

ws2.merge_cells('A1:H1')
ws2['A1'] = 'HAT VERİMLİLİK ANALİZİ – HANGİ HAT HANGİ ÜRÜNDE DAHA VERİMLİ?'
ws2['A1'].font = Font(bold=True, size=13, color='FFFFFF', name='Calibri')
ws2['A1'].fill = fill(C_DARK)
ws2['A1'].alignment = align()
ws2.row_dimensions[1].height = 28

ws2.merge_cells('A2:H2')
ws2['A2'] = 'Fiili Saha Ortalamaları (Ağustos–Ekim 2026 Dönemi – 600+ Vardiya Kaydı)'
ws2['A2'].font = Font(italic=True, size=9, color='FFFFFF', name='Calibri')
ws2['A2'].fill = fill(C_MED)
ws2['A2'].alignment = align()
ws2.row_dimensions[2].height = 14

v2_headers = [
    'Hat', 'Ürün Grubu', 'Ürün Adı', 'Çekici Hızı\n(m/dak)', 'Saatlik Çıktı\n(boy/saat)',
    '12s Vardiya\nÇıktısı', '24s Günlük\nÇıktısı', 'Verimliliği Sıralaması'
]
v2_widths = [10, 12, 28, 13, 15, 14, 14, 26]
for ci, (h, w) in enumerate(zip(v2_headers, v2_widths), 1):
    c = ws2.cell(3, ci, h)
    c.font = Font(bold=True, size=9, color='FFFFFF', name='Calibri')
    c.fill = fill(C_DARK)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = thin_border()
    ws2.column_dimensions[get_column_letter(ci)].width = w
ws2.row_dimensions[3].height = 36

verimlilik_data = [
    # Hat, Grup, Ürün, Hız, b/s, 12s, 24s, sıralama, renk
    ('Hat 109', 'Pervaz', '50×80 Pervaz',   1.34, 50.2,  602, 1204, '🥇 1. – EN VERİMLİ (Pervaz)', 'D4EDDA', False),
    ('Hat 103', 'Pervaz', '50×80 Pervaz',   0.78, 31.5,  378,  756, '🥈 2. – İkinci Pervaz Hattı', 'E2EFDA', False),
    ('Hat 104', 'Seren',  'Seren 38×45 mm', 0.65, 28.2,  338,  676, '🥇 1. – EN VERİMLİ (Seren)',  'D4EDDA', False),
    ('Hat 105', 'Seren',  'Seren 38×45 mm', 0.70, 26.7,  320,  640, '🥈 2. – İkinci Seren Hattı',  'E2EFDA', False),
    ('Hat 108', 'Kasa',   '140 mm Kasa',    0.70, 14.7,  176,  352, '🥇 1. – KASA için En Hızlı',  'D4EDDA', True),
    ('Hat 102', 'Kasa',   '100 mm Kasa',    0.77, 15.7,  188,  376, '🥇 1. – 100mm Kasa 1.Hattı',  'D4EDDA', False),
    ('Hat 106', 'Kasa',   '140 mm Kasa',    0.71, 14.0,  168,  336, '🥈 2. – 140mm Kasa 2.Hattı',  'E2EFDA', False),
    ('Hat 201', 'Levha',  'WPC Kapı Levhası', 0.0, 41.7, 500, 1000, '🥇 1. – Tek Levha Hattı',     'D4EDDA', False),
]

ri2 = 4
group_colors2 = {'Kasa': C_KASA, 'Pervaz': C_PERVAZ, 'Seren': 'E2EFDA', 'Levha': 'FCE4D6'}
for hat, grup, urun, hiz, bs, v12, v24, sirala, acc_c, is_star in verimlilik_data:
    gc = group_colors2.get(grup, 'FFFFFF')
    vals = [hat, grup, urun, f'{hiz:.2f} m/dak' if hiz > 0 else '–', f'{bs:.1f}', v12, v24, sirala]
    for ci, val in enumerate(vals, 1):
        c = ws2.cell(ri2, ci, val)
        c.fill = fill(acc_c if ci == 8 else gc)
        c.border = thin_border()
        c.alignment = Alignment(horizontal='left' if ci in (3,8) else 'center', vertical='center', wrap_text=(ci==8))
        if ci == 1:
            c.font = Font(bold=True, size=10, name='Calibri')
        elif ci in (4,5,6,7):
            c.font = Font(bold=True, size=11, name='Calibri')
        else:
            c.font = Font(size=10, name='Calibri')
    ws2.row_dimensions[ri2].height = 22
    ri2 += 1

# Kutu notu: Kasa hakkında özel yorum
ri2 += 1
ws2.merge_cells(f'A{ri2}:H{ri2}')
ws2.cell(ri2, 1, (
    '⚠  KASA DARBOĞAZI: Toplam Kasım Kasa Siparişi = 6.678 Takım = 33.390 Boy (5 boy/kapı). '
    'Hat 106+108+102 kombine günlük 1.064 boy üretiyor → '
    '33.390 / 1.064 = 31.4 iş günü gerekir. '
    'Hat 108 (140mm) ve Hat 102 (100mm) en verimli hatlar; Hat 106 destek.'
))
ws2.cell(ri2, 1).font = Font(bold=True, size=10, color='8B0000', name='Calibri')
ws2.cell(ri2, 1).fill = fill('FFDCE0')
ws2.cell(ri2, 1).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
ws2.row_dimensions[ri2].height = 40
ri2 += 1

ws2.merge_cells(f'A{ri2}:H{ri2}')
ws2.cell(ri2, 1, (
    '✅  PERVAZ: Toplam Kasım Pervaz = 7.554 Takım = 37.770 Boy. '
    'Hat 109+103 kombine 1.960 boy/gün → 37.770 / 1.960 = 19.3 iş günü. '
    'Hat 109 (50x80) → 50.2 boy/saat ile FABRİKANIN EN VERİMLİ HATTI!'
))
ws2.cell(ri2, 1).font = Font(bold=True, size=10, color='375623', name='Calibri')
ws2.cell(ri2, 1).fill = fill(C_GREEN)
ws2.cell(ri2, 1).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
ws2.row_dimensions[ri2].height = 40

ws2.freeze_panes = 'B4'

# ─── KAYDET ──────────────────────────────────────────────────────
out_path = r'C:\Users\llboz\OneDrive\Desktop\Ergunbas_Kasim_Uretim_Plani_Termin_Sirali.xlsx'
wb.save(out_path)
print(f'Saved: {out_path}')
