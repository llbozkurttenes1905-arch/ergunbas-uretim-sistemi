# -*- coding: utf-8 -*-
"""
Kasa ve Pervaz Sekmelerini İnceleyerek Oluşturulan Detaylı Üretim ve Termin Planı
Kaynak: sipariş_içeriği_düzenli (2).xlsx
"""
import openpyxl, datetime, math
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def build_kasa_pervaz_plan():
    src_path = r'C:\Users\llboz\OneDrive\Desktop\siparis_icerik_temp.xlsx'
    wb_src = openpyxl.load_workbook(src_path, data_only=True)

    # ─────────────────────────────────────────────────────────────
    # 1. PARSE KASA
    # ─────────────────────────────────────────────────────────────
    ws_kasa = wb_src['kasa']
    kasa_rows = []
    curr_sip = None
    curr_term = None

    for r in range(2, 95):
        sno = ws_kasa.cell(r, 1).value
        desc = ws_kasa.cell(r, 2).value
        qty = ws_kasa.cell(r, 3).value or 0
        term = ws_kasa.cell(r, 4).value
        
        if sno is not None and str(sno).strip():
            curr_sip = str(sno).strip()
        if term is not None:
            if isinstance(term, datetime.datetime):
                curr_term = term
            elif str(term).strip():
                try:
                    curr_term = datetime.datetime.strptime(str(term).strip(), '%Y-%m-%d %H:%M:%S')
                except:
                    curr_term = term

        if desc and qty:
            d_upper = str(desc).upper()
            size = '140 MM'
            unit_kg = 12.5
            line_pref = 'Hat 106 & Hat 102'
            line_rate_day = 270 # tk/gun cift hat
            if '100X' in d_upper or ' 100 ' in d_upper:
                size = '100 MM'
                unit_kg = 10.0
                line_pref = 'Hat 108'
                line_rate_day = 148 # tk/gun
            elif '200X' in d_upper or ' 200 ' in d_upper:
                size = '200 MM'
                unit_kg = 16.5
                line_pref = 'Hat 107 (200mm)'
                line_rate_day = 160 # tk/gun
            
            # color detection
            color = 'D.BEYAZ'
            for c_name in ['D.BEYAZ', 'DISBUDAK BEYAZ', 'BEYAZ TEAK', 'B.TEAK', 'VİZON', 'VIZON', 'ANTRASİT', 'ANTRASIT', 'SOMONO', 'TEAK', 'K.CEVİZ', 'K.CEVIZ']:
                if c_name in d_upper:
                    color = c_name.replace('VİZON', 'VIZON').replace('ANTRASİT', 'ANTRASIT').replace('K.CEVİZ', 'K.CEVIZ').replace('B.TEAK', 'BEYAZ TEAK')
                    break

            tot_kg = qty * unit_kg
            profil_boy = round(qty * 2.5, 1) # 1 takim kasa = 2.5 boy profil (2 dikme + 0.5 ust baslik)
            est_days = round(qty / line_rate_day, 2) if line_rate_day > 0 else 0

            kasa_rows.append({
                'row_id': r,
                'sip': curr_sip,
                'desc': str(desc).strip(),
                'size': size,
                'color': color,
                'qty': int(qty),
                'profil_boy': profil_boy,
                'unit_kg': unit_kg,
                'tot_kg': tot_kg,
                'term': curr_term,
                'line': line_pref,
                'est_days': est_days
            })

    # ─────────────────────────────────────────────────────────────
    # 2. PARSE PERVAZ
    # ─────────────────────────────────────────────────────────────
    ws_pervaz = wb_src['pervaz']
    pervaz_rows = []
    curr_sip = None
    curr_term = None

    for r in range(2, 57):
        sno = ws_pervaz.cell(r, 1).value
        desc = ws_pervaz.cell(r, 2).value
        qty = ws_pervaz.cell(r, 3).value or 0
        term = ws_pervaz.cell(r, 4).value
        
        if sno is not None and str(sno).strip():
            curr_sip = str(sno).strip()
        if term is not None:
            if isinstance(term, datetime.datetime):
                curr_term = term
            elif str(term).strip():
                try:
                    curr_term = datetime.datetime.strptime(str(term).strip(), '%Y-%m-%d %H:%M:%S')
                except:
                    curr_term = term

        if desc and qty:
            d_upper = str(desc).upper()
            ptype = '50X80'
            unit_kg = 1.55
            line_pref = 'Hat 109 & Hat 103'
            line_rate_day = 392 # tk/gun cift hat (1960 boy / 5)
            if '80X80' in d_upper:
                ptype = '80X80'
                unit_kg = 2.22
                line_pref = 'Hat 101 / 109'
                line_rate_day = 90 # tk/gun
            
            color = 'D.BEYAZ'
            for c_name in ['D.BEYAZ', 'DISBUDAK BEYAZ', 'BEYAZ TEAK', 'B.TEAK', 'VİZON', 'VIZON', 'ANTRASİT', 'ANTRASIT', 'SOMONO', 'TEAK', 'K.CEVİZ', 'K.CEVIZ']:
                if c_name in d_upper:
                    color = c_name.replace('VİZON', 'VIZON').replace('ANTRASİT', 'ANTRASIT').replace('K.CEVİZ', 'K.CEVIZ').replace('B.TEAK', 'BEYAZ TEAK')
                    break

            tot_kg = qty * unit_kg
            profil_boy = qty * 5.0
            est_days = round(qty / line_rate_day, 2) if line_rate_day > 0 else 0

            pervaz_rows.append({
                'row_id': r,
                'sip': curr_sip,
                'desc': str(desc).strip(),
                'ptype': ptype,
                'color': color,
                'qty': int(qty),
                'profil_boy': int(profil_boy),
                'unit_kg': unit_kg,
                'tot_kg': tot_kg,
                'term': curr_term,
                'line': line_pref,
                'est_days': est_days
            })

    # ─────────────────────────────────────────────────────────────
    # EXCEL FORMATTING HELPERS
    # ─────────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()

    C_DARK    = '1F3864' # koyu lacivert
    C_MED     = '2E75B6' # orta mavi
    C_KASA    = 'DDEEFF' # acik mavi
    C_PERVAZ  = 'FFF2CC' # acik sari
    C_RED     = 'FFDCE0' # acil kirmizi
    C_GREEN   = 'C6EFCE' # normal yesil
    C_WARN    = 'FFEB9C' # sari uyari
    C_HEADER  = 'F2F2F2'
    C_ZEBRA   = 'F9FBFD'

    def fill(h): return PatternFill('solid', fgColor=h)
    def font(bold=False, size=10, color='000000', italic=False):
        return Font(bold=bold, size=size, color=color, italic=italic, name='Calibri')
    def align(h='center', v='center', wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    def thin_border():
        s = Side(style='thin', color='CCCCCC')
        return Border(left=s, right=s, top=s, bottom=s)
    def double_bottom_border():
        t = Side(style='thin', color='1F3864')
        b = Side(style='double', color='1F3864')
        return Border(top=t, bottom=b)

    today = datetime.datetime(2026, 10, 24)

    # ═════════════════════════════════════════════════════════════
    # SAYFA 1: KASA ÜRETİM VE TERMİN PLANI
    # ═════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = 'Kasa_Uretim_ve_Termin_Plani'
    ws1.sheet_view.showGridLines = True

    # Title Banner
    ws1.merge_cells('A1:M1')
    ws1['A1'] = 'ERGÜNBAŞ WPC – KASIM AYI KASA ÜRETİM VE TERMİN PLANLAMA MATRİSİ'
    ws1['A1'].font = Font(bold=True, size=15, color='FFFFFF', name='Calibri')
    ws1['A1'].fill = fill(C_DARK)
    ws1['A1'].alignment = align()
    ws1.row_dimensions[1].height = 32

    ws1.merge_cells('A2:M2')
    ws1['A2'] = ('Reçete: 1 Takım Kasa = 2.5 Boy Profil (2 Dikme + 0.5 Üst Başlık)  |  '
                 '100 MM (10 kg/tk): Hat 108  |  140 MM (12.5 kg/tk): Hat 106 & Hat 102  |  '
                 '200 MM (16.5 kg/tk): Hat 107 [20lik Şu An Çalışmıyor Uyarısı]')
    ws1['A2'].font = Font(italic=True, size=9, color='FFFFFF', name='Calibri')
    ws1['A2'].fill = fill(C_MED)
    ws1['A2'].alignment = align()
    ws1.row_dimensions[2].height = 18

    # Headers
    k_headers = [
        'Sıra', 'Sipariş No', 'Termin Tarihi', 'Aciliyet / Kalan',
        'Kasa Kesiti', 'Renk', 'Ürün Tam Tanımı (Kalıp / Boy / Kesit)',
        'Sipariş (Takım)', 'Gereken Profil (Boy)', 'Birim (kg/tk)', 'Toplam Ağırlık (kg)',
        'Atanan Hat(lar)', 'Planlama / Operasyon Notu'
    ]
    k_widths = [6, 12, 14, 16, 12, 14, 38, 14, 16, 13, 16, 20, 28]

    for ci, (h, w) in enumerate(zip(k_headers, k_widths), 1):
        c = ws1.cell(3, ci, h)
        c.font = Font(bold=True, size=9, color='FFFFFF', name='Calibri')
        c.fill = fill(C_DARK)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = thin_border()
        ws1.column_dimensions[get_column_letter(ci)].width = w
    ws1.row_dimensions[3].height = 36

    # Sort kasa rows by termin
    kasa_rows_sorted = sorted(kasa_rows, key=lambda x: (
        x['term'] is None or not isinstance(x['term'], datetime.datetime),
        x['term'] if isinstance(x['term'], datetime.datetime) else datetime.datetime(2099,12,31)
    ))

    r_idx = 4
    for idx, kr in enumerate(kasa_rows_sorted, 1):
        t_val = kr['term']
        if isinstance(t_val, datetime.datetime):
            t_str = t_val.strftime('%d.%m.%Y')
            diff = (t_val - today).days
            if diff <= 5:
                status_str = f'⚠️ ACİL ({diff} gün)'
                s_fill = fill(C_RED)
                s_color = 'CC0000'
            elif diff <= 10:
                status_str = f'🔶 ÖNCELİKLİ ({diff} gün)'
                s_fill = fill(C_WARN)
                s_color = '9C5700'
            else:
                status_str = f'🟢 Normal ({diff} gün)'
                s_fill = fill(C_GREEN)
                s_color = '276A3C'
        elif t_val:
            t_str = str(t_val)
            status_str = 'Planlı'
            s_fill = fill(C_HEADER)
            s_color = '333333'
        else:
            t_str = '–'
            status_str = 'Termin Yok'
            s_fill = fill(C_HEADER)
            s_color = '666666'

        # Note
        op_note = ''
        if kr['size'] == '200 MM':
            op_note = '⚠️ 20lik kalıp/bakım kontrol edilmeli'
        elif kr['size'] == '100 MM':
            op_note = 'Hat 108 hızlı rota'
        elif kr['qty'] >= 200:
            op_note = 'Yüksek hacim – Çift hat eşzamanlı'

        row_vals = [
            idx,
            kr['sip'],
            t_str,
            status_str,
            kr['size'],
            kr['color'],
            kr['desc'],
            kr['qty'],
            kr['profil_boy'],
            kr['unit_kg'],
            kr['tot_kg'],
            kr['line'],
            op_note
        ]

        zebra_fill = fill(C_ZEBRA) if r_idx % 2 == 0 else fill('FFFFFF')

        for ci, val in enumerate(row_vals, 1):
            cell = ws1.cell(r_idx, ci, val)
            cell.border = thin_border()
            cell.alignment = Alignment(
                horizontal='left' if ci in (7, 13) else 'center',
                vertical='center',
                wrap_text=(ci in (7, 13))
            )
            
            if ci == 4:
                cell.fill = s_fill
                cell.font = Font(bold=True, size=9, color=s_color, name='Calibri')
            elif ci == 5:
                # Kesit rengi
                cell.fill = fill('DDEEFF') if val == '140 MM' else (fill('E2EFDA') if val == '100 MM' else fill('FFF2CC'))
                cell.font = Font(bold=True, size=10, name='Calibri')
            elif ci in (8, 9, 11):
                cell.fill = zebra_fill
                cell.font = Font(bold=True, size=10, name='Calibri')
                cell.number_format = '#,##0.0' if ci == 11 else '#,##0'
            elif ci == 10:
                cell.fill = zebra_fill
                cell.font = Font(size=9, name='Calibri')
                cell.number_format = '0.0 "kg"'
            elif ci == 13 and '⚠️' in str(val):
                cell.fill = fill(C_RED)
                cell.font = Font(bold=True, size=9, color='CC0000', name='Calibri')
            else:
                cell.fill = zebra_fill
                cell.font = Font(bold=(ci in (1, 2)), size=10, name='Calibri')

        ws1.row_dimensions[r_idx].height = 20
        r_idx += 1

    # Totals row for Kasa
    last_k_row = r_idx - 1
    ws1.cell(r_idx, 1, 'GENEL KASA TOPLAMI').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws1.cell(r_idx, 1).fill = fill(C_DARK)
    ws1.cell(r_idx, 1).alignment = align('left')

    ws1.cell(r_idx, 8, f'=SUM(H4:H{last_k_row})').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws1.cell(r_idx, 8).fill = fill(C_DARK)
    ws1.cell(r_idx, 8).number_format = '#,##0 " Takım"'
    ws1.cell(r_idx, 8).alignment = align()

    ws1.cell(r_idx, 9, f'=SUM(I4:I{last_k_row})').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws1.cell(r_idx, 9).fill = fill(C_DARK)
    ws1.cell(r_idx, 9).number_format = '#,##0 " Boy"'
    ws1.cell(r_idx, 9).alignment = align()

    ws1.cell(r_idx, 11, f'=SUM(K4:K{last_k_row})').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws1.cell(r_idx, 11).fill = fill(C_DARK)
    ws1.cell(r_idx, 11).number_format = '#,##0.0 " kg"'
    ws1.cell(r_idx, 11).alignment = align()

    for ci in [2, 3, 4, 5, 6, 7, 10, 12, 13]:
        c = ws1.cell(r_idx, ci)
        c.fill = fill(C_DARK)
        c.border = thin_border()
    ws1.row_dimensions[r_idx].height = 25
    r_idx += 2

    # ── KASA KESİT BAZLI YÖNETİCİ ÖZET TABLOSU ──
    ws1.merge_cells(f'B{r_idx}:L{r_idx}')
    ws1.cell(r_idx, 2, 'KASA KESİT BAZLI KAPASİTE VE HAT ATAMA ÖZETİ (kasa Sekmesi Verileri)').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws1.cell(r_idx, 2).fill = fill(C_MED)
    ws1.cell(r_idx, 2).alignment = align('left')
    r_idx += 1

    sum_hdrs = ['Kasa Kesiti', 'Toplam Sipariş (Takım)', 'Toplam Ağırlık (kg)', '1 Takım Ağırlığı', 'Atanan Hat(lar)', 'Hat Günlük Gücü', 'Süre (İş Günü)', 'Kritik Durum / Not']
    for ci, h in enumerate(sum_hdrs, 2):
        c = ws1.cell(r_idx, ci, h)
        c.font = Font(bold=True, size=9, color='FFFFFF', name='Calibri')
        c.fill = fill(C_DARK)
        c.alignment = align()
        c.border = thin_border()
    ws1.row_dimensions[r_idx].height = 22
    r_idx += 1

    k_summary_data = [
        ('100 MM KASA', '=COUNTIF(E4:E' + str(last_k_row) + ',"100 MM")', '=SUMIF(E4:E' + str(last_k_row) + ',"100 MM",H4:H' + str(last_k_row) + ')', '10.0 kg/tk', 'Hat 108 (Öncelikli) & Hat 102', '148 takım/gün (1.300 kg/gün)', '~8.1 iş günü', 'Seri ve istikrarlı hat'),
        ('140 MM KASA', '=COUNTIF(E4:E' + str(last_k_row) + ',"140 MM")', '=SUMIF(E4:E' + str(last_k_row) + ',"140 MM",H4:H' + str(last_k_row) + ')', '12.5 kg/tk', 'Hat 106 (Ana) + Hat 102 (Destek)', '270 takım/gün (çift hat)', '~16 - 19 iş günü (Tek hat: 38 gün)', 'En yüksek hacim – Çift hat şart'),
        ('200 MM KASA', '=COUNTIF(E4:E' + str(last_k_row) + ',"200 MM")', '=SUMIF(E4:E' + str(last_k_row) + ',"200 MM",H4:H' + str(last_k_row) + ')', '16.5 kg/tk', 'Hat 107', '160 takım/gün (2.500 kg/gün)', '~4.8 - 6.9 iş günü', '⚠️ DİKKAT: "20 lik şu an çalışmıyor"'),
    ]

    for k_item in k_summary_data:
        k_sz = k_item[0]
        # Calculate actual numbers for display
        sub_qty = sum(x['qty'] for x in kasa_rows if x['size'] == k_sz)
        sub_kg = sum(x['tot_kg'] for x in kasa_rows if x['size'] == k_sz)
        
        is_warn = '⚠️' in k_item[7]
        r_f = fill(C_RED) if is_warn else (fill('E2EFDA') if '100' in k_sz else fill(C_KASA))
        
        row_vals = [
            k_sz,
            sub_qty,
            sub_kg,
            k_item[3],
            k_item[4],
            k_item[5],
            k_item[6],
            k_item[7]
        ]
        for ci, v in enumerate(row_vals, 2):
            cell = ws1.cell(r_idx, ci, v)
            cell.fill = r_f
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal='left' if ci in (2, 6, 9) else 'center', vertical='center')
            if ci == 2:
                cell.font = Font(bold=True, size=10, name='Calibri')
            elif ci in (3, 4):
                cell.font = Font(bold=True, size=10, name='Calibri')
                cell.number_format = '#,##0 " Takım"' if ci == 3 else '#,##0.0 " kg"'
            elif ci == 9 and is_warn:
                cell.font = Font(bold=True, size=9, color='CC0000', name='Calibri')
            else:
                cell.font = Font(size=9, name='Calibri')
        ws1.row_dimensions[r_idx].height = 22
        r_idx += 1

    ws1.auto_filter.ref = f'A3:M{last_k_row}'
    ws1.freeze_panes = 'C4'

    # ═════════════════════════════════════════════════════════════
    # SAYFA 2: PERVAZ ÜRETİM VE TERMİN PLANI
    # ═════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet('Pervaz_Uretim_ve_Termin_Plani')
    ws2.sheet_view.showGridLines = True

    # Title Banner
    ws2.merge_cells('A1:L1')
    ws2['A1'] = 'ERGÜNBAŞ WPC – KASIM AYI PERVAZ ÜRETİM VE TERMİN PLANLAMA MATRİSİ'
    ws2['A1'].font = Font(bold=True, size=15, color='FFFFFF', name='Calibri')
    ws2['A1'].fill = fill(C_DARK)
    ws2['A1'].alignment = align()
    ws2.row_dimensions[1].height = 32

    ws2.merge_cells('A2:L2')
    ws2['A2'] = ('Kaynak: sipariş_içeriği_düzenli (2).xlsx "pervaz" Sekmesi Analizi  |  '
                 '50X80 Pervaz (1.55 kg/tk): Hat 109 (Ana - 1.204 boy/24s) & Hat 103 (Destek - 756 boy/24s)  |  '
                 '80X80 Pervaz (2.22 kg/tk): Hat 101 / 109 (90 Takım - ~1 Gün)')
    ws2['A2'].font = Font(italic=True, size=9, color='FFFFFF', name='Calibri')
    ws2['A2'].fill = fill(C_MED)
    ws2['A2'].alignment = align()
    ws2.row_dimensions[2].height = 18

    # Headers
    p_headers = [
        'Sıra', 'Sipariş No', 'Termin Tarihi', 'Aciliyet / Kalan',
        'Pervaz Tipi', 'Renk', 'Ürün Tam Tanımı (Kalıp / Boy)',
        'Sipariş (Takım)', 'Gereken Profil (Boy)', 'Birim (kg/tk)', 'Toplam Ağırlık (kg)',
        'Atanan Hat(lar)'
    ]
    p_widths = [6, 12, 14, 16, 12, 14, 38, 14, 16, 13, 16, 22]

    for ci, (h, w) in enumerate(zip(p_headers, p_widths), 1):
        c = ws2.cell(3, ci, h)
        c.font = Font(bold=True, size=9, color='FFFFFF', name='Calibri')
        c.fill = fill(C_DARK)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = thin_border()
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.row_dimensions[3].height = 36

    # Sort pervaz rows by termin
    pervaz_rows_sorted = sorted(pervaz_rows, key=lambda x: (
        x['term'] is None or not isinstance(x['term'], datetime.datetime),
        x['term'] if isinstance(x['term'], datetime.datetime) else datetime.datetime(2099,12,31)
    ))

    r_idx2 = 4
    for idx, pr in enumerate(pervaz_rows_sorted, 1):
        t_val = pr['term']
        if isinstance(t_val, datetime.datetime):
            t_str = t_val.strftime('%d.%m.%Y')
            diff = (t_val - today).days
            if diff <= 5:
                status_str = f'⚠️ ACİL ({diff} gün)'
                s_fill = fill(C_RED)
                s_color = 'CC0000'
            elif diff <= 10:
                status_str = f'🔶 ÖNCELİKLİ ({diff} gün)'
                s_fill = fill(C_WARN)
                s_color = '9C5700'
            else:
                status_str = f'🟢 Normal ({diff} gün)'
                s_fill = fill(C_GREEN)
                s_color = '276A3C'
        elif t_val:
            t_str = str(t_val)
            status_str = 'Planlı'
            s_fill = fill(C_HEADER)
            s_color = '333333'
        else:
            t_str = '–'
            status_str = 'Termin Yok'
            s_fill = fill(C_HEADER)
            s_color = '666666'

        row_vals = [
            idx,
            pr['sip'],
            t_str,
            status_str,
            pr['ptype'],
            pr['color'],
            pr['desc'],
            pr['qty'],
            pr['profil_boy'],
            pr['unit_kg'],
            pr['tot_kg'],
            pr['line']
        ]

        zebra_fill = fill(C_ZEBRA) if r_idx2 % 2 == 0 else fill('FFFFFF')

        for ci, val in enumerate(row_vals, 1):
            cell = ws2.cell(r_idx2, ci, val)
            cell.border = thin_border()
            cell.alignment = Alignment(
                horizontal='left' if ci == 7 else 'center',
                vertical='center',
                wrap_text=(ci == 7)
            )
            
            if ci == 4:
                cell.fill = s_fill
                cell.font = Font(bold=True, size=9, color=s_color, name='Calibri')
            elif ci == 5:
                cell.fill = fill(C_PERVAZ) if val == '50X80' else fill('FFEB9C')
                cell.font = Font(bold=True, size=10, name='Calibri')
            elif ci in (8, 9, 11):
                cell.fill = zebra_fill
                cell.font = Font(bold=True, size=10, name='Calibri')
                cell.number_format = '#,##0.0' if ci == 11 else '#,##0'
            elif ci == 10:
                cell.fill = zebra_fill
                cell.font = Font(size=9, name='Calibri')
                cell.number_format = '0.00 "kg"'
            else:
                cell.fill = zebra_fill
                cell.font = Font(bold=(ci in (1, 2)), size=10, name='Calibri')

        ws2.row_dimensions[r_idx2].height = 20
        r_idx2 += 1

    # Totals row for Pervaz
    last_p_row = r_idx2 - 1
    ws2.cell(r_idx2, 1, 'GENEL PERVAZ TOPLAMI').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws2.cell(r_idx2, 1).fill = fill(C_DARK)
    ws2.cell(r_idx2, 1).alignment = align('left')

    ws2.cell(r_idx2, 8, f'=SUM(H4:H{last_p_row})').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws2.cell(r_idx2, 8).fill = fill(C_DARK)
    ws2.cell(r_idx2, 8).number_format = '#,##0 " Takım"'
    ws2.cell(r_idx2, 8).alignment = align()

    ws2.cell(r_idx2, 9, f'=SUM(I4:I{last_p_row})').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws2.cell(r_idx2, 9).fill = fill(C_DARK)
    ws2.cell(r_idx2, 9).number_format = '#,##0 " Boy"'
    ws2.cell(r_idx2, 9).alignment = align()

    ws2.cell(r_idx2, 11, f'=SUM(K4:K{last_p_row})').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws2.cell(r_idx2, 11).fill = fill(C_DARK)
    ws2.cell(r_idx2, 11).number_format = '#,##0.0 " kg"'
    ws2.cell(r_idx2, 11).alignment = align()

    for ci in [2, 3, 4, 5, 6, 7, 10, 12]:
        c = ws2.cell(r_idx2, ci)
        c.fill = fill(C_DARK)
        c.border = thin_border()
    ws2.row_dimensions[r_idx2].height = 25
    r_idx2 += 2

    # ── PERVAZ TİP BAZLI YÖNETİCİ ÖZETİ ──
    ws2.merge_cells(f'B{r_idx2}:K{r_idx2}')
    ws2.cell(r_idx2, 2, 'PERVAZ TİP BAZLI KAPASİTE VE HAT ATAMA ÖZETİ (pervaz Sekmesi Verileri)').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws2.cell(r_idx2, 2).fill = fill(C_MED)
    ws2.cell(r_idx2, 2).alignment = align('left')
    r_idx2 += 1

    p_sum_hdrs = ['Pervaz Tipi', 'Toplam Sipariş (Takım)', 'Toplam Ağırlık (kg)', '1 Takım Ağırlığı', 'Atanan Hat(lar)', 'Hat Günlük Gücü', 'Süre (İş Günü)', 'Operasyon Notu']
    for ci, h in enumerate(p_sum_hdrs, 2):
        c = ws2.cell(r_idx2, ci, h)
        c.font = Font(bold=True, size=9, color='FFFFFF', name='Calibri')
        c.fill = fill(C_DARK)
        c.alignment = align()
        c.border = thin_border()
    ws2.row_dimensions[r_idx2].height = 22
    r_idx2 += 1

    p_summary_data = [
        ('50X80 PERVAZ', sum(x['qty'] for x in pervaz_rows if x['ptype'] == '50X80'), sum(x['tot_kg'] for x in pervaz_rows if x['ptype'] == '50X80'), '1.55 kg/tk', 'Hat 109 (Lokomotif) + Hat 103 (Destek)', '392 takım/gün (1.960 boy)', '~19.0 iş günü', 'Hat 109 saatte 50.2 boy ile en hızlı hat'),
        ('80X80 PERVAZ', sum(x['qty'] for x in pervaz_rows if x['ptype'] == '80X80'), sum(x['tot_kg'] for x in pervaz_rows if x['ptype'] == '80X80'), '2.22 kg/tk', 'Hat 101 / Hat 109', '90 takım/gün', '~1.0 iş günü', 'Geniş kulaklı pervaz – Kısa sürede tamamlanır'),
    ]

    for p_item in p_summary_data:
        p_t = p_item[0]
        r_f = fill(C_PERVAZ) if '50X80' in p_t else fill('FFEB9C')
        row_vals = [
            p_t,
            p_item[1],
            p_item[2],
            p_item[3],
            p_item[4],
            p_item[5],
            p_item[6],
            p_item[7]
        ]
        for ci, v in enumerate(row_vals, 2):
            cell = ws2.cell(r_idx2, ci, v)
            cell.fill = r_f
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal='left' if ci in (2, 6, 9) else 'center', vertical='center')
            if ci == 2:
                cell.font = Font(bold=True, size=10, name='Calibri')
            elif ci in (3, 4):
                cell.font = Font(bold=True, size=10, name='Calibri')
                cell.number_format = '#,##0 " Takım"' if ci == 3 else '#,##0.0 " kg"'
            else:
                cell.font = Font(size=9, name='Calibri')
        ws2.row_dimensions[r_idx2].height = 22
        r_idx2 += 1

    ws2.auto_filter.ref = f'A3:L{last_p_row}'
    ws2.freeze_panes = 'C4'

    # ═════════════════════════════════════════════════════════════
    # SAYFA 3: SİPARİŞ BAZINDA KASA & PERVAZ SENKRON MASTER PLAN
    # ═════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet('Siparis_Senkron_Termin_Plani')
    ws3.sheet_view.showGridLines = True

    # Title Banner
    ws3.merge_cells('A1:N1')
    ws3['A1'] = 'ERGÜNBAŞ WPC – SİPARİŞ BAZINDA KASA VE PERVAZ SENKRON ÜRETİM & TERMİN TAKVİMİ'
    ws3['A1'].font = Font(bold=True, size=15, color='FFFFFF', name='Calibri')
    ws3['A1'].fill = fill(C_DARK)
    ws3['A1'].alignment = align()
    ws3.row_dimensions[1].height = 32

    ws3.merge_cells('A2:N2')
    ws3['A2'] = ('Bu sayfa her siparişin 100mm, 140mm, 200mm Kasa ve 50x80, 80x80 Pervaz ihtiyaçlarını tek satırda birleştirir.  '
                 'Hangi siparişin ne zaman kapı montajına hazır olacağını gösterir.')
    ws3['A2'].font = Font(italic=True, size=9, color='FFFFFF', name='Calibri')
    ws3['A2'].fill = fill(C_MED)
    ws3['A2'].alignment = align()
    ws3.row_dimensions[2].height = 18

    # Headers
    m_headers = [
        'Sıra', 'Sipariş No', 'Termin Tarihi', 'Aciliyet / Kalan',
        '100 Kasa\n(Takım)', '140 Kasa\n(Takım)', '200 Kasa\n(Takım)', 'TOPLAM KASA\n(Takım)',
        '50x80 Pervaz\n(Takım)', '80x80 Pervaz\n(Takım)', 'TOPLAM PERVAZ\n(Takım)',
        'Kasa İhtiyaç\n(İş Günü)', 'Pervaz İhtiyaç\n(İş Günü)', 'Kritik Durum / Sevkiyat Notu'
    ]
    m_widths = [6, 12, 14, 16, 12, 12, 12, 15, 13, 13, 15, 13, 13, 30]

    for ci, (h, w) in enumerate(zip(m_headers, m_widths), 1):
        c = ws3.cell(3, ci, h)
        c.font = Font(bold=True, size=9, color='FFFFFF', name='Calibri')
        c.fill = fill(C_DARK)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = thin_border()
        ws3.column_dimensions[get_column_letter(ci)].width = w
    ws3.row_dimensions[3].height = 36

    # Aggregate by order
    all_order_ids = set([x['sip'] for x in kasa_rows if x['sip']] + [x['sip'] for x in pervaz_rows if x['sip']])
    
    order_master = []
    for sno in all_order_ids:
        k_items = [x for x in kasa_rows if x['sip'] == sno]
        p_items = [x for x in pervaz_rows if x['sip'] == sno]

        t_date = None
        for it in (k_items + p_items):
            if it['term'] and isinstance(it['term'], datetime.datetime):
                t_date = it['term']
                break

        q_100 = sum(x['qty'] for x in k_items if x['size'] == '100 MM')
        q_140 = sum(x['qty'] for x in k_items if x['size'] == '140 MM')
        q_200 = sum(x['qty'] for x in k_items if x['size'] == '200 MM')
        tot_k = q_100 + q_140 + q_200

        q_p50 = sum(x['qty'] for x in p_items if x['ptype'] == '50X80')
        q_p80 = sum(x['qty'] for x in p_items if x['ptype'] == '80X80')
        tot_p = q_p50 + q_p80

        # Estimate production days
        # Kasa rate ~ 270 tk/day
        k_days = round(tot_k / 270.0, 1) if tot_k > 0 else 0.0
        # Pervaz rate ~ 392 tk/day
        p_days = round(tot_p / 392.0, 1) if tot_p > 0 else 0.0

        order_master.append({
            'sip': sno,
            'term': t_date,
            'q_100': q_100,
            'q_140': q_140,
            'q_200': q_200,
            'tot_k': tot_k,
            'q_p50': q_p50,
            'q_p80': q_p80,
            'tot_p': tot_p,
            'k_days': k_days,
            'p_days': p_days
        })

    # Sort master orders by termin
    order_master_sorted = sorted(order_master, key=lambda x: (
        x['term'] is None or not isinstance(x['term'], datetime.datetime),
        x['term'] if isinstance(x['term'], datetime.datetime) else datetime.datetime(2099,12,31)
    ))

    r_idx3 = 4
    for idx, om in enumerate(order_master_sorted, 1):
        t_val = om['term']
        if isinstance(t_val, datetime.datetime):
            t_str = t_val.strftime('%d.%m.%Y')
            diff = (t_val - today).days
            if diff <= 5:
                status_str = f'⚠️ ACİL ({diff} gün)'
                s_fill = fill(C_RED)
                s_color = 'CC0000'
            elif diff <= 10:
                status_str = f'🔶 ÖNCELİKLİ ({diff} gün)'
                s_fill = fill(C_WARN)
                s_color = '9C5700'
            else:
                status_str = f'🟢 Normal ({diff} gün)'
                s_fill = fill(C_GREEN)
                s_color = '276A3C'
        else:
            t_str = '–'
            status_str = 'Termin Yok'
            s_fill = fill(C_HEADER)
            s_color = '666666'

        sevkiyat_notu = ''
        if om['q_200'] > 0:
            sevkiyat_notu = f"⚠️ {om['q_200']} Tk 200'lük Kasa var (Kalıp kontrol!)"
        elif om['tot_k'] > om['tot_p'] * 1.5:
            sevkiyat_notu = 'Kasa ağırlıklı sipariş'
        elif om['tot_p'] > om['tot_k'] * 1.5:
            sevkiyat_notu = 'Pervaz ağırlıklı sipariş'
        else:
            sevkiyat_notu = 'Dengeli set'

        row_vals = [
            idx,
            om['sip'],
            t_str,
            status_str,
            om['q_100'] if om['q_100'] > 0 else '–',
            om['q_140'] if om['q_140'] > 0 else '–',
            om['q_200'] if om['q_200'] > 0 else '–',
            om['tot_k'],
            om['q_p50'] if om['q_p50'] > 0 else '–',
            om['q_p80'] if om['q_p80'] > 0 else '–',
            om['tot_p'],
            om['k_days'],
            om['p_days'],
            sevkiyat_notu
        ]

        zebra_fill = fill(C_ZEBRA) if r_idx3 % 2 == 0 else fill('FFFFFF')

        for ci, val in enumerate(row_vals, 1):
            cell = ws3.cell(r_idx3, ci, val)
            cell.border = thin_border()
            cell.alignment = Alignment(
                horizontal='left' if ci == 14 else 'center',
                vertical='center',
                wrap_text=(ci == 14)
            )

            if ci == 4:
                cell.fill = s_fill
                cell.font = Font(bold=True, size=9, color=s_color, name='Calibri')
            elif ci in (5, 6, 7):
                cell.fill = fill(C_KASA)
                cell.font = Font(bold=(val != '–'), size=10, name='Calibri')
            elif ci == 8:
                cell.fill = fill('D2E4F8')
                cell.font = Font(bold=True, size=10, name='Calibri')
                cell.number_format = '#,##0'
            elif ci in (9, 10):
                cell.fill = fill(C_PERVAZ)
                cell.font = Font(bold=(val != '–'), size=10, name='Calibri')
            elif ci == 11:
                cell.fill = fill('FFE899')
                cell.font = Font(bold=True, size=10, name='Calibri')
                cell.number_format = '#,##0'
            elif ci in (12, 13):
                cell.fill = zebra_fill
                cell.font = Font(bold=True, size=10, name='Calibri')
                cell.number_format = '0.0 "gün"'
            elif ci == 14 and '⚠️' in str(val):
                cell.fill = fill(C_RED)
                cell.font = Font(bold=True, size=9, color='CC0000', name='Calibri')
            else:
                cell.fill = zebra_fill
                cell.font = Font(bold=(ci in (1, 2)), size=10, name='Calibri')

        ws3.row_dimensions[r_idx3].height = 20
        r_idx3 += 1

    # Totals row for Master
    last_m_row = r_idx3 - 1
    ws3.cell(r_idx3, 1, 'TOPLAM').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws3.cell(r_idx3, 1).fill = fill(C_DARK)
    ws3.cell(r_idx3, 1).alignment = align('left')

    ws3.cell(r_idx3, 8, f'=SUM(H4:H{last_m_row})').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws3.cell(r_idx3, 8).fill = fill(C_DARK)
    ws3.cell(r_idx3, 8).number_format = '#,##0 " Tk Kasa"'
    ws3.cell(r_idx3, 8).alignment = align()

    ws3.cell(r_idx3, 11, f'=SUM(K4:K{last_m_row})').font = Font(bold=True, size=11, color='FFFFFF', name='Calibri')
    ws3.cell(r_idx3, 11).fill = fill(C_DARK)
    ws3.cell(r_idx3, 11).number_format = '#,##0 " Tk Pervaz"'
    ws3.cell(r_idx3, 11).alignment = align()

    for ci in [2, 3, 4, 5, 6, 7, 9, 10, 12, 13, 14]:
        c = ws3.cell(r_idx3, ci)
        c.fill = fill(C_DARK)
        c.border = thin_border()
    ws3.row_dimensions[r_idx3].height = 25

    ws3.auto_filter.ref = f'A3:N{last_m_row}'
    ws3.freeze_panes = 'C4'

    # Save workbook
    out_desktop = r'C:\Users\llboz\OneDrive\Desktop\Ergunbas_Kasa_ve_Pervaz_Detayli_Uretim_Plani.xlsx'
    out_alt = r'C:\Users\llboz\OneDrive\Desktop\Ergunbas_Kasa_ve_Pervaz_Detayli_Uretim_Plani (Guncel_2.5_Kasa).xlsx'
    try:
        wb.save(out_desktop)
        print(f'Successfully created: {out_desktop}')
    except Exception as e:
        print(f'Desktop ana dosya açık ({e}), alternatif kaydediliyor...')
        wb.save(out_alt)
        print(f'Successfully created: {out_alt}')
    return out_desktop

if __name__ == '__main__':
    build_kasa_pervaz_plan()
