import json
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from collections import defaultdict

APP_DIR = os.path.dirname(os.path.abspath(__file__))
FILES = [
    os.path.join(APP_DIR, 'data_days_2026-08.json'),
    os.path.join(APP_DIR, 'data_days_2026-09.json'),
    os.path.join(APP_DIR, 'data_days_2026-10.json')
]

records = []
prod_dim_stats = defaultdict(lambda: defaultdict(lambda: {'qty': 0.0, 'kg': 0.0, 'fire_kg': 0.0, 'hours': 0.0, 'hats': set(), 'weeks': set()}))

def norm_product_name(p):
    p_u = p.strip().upper()
    if '50X80' in p_u: return '50X80 PERVAZ'
    if '80X80' in p_u: return '80X80 PERVAZ'
    if '35X100' in p_u: return '35X100 PERVAZ'
    if '35X150' in p_u: return '35X150 PERVAZ'
    if '35X80' in p_u: return '35X80 PERVAZ'
    if '100 MM KASA' in p_u or '100MM KASA' in p_u: return '100 mm KASA'
    if '140 MM KASA' in p_u or '140MM KASA' in p_u: return '140 mm KASA'
    if '200 MM KASA' in p_u or '200MM KASA' in p_u: return '200 mm KASA'
    if 'SEREN(39MM)' in p_u: return 'SEREN(39mm)'
    if 'SEREN' in p_u and '39' not in p_u: return 'SEREN'
    if 'CAM' in p_u and 'ITA' in p_u: return 'CAM ÇITASI'
    if 'U KAPAMA' in p_u: return 'U KAPAMA'
    if 'KT100' in p_u: return 'KT100-220x24-ANT'
    if 'KT102' in p_u and 'KAH' in p_u: return 'KT102-219x26-KAH'
    if 'KT102' in p_u: return 'KT102-219x26-ANT'
    return p.strip()

for fpath in FILES:
    if not os.path.exists(fpath):
        continue
    with open(fpath, 'r', encoding='utf-8') as f:
        data = json.load(f)
    days = data.get('days', {})
    for dk, dval in days.items():
        date_str = dval.get('date', '').strip()
        if not date_str:
            continue
        dt = None
        for fmt in ('%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y'):
            try:
                dt = datetime.strptime(date_str, fmt)
                break
            except Exception:
                pass
        if not dt:
            continue

        iso_year, iso_week, _ = dt.isocalendar()
        week_key = f"{iso_year}-W{iso_week:02d}"

        for shift in ['gunduz', 'gece']:
            sh = dval.get(shift, {})
            # Extruders
            for row in sh.get('extruders', []):
                p_raw = row.get('product', '').strip()
                if not p_raw:
                    continue
                p = norm_product_name(p_raw)
                hat_no = str(row.get('hat', '')).strip()
                h_name = f"Hat {hat_no}" if hat_no.isdigit() else hat_no
                prod_kg = float(row.get('prod_kg') or 0.0)
                fire_kg = float(row.get('fire_kg') or 0.0)
                qty = float(row.get('qty') or 0.0)
                hours = float(row.get('hours') or 0.0)
                heads = int(row.get('heads', 1.0) or 1)
                l = row.get('length')

                if l is None or l <= 0:
                    dim_str = "2.10 m (210 cm)" if "SEREN" in p else ("2.15 m (215 cm)" if "KASA" in p else "2.20 m (220 cm)")
                else:
                    dim_str = f"{l:.2f} m ({int(round(l*100))} cm)"

                records.append({
                    'date': dt, 'week': week_key, 'hat': h_name, 'product': p,
                    'dim': dim_str, 'heads': heads, 'prod_kg': prod_kg,
                    'fire_kg': fire_kg, 'qty': qty, 'hours': hours, 'type': 'profil'
                })

                s = prod_dim_stats[p][dim_str]
                s['qty'] += qty
                s['kg'] += prod_kg
                s['fire_kg'] += fire_kg
                s['hours'] += hours
                s['hats'].add(h_name)
                s['weeks'].add(week_key)

            # Levha
            for row in sh.get('levha', []):
                c = row.get('color', '').strip()
                if not c:
                    continue
                p = norm_product_name(f"{c} Levha" if "LEVHA" not in c.upper() else c)
                prod_kg = float(row.get('total_kg') or row.get('prod_kg') or 0.0)
                fire_kg = float(row.get('dead_fire_kg') or row.get('fire_kg') or 0.0)
                qty = float(row.get('qty') or 0.0)
                hours = float(row.get('hours') or 0.0)
                w = row.get('width')
                l = row.get('length')

                if w and l:
                    dim_str = f"{int(w)}x{int(l)} cm ({w/100:.2f}x{l/100:.2f} m)"
                else:
                    dim_str = "94x207 cm (0.94x2.07 m)"

                records.append({
                    'date': dt, 'week': week_key, 'hat': 'Hat 201', 'product': p,
                    'dim': dim_str, 'heads': 1, 'prod_kg': prod_kg,
                    'fire_kg': fire_kg, 'qty': qty, 'hours': hours, 'type': 'levha'
                })

                s = prod_dim_stats[p][dim_str]
                s['qty'] += qty
                s['kg'] += prod_kg
                s['fire_kg'] += fire_kg
                s['hours'] += hours
                s['hats'].add('Hat 201')
                s['weeks'].add(week_key)

all_weeks = sorted(list(set(r['week'] for r in records)))

WEEK_LABELS = {
    '2026-W31': 'W31 (27 Tem - 02 Ağu)',
    '2026-W32': 'W32 (03 Ağu - 09 Ağu)',
    '2026-W33': 'W33 (10 Ağu - 16 Ağu)',
    '2026-W34': 'W34 (17 Ağu - 23 Ağu)',
    '2026-W35': 'W35 (24 Ağu - 30 Ağu)',
    '2026-W36': 'W36 (31 Ağu - 06 Eyl)',
    '2026-W37': 'W37 (07 Eyl - 13 Eyl)',
    '2026-W38': 'W38 (14 Eyl - 20 Eyl)',
    '2026-W39': 'W39 (21 Eyl - 27 Eyl)',
    '2026-W40': 'W40 (28 Eyl - 04 Eki)',
    '2026-W41': 'W41 (05 Eki - 11 Eki)',
    '2026-W42': 'W42 (12 Eki - 18 Eki)',
    '2026-W43': 'W43 (19 Eki - 25 Eki)'
}

PRODUCT_CONFIG = {
    '50X80 PERVAZ': {'cat': 'Pervaz', 'nominal_kg': 1.55, 'heads': 'Çift Göz (2 Kafa)', 'hat': 'Hat 109 & Hat 103', 'recipe': 'Kasa Pervaz', 'fire': 0.289},
    '80X80 PERVAZ': {'cat': 'Pervaz', 'nominal_kg': 2.20, 'heads': 'Tek/Çift Göz', 'hat': 'Hat 101 & Hat 109', 'recipe': 'Kasa Pervaz', 'fire': 0.194},
    '35X100 PERVAZ': {'cat': 'Pervaz', 'nominal_kg': 2.08, 'heads': 'Tek Göz (1 Kafa)', 'hat': 'Hat 101', 'recipe': 'Kasa Pervaz', 'fire': 0.125},
    '35X150 PERVAZ': {'cat': 'Pervaz', 'nominal_kg': 2.97, 'heads': 'Tek Göz (1 Kafa)', 'hat': 'Hat 101', 'recipe': 'Kasa Pervaz', 'fire': 0.177},
    '35X80 PERVAZ': {'cat': 'Pervaz', 'nominal_kg': 2.94, 'heads': 'Tek Göz (1 Kafa)', 'hat': 'Hat 101', 'recipe': 'Kasa Pervaz', 'fire': 0.113},
    '100 mm KASA': {'cat': 'Kasa', 'nominal_kg': 4.00, 'heads': 'Tek Göz (1 Kafa)', 'hat': 'Hat 108 & Hat 102', 'recipe': 'Kasa Pervaz', 'fire': 0.300},
    '140 mm KASA': {'cat': 'Kasa', 'nominal_kg': 4.95, 'heads': 'Tek Göz (1 Kafa)', 'hat': 'Hat 106 & Hat 108', 'recipe': 'Kasa Pervaz', 'fire': 0.300},
    '200 mm KASA': {'cat': 'Kasa', 'nominal_kg': 6.66, 'heads': 'Tek Göz (1 Kafa)', 'hat': 'Hat 107', 'recipe': 'Kasa Pervaz', 'fire': 0.300},
    'SEREN': {'cat': 'Seren', 'nominal_kg': 3.39, 'heads': 'Çift Göz (2 Kafa)', 'hat': 'Hat 104 & Hat 105', 'recipe': 'Seren', 'fire': 0.260},
    'SEREN(39mm)': {'cat': 'Seren', 'nominal_kg': 3.81, 'heads': 'Tek Göz Özel', 'hat': 'Hat 107', 'recipe': 'Seren', 'fire': 0.220},
    'KT100-220x24-ANT': {'cat': 'Deck', 'nominal_kg': 10.33, 'heads': 'Tek Göz', 'hat': 'Hat 107', 'recipe': 'Deck Antrasit', 'fire': 0.250},
    'KT102-219x26-ANT': {'cat': 'Deck', 'nominal_kg': 10.52, 'heads': 'Tek Göz', 'hat': 'Hat 107', 'recipe': 'Deck Antrasit', 'fire': 0.250},
    'KT102-219x26-KAH': {'cat': 'Deck', 'nominal_kg': 9.26, 'heads': 'Tek Göz', 'hat': 'Hat 107', 'recipe': 'Deck Kahverengi', 'fire': 0.250},
    'CAM ÇITASI': {'cat': 'Çıta', 'nominal_kg': 1.50, 'heads': 'Tek Göz', 'hat': 'Hat 102', 'recipe': 'Kasa Pervaz', 'fire': 0.150},
    'U KAPAMA': {'cat': 'Profil', 'nominal_kg': 0.48, 'heads': 'Tek Göz', 'hat': 'Hat 102', 'recipe': 'Kasa Pervaz', 'fire': 0.150},
    'DIŞBUDAK Levha': {'cat': 'Levha', 'nominal_kg': 4.77, 'heads': 'Levha Kalıbı', 'hat': 'Hat 201', 'recipe': 'Levha Renkli', 'fire': 0.050},
    'TEAK Levha': {'cat': 'Levha', 'nominal_kg': 4.25, 'heads': 'Levha Kalıbı', 'hat': 'Hat 201', 'recipe': 'Levha Renkli', 'fire': 0.050},
    'BEYAZ TEAK Levha': {'cat': 'Levha', 'nominal_kg': 4.41, 'heads': 'Levha Kalıbı', 'hat': 'Hat 201', 'recipe': 'Levha Beyaz', 'fire': 0.050},
    'MIX TEAK Levha': {'cat': 'Levha', 'nominal_kg': 4.25, 'heads': 'Levha Kalıbı', 'hat': 'Hat 201', 'recipe': 'Levha Renkli', 'fire': 0.050},
    'VIZON Levha': {'cat': 'Levha', 'nominal_kg': 4.77, 'heads': 'Levha Kalıbı', 'hat': 'Hat 201', 'recipe': 'Levha Renkli', 'fire': 0.050},
    'MIX BEYAZ TEAK Levha': {'cat': 'Levha', 'nominal_kg': 4.41, 'heads': 'Levha Kalıbı', 'hat': 'Hat 201', 'recipe': 'Levha Beyaz', 'fire': 0.050},
    'SONOMO Levha': {'cat': 'Levha', 'nominal_kg': 4.33, 'heads': 'Levha Kalıbı', 'hat': 'Hat 201', 'recipe': 'Levha Renkli', 'fire': 0.050},
    'AY GRİ Levha': {'cat': 'Levha', 'nominal_kg': 4.64, 'heads': 'Levha Kalıbı', 'hat': 'Hat 201', 'recipe': 'Levha Renkli', 'fire': 0.050},
    'ANTRASİT Levha': {'cat': 'Levha', 'nominal_kg': 4.77, 'heads': 'Levha Kalıbı', 'hat': 'Hat 201', 'recipe': 'Levha Renkli', 'fire': 0.050},
    'KOYU CEVİZ Levha': {'cat': 'Levha', 'nominal_kg': 4.79, 'heads': 'Levha Kalıbı', 'hat': 'Hat 201', 'recipe': 'Levha Renkli', 'fire': 0.050},
    'COCO Levha': {'cat': 'Levha', 'nominal_kg': 4.79, 'heads': 'Levha Kalıbı', 'hat': 'Hat 201', 'recipe': 'Levha Renkli', 'fire': 0.050}
}

# Aggregate by product
prod_totals = defaultdict(lambda: {
    'prod_kg': 0.0, 'fire_kg': 0.0, 'qty': 0.0, 'hours': 0.0,
    'weekly_kg': defaultdict(float), 'weekly_qty': defaultdict(float)
})

line_totals = defaultdict(lambda: {
    'prod_kg': 0.0, 'fire_kg': 0.0, 'qty': 0.0, 'hours': 0.0,
    'products': set(), 'weekly_kg': defaultdict(float), 'weekly_qty': defaultdict(float)
})

for r in records:
    p = r['product']
    w = r['week']
    pkg = r['prod_kg']
    fkg = r['fire_kg']
    qty = r['qty']
    hrs = r['hours']
    h = r['hat']

    s = prod_totals[p]
    s['prod_kg'] += pkg
    s['fire_kg'] += fkg
    s['qty'] += qty
    s['hours'] += hrs
    s['weekly_kg'][w] += pkg
    s['weekly_qty'][w] += qty

    ld = line_totals[h]
    ld['prod_kg'] += pkg
    ld['fire_kg'] += fkg
    ld['qty'] += qty
    ld['hours'] += hrs
    ld['products'].add(p)
    ld['weekly_kg'][w] += pkg
    ld['weekly_qty'][w] += qty

sorted_prods = sorted(prod_totals.items(), key=lambda x: x[1]['prod_kg'], reverse=True)

# Determine dominant dimension & variations for each product
prod_dim_summary = {}
for pname, dim_map in prod_dim_stats.items():
    total_p_kg = sum(x['kg'] for x in dim_map.values())
    sorted_dims = sorted(dim_map.items(), key=lambda x: x[1]['kg'], reverse=True)
    if not sorted_dims:
        continue
    top_dim, top_stat = sorted_dims[0]
    top_pct = (top_stat['kg'] / total_p_kg * 100) if total_p_kg > 0 else 0
    top_label = f"{top_dim} (%{top_pct:.1f})"
    
    other_labels = []
    for d, st in sorted_dims[1:]:
        p_pct = (st['kg'] / total_p_kg * 100) if total_p_kg > 0 else 0
        if p_pct >= 0.5:
            other_labels.append(f"{d} (%{p_pct:.1f})")
    
    prod_dim_summary[pname] = {
        'top_dim': top_dim,
        'top_label': top_label,
        'other_label': ", ".join(other_labels) if other_labels else "Tek Ölçü Çekildi"
    }

# -------------------------------------------------------------
# OPENPYXL FORMATLAMA VE ÇALIŞMA KİTABI
# -------------------------------------------------------------
wb = openpyxl.Workbook()
wb.remove(wb.active)

CLR_PRIMARY_DARK = "0F172A"
CLR_HEADER_FILL = "1E293B"
CLR_EMERALD_DARK = "065F46"
CLR_EMERALD_LIGHT = "D1FAE5"
CLR_BLUE_LIGHT = "EFF6FF"
CLR_ZEBRA = "F8FAFC"
CLR_BORDER = "CBD5E1"

font_title = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="94A3B8")
font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
font_data = Font(name="Segoe UI", size=10, color="0F172A")
font_bold = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
font_total = Font(name="Segoe UI", size=11, bold=True, color="1E3A8A")

fill_title = PatternFill(start_color=CLR_PRIMARY_DARK, end_color=CLR_PRIMARY_DARK, fill_type="solid")
fill_header = PatternFill(start_color=CLR_HEADER_FILL, end_color=CLR_HEADER_FILL, fill_type="solid")
fill_header_green = PatternFill(start_color=CLR_EMERALD_DARK, end_color=CLR_EMERALD_DARK, fill_type="solid")
fill_zebra = PatternFill(start_color=CLR_ZEBRA, end_color=CLR_ZEBRA, fill_type="solid")
fill_highlight = PatternFill(start_color=CLR_EMERALD_LIGHT, end_color=CLR_EMERALD_LIGHT, fill_type="solid")
fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

thin_side = Side(border_style="thin", color=CLR_BORDER)
thick_bottom = Side(border_style="medium", color="475569")
double_bottom = Side(border_style="double", color="1E3A8A")

border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
border_header = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thick_bottom)
border_total = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_bottom)

align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")
align_center = Alignment(horizontal="center", vertical="center")

def auto_fit_columns(ws, min_width=12, max_width=50):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row in [1, 2]:
                continue
            val_str = str(cell.value or '')
            if cell.number_format and '%' in cell.number_format:
                val_str += ' %'
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(min_width, min(max_len + 3, max_width))

# ============================================================================
# SAYFA 1: HAFTALIK PLANLAMA MASTER (GERÇEK ÇEKİLEN ÖLÇÜ DETAYLI)
# ============================================================================
ws1 = wb.create_sheet(title="Haftalık Planlama Master")
ws1.views.sheetView[0].showGridLines = True

ws1.merge_cells("A1:R1")
ws1["A1"] = "ERGUNBAS GROUP — HAFTALIK PLANLAMA & FİİLİ ÇEKİLEN ÖLÇÜ KAPASİTE TABLOSU"
ws1["A1"].font = font_title
ws1["A1"].fill = fill_title
ws1["A1"].alignment = align_left

ws1.merge_cells("A2:R2")
ws1["A2"] = f"Ağustos, Eylül ve Ekim 2026 fiili kantar tartım kayıtları ({len(all_weeks)} hafta / 1.532 vardiya) — Gerçek kesim boyları ve haftalık planlama kapasiteleri"
ws1["A2"].font = font_subtitle
ws1["A2"].fill = fill_title
ws1["A2"].alignment = align_left

headers_s1 = [
    "Ürün Adı",
    "Kategori",
    "Fiilen En Çok Çekilen Ana Ölçü",
    "Çekilen Diğer Ölçü Varyasyonları",
    "1 Boy Net Ağırlığı (kg)",
    "Nominal Katalog (kg)",
    "Geçmiş Toplam Net (kg)",
    "Geçmiş Toplam (Boy/Adet)",
    "Aktif Hafta",
    "Haftalık Ort. KG",
    "Haftalık Ort. Boy",
    "Haftalık Max KG",
    "Haftalık Min KG",
    "Önerilen Güvenli Plan (kg/hft)",
    "Önerilen Güvenli Boy (boy/hft)",
    "Ort. Fire (%)",
    "Hammadde Brüt İhtiyaç (kg)",
    "Öncelikli Hat & Kalıp Çıkışı"
]

row_num = 4
for c_idx, h in enumerate(headers_s1, 1):
    cell = ws1.cell(row_num, c_idx, h)
    cell.font = font_header
    cell.fill = fill_header_green if c_idx in [3, 5, 10, 11, 14, 15] else fill_header
    cell.alignment = align_center
    cell.border = border_header

data_start_row = 5
for idx, (pname, pstat) in enumerate(sorted_prods):
    r_idx = data_start_row + idx
    cfg = PRODUCT_CONFIG.get(pname, {})
    tot_kg = pstat['prod_kg']
    tot_qty = pstat['qty']
    actual_unit_kg = (tot_kg / tot_qty) if tot_qty > 0 else cfg.get('nominal_kg', 1.0)
    nominal_kg = cfg.get('nominal_kg', actual_unit_kg)

    dim_info = prod_dim_summary.get(pname, {'top_label': '2.20 m (220 cm)', 'other_label': '-'})
    
    w_kgs = [kg for w, kg in pstat['weekly_kg'].items() if kg > 0]
    active_weeks_count = len(w_kgs)
    avg_w_kg = tot_kg / active_weeks_count if active_weeks_count > 0 else 0.0
    avg_w_qty = tot_qty / active_weeks_count if active_weeks_count > 0 else 0.0
    max_w_kg = max(w_kgs) if w_kgs else 0.0
    min_w_kg = min(w_kgs) if w_kgs else 0.0

    safe_w_kg = round(avg_w_kg * 0.90, 1)
    safe_w_qty = round(safe_w_kg / actual_unit_kg) if actual_unit_kg > 0 else 0
    fire_rate = cfg.get('fire', 0.20)
    gross_w_kg = round(safe_w_kg / (1.0 - fire_rate), 1) if fire_rate < 1.0 else safe_w_kg
    line_info = f"{cfg.get('hat', 'Hat 101')} ({cfg.get('heads', '1 Göz')})"

    row_vals = [
        pname,
        cfg.get('cat', 'Profil'),
        dim_info['top_label'],
        dim_info['other_label'],
        round(actual_unit_kg, 2),
        nominal_kg,
        round(tot_kg, 1),
        round(tot_qty),
        f"{active_weeks_count} / {len(all_weeks)} hft",
        round(avg_w_kg, 1),
        round(avg_w_qty),
        round(max_w_kg, 1),
        round(min_w_kg, 1),
        safe_w_kg,
        safe_w_qty,
        fire_rate,
        gross_w_kg,
        line_info
    ]

    fill_row = fill_zebra if idx % 2 == 1 else PatternFill(fill_type=None)

    for c_idx, val in enumerate(row_vals, 1):
        cell = ws1.cell(r_idx, c_idx, val)
        cell.font = font_bold if c_idx in [1, 3, 5, 14, 15] else font_data
        cell.border = border_cell
        cell.fill = fill_highlight if c_idx in [3, 5, 14, 15] else fill_row

        if c_idx in [1, 2, 3, 4, 18]:
            cell.alignment = align_left
        elif c_idx in [9]:
            cell.alignment = align_center
        else:
            cell.alignment = align_right

        if c_idx in [5, 6]:
            cell.number_format = '#,##0.00 "kg"'
        elif c_idx in [7, 10, 12, 13, 14, 17]:
            cell.number_format = '#,##0.0 "kg"'
        elif c_idx in [8, 11, 15]:
            cell.number_format = '#,##0 "boy"' if 'Levha' not in pname else '#,##0 "plk"'
        elif c_idx == 16:
            cell.number_format = '0.0%'

tot_row_idx = data_start_row + len(sorted_prods)
ws1.cell(tot_row_idx, 1, "FABRİKA GENEL TOPLAMI / HAFTALIK ORTALAMA").font = font_total
ws1.cell(tot_row_idx, 1).fill = fill_total
ws1.cell(tot_row_idx, 1).border = border_total

for c in range(2, 19):
    cell = ws1.cell(tot_row_idx, c)
    cell.fill = fill_total
    cell.border = border_total
    cell.font = font_total

ws1.cell(tot_row_idx, 7, f"=SUM(G{data_start_row}:G{tot_row_idx-1})").number_format = '#,##0.0 "kg"'
ws1.cell(tot_row_idx, 8, f"=SUM(H{data_start_row}:H{tot_row_idx-1})").number_format = '#,##0 "adet"'
ws1.cell(tot_row_idx, 10, f"=SUM(J{data_start_row}:J{tot_row_idx-1})").number_format = '#,##0.0 "kg"'
ws1.cell(tot_row_idx, 11, f"=SUM(K{data_start_row}:K{tot_row_idx-1})").number_format = '#,##0 "adet"'
ws1.cell(tot_row_idx, 14, f"=SUM(N{data_start_row}:N{tot_row_idx-1})").number_format = '#,##0.0 "kg"'
ws1.cell(tot_row_idx, 15, f"=SUM(O{data_start_row}:O{tot_row_idx-1})").number_format = '#,##0 "adet"'
ws1.cell(tot_row_idx, 17, f"=SUM(Q{data_start_row}:Q{tot_row_idx-1})").number_format = '#,##0.0 "kg"'

auto_fit_columns(ws1)
ws1.freeze_panes = "E5"

# ============================================================================
# SAYFA 2: ÖLÇÜ & BOY KESİM DETAYLARI (HER ÖLÇÜNÜN AYRI AYRI KG VE BOYU)
# ============================================================================
ws_detail = wb.create_sheet(title="Ölçü & Boy Kesim Detayları")
ws_detail.views.sheetView[0].showGridLines = True

ws_detail.merge_cells("A1:K1")
ws_detail["A1"] = "ERGUNBAS GROUP — FİİLEN ÇEKİLEN ÖLÇÜ BAZINDA ÜRETİM VE AĞIRLIK DETAYLARI"
ws_detail["A1"].font = font_title
ws_detail["A1"].fill = fill_title
ws_detail["A1"].alignment = align_left

ws_detail.merge_cells("A2:K2")
ws_detail["A2"] = "Sistem kayıtlarındaki tüm kesim boyları ve plaka ebatlarının (cm/metre) fiili gerçekleşme tonajları, adetleri ve 1 boy ağırlıkları"
ws_detail["A2"].font = font_subtitle
ws_detail["A2"].fill = fill_title
ws_detail["A2"].alignment = align_left

headers_detail = [
    "Ürün Adı",
    "Kategori",
    "Fiilen Çekilen Kesim Ölçüsü (cm / m)",
    "Toplam Çekilen Boy / Adet",
    "Toplam Net KG",
    "Bu Ölçünün 1 Boy Ağırlığı (kg)",
    "Ürün İçi Payı (%)",
    "Çekildiği Hat(lar)",
    "Aktif Hafta Sayısı",
    "Haftalık Ort. Boy",
    "Haftalık Ort. KG"
]

row_num = 4
for c_idx, h in enumerate(headers_detail, 1):
    cell = ws_detail.cell(row_num, c_idx, h)
    cell.font = font_header
    cell.fill = fill_header_green if c_idx in [3, 4, 5, 6] else fill_header
    cell.alignment = align_center
    cell.border = border_header

# Flatten all (product, dim) rows
flat_dim_rows = []
for pname, dim_map in prod_dim_stats.items():
    tot_p_kg = prod_totals[pname]['prod_kg']
    cfg = PRODUCT_CONFIG.get(pname, {})
    for dim_str, stat in dim_map.items():
        flat_dim_rows.append({
            'product': pname,
            'cat': cfg.get('cat', 'Profil'),
            'dim': dim_str,
            'qty': stat['qty'],
            'kg': stat['kg'],
            'unit_kg': (stat['kg'] / stat['qty']) if stat['qty'] > 0 else 0.0,
            'pct': (stat['kg'] / tot_p_kg) if tot_p_kg > 0 else 0.0,
            'hats': ", ".join(sorted(list(stat['hats']))),
            'weeks_count': len(stat['weeks']),
            'avg_qty': stat['qty'] / len(stat['weeks']) if stat['weeks'] else 0.0,
            'avg_kg': stat['kg'] / len(stat['weeks']) if stat['weeks'] else 0.0
        })

# Sort by product total kg desc, then by dim kg desc
flat_dim_rows.sort(key=lambda x: (prod_totals[x['product']]['prod_kg'], x['kg']), reverse=True)

data_start_row = 5
for idx, r in enumerate(flat_dim_rows):
    r_idx = data_start_row + idx
    
    row_vals = [
        r['product'],
        r['cat'],
        r['dim'],
        round(r['qty']),
        round(r['kg'], 1),
        round(r['unit_kg'], 2),
        r['pct'],
        r['hats'],
        f"{r['weeks_count']} hft",
        round(r['avg_qty']),
        round(r['avg_kg'], 1)
    ]

    fill_row = fill_zebra if idx % 2 == 1 else PatternFill(fill_type=None)

    for c_idx, val in enumerate(row_vals, 1):
        cell = ws_detail.cell(r_idx, c_idx, val)
        cell.font = font_bold if c_idx in [1, 3, 6] else font_data
        cell.border = border_cell
        cell.fill = fill_highlight if c_idx in [3, 6] else fill_row

        if c_idx in [1, 2, 3, 8]:
            cell.alignment = align_left
        elif c_idx == 9:
            cell.alignment = align_center
        else:
            cell.alignment = align_right

        if c_idx in [4, 10]:
            cell.number_format = '#,##0 "boy"' if 'Levha' not in r['product'] else '#,##0 "plk"'
        elif c_idx in [5, 11]:
            cell.number_format = '#,##0.0 "kg"'
        elif c_idx == 6:
            cell.number_format = '#,##0.00 "kg"'
        elif c_idx == 7:
            cell.number_format = '0.0%'

tot_row_idx = data_start_row + len(flat_dim_rows)
ws_detail.cell(tot_row_idx, 1, "GENEL TOPLAM").font = font_total
ws_detail.cell(tot_row_idx, 1).fill = fill_total
ws_detail.cell(tot_row_idx, 1).border = border_total

for c in range(2, 12):
    cell = ws_detail.cell(tot_row_idx, c)
    cell.fill = fill_total
    cell.border = border_total
    cell.font = font_total

ws_detail.cell(tot_row_idx, 4, f"=SUM(D{data_start_row}:D{tot_row_idx-1})").number_format = '#,##0 "adet"'
ws_detail.cell(tot_row_idx, 5, f"=SUM(E{data_start_row}:E{tot_row_idx-1})").number_format = '#,##0.0 "kg"'

auto_fit_columns(ws_detail)
ws_detail.freeze_panes = "D5"

# ============================================================================
# SAYFA 3: HAFTA HAFTA ÜRETİM (KG)
# ============================================================================
ws2 = wb.create_sheet(title="Hafta Hafta Üretim (KG)")
ws2.views.sheetView[0].showGridLines = True

ws2.merge_cells(f"A1:{get_column_letter(len(all_weeks)+5)}1")
ws2["A1"] = "ERGUNBAS GROUP — ÜRÜN BAZINDA HAFTA HAFTA FİİLİ ÜRETİM TONAJI (NET KG)"
ws2["A1"].font = font_title
ws2["A1"].fill = fill_title
ws2["A1"].alignment = align_left

ws2.merge_cells(f"A2:{get_column_letter(len(all_weeks)+5)}2")
ws2["A2"] = "W31'den W43'e kadar her hafta hangi üründen kaç kg net üretim çıktığının geçmiş trend analizi"
ws2["A2"].font = font_subtitle
ws2["A2"].fill = fill_title
ws2["A2"].alignment = align_left

headers_s2 = ["Ürün Adı", "Kategori", "Ana Kesim Ölçüsü", "1 Boy (kg)"] + [WEEK_LABELS.get(w, w) for w in all_weeks] + ["Toplam (kg)", "Haftalık Ort (kg)"]

row_num = 4
for c_idx, h in enumerate(headers_s2, 1):
    cell = ws2.cell(row_num, c_idx, h)
    cell.font = font_header
    cell.fill = fill_header_green if "Toplam" in h or "Ort" in h or "Ölçü" in h else fill_header
    cell.alignment = align_center
    cell.border = border_header

data_start_row = 5
for idx, (pname, pstat) in enumerate(sorted_prods):
    r_idx = data_start_row + idx
    cfg = PRODUCT_CONFIG.get(pname, {})
    tot_kg = pstat['prod_kg']
    tot_qty = pstat['qty']
    actual_unit_kg = (tot_kg / tot_qty) if tot_qty > 0 else cfg.get('nominal_kg', 1.0)
    dim_info = prod_dim_summary.get(pname, {'top_label': '-'})

    row_vals = [pname, cfg.get('cat', 'Profil'), dim_info['top_label'], round(actual_unit_kg, 2)]
    for w in all_weeks:
        row_vals.append(round(pstat['weekly_kg'].get(w, 0.0), 1))

    sum_col_start = get_column_letter(5)
    sum_col_end = get_column_letter(4 + len(all_weeks))

    fill_row = fill_zebra if idx % 2 == 1 else PatternFill(fill_type=None)

    for c_idx, val in enumerate(row_vals, 1):
        cell = ws2.cell(r_idx, c_idx, val)
        cell.font = font_bold if c_idx in [1, 3] else font_data
        cell.border = border_cell
        cell.fill = fill_row

        if c_idx in [1, 2, 3]:
            cell.alignment = align_left
        elif c_idx == 4:
            cell.alignment = align_right
            cell.number_format = '#,##0.00 "kg"'
        else:
            cell.alignment = align_right
            cell.number_format = '#,##0.0' if val > 0 else '-'

    tot_col_idx = len(all_weeks) + 5
    avg_col_idx = len(all_weeks) + 6

    c_tot = ws2.cell(r_idx, tot_col_idx, f"=SUM({sum_col_start}{r_idx}:{sum_col_end}{r_idx})")
    c_tot.font = font_bold
    c_tot.alignment = align_right
    c_tot.number_format = '#,##0.0 "kg"'
    c_tot.border = border_cell
    c_tot.fill = fill_highlight

    c_avg = ws2.cell(r_idx, avg_col_idx, f"=AVERAGEIF({sum_col_start}{r_idx}:{sum_col_end}{r_idx}, \">0\")")
    c_avg.font = font_bold
    c_avg.alignment = align_right
    c_avg.number_format = '#,##0.0 "kg"'
    c_avg.border = border_cell
    c_avg.fill = fill_highlight

tot_row_idx = data_start_row + len(sorted_prods)
ws2.cell(tot_row_idx, 1, "HAFTALIK TOPLAM FABRİKA ÜRETİMİ (KG)").font = font_total
ws2.cell(tot_row_idx, 1).fill = fill_total
ws2.cell(tot_row_idx, 1).border = border_total

for c in range(2, len(headers_s2) + 1):
    cell = ws2.cell(tot_row_idx, c)
    cell.fill = fill_total
    cell.border = border_total
    cell.font = font_total
    if c >= 5:
        col_let = get_column_letter(c)
        cell.value = f"=SUM({col_let}{data_start_row}:{col_let}{tot_row_idx-1})"
        cell.alignment = align_right
        cell.number_format = '#,##0.0 "kg"'

auto_fit_columns(ws2)
ws2.freeze_panes = "E5"

# ============================================================================
# SAYFA 4: HAFTA HAFTA ÜRETİM (BOY-ADET)
# ============================================================================
ws3 = wb.create_sheet(title="Hafta Hafta Üretim (Boy-Adet)")
ws3.views.sheetView[0].showGridLines = True

ws3.merge_cells(f"A1:{get_column_letter(len(all_weeks)+5)}1")
ws3["A1"] = "ERGUNBAS GROUP — ÜRÜN BAZINDA HAFTA HAFTA ÇEKİLEN BOY & ADET SAYILARI"
ws3["A1"].font = font_title
ws3["A1"].fill = fill_title
ws3["A1"].alignment = align_left

ws3.merge_cells(f"A2:{get_column_letter(len(all_weeks)+5)}2")
ws3["A2"] = "Kapı setleme için gerekli boy adetlerinin 13 haftalık gerçekleşme sayıları ve ana kesim ölçüleri"
ws3["A2"].font = font_subtitle
ws3["A2"].fill = fill_title
ws3["A2"].alignment = align_left

headers_s3 = ["Ürün Adı", "Kategori", "Ana Kesim Ölçüsü", "1 Boy (kg)"] + [WEEK_LABELS.get(w, w) for w in all_weeks] + ["Toplam Boy", "Haftalık Ort. Boy"]

row_num = 4
for c_idx, h in enumerate(headers_s3, 1):
    cell = ws3.cell(row_num, c_idx, h)
    cell.font = font_header
    cell.fill = fill_header_green if "Toplam" in h or "Ort" in h or "Ölçü" in h else fill_header
    cell.alignment = align_center
    cell.border = border_header

for idx, (pname, pstat) in enumerate(sorted_prods):
    r_idx = data_start_row + idx
    cfg = PRODUCT_CONFIG.get(pname, {})
    tot_kg = pstat['prod_kg']
    tot_qty = pstat['qty']
    actual_unit_kg = (tot_kg / tot_qty) if tot_qty > 0 else cfg.get('nominal_kg', 1.0)
    dim_info = prod_dim_summary.get(pname, {'top_label': '-'})

    row_vals = [pname, cfg.get('cat', 'Profil'), dim_info['top_label'], round(actual_unit_kg, 2)]
    for w in all_weeks:
        row_vals.append(round(pstat['weekly_qty'].get(w, 0.0)))

    sum_col_start = get_column_letter(5)
    sum_col_end = get_column_letter(4 + len(all_weeks))

    fill_row = fill_zebra if idx % 2 == 1 else PatternFill(fill_type=None)

    for c_idx, val in enumerate(row_vals, 1):
        cell = ws3.cell(r_idx, c_idx, val)
        cell.font = font_bold if c_idx in [1, 3] else font_data
        cell.border = border_cell
        cell.fill = fill_row

        if c_idx in [1, 2, 3]:
            cell.alignment = align_left
        elif c_idx == 4:
            cell.alignment = align_right
            cell.number_format = '#,##0.00 "kg"'
        else:
            cell.alignment = align_right
            cell.number_format = '#,##0 "boy"' if val > 0 else '-'

    tot_col_idx = len(all_weeks) + 5
    avg_col_idx = len(all_weeks) + 6

    c_tot = ws3.cell(r_idx, tot_col_idx, f"=SUM({sum_col_start}{r_idx}:{sum_col_end}{r_idx})")
    c_tot.font = font_bold
    c_tot.alignment = align_right
    c_tot.number_format = '#,##0 "boy"'
    c_tot.border = border_cell
    c_tot.fill = fill_highlight

    c_avg = ws3.cell(r_idx, avg_col_idx, f"=AVERAGEIF({sum_col_start}{r_idx}:{sum_col_end}{r_idx}, \">0\")")
    c_avg.font = font_bold
    c_avg.alignment = align_right
    c_avg.number_format = '#,##0 "boy"'
    c_avg.border = border_cell
    c_avg.fill = fill_highlight

tot_row_idx = data_start_row + len(sorted_prods)
ws3.cell(tot_row_idx, 1, "HAFTALIK TOPLAM ÇEKİLEN BOY / ADET").font = font_total
ws3.cell(tot_row_idx, 1).fill = fill_total
ws3.cell(tot_row_idx, 1).border = border_total

for c in range(2, len(headers_s3) + 1):
    cell = ws3.cell(tot_row_idx, c)
    cell.fill = fill_total
    cell.border = border_total
    cell.font = font_total
    if c >= 5:
        col_let = get_column_letter(c)
        cell.value = f"=SUM({col_let}{data_start_row}:{col_let}{tot_row_idx-1})"
        cell.alignment = align_right
        cell.number_format = '#,##0 "adet"'

auto_fit_columns(ws3)
ws3.freeze_panes = "E5"

# ============================================================================
# SAYFA 5: HAT BAZINDA HAFTALIK KAPASİTE
# ============================================================================
ws4 = wb.create_sheet(title="Hat Bazında Haftalık Kapasite")
ws4.views.sheetView[0].showGridLines = True

ws4.merge_cells("A1:K1")
ws4["A1"] = "ERGUNBAS GROUP — 10 ÜRETİM HATTININ HAFTALIK FİİLİ KAPASİTE VE YÜK ANALİZİ"
ws4["A1"].font = font_title
ws4["A1"].fill = fill_title
ws4["A1"].alignment = align_left

ws4.merge_cells("A2:K2")
ws4["A2"] = "Hat 101 - 109 ve Hat 201'in 13 haftalık gerçekleşen saatlik hızları, haftalık tonaj ve boy kapasiteleri"
ws4["A2"].font = font_subtitle
ws4["A2"].fill = fill_title
ws4["A2"].alignment = align_left

headers_s4 = [
    "Hat No",
    "Hat Fonksiyonu & Kalıp Tipi",
    "Çektiği Ana Ürünler",
    "Toplam Çalışma (Saat)",
    "Haftalık Ort. Saat",
    "Haftalık Ort. Vardiya",
    "Fiili Saatlik Hız (kg/sa)",
    "Toplam Net KG",
    "Haftalık Ort. KG",
    "Haftalık Ort. Boy",
    "Ort. Fire Oranı (%)"
]

row_num = 4
for c_idx, h in enumerate(headers_s4, 1):
    cell = ws4.cell(row_num, c_idx, h)
    cell.font = font_header
    cell.fill = fill_header_green if c_idx in [5, 6, 9, 10] else fill_header
    cell.alignment = align_center
    cell.border = border_header

def sort_hat_key(h):
    digits = ''.join(c for c in h if c.isdigit())
    return int(digits) if digits else 999

sorted_lines = sorted(line_totals.items(), key=lambda x: sort_hat_key(x[0]))

LINE_DESC = {
    'Hat 101': 'Pervaz & Profil Hattı (Tek Göz)',
    'Hat 102': 'Kasa & Çıta & U Kapama (Tek Göz)',
    'Hat 103': 'Pervaz Hattı (Çift Göz)',
    'Hat 104': 'Seren Hattı (Çift Göz Yüksek Hızlı)',
    'Hat 105': 'Seren Hattı (Çift Göz)',
    'Hat 106': '140 Kasa Ana Hattı (Tek Göz)',
    'Hat 107': '200 Kasa & 39mm Seren & Deck (Tek Göz)',
    'Hat 108': '100 & 140 Kasa Hattı (Tek Göz)',
    'Hat 109': '50x80 & 80x80 Pervaz Çift Göz (Yüksek Hız)',
    'Hat 201': 'WPC Levha Ekstrüder Makinesi'
}

for idx, (hname, hstat) in enumerate(sorted_lines):
    r_idx = data_start_row + idx
    tot_h = hstat['hours']
    tot_kg = hstat['prod_kg']
    tot_fire = hstat['fire_kg']
    tot_qty = hstat['qty']

    active_weeks_h = len([w for w, kg in hstat['weekly_kg'].items() if kg > 0]) or len(all_weeks)
    avg_h_w = tot_h / active_weeks_h if active_weeks_h > 0 else 0.0
    avg_shifts_w = avg_h_w / 12.0
    speed_kg_h = tot_kg / tot_h if tot_h > 0 else 0.0
    avg_w_kg = tot_kg / active_weeks_h if active_weeks_h > 0 else 0.0
    avg_w_qty = tot_qty / active_weeks_h if active_weeks_h > 0 else 0.0
    fire_pct = (tot_fire / (tot_kg + tot_fire)) if (tot_kg + tot_fire) > 0 else 0.0

    top_prods = ", ".join(list(hstat['products'])[:3])

    row_vals = [
        hname,
        LINE_DESC.get(hname, 'Ekstrüder Profil Hattı'),
        top_prods,
        round(tot_h, 1),
        round(avg_h_w, 1),
        round(avg_shifts_w, 1),
        round(speed_kg_h, 1),
        round(tot_kg, 1),
        round(avg_w_kg, 1),
        round(avg_w_qty),
        fire_pct
    ]

    fill_row = fill_zebra if idx % 2 == 1 else PatternFill(fill_type=None)

    for c_idx, val in enumerate(row_vals, 1):
        cell = ws4.cell(r_idx, c_idx, val)
        cell.font = font_bold if c_idx in [1, 9, 10] else font_data
        cell.border = border_cell
        cell.fill = fill_highlight if c_idx in [9, 10] else fill_row

        if c_idx in [1, 2, 3]:
            cell.alignment = align_left
        elif c_idx in [6]:
            cell.alignment = align_center
            cell.number_format = '0.0 "vrd"'
        else:
            cell.alignment = align_right

        if c_idx in [4, 5]:
            cell.number_format = '#,##0.0 "sa"'
        elif c_idx == 7:
            cell.number_format = '#,##0.0 "kg/sa"'
        elif c_idx in [8, 9]:
            cell.number_format = '#,##0.0 "kg"'
        elif c_idx == 10:
            cell.number_format = '#,##0 "boy"'
        elif c_idx == 11:
            cell.number_format = '0.0%'

tot_row_idx = data_start_row + len(sorted_lines)
ws4.cell(tot_row_idx, 1, "TOPLAM FABRİKA KAPASİTESİ (10 HAT)").font = font_total
ws4.cell(tot_row_idx, 1).fill = fill_total
ws4.cell(tot_row_idx, 1).border = border_total

for c in range(2, 12):
    cell = ws4.cell(tot_row_idx, c)
    cell.fill = fill_total
    cell.border = border_total
    cell.font = font_total

ws4.cell(tot_row_idx, 4, f"=SUM(D{data_start_row}:D{tot_row_idx-1})").number_format = '#,##0.0 "sa"'
ws4.cell(tot_row_idx, 5, f"=SUM(E{data_start_row}:E{tot_row_idx-1})").number_format = '#,##0.0 "sa"'
ws4.cell(tot_row_idx, 6, f"=SUM(F{data_start_row}:F{tot_row_idx-1})").number_format = '0.0 "vrd"'
ws4.cell(tot_row_idx, 7, f"=AVERAGE(G{data_start_row}:G{tot_row_idx-1})").number_format = '#,##0.0 "kg/sa"'
ws4.cell(tot_row_idx, 8, f"=SUM(H{data_start_row}:H{tot_row_idx-1})").number_format = '#,##0.0 "kg"'
ws4.cell(tot_row_idx, 9, f"=SUM(I{data_start_row}:I{tot_row_idx-1})").number_format = '#,##0.0 "kg"'
ws4.cell(tot_row_idx, 10, f"=SUM(J{data_start_row}:J{tot_row_idx-1})").number_format = '#,##0 "adet"'

auto_fit_columns(ws4)
ws4.freeze_panes = "D5"

# ============================================================================
# SAYFA 6: HAFTALIK PLANLAMA ŞABLONU (İNTERAKTİF)
# ============================================================================
ws5 = wb.create_sheet(title="Haftalık Planlama Şablonu")
ws5.views.sheetView[0].showGridLines = True

ws5.merge_cells("A1:M1")
ws5["A1"] = "ERGUNBAS GROUP — SIFIR RİSK HAFTALIK SİPARİŞ & ÜRETİM PLANLAMA ŞABLONU"
ws5["A1"].font = font_title
ws5["A1"].fill = fill_title
ws5["A1"].alignment = align_left

ws5.merge_cells("A2:M2")
ws5["A2"] = "Mavi sütuna ('Haftalık Hedef Boy') istediğiniz boy miktarını yazın; formüller net kg, brüt fireli kg, hat saati ve mikser şarjını otomatik hesaplar."
ws5["A2"].font = font_subtitle
ws5["A2"].fill = fill_title
ws5["A2"].alignment = align_left

headers_s5 = [
    "Ürün Adı",
    "Fiilen Çekilen Ana Ölçü",
    "1 Boy Ağırlığı (kg)",
    "Önerilen Hat & Göz",
    "Ort. Hız (boy/sa)",
    "Haftalık Hedef (BOY GİRİNİZ)",
    "Gerekli Net KG",
    "Fire Oranı (%)",
    "Fireli Brüt KG",
    "Gerekli Makine Saati",
    "Gerekli Vardiya Sayısı",
    "Mikser Şarj İhtiyacı (275kg)",
    "Haftalık Kapasite Kontrolü"
]

row_num = 4
for c_idx, h in enumerate(headers_s5, 1):
    cell = ws5.cell(row_num, c_idx, h)
    cell.font = font_header
    if c_idx == 6:
        cell.fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    elif c_idx in [7, 9, 10, 11, 12]:
        cell.fill = fill_header_green
    else:
        cell.fill = fill_header
    cell.alignment = align_center
    cell.border = border_header

sample_targets = {
    'SEREN': 6000,
    '50X80 PERVAZ': 9000,
    '140 mm KASA': 2500,
    '100 mm KASA': 2500,
    '200 mm KASA': 1200,
    '80X80 PERVAZ': 1200,
    'SEREN(39mm)': 2500,
    '35X100 PERVAZ': 1500,
    '35X150 PERVAZ': 1500,
    'CAM ÇITASI': 2500,
    'KT100-220x24-ANT': 500,
    'DIŞBUDAK Levha': 3500,
    'TEAK Levha': 1500,
    'BEYAZ TEAK Levha': 1000
}

for idx, (pname, pstat) in enumerate(sorted_prods):
    r_idx = data_start_row + idx
    cfg = PRODUCT_CONFIG.get(pname, {})
    tot_kg = pstat['prod_kg']
    tot_qty = pstat['qty']
    actual_unit_kg = (tot_kg / tot_qty) if tot_qty > 0 else cfg.get('nominal_kg', 1.0)
    dim_info = prod_dim_summary.get(pname, {'top_label': '-'})

    tot_h = pstat['hours']
    speed_boy_h = (tot_qty / tot_h) if tot_h > 0 else 25.0
    if speed_boy_h <= 0: speed_boy_h = 25.0

    def_target = sample_targets.get(pname, 0)
    fire_rate = cfg.get('fire', 0.20)

    ws5.cell(r_idx, 1, pname).font = font_bold
    ws5.cell(r_idx, 1).alignment = align_left
    ws5.cell(r_idx, 1).border = border_cell

    ws5.cell(r_idx, 2, dim_info['top_label']).alignment = align_left
    ws5.cell(r_idx, 2).border = border_cell

    c3 = ws5.cell(r_idx, 3, round(actual_unit_kg, 2))
    c3.number_format = '#,##0.00 "kg"'
    c3.alignment = align_right
    c3.border = border_cell

    ws5.cell(r_idx, 4, f"{cfg.get('hat', 'Hat 101')} ({cfg.get('heads', '1 Göz')})").alignment = align_left
    ws5.cell(r_idx, 4).border = border_cell

    c5 = ws5.cell(r_idx, 5, round(speed_boy_h, 1))
    c5.number_format = '#,##0.0'
    c5.alignment = align_right
    c5.border = border_cell

    c6 = ws5.cell(r_idx, 6, def_target)
    c6.font = Font(name="Segoe UI", size=10, bold=True, color="1E3A8A")
    c6.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    c6.alignment = align_right
    c6.number_format = '#,##0 "boy"'
    c6.border = Border(left=Side(style="medium", color="2563EB"), right=Side(style="medium", color="2563EB"), top=thin_side, bottom=thin_side)

    c7 = ws5.cell(r_idx, 7, f"=F{r_idx}*C{r_idx}")
    c7.number_format = '#,##0.0 "kg"'
    c7.alignment = align_right
    c7.border = border_cell

    c8 = ws5.cell(r_idx, 8, fire_rate)
    c8.number_format = '0.0%'
    c8.alignment = align_right
    c8.border = border_cell

    c9 = ws5.cell(r_idx, 9, f"=IF(H{r_idx}<1, G{r_idx}/(1-H{r_idx}), G{r_idx})")
    c9.number_format = '#,##0.0 "kg"'
    c9.font = font_bold
    c9.alignment = align_right
    c9.border = border_cell
    c9.fill = fill_highlight

    c10 = ws5.cell(r_idx, 10, f"=IF(E{r_idx}>0, F{r_idx}/E{r_idx}, 0)")
    c10.number_format = '#,##0.0 "sa"'
    c10.alignment = align_right
    c10.border = border_cell

    c11 = ws5.cell(r_idx, 11, f"=J{r_idx}/12")
    c11.number_format = '0.0 "vrd"'
    c11.alignment = align_right
    c11.border = border_cell

    c12 = ws5.cell(r_idx, 12, f"=ROUNDUP(I{r_idx}/275.65, 0)")
    c12.number_format = '#,##0 "şarj"'
    c12.alignment = align_right
    c12.border = border_cell
    c12.font = font_bold

    c13 = ws5.cell(r_idx, 13, f'=IF(K{r_idx}<=14, "🟢 Güvenli Plan", "🔴 Kapasite Aşımı!")')
    c13.alignment = align_center
    c13.font = font_bold
    c13.border = border_cell

tot_row_idx = data_start_row + len(sorted_prods)
ws5.cell(tot_row_idx, 1, "HAFTALIK PLANLANAN TOPLAM").font = font_total
ws5.cell(tot_row_idx, 1).fill = fill_total
ws5.cell(tot_row_idx, 1).border = border_total

for c in range(2, 14):
    cell = ws5.cell(tot_row_idx, c)
    cell.fill = fill_total
    cell.border = border_total
    cell.font = font_total

ws5.cell(tot_row_idx, 6, f"=SUM(F{data_start_row}:F{tot_row_idx-1})").number_format = '#,##0 "boy"'
ws5.cell(tot_row_idx, 7, f"=SUM(G{data_start_row}:G{tot_row_idx-1})").number_format = '#,##0.0 "kg"'
ws5.cell(tot_row_idx, 9, f"=SUM(I{data_start_row}:I{tot_row_idx-1})").number_format = '#,##0.0 "kg"'
ws5.cell(tot_row_idx, 10, f"=SUM(J{data_start_row}:J{tot_row_idx-1})").number_format = '#,##0.0 "sa"'
ws5.cell(tot_row_idx, 11, f"=SUM(K{data_start_row}:K{tot_row_idx-1})").number_format = '0.0 "vrd"'
ws5.cell(tot_row_idx, 12, f"=SUM(L{data_start_row}:L{tot_row_idx-1})").number_format = '#,##0 "şarj"'

auto_fit_columns(ws5)
ws5.freeze_panes = "G5"

# -------------------------------------------------------------
# KAYIT: 1) KULLANICIYA ÖZEL İNTERAKTİF PLANLAMA ARACI
#        2) YÖNETİCİYE GÖNDERİLECEK RESMİ KAPASİTE RAPORU
# -------------------------------------------------------------
desktop_ozel_path = r"C:\Users\llboz\OneDrive\Desktop\Ergunbas_Ozel_Haftalik_Planlama_Hesaplayici.xlsx"
desktop_yonetici_path = r"C:\Users\llboz\OneDrive\Desktop\Ergunbas_Haftalik_Uretim_ve_Kapasite_Yonetici_Raporu.xlsx"
desktop_main_path = r"C:\Users\llboz\OneDrive\Desktop\Ergunbas_Haftalik_Uretim_Planlama_ve_Kapasite.xlsx"
static_path = os.path.join(APP_DIR, "static", "Ergunbas_Haftalik_Uretim_Planlama_ve_Kapasite.xlsx")

# 1. Önce şablon dahil tüm sayfaları KULLANICIYA ÖZEL dosyaya kaydet
try:
    wb.save(desktop_ozel_path)
    print("1. Kullanıcıya Özel İnteraktif Dosya kaydedildi:", desktop_ozel_path)
except Exception as e:
    print("Özel dosya kayıt hatası:", e)

# 2. Şimdi 'Haftalık Planlama Şablonu' sayfasını çalışma kitabından kaldırarak YÖNETİCİ RAPORU oluştur
if "Haftalık Planlama Şablonu" in wb.sheetnames:
    wb.remove(wb["Haftalık Planlama Şablonu"])
    print("Yönetici raporundan 'Haftalık Planlama Şablonu' sayfası kaldırıldı.")

# Yönetici dosyasını kaydet
try:
    wb.save(desktop_yonetici_path)
    print("2. Yöneticiye Gönderilecek Resmi Rapor kaydedildi:", desktop_yonetici_path)
except Exception as e:
    print("Yönetici raporu kayıt hatası:", e)

try:
    wb.save(desktop_main_path)
    print("3. Ana masaüstü dosyası güncellendi:", desktop_main_path)
except Exception as e:
    print("Ana masaüstü açık, pas geçildi (Yönetici raporu adıyla hazır):", e)

try:
    wb.save(static_path)
    print("4. Static klasörüne yönetici formatı kaydedildi:", static_path)
except Exception as e:
    print("Static kayıt hatası:", e)

print("YÖNETİCİ RAPORU VE ÖZEL PLANLAMA DOSYASI BAŞARIYLA AYRILDI!")
