import sys
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)
import app_backend

PRODUCT_DIMENSIONS = {
    # Pervazlar (Standart Boy: 2.20 Metre)
    "50X80 PERVAZ": {"boy": "2.20 Metre", "kesit": "50 x 80 mm", "unit_name": "Boy", "desc": "50x80 Oval/Düz Kulaklı Pervaz Profili"},
    "80X80 PERVAZ": {"boy": "2.20 Metre", "kesit": "80 x 80 mm", "unit_name": "Boy", "desc": "80x80 Geniş Kulaklı Pervaz Profili"},
    "35X100 PERVAZ": {"boy": "2.20 Metre", "kesit": "35 x 100 mm", "unit_name": "Boy", "desc": "35x100 Düz Pervaz Profili"},
    "35X150 PERVAZ": {"boy": "2.20 Metre", "kesit": "35 x 150 mm", "unit_name": "Boy", "desc": "35x150 Ekstra Geniş Pervaz"},
    "35X80 PERVAZ": {"boy": "2.20 Metre", "kesit": "35 x 80 mm", "unit_name": "Boy", "desc": "35x80 Standart Pervaz Profili"},
    
    # Kasalar (Standart Ebatlama / Kesme Boyu: 2.15 Metre - 2150 mm)
    "100 mm KASA": {"boy": "2.15 Metre (215 cm)", "kesit": "100 x 36 mm", "unit_name": "Boy", "desc": "100x36 mm Düz WPC Kapı Kasası (Standart Boy: 2.15 m)"},
    "140 mm KASA": {"boy": "2.15 Metre (215 cm)", "kesit": "140 x 36 mm", "unit_name": "Boy", "desc": "140x36 mm Düz WPC Kapı Kasası (Standart Boy: 2.15 m)"},
    "200 mm KASA": {"boy": "2.15 Metre (215 cm)", "kesit": "200 x 36 mm", "unit_name": "Boy", "desc": "200x36 mm Geniş WPC Kapı Kasası (Standart Boy: 2.15 m)"},

    # Serenler (Standart Boy: 2.10 Metre)
    "SEREN": {"boy": "2.10 Metre", "kesit": "38 x 45 mm", "unit_name": "Boy", "desc": "38x45 Standart Masif WPC Kapı Sereni"},
    "SEREN(39mm)": {"boy": "2.10 Metre", "kesit": "39 x 45 mm", "unit_name": "Boy", "desc": "39x45 Özel Kalın Masif Kapı Sereni"},

    # Levhalar (Standart Boy: 207 cm / 2.07 Metre, En: 94 cm)
    "DİŞBUDAK Levha": {"boy": "207 cm (2.07 m)", "kesit": "En: 94 cm (2 mm)", "unit_name": "Plaka", "desc": "94x207 cm Desenli WPC Kapı Yüzey Levhası"},
    "DIŞBUDAK Levha": {"boy": "207 cm (2.07 m)", "kesit": "En: 94 cm (2 mm)", "unit_name": "Plaka", "desc": "94x207 cm Desenli WPC Kapı Yüzey Levhası"},
    "TEAK Levha": {"boy": "207 cm (2.07 m)", "kesit": "En: 94 cm (2 mm)", "unit_name": "Plaka", "desc": "94x207 cm Teak Desenli Kapı Levhası"},
    "BEYAZ TEAK Levha": {"boy": "207 cm (2.07 m)", "kesit": "En: 94 cm (2 mm)", "unit_name": "Plaka", "desc": "94x207 cm Beyaz Teak Kapı Levhası"},
    "VIZON Levha": {"boy": "207 cm (2.07 m)", "kesit": "En: 94 cm (2 mm)", "unit_name": "Plaka", "desc": "94x207 cm Vizon Renkli Kapı Levhası"},
    "AY GRİ Levha": {"boy": "207 cm (2.07 m)", "kesit": "En: 94 cm (2 mm)", "unit_name": "Plaka", "desc": "94x207 cm Ay Gri Kapı Levhası"},
    "ANTRASİT Levha": {"boy": "207 cm (2.07 m)", "kesit": "En: 94 cm (2 mm)", "unit_name": "Plaka", "desc": "94x207 cm Antrasit Kapı Levhası"},
    "SONOMO Levha": {"boy": "207 cm (2.07 m)", "kesit": "En: 94 cm (2 mm)", "unit_name": "Plaka", "desc": "94x207 cm Sonomo Meşe Kapı Levhası"},
    "MIX TEAK Levha": {"boy": "207 cm (2.07 m)", "kesit": "En: 94 cm (2 mm)", "unit_name": "Plaka", "desc": "94x207 cm Mix Teak Kapı Levhası"},
    "MIX BEYAZ TEAK Levha": {"boy": "207 cm (2.07 m)", "kesit": "En: 94 cm (2 mm)", "unit_name": "Plaka", "desc": "94x207 cm Mix Beyaz Teak Kapı Levhası"},
    "KOYU CEVİZ Levha": {"boy": "207 cm (2.07 m)", "kesit": "En: 94 cm (2 mm)", "unit_name": "Plaka", "desc": "94x207 cm Koyu Ceviz Kapı Levhası"},
    "COCO Levha": {"boy": "207 cm (2.07 m)", "kesit": "En: 94 cm (2 mm)", "unit_name": "Plaka", "desc": "94x207 cm Coco Kahve Kapı Levhası"},

    # Deck / WPC Zemin (Standart Boy: 3.00 Metre)
    "KT100-220x24-ANT": {"boy": "3.00 Metre", "kesit": "220 x 24 mm", "unit_name": "Boy", "desc": "220x24 mm Antrasit WPC Zemin Deck Profili"},
    "KT102-219x26-ANT": {"boy": "3.00 Metre", "kesit": "219 x 26 mm", "unit_name": "Boy", "desc": "219x26 mm Antrasit Ağır Hizmet Deck Profili"},
    "KT102-219x26-KAH": {"boy": "3.00 Metre", "kesit": "219 x 26 mm", "unit_name": "Boy", "desc": "219x26 mm Kahverengi Ağır Hizmet Deck"},

    # Yardımcı Profiller (Standart Boy: 2.20 Metre)
    "CAM ÇITASI": {"boy": "2.20 Metre", "kesit": "15 x 20 mm", "unit_name": "Boy", "desc": "WPC Fitilli Cam Tutucu Çıta Profili"},
    "U KAPAMA": {"boy": "2.20 Metre", "kesit": "30 x 15 mm", "unit_name": "Boy", "desc": "WPC Duvar / Kasa U Kapama Profili"}
}

# Fabrika Resmi Parametreleri (erp extrüder.xlsx / PARAMETRELER Tablosu)
STD_SPEEDS = {
    ("Hat 101", "35X100 PERVAZ"): 46.8,
    ("Hat 101", "35X150 PERVAZ"): 41.2,
    ("Hat 101", "35X80 PERVAZ"): 44.2,
    ("Hat 101", "80X80 PERVAZ"): 46.5,
    ("Hat 102", "100 mm KASA"): 63.3,
    ("Hat 102", "140 mm KASA"): 51.8,
    ("Hat 102", "CAM ÇITASI"): 50.0,
    ("Hat 102", "U KAPAMA"): 40.7,
    ("Hat 103", "50X80 PERVAZ"): 53.1,
    ("Hat 103", "SEREN"): 63.3,
    ("Hat 104", "SEREN"): 99.2,
    ("Hat 105", "140 mm KASA"): 59.7,
    ("Hat 105", "SEREN"): 92.2,
    ("Hat 106", "100 mm KASA"): 70.1,
    ("Hat 106", "140 mm KASA"): 74.1,
    ("Hat 106", "200 mm KASA"): 100.9,
    ("Hat 107", "100 mm KASA"): 96.1,
    ("Hat 107", "140 mm KASA"): 45.7,
    ("Hat 107", "200 mm KASA"): 84.6,
    ("Hat 107", "KT100-220x24-ANT"): 84.7,
    ("Hat 107", "KT102-219x26-ANT"): 108.2,
    ("Hat 107", "KT102-219x26-KAH"): 115.8,
    ("Hat 107", "SEREN"): 82.6,
    ("Hat 107", "SEREN(39mm)"): 96.1,
    ("Hat 108", "100 mm KASA"): 69.5,
    ("Hat 108", "140 mm KASA"): 64.9,
    ("Hat 108", "200 mm KASA"): 78.5,
    ("Hat 109", "50X80 PERVAZ"): 73.6,
    ("Hat 109", "80X80 PERVAZ"): 85.1,
    ("Hat 201", "DIŞBUDAK Levha"): 187.9,
    ("Hat 201", "DİŞBUDAK Levha"): 187.9,
    ("Hat 201", "TEAK Levha"): 145.5,
    ("Hat 201", "BEYAZ TEAK Levha"): 148.4,
    ("Hat 201", "MIX TEAK Levha"): 131.5,
    ("Hat 201", "VIZON Levha"): 116.2,
    ("Hat 201", "SONOMO Levha"): 113.1,
    ("Hat 201", "AY GRİ Levha"): 144.9,
    ("Hat 201", "ANTRASİT Levha"): 111.8,
    ("Hat 201", "MIX BEYAZ TEAK Levha"): 118.6,
    ("Hat 201", "KOYU CEVİZ Levha"): 79.2,
    ("Hat 201", "COCO Levha"): 273.6,
}

def create_product_rotas_excel_wb():
    data = app_backend.load_data()
    daily = data.get("daily_data", {})
    prod_catalog = app_backend.PRODUCT_CATALOG
    hat_caps = app_backend.HAT_CAPABILITIES
    machines = data.get("machines", [])

    # 1. 70 Günlük Fiili Üretim Verilerini Çıkar (Gerçek Tartı Kiloları ile)
    actual_runs = {}
    machine_totals = {}

    for m in machines:
        m_name = m.get("name")
        norm_m = f"Hat {m_name}" if m_name.isdigit() else m_name
        machine_totals[norm_m] = {"total_kg": 0.0, "shifts": 0, "hours": 0.0, "products": set(), "type": m.get("type")}

    for day_key, day in daily.items():
        for sh_key in ["gunduz", "gece"]:
            sh = day.get(sh_key, {})
            # Extrüder
            for ext in sh.get("extruders", []):
                p = ext.get("product", "").strip()
                h_raw = str(ext.get("hat", "")).strip()
                qty = float(ext.get("qty", 0) or 0)
                pkg = float(ext.get("prod_kg", 0) or 0)
                hrs = float(ext.get("hours", 0) or 0)
                hrs = hrs if hrs > 0 else 12.0
                if not p or (qty <= 0 and pkg <= 0):
                    continue
                
                u_kg = 1.0
                matched_cat = prod_catalog.get(p)
                if not matched_cat:
                    for cp, info in prod_catalog.items():
                        if app_backend.match_product_daily(cp, "", p):
                            matched_cat = info
                            break
                if matched_cat:
                    u_kg = matched_cat.get("unit_kg", 1.0)

                # Gerçek kantar / tartı kilosu varsa onu al
                kg = pkg if pkg > 0 else (qty * u_kg)

                hat_name = f"Hat {h_raw}" if h_raw.isdigit() else h_raw
                if hat_name not in machine_totals:
                    machine_totals[hat_name] = {"total_kg": 0.0, "shifts": 0, "hours": 0.0, "products": set(), "type": "extruder"}
                
                machine_totals[hat_name]["total_kg"] += kg
                machine_totals[hat_name]["shifts"] += 1
                machine_totals[hat_name]["hours"] += hrs
                machine_totals[hat_name]["products"].add(p)

                actual_runs.setdefault(p, {}).setdefault(hat_name, {"shifts": 0, "hours": 0.0, "qty": 0.0, "kg": 0.0})
                actual_runs[p][hat_name]["shifts"] += 1
                actual_runs[p][hat_name]["hours"] += hrs
                actual_runs[p][hat_name]["qty"] += qty
                actual_runs[p][hat_name]["kg"] += kg

            # Levha
            for lev in sh.get("levha", []):
                c = lev.get("color", "").strip()
                qty = float(lev.get("qty", 0) or 0)
                hrs = float(lev.get("hours", 0) or 0)
                hrs = hrs if hrs > 0 else 12.0
                if not c or qty <= 0:
                    continue
                if "BUDAK" in c.upper():
                    c = "DİŞBUDAK"
                p = f"{c} Levha" if not c.lower().endswith("levha") else c
                h_raw = str(lev.get("hat", "201")).strip()

                u_kg = 4.77
                matched_cat = prod_catalog.get(p)
                if not matched_cat:
                    for cp, info in prod_catalog.items():
                        if app_backend.match_product_daily(cp, "", p):
                            matched_cat = info
                            break
                if matched_cat:
                    u_kg = matched_cat.get("unit_kg", 4.77)
                kg = qty * u_kg

                hat_name = f"Hat {h_raw}" if h_raw.isdigit() else h_raw
                if hat_name not in machine_totals:
                    machine_totals[hat_name] = {"total_kg": 0.0, "shifts": 0, "hours": 0.0, "products": set(), "type": "levha"}

                machine_totals[hat_name]["total_kg"] += kg
                machine_totals[hat_name]["shifts"] += 1
                machine_totals[hat_name]["hours"] += hrs
                machine_totals[hat_name]["products"].add(p)

                actual_runs.setdefault(p, {}).setdefault(hat_name, {"shifts": 0, "hours": 0.0, "qty": 0.0, "kg": 0.0})
                actual_runs[p][hat_name]["shifts"] += 1
                actual_runs[p][hat_name]["hours"] += hrs
                actual_runs[p][hat_name]["qty"] += qty
                actual_runs[p][hat_name]["kg"] += kg

    all_products_set = set(prod_catalog.keys()) | set(hat_caps.keys()) | set(actual_runs.keys())

    def get_product_category(p_name):
        p_up = p_name.upper()
        if "PERVAZ" in p_up: return "Pervaz Profili"
        if "KASA" in p_up: return "Kasa Profili"
        if "SEREN" in p_up: return "Seren Çıtası"
        if "LEVHA" in p_up: return "WPC Levha Plaka"
        if "KT100" in p_up or "KT102" in p_up or "DECK" in p_up: return "Deck / WPC Zemin"
        if "ÇITA" in p_up or "CITA" in p_up: return "Cam Çıtası"
        if "KAPAMA" in p_up: return "U Kapama / Lambiri"
        return "Özel Profil"

    def get_standard_capacity(hat_name, p_name, fallback=70.0):
        # 1. Doğrudan fabrika ERP parametre tablosundan eşle
        val = STD_SPEEDS.get((hat_name, p_name))
        if val is not None:
            return val
        for (h_k, p_k), speed in STD_SPEEDS.items():
            if h_k == hat_name and app_backend.match_product_daily(p_k, "", p_name):
                return speed
        # 2. Kalıp kataloğundan eşle
        caps = hat_caps.get(p_name, [])
        for c in caps:
            if f"Hat {c.get('hat')}" == hat_name:
                return float(c.get("speed_kg_h", fallback))
        return fallback

    product_rows = []

    for p in sorted(all_products_set):
        cat_info = prod_catalog.get(p)
        if not cat_info:
            for cp, info in prod_catalog.items():
                if app_backend.match_product_daily(cp, "", p):
                    cat_info = info
                    break
        unit_kg = cat_info.get("unit_kg", 1.0) if cat_info else 1.0
        category = get_product_category(p)

        dim_info = PRODUCT_DIMENSIONS.get(p)
        if not dim_info:
            for kp, kv in PRODUCT_DIMENSIONS.items():
                if app_backend.match_product_daily(kp, "", p):
                    dim_info = kv
                    break
        boy_str = dim_info["boy"] if dim_info else "2.20 m"
        kesit_str = dim_info["kesit"] if dim_info else "Standart Profil"
        unit_label = dim_info["unit_name"] if dim_info else ("Plaka" if "Levha" in category else "Boy")

        caps = hat_caps.get(p)
        if not caps:
            for cp, c_list in hat_caps.items():
                if app_backend.match_product_daily(cp, "", p):
                    caps = c_list
                    break
        if not caps:
            caps = []

        act_h = actual_runs.get(p, {})

        ana_hat_str = ""
        ana_hours = 0.0
        ana_shifts = 0
        ana_qty = 0.0
        ana_kg = 0.0
        ana_act_speed = 0.0
        ana_tek_speed = 0.0
        ana_setup = 45
        ana_mold = "—"

        ydk1_hat_str = ""
        ydk1_hours = 0.0
        ydk1_shifts = 0
        ydk1_qty = 0.0
        ydk1_kg = 0.0
        ydk1_act_speed = 0.0
        ydk1_tek_speed = 0.0

        ydk2_hat_str = ""
        ydk2_hours = 0.0
        ydk2_act_speed = 0.0
        ydk2_tek_speed = 0.0

        sorted_acts = sorted(act_h.items(), key=lambda x: -x[1]["kg"]) if act_h else []

        if sorted_acts:
            top_h, top_v = sorted_acts[0]
            ana_hat_str = top_h
            ana_hours = top_v["hours"]
            ana_shifts = top_v["shifts"]
            ana_qty = top_v["qty"]
            ana_kg = top_v["kg"]
            ana_act_speed = round(ana_kg / ana_hours, 1) if ana_hours > 0 else 0.0
            ana_tek_speed = get_standard_capacity(ana_hat_str, p, cat_info.get("speed_kg_h", 70.0) if cat_info else 70.0)

            top_cap = next((c for c in caps if f"Hat {c.get('hat')}" == top_h), None)
            if top_cap:
                ana_setup = top_cap.get("default_setup", 45)
                ana_mold = top_cap.get("mold_code", "—")
            else:
                ana_setup = 45
                ana_mold = "—"

            # Yedek Rota 1
            if len(sorted_acts) > 1:
                ydk1_h, ydk1_v = sorted_acts[1]
                ydk1_hat_str = ydk1_h
                ydk1_hours = ydk1_v["hours"]
                ydk1_shifts = ydk1_v["shifts"]
                ydk1_qty = ydk1_v["qty"]
                ydk1_kg = ydk1_v["kg"]
                ydk1_act_speed = round(ydk1_kg / ydk1_hours, 1) if ydk1_hours > 0 else 0.0
                ydk1_tek_speed = get_standard_capacity(ydk1_hat_str, p, ana_tek_speed * 0.9)
            elif len(caps) > 1:
                c1_candidates = [c for c in caps if f"Hat {c.get('hat')}" != top_h]
                if c1_candidates:
                    c1 = c1_candidates[0]
                    ydk1_hat_str = f"Hat {c1['hat']}"
                    ydk1_tek_speed = get_standard_capacity(ydk1_hat_str, p, float(c1.get("speed_kg_h", 0.0)))

            # Yedek Rota 2
            if len(sorted_acts) > 2:
                ydk2_h, ydk2_v = sorted_acts[2]
                ydk2_hat_str = ydk2_h
                ydk2_hours = ydk2_v["hours"]
                ydk2_act_speed = round(ydk2_v["kg"] / ydk2_hours, 1) if ydk2_hours > 0 else 0.0
                ydk2_tek_speed = get_standard_capacity(ydk2_hat_str, p, 0.0)
            elif len(caps) > 2:
                remain_caps = [c for c in caps if f"Hat {c.get('hat')}" not in [ana_hat_str, ydk1_hat_str]]
                if remain_caps:
                    c2 = remain_caps[0]
                    ydk2_hat_str = f"Hat {c2['hat']}"
                    ydk2_tek_speed = get_standard_capacity(ydk2_hat_str, p, float(c2.get("speed_kg_h", 0.0)))
        elif caps:
            c0 = caps[0]
            ana_hat_str = f"Hat {c0['hat']}"
            ana_tek_speed = get_standard_capacity(ana_hat_str, p, float(c0.get("speed_kg_h", 70.0)))
            ana_setup = c0.get("default_setup", 45)
            ana_mold = c0.get("mold_code", "—")
            if len(caps) > 1:
                c1 = caps[1]
                ydk1_hat_str = f"Hat {c1['hat']}"
                ydk1_tek_speed = get_standard_capacity(ydk1_hat_str, p, float(c1.get("speed_kg_h", 60.0)))

        # Hızları boy/ölçü olarak hesapla
        ana_act_qty_speed = round(ana_act_speed / unit_kg, 1) if unit_kg > 0 and ana_act_speed > 0 else 0.0
        ana_tek_qty_speed = round(ana_tek_speed / unit_kg, 1) if unit_kg > 0 and ana_tek_speed > 0 else 0.0
        ydk1_act_qty_speed = round(ydk1_act_speed / unit_kg, 1) if unit_kg > 0 and ydk1_act_speed > 0 else 0.0
        ydk1_tek_qty_speed = round(ydk1_tek_speed / unit_kg, 1) if unit_kg > 0 and ydk1_tek_speed > 0 else 0.0

        ana_act_speed_text = f"Saatte ~{int(round(ana_act_qty_speed))} {unit_label}" if ana_act_qty_speed > 0 else "—"
        ana_tek_speed_text = f"Saatte ~{int(round(ana_tek_qty_speed))} {unit_label}" if ana_tek_qty_speed > 0 else "—"
        ydk1_act_speed_text = f"Saatte ~{int(round(ydk1_act_qty_speed))} {unit_label}" if ydk1_act_qty_speed > 0 else "—"
        ydk1_tek_speed_text = f"Saatte ~{int(round(ydk1_tek_qty_speed))} {unit_label}" if ydk1_tek_qty_speed > 0 else "—"

        ana_daily_act = round(ana_act_speed * 24.0, 1) if ana_act_speed > 0 else 0.0
        ana_daily_tek = round(ana_tek_speed * 24.0, 1) if ana_tek_speed > 0 else 0.0
        ana_eff_pct = round((ana_act_speed / ana_tek_speed * 100), 1) if ana_tek_speed > 0 and ana_act_speed > 0 else (100.0 if ana_act_speed > 0 else 0.0)

        strategy = ""
        if ana_hours > 0 and ydk1_hours > 0:
            strategy = "Aktif Çift Hatlı Üretim (Yüksek Hacimli)"
        elif ana_hours > 0 and ydk1_hat_str:
            strategy = "Ana Hat Devrede (Yedek Hat Hazır)"
        elif not ydk1_hat_str and "Levha" in category:
            strategy = "Levha Özel Hattı (Hat 201)"
        elif not ydk1_hat_str:
            strategy = "Tek Hat Kalıp Atamalı"
        else:
            strategy = "Standart Planlama"

        product_rows.append({
            "product": p,
            "category": category,
            "boy": boy_str,
            "kesit": kesit_str,
            "unit_label": unit_label,
            "unit_kg": unit_kg,
            "total_kg": sum(v["kg"] for v in act_h.values()),
            "ana_hat": ana_hat_str,
            "ana_hours": ana_hours,
            "ana_shifts": ana_shifts,
            "ana_qty": ana_qty,
            "ana_kg": ana_kg,
            "ana_act_speed": ana_act_speed,
            "ana_act_speed_text": ana_act_speed_text,
            "ana_tek_speed": ana_tek_speed,
            "ana_tek_speed_text": ana_tek_speed_text,
            "ana_daily_act": ana_daily_act,
            "ana_daily_tek": ana_daily_tek,
            "ana_eff_pct": ana_eff_pct,
            "ana_setup": ana_setup,
            "ana_mold": ana_mold,
            "ydk1_hat": ydk1_hat_str,
            "ydk1_hours": ydk1_hours,
            "ydk1_shifts": ydk1_shifts,
            "ydk1_act_speed": ydk1_act_speed,
            "ydk1_act_speed_text": ydk1_act_speed_text,
            "ydk1_tek_speed": ydk1_tek_speed,
            "ydk1_tek_speed_text": ydk1_tek_speed_text,
            "ydk2_hat": ydk2_hat_str,
            "ydk2_hours": ydk2_hours,
            "strategy": strategy
        })

    # Sadece 70 günlük fiili üretimde gerçek kaydı olan ürünleri al (0 kg olan mükerrer ve hayalet satırları temizle)
    product_rows = [r for r in product_rows if r["total_kg"] > 0]
    product_rows.sort(key=lambda x: -x["total_kg"])

    # 3. OpenPyXL Workbook Oluştur
    wb = openpyxl.Workbook()
    
    font_title = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    font_subtitle = Font(name="Segoe UI", size=9.5, italic=True, color="E2E8F0")
    font_head = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Segoe UI", size=9.5, bold=True, color="0F172A")
    font_regular = Font(name="Segoe UI", size=9.5, color="1E293B")
    font_muted = Font(name="Segoe UI", size=9, color="64748B")
    font_green = Font(name="Segoe UI", size=9.5, bold=True, color="047857")
    font_amber = Font(name="Segoe UI", size=9.5, bold=True, color="B45309")
    font_blue = Font(name="Segoe UI", size=9.5, bold=True, color="0369A1")
    font_highlight_boy = Font(name="Segoe UI", size=9.5, bold=True, color="0F766E")

    fill_brand = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    fill_ana = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_ydk = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    fill_info = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    fill_boy = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid") # Hafif yeşil boy vurgusu
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    double_bottom_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='94A3B8'),
        bottom=Side(style='double', color='0F172A')
    )

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    # =========================================================================
    # SHEET 1: ÜRÜN - MAKİNE ROTALARI & KAPASİTE
    # =========================================================================
    ws1 = wb.active
    ws1.title = "Ürün Rotaları & Kapasite"
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:X1")
    ws1["A1"] = "ERGUNBAS GROUP — ÜRÜN MAKİNE ROTALARI, BOY/KESİM ÖLÇÜLERİ VE KAPASİTE MATRİSİ"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_brand
    ws1["A1"].alignment = align_center
    ws1.row_dimensions[1].height = 36

    ws1.merge_cells("A2:X2")
    ws1["A2"] = "Standart Boy Ölçüleri (2.20m Pervaz, 2.10m Seren, 207cm Levha) | Saatteki Boy Hızları | Ana Rota vs Yedek Rota Kapasiteleri"
    ws1["A2"].font = font_subtitle
    ws1["A2"].fill = fill_brand
    ws1["A2"].alignment = align_center
    ws1.row_dimensions[2].height = 20

    group_headers = [
        ("A3:F3", "ÜRÜN & BOY - KESİM ÖLÇÜLERİ", fill_info),
        ("G3:P3", "ANA ROTA (BİRİNCİL MAKİNE)", fill_ana),
        ("Q3:U3", "YEDEK ROTA 1 (İKİNCİL MAKİNE)", fill_ydk),
        ("V3:W3", "YEDEK ROTA 2 & KALIP", fill_info),
        ("X3:X3", "STRATEJİ", fill_info),
    ]
    for cell_range, title, fill_color in group_headers:
        ws1.merge_cells(cell_range)
        top_cell = ws1[cell_range.split(":")[0]]
        top_cell.value = title
        top_cell.font = font_head
        top_cell.fill = fill_color
        top_cell.alignment = align_center
    ws1.row_dimensions[3].height = 24

    headers_s1 = [
        ("Ürün Kodu & Adı", align_left, 24),
        ("Kategori", align_left, 16),
        ("Standart Boy (Uzunluk)", align_center, 22), # ÖZEL BOY KOLONU
        ("Kesit Ölçüsü", align_left, 18),
        ("Birim Ağırlık", align_right, 14),
        ("70G Toplam KG", align_right, 15),
        ("Ana Rota Hat", align_center, 14),
        ("Fiili Çalışma (Saat)", align_right, 15),
        ("Vardiya Sayısı", align_right, 13),
        ("Fiili Hız (kg/sa)", align_right, 14),
        ("Fiili Hız (Saatteki Boy/Adet)", align_center, 24), # BOY BAZLI HIZ
        ("24s Fiili Kapasite (kg/gün)", align_right, 18),
        ("Standart Kapasite (kg/sa)", align_right, 16),
        ("Standart Hız (Saatteki Boy)", align_center, 24), # HEDEF BOY BAZLI HIZ
        ("Kapasite Gerçekleşme (%)", align_right, 14),
        ("Kalıp Kodu", align_center, 13),
        ("Yedek Rota 1 Hat", align_center, 15),
        ("Yedek 1 Fiili Saat", align_right, 14),
        ("Yedek 1 Fiili Hız (kg/sa)", align_right, 16),
        ("Yedek 1 Hız (Saatteki Boy)", align_center, 24),
        ("Yedek 1 Standart Kapasite (kg/sa)", align_right, 16),
        ("Yedek Rota 2 Hat", align_center, 14),
        ("Setup Süresi (dk)", align_right, 13),
        ("Planlama & Rota Durumu", align_left, 28)
    ]

    for col_idx, (h_name, h_align, col_w) in enumerate(headers_s1, 1):
        cell = ws1.cell(row=4, column=col_idx, value=h_name)
        cell.font = font_head
        cell.fill = fill_brand
        cell.alignment = align_center
        cell.border = thin_border
        col_letter = get_column_letter(col_idx)
        ws1.column_dimensions[col_letter].width = col_w
    ws1.row_dimensions[4].height = 28

    row_num = 5
    tot_produced_kg = 0.0
    tot_ana_hours = 0.0
    tot_ydk_hours = 0.0

    for r in product_rows:
        tot_produced_kg += r["total_kg"]
        tot_ana_hours += r["ana_hours"]
        tot_ydk_hours += r["ydk1_hours"]

        row_fill = fill_zebra if row_num % 2 == 0 else PatternFill(fill_type=None)
        unit_weight_str = f"{r['unit_kg']:.2f} kg/{r['unit_label'].lower()}"

        vals = [
            (r["product"], font_bold, align_left, "@"),
            (r["category"], font_regular, align_left, "@"),
            (r["boy"], font_highlight_boy, align_center, "@"), # BOY KOLONU
            (r["kesit"], font_regular, align_left, "@"),
            (unit_weight_str, font_regular, align_right, "@"),
            (r["total_kg"], font_bold, align_right, "#,##0"),
            (r["ana_hat"], font_blue, align_center, "@"),
            (r["ana_hours"], font_bold, align_right, "#,##0.0"),
            (r["ana_shifts"], font_regular, align_right, "#,##0"),
            (r["ana_act_speed"], font_green if r["ana_act_speed"] > 0 else font_muted, align_right, "#,##0.0"),
            (r["ana_act_speed_text"], font_green if r["ana_act_speed"] > 0 else font_muted, align_center, "@"),
            (r["ana_daily_act"], font_bold, align_right, "#,##0"),
            (r["ana_tek_speed"], font_bold, align_right, "#,##0.0"),
            (r["ana_tek_speed_text"], font_regular, align_center, "@"),
            (r["ana_eff_pct"] / 100.0 if r["ana_eff_pct"] > 0 else 0, font_green if 80 <= r["ana_eff_pct"] <= 120 else (font_amber if r["ana_eff_pct"] > 0 else font_muted), align_right, "0.0%"),
            (r["ana_mold"], font_regular, align_center, "@"),
            (r["ydk1_hat"] or "—", font_amber if r["ydk1_hat"] else font_muted, align_center, "@"),
            (r["ydk1_hours"] if r["ydk1_hours"] > 0 else "—", font_regular, align_right, "#,##0.0" if r["ydk1_hours"] > 0 else "@"),
            (r["ydk1_act_speed"] if r["ydk1_act_speed"] > 0 else "—", font_regular, align_right, "#,##0.0" if r["ydk1_act_speed"] > 0 else "@"),
            (r["ydk1_act_speed_text"], font_regular, align_center, "@"),
            (r["ydk1_tek_speed"] if r["ydk1_tek_speed"] > 0 else "—", font_regular, align_right, "#,##0.0" if r["ydk1_tek_speed"] > 0 else "@"),
            (r["ydk2_hat"] or "—", font_muted, align_center, "@"),
            (r["ana_setup"], font_regular, align_right, "#,##0"),
            (r["strategy"], font_regular, align_left, "@")
        ]

        for col_idx, (val, c_font, c_align, num_fmt) in enumerate(vals, 1):
            cell = ws1.cell(row=row_num, column=col_idx, value=val)
            cell.font = c_font
            cell.alignment = c_align
            cell.border = thin_border
            if col_idx == 3:
                cell.fill = fill_boy
            elif row_fill.fill_type:
                cell.fill = row_fill
            if num_fmt and num_fmt != "@": cell.number_format = num_fmt

        ws1.row_dimensions[row_num].height = 20
        row_num += 1

    ws1.cell(row=row_num, column=1, value="GENEL TOPLAM / FABRİKA").font = font_bold
    ws1.cell(row=row_num, column=1).alignment = align_left
    ws1.cell(row=row_num, column=6, value=tot_produced_kg).font = font_bold
    ws1.cell(row=row_num, column=6).number_format = "#,##0"
    ws1.cell(row=row_num, column=8, value=tot_ana_hours).font = font_bold
    ws1.cell(row=row_num, column=8).number_format = "#,##0.0"
    ws1.cell(row=row_num, column=18, value=tot_ydk_hours).font = font_bold
    ws1.cell(row=row_num, column=18).number_format = "#,##0.0"

    for c in range(1, 25):
        cell = ws1.cell(row=row_num, column=c)
        cell.fill = fill_total
        cell.border = double_bottom_border
    ws1.row_dimensions[row_num].height = 24
    ws1.freeze_panes = "A5"

    # =========================================================================
    # SHEET 2: HAT & MAKİNE ÇALIŞMA ÖZETİ
    # =========================================================================
    ws2 = wb.create_sheet(title="Hat & Makine Çalışma Özeti")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:K1")
    ws2["A1"] = "ERGUNBAS GROUP — MAKİNE & HATLARIN TOPLAM ÇALIŞMA SAATLERİ VE KAPASİTE PERFORMANSI"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_brand
    ws2["A1"].alignment = align_center
    ws2.row_dimensions[1].height = 36

    ws2.merge_cells("A2:K2")
    ws2["A2"] = "Son 70 günde (1.319 vardiya) her bir makinenin fiilen kaç saat çalıştığı, toplam net üretimi ve fiili ortalama hızları"
    ws2["A2"].font = font_subtitle
    ws2["A2"].fill = fill_brand
    ws2["A2"].alignment = align_center
    ws2.row_dimensions[2].height = 20

    headers_s2 = [
        ("Makine / Hat Kodu", align_center, 18),
        ("Makine Adı & Tipi", align_left, 24),
        ("Bölüm / İstasyon", align_left, 18),
        ("Toplam Çalışma (Saat)", align_right, 20),
        ("Tamamlanan Vardiya", align_right, 18),
        ("Toplam Net Üretim (kg)", align_right, 22),
        ("Ortalama Fiili Hız (kg/sa)", align_right, 22),
        ("Aktif Üretilen Ürün Çeşidi", align_center, 24),
        ("Üretilen Başlıca Ürünler", align_left, 38),
        ("Kalıp / Setup Tipi", align_center, 18),
        ("Hat Çalışma Durumu", align_center, 18)
    ]

    for col_idx, (h_name, h_align, col_w) in enumerate(headers_s2, 1):
        cell = ws2.cell(row=3, column=col_idx, value=h_name)
        cell.font = font_head
        cell.fill = fill_brand
        cell.alignment = align_center
        cell.border = thin_border
        ws2.column_dimensions[get_column_letter(col_idx)].width = col_w
    ws2.row_dimensions[3].height = 26

    row_s2 = 4
    tot_s2_hours = 0.0
    tot_s2_shifts = 0
    tot_s2_kg = 0.0

    machine_order = [
        "Hat 101", "Hat 102", "Hat 103", "Hat 104", "Hat 105", 
        "Hat 106", "Hat 107", "Hat 108", "Hat 109", "Hat 201"
    ]
    for m in machines:
        m_name = m.get("name")
        norm_m = f"Hat {m_name}" if m_name.isdigit() else m_name
        if norm_m not in machine_order and m.get("type") in ["extruder", "levha", "mixer", "kirim", "mikronize"]:
            machine_order.append(norm_m)

    for m_key in machine_order:
        m_stat = machine_totals.get(m_key, {"total_kg": 0.0, "shifts": 0, "hours": 0.0, "products": set(), "type": "extruder"})
        tot_kg = m_stat["total_kg"]
        shifts = m_stat["shifts"]
        hours = m_stat["hours"]
        prods = sorted(list(m_stat["products"]))
        avg_spd = round(tot_kg / hours, 1) if hours > 0 else 0.0

        tot_s2_hours += hours
        tot_s2_shifts += shifts
        tot_s2_kg += tot_kg

        dept = "Profil Ekstrüzyon" if "201" not in m_key else "Levha Ekstrüzyon"
        if "Mikser" in m_key: dept = "Hammadde & Mikser"
        if "Kırım" in m_key or "Mikronize" in m_key: dept = "Geri Dönüşüm"

        m_type_name = "Profil Ekstrüder Hattı"
        if "201" in m_key: m_type_name = "WPC Levha Hattı (Hat 201)"
        elif "104" in m_key or "105" in m_key: m_type_name = "Yüksek Hızlı Seren Hattı"
        elif "107" in m_key: m_type_name = "Geniş Kasa & Deck Hattı"
        elif "109" in m_key: m_type_name = "Pervaz Hattı"

        setup_desc = "Standart Kalıp"
        if "107" in m_key: setup_desc = "Kasa / Deck / Seren"
        elif "101" in m_key: setup_desc = "Pervaz Kalıpları"
        elif "201" in m_key: setup_desc = "Levha Düz Kalıp"

        status_str = "Aktif Üretimde" if hours > 0 else "Standby / Bakım"

        row_fill = fill_zebra if row_s2 % 2 == 0 else PatternFill(fill_type=None)

        s2_vals = [
            (m_key, font_bold, align_center, "@"),
            (m_type_name, font_regular, align_left, "@"),
            (dept, font_regular, align_left, "@"),
            (hours, font_bold, align_right, "#,##0.0"),
            (shifts, font_regular, align_right, "#,##0"),
            (tot_kg, font_bold, align_right, "#,##0"),
            (avg_spd, font_green if avg_spd > 0 else font_muted, align_right, "#,##0.0"),
            (len(prods), font_regular, align_center, "#,##0"),
            (", ".join(prods) if prods else "—", font_muted if not prods else font_regular, align_left, "@"),
            (setup_desc, font_regular, align_center, "@"),
            (status_str, font_blue if hours > 0 else font_muted, align_center, "@")
        ]

        for col_idx, (val, c_font, c_align, num_fmt) in enumerate(s2_vals, 1):
            cell = ws2.cell(row=row_s2, column=col_idx, value=val)
            cell.font = c_font
            cell.alignment = c_align
            cell.border = thin_border
            if row_fill.fill_type: cell.fill = row_fill
            if num_fmt and num_fmt != "@": cell.number_format = num_fmt

        ws2.row_dimensions[row_s2].height = 20
        row_s2 += 1

    ws2.cell(row=row_s2, column=1, value="GENEL TOPLAM").font = font_bold
    ws2.cell(row=row_s2, column=4, value=tot_s2_hours).font = font_bold
    ws2.cell(row=row_s2, column=4).number_format = "#,##0.0"
    ws2.cell(row=row_s2, column=5, value=tot_s2_shifts).font = font_bold
    ws2.cell(row=row_s2, column=5).number_format = "#,##0"
    ws2.cell(row=row_s2, column=6, value=tot_s2_kg).font = font_bold
    ws2.cell(row=row_s2, column=6).number_format = "#,##0"
    avg_factory_spd = round(tot_s2_kg / tot_s2_hours, 1) if tot_s2_hours > 0 else 0.0
    ws2.cell(row=row_s2, column=7, value=avg_factory_spd).font = font_bold
    ws2.cell(row=row_s2, column=7).number_format = "#,##0.0"

    for c in range(1, 12):
        cell = ws2.cell(row=row_s2, column=c)
        cell.fill = fill_total
        cell.border = double_bottom_border
    ws2.row_dimensions[row_s2].height = 24
    ws2.freeze_panes = "A4"

    # =========================================================================
    # SHEET 3: FİİLİ ÜRETİM DÖKÜMÜ (ÜRÜN x HAT EŞLEŞMESİ)
    # =========================================================================
    ws3 = wb.create_sheet(title="Ürün x Hat Fiili Üretim")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:N1")
    ws3["A1"] = "ERGUNBAS GROUP — ÜRÜN BAZINDA HAT ÜRETİM DAĞILIMI VE KAPASİTE PERFORMANSI"
    ws3["A1"].font = font_title
    ws3["A1"].fill = fill_brand
    ws3["A1"].alignment = align_center
    ws3.row_dimensions[1].height = 36

    ws3.merge_cells("A2:N2")
    ws3["A2"] = "Her bir ürünün çalıştığı makineler, standart boyları, çekilen adetleri, net tonajları ve saatteki net hızları"
    ws3["A2"].font = font_subtitle
    ws3["A2"].fill = fill_brand
    ws3["A2"].alignment = align_center
    ws3.row_dimensions[2].height = 20

    headers_s3 = [
        ("Ürün Adı", align_left, 24),
        ("Kategori", align_left, 16),
        ("Standart Boy (Uzunluk)", align_center, 22),
        ("Kesit Detayı", align_left, 18),
        ("Çalışılan Hat", align_center, 14),
        ("Rota Sınıfı", align_center, 15),
        ("Fiili Çalışma (Saat)", align_right, 17),
        ("Vardiya Sayısı", align_right, 14),
        ("Üretilen Miktar (Adet/Boy)", align_right, 18),
        ("Birim Ağırlık", align_right, 14),
        ("Toplam Üretim (kg)", align_right, 17),
        ("Fiili Hız (kg/sa)", align_right, 15),
        ("Fiili Hız (Saatteki Boy/Adet)", align_center, 24),
        ("Kapasite Payı (%)", align_right, 15)
    ]

    for col_idx, (h_name, h_align, col_w) in enumerate(headers_s3, 1):
        cell = ws3.cell(row=3, column=col_idx, value=h_name)
        cell.font = font_head
        cell.fill = fill_brand
        cell.alignment = align_center
        cell.border = thin_border
        ws3.column_dimensions[get_column_letter(col_idx)].width = col_w
    ws3.row_dimensions[3].height = 26

    row_s3 = 4
    flat_pairs = []
    for p, hats in actual_runs.items():
        cat = get_product_category(p)
        cat_info = prod_catalog.get(p)
        if not cat_info:
            for cp, info in prod_catalog.items():
                if app_backend.match_product_daily(cp, "", p):
                    cat_info = info
                    break
        unit_kg = cat_info.get("unit_kg", 1.0) if cat_info else 1.0

        dim_info = PRODUCT_DIMENSIONS.get(p)
        if not dim_info:
            for kp, kv in PRODUCT_DIMENSIONS.items():
                if app_backend.match_product_daily(kp, "", p):
                    dim_info = kv
                    break
        boy_str = dim_info["boy"] if dim_info else "2.20 Metre"
        kesit_str = dim_info["kesit"] if dim_info else "Standart Profil"
        unit_label = dim_info["unit_name"] if dim_info else ("Plaka" if "Levha" in category else "Boy")

        p_total_kg = sum(v["kg"] for v in hats.values())

        sorted_h = sorted(hats.items(), key=lambda x: -x[1]["kg"])
        for idx, (h, v) in enumerate(sorted_h, 1):
            role = "ANA ROTA" if idx == 1 else (f"YEDEK ROTA {idx-1}")
            share = (v["kg"] / p_total_kg) if p_total_kg > 0 else 0.0
            act_spd = round(v["kg"] / v["hours"], 1) if v["hours"] > 0 else 0.0
            act_qty_spd = round(v["qty"] / v["hours"], 1) if v["hours"] > 0 else 0.0
            act_speed_text = f"Saatte ~{int(round(act_qty_spd))} {unit_label}" if act_qty_spd > 0 else "—"

            flat_pairs.append({
                "product": p,
                "category": cat,
                "boy": boy_str,
                "kesit": kesit_str,
                "unit_label": unit_label,
                "hat": h,
                "role": role,
                "hours": v["hours"],
                "shifts": v["shifts"],
                "qty": v["qty"],
                "unit_kg": unit_kg,
                "kg": v["kg"],
                "act_spd": act_spd,
                "act_speed_text": act_speed_text,
                "share": share
            })

    flat_pairs.sort(key=lambda x: -x["kg"])

    tot_s3_hours = 0.0
    tot_s3_qty = 0.0
    tot_s3_kg = 0.0

    for item in flat_pairs:
        tot_s3_hours += item["hours"]
        tot_s3_qty += item["qty"]
        tot_s3_kg += item["kg"]

        row_fill = fill_zebra if row_s3 % 2 == 0 else PatternFill(fill_type=None)
        role_font = font_blue if "ANA" in item["role"] else font_amber
        unit_weight_str = f"{item['unit_kg']:.2f} kg/{item['unit_label'].lower()}"

        s3_vals = [
            (item["product"], font_bold, align_left, "@"),
            (item["category"], font_regular, align_left, "@"),
            (item["boy"], font_highlight_boy, align_center, "@"),
            (item["kesit"], font_regular, align_left, "@"),
            (item["hat"], font_bold, align_center, "@"),
            (item["role"], role_font, align_center, "@"),
            (item["hours"], font_bold, align_right, "#,##0.0"),
            (item["shifts"], font_regular, align_right, "#,##0"),
            (item["qty"], font_regular, align_right, "#,##0"),
            (unit_weight_str, font_regular, align_right, "@"),
            (item["kg"], font_bold, align_right, "#,##0"),
            (item["act_spd"], font_green, align_right, "#,##0.0"),
            (item["act_speed_text"], font_green, align_center, "@"),
            (item["share"], font_regular, align_right, "0.0%")
        ]

        for col_idx, (val, c_font, c_align, num_fmt) in enumerate(s3_vals, 1):
            cell = ws3.cell(row=row_s3, column=col_idx, value=val)
            cell.font = c_font
            cell.alignment = c_align
            cell.border = thin_border
            if col_idx == 3:
                cell.fill = fill_boy
            elif row_fill.fill_type:
                cell.fill = row_fill
            if num_fmt and num_fmt != "@": cell.number_format = num_fmt

        ws3.row_dimensions[row_s3].height = 20
        row_s3 += 1

    ws3.cell(row=row_s3, column=1, value="GENEL TOPLAM").font = font_bold
    ws3.cell(row=row_s3, column=7, value=tot_s3_hours).font = font_bold
    ws3.cell(row=row_s3, column=7).number_format = "#,##0.0"
    ws3.cell(row=row_s3, column=9, value=tot_s3_qty).font = font_bold
    ws3.cell(row=row_s3, column=9).number_format = "#,##0"
    ws3.cell(row=row_s3, column=11, value=tot_s3_kg).font = font_bold
    ws3.cell(row=row_s3, column=11).number_format = "#,##0"

    for c in range(1, 15):
        cell = ws3.cell(row=row_s3, column=c)
        cell.fill = fill_total
        cell.border = double_bottom_border
    ws3.row_dimensions[row_s3].height = 24
    ws3.freeze_panes = "A4"

    return wb

def create_product_rotas_excel():
    wb = create_product_rotas_excel_wb()
    desktop_dir = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
    if not os.path.exists(desktop_dir):
        desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    dest_path_desktop = os.path.join(desktop_dir, "Ergunbas_Urun_Hat_Rotalari_ve_Kapasite_Analizi.xlsx")
    dest_path_static = os.path.join(APP_DIR, "static", "Ergunbas_Urun_Hat_Rotalari_ve_Kapasite_Analizi.xlsx")

    try:
        wb.save(dest_path_desktop)
        print(f"Excel Masaüstüne kaydedildi: {dest_path_desktop}")
    except Exception as e:
        print(f"Desktop save error (muhtemelen Excel açık): {e}")
        try:
            alt_path = os.path.join(desktop_dir, "Ergunbas_Urun_Hat_Rotalari_ve_Kapasite_Analizi_Guncel.xlsx")
            wb.save(alt_path)
            print(f"Alternatif dosya kaydedildi: {alt_path}")
        except Exception as e2:
            print(f"Alt save error: {e2}")

    try:
        wb.save(dest_path_static)
        print(f"Excel static klasörüne kaydedildi: {dest_path_static}")
    except Exception as e:
        print(f"Static save error: {e}")

    return dest_path_desktop

if __name__ == "__main__":
    create_product_rotas_excel()
