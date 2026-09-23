"""
Ağustos 2026 Excel verisini (Mixer, Kırım, Mikronize) sisteme aktaran ve
aylık JSON parçalarını (data_days_2026-08.json) güncelleyen script.
Çalıştırmak için: python import_agustus_2026.py
"""

import openpyxl
import sys
import json
import os

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)
import app_backend as backend

EXCEL_PATH = r"C:\Users\llboz\OneDrive\Desktop\AGUSTUS_2026_MIXER_VE_GERI_DONUSUM_VERI_GIRIS_TABLOSU.xlsx"

def parse_val(v):
    if v is None:
        return 0.0, "normal"
    s = str(v).strip().lower()
    if "ariza" in s:
        return 0.0, "arizali"
    if "personel" in s:
        return 0.0, "personel_yok"
    try:
        return float(v), "normal"
    except (ValueError, TypeError):
        return 0.0, "normal"

def s(ws, row, col):
    val = ws.cell(row=row, column=col).value
    if val is None:
        return ""
    return str(val).strip()

def parse_sheet(ws, sheet_name):
    """Parse one daily sheet and return dict for mixer, kirim, mikronize."""
    # Row 4/5: employees (J4 / J5 = Gündüz, K4 / K5 = Gece)
    emp_gunduz_val, _ = parse_val(ws.cell(row=5, column=10).value)
    emp_gece_val, _ = parse_val(ws.cell(row=5, column=11).value)
    emp_gunduz = int(emp_gunduz_val)
    emp_gece = int(emp_gece_val)

    # ---- KIRIM ----
    kirim_list = []
    for hat_name, row_idx in [
        ('Kırım 1', 10),
        ('Kırım 2', 11),
        ('Kırım 3', 12),
        ('Kırım 4', 13),
    ]:
        gunduz_kg, sg = parse_val(ws.cell(row=row_idx, column=2).value)
        gece_kg, sn = parse_val(ws.cell(row=row_idx, column=3).value)
        status = sg if sg != "normal" else sn
        kirim_list.append({
            "hat": hat_name,
            "gunduz": round(gunduz_kg, 2),
            "gece": round(gece_kg, 2),
            "gunduz_kg": round(gunduz_kg, 2),
            "gece_kg": round(gece_kg, 2),
            "toplam_kg": round(gunduz_kg + gece_kg, 2),
            "status": status
        })

    # ---- MIKRONIZE ----
    mikronize_list = []
    for hat_name, row_idx in [
        ('Mikronize 1', 19),
        ('Mikronize 2', 20),
        ('Mikronize 3', 21),
        ('Mikronize 4', 22),
        ('Mikronize 5', 23),
        ('Mikronize 6', 24),
    ]:
        gunduz_kg, sg = parse_val(ws.cell(row=row_idx, column=2).value)
        gece_kg, sn = parse_val(ws.cell(row=row_idx, column=3).value)
        status = sg if sg != "normal" else sn
        mikronize_list.append({
            "hat": hat_name,
            "gunduz": round(gunduz_kg, 2),
            "gece": round(gece_kg, 2),
            "gunduz_kg": round(gunduz_kg, 2),
            "gece_kg": round(gece_kg, 2),
            "toplam_kg": round(gunduz_kg + gece_kg, 2),
            "status": status
        })

    # ---- MIXER ----
    mixer_list = []
    for makine, row_range in [
        ('Mixer 1', range(10, 18)),
        ('Mixer 2', range(20, 28)),
        ('Mixer 3', range(30, 38)),
    ]:
        for r in row_range:
            recipe = s(ws, r, 7)
            if not recipe or 'TOPLAM' in recipe.upper() or 'MAKINE' in recipe.upper():
                continue
            batch_kg, _ = parse_val(ws.cell(r, 10).value)
            gunduz_sarj, _ = parse_val(ws.cell(r, 8).value)
            gece_sarj, _ = parse_val(ws.cell(r, 9).value)
            gunduz_kg, _ = parse_val(ws.cell(r, 11).value)
            gece_kg, _ = parse_val(ws.cell(r, 12).value)
            toplam_kg = gunduz_kg + gece_kg
            if gunduz_sarj > 0 or gece_sarj > 0 or gunduz_kg > 0 or gece_kg > 0 or toplam_kg > 0:
                mixer_list.append({
                    "makine": makine,
                    "recipe": recipe,
                    "batch_kg": round(batch_kg, 4),
                    "gunduz_sarj": int(gunduz_sarj),
                    "gece_sarj": int(gece_sarj),
                    "toplam_sarj": int(gunduz_sarj + gece_sarj),
                    "gunduz_kg": round(gunduz_kg, 4),
                    "gece_kg": round(gece_kg, 4),
                    "toplam_kg": round(toplam_kg, 4),
                })

    return {
        "date": sheet_name,
        "kirim": kirim_list,
        "mikronize": mikronize_list,
        "mixer": mixer_list,
        "mixer_emp_gunduz": emp_gunduz,
        "mixer_emp_gece": emp_gece,
    }


def main():
    print(f"Excel okunuyor: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

    date_sheets = [sn for sn in wb.sheetnames if len(sn) == 10 and sn[2] == '.' and sn[5] == '.']
    print(f"Toplam gün sayfası: {len(date_sheets)}")

    # Verileri app_backend üzerinden yükle (bölünmüş dosyaları da birleştirir)
    data = backend.load_data()
    if "daily_data" not in data:
        data["daily_data"] = {}

    existing_keys = [int(k) for k in data["daily_data"].keys() if str(k).isdigit()]
    next_key = max(existing_keys, default=0) + 1

    updated_count = 0
    added_count = 0

    for sheet_name in sorted(date_sheets):
        ws = wb[sheet_name]
        parsed = parse_sheet(ws, sheet_name)
        date_str = parsed["date"]

        # Bu tarihe ait kayıt var mı kontrol et
        target_k = None
        for k, v in data["daily_data"].items():
            if v.get("date") == date_str:
                target_k = k
                break

        if target_k:
            # Mevcut günün içine mixer / kırım / mikronize verilerini entegre et
            data["daily_data"][target_k]["kirim"] = parsed["kirim"]
            data["daily_data"][target_k]["mikronize"] = parsed["mikronize"]
            data["daily_data"][target_k]["mixer"] = parsed["mixer"]
            data["daily_data"][target_k]["mixer_emp_gunduz"] = parsed["mixer_emp_gunduz"]
            data["daily_data"][target_k]["mixer_emp_gece"] = parsed["mixer_emp_gece"]
            updated_count += 1
            print(f"  [OK] Guncellendi: {date_str} (key={target_k}, {len(parsed['mixer'])} mikser satiri)")
        else:
            # Yeni gün oluştur
            new_key = str(next_key)
            data["daily_data"][new_key] = {
                "day": next_key,
                "date": date_str,
                "gunduz": {"employees": 0, "hours": 12.0, "operator": "", "extruders": [], "levha": []},
                "gece": {"employees": 0, "hours": 12.0, "operator": "", "extruders": [], "levha": []},
                "downtimes": [],
                "kirim": parsed["kirim"],
                "mikronize": parsed["mikronize"],
                "mixer": parsed["mixer"],
                "mixer_emp_gunduz": parsed["mixer_emp_gunduz"],
                "mixer_emp_gece": parsed["mixer_emp_gece"],
            }
            next_key += 1
            added_count += 1
            print(f"  + Yeni Gün Eklendi: {date_str} (key={new_key})")

    print(f"\nİşlem Tamamlandı: {updated_count} gün güncellendi, {added_count} gün eklendi.")

    # Tüm yerel dosyalara (data_days_2026-08.json vb.) ve GitHub'a kaydet
    backend.save_data(data)
    print("[OK] backend.save_data() tamamlandi (data_days_2026-08.json dahil tum dosyalar yerelde ve bellekte guncellendi)!")

    # Özet istatistik
    total_mx = 0.0
    total_k = 0.0
    total_m = 0.0
    for k, v in data["daily_data"].items():
        if ".08.2026" in v.get("date", ""):
            for mx in v.get("mixer", []):
                total_mx += mx.get("toplam_kg", 0)
            for kr in v.get("kirim", []):
                total_k += kr.get("toplam_kg", 0)
            for mk in v.get("mikronize", []):
                total_m += mk.get("toplam_kg", 0)

    print("\n=== AĞUSTOS 2026 AKTARIM ÖZETİ ===")
    print(f"  Toplam Mikser:    {total_mx/1000:,.2f} Ton ({total_mx:,.1f} kg)")
    print(f"  Toplam Kırım:     {total_k/1000:,.2f} Ton ({total_k:,.1f} kg)")
    print(f"  Toplam Mikronize: {total_m/1000:,.2f} Ton ({total_m:,.1f} kg)")
    print(f"  Genel Toplam:     {(total_mx+total_k+total_m)/1000:,.2f} Ton")


if __name__ == "__main__":
    main()
