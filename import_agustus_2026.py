"""
Ağustos 2026 Excel verisini sisteme aktaran script.
Çalıştırmak için: python import_agustus_2026.py
"""

import openpyxl
import sys
import json
import os

EXCEL_PATH = r"C:\Users\llboz\OneDrive\Desktop\AGUSTUS_2026_MIXER_VE_GERI_DONUSUM_VERI_GIRIS_TABLOSU.xlsx"

# Kirim hat isimleri → satır karşılıkları (row 10-13 = K1-K4)
KIRIM_ROWS = {
    'Kırım 1': 10, 'Kirim 1': 10, 'K\u0131r\u0131m 1': 10,
    'Kırım 2': 11, 'Kirim 2': 11, 'K\u0131r\u0131m 2': 11,
    'Kırım 3': 12, 'Kirim 3': 12, 'K\u0131r\u0131m 3': 12,
    'Kırım 4': 13, 'Kirim 4': 13, 'K\u0131r\u0131m 4': 13,
}
# Mikronize satırları (row 19-24)
MIKRONIZE_ROWS = {
    'Mikronize 1': 19,
    'Mikronize 2': 20,
    'Mikronize 3': 21,
    'Mikronize 4': 22,
    'Mikronize 5': 23,
    'Mikronize 6': 24,
}
# Mixer makine → satır aralıkları  (recipe rows per machine)
# Mixer1: rows 10-17 (recipes F10:F17), Mixer2: 20-27, Mixer3: 30-37
MIXER_MACHINE_RECIPE_RANGES = {
    'Mixer 1': range(10, 18),
    'Mixer 2': range(20, 28),
    'Mixer 3': range(30, 38),
}

def v(ws, row, col):
    """Return numeric value or 0 from cell."""
    val = ws.cell(row=row, column=col).value
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0

def s(ws, row, col):
    """Return string value from cell."""
    val = ws.cell(row=row, column=col).value
    if val is None:
        return ""
    return str(val).strip()

def parse_sheet(ws, sheet_name):
    """Parse one daily sheet and return a dict for daily_data entry."""
    # Row 4/5: employees
    # J4='GÜNDÜZ ÇALIŞAN', K4='GECE ÇALIŞAN'
    emp_gunduz = int(v(ws, 5, 10)) if ws.cell(row=5, column=10).value else 0
    emp_gece   = int(v(ws, 5, 11)) if ws.cell(row=5, column=11).value else 0

    # ---- KIRIM ----
    kirim_list = []
    for hat_name, row_idx in [
        ('Kırım 1', 10),
        ('Kırım 2', 11),
        ('Kırım 3', 12),
        ('Kırım 4', 13),
    ]:
        gunduz_kg = v(ws, row_idx, 2)  # col B
        gece_kg   = v(ws, row_idx, 3)  # col C
        # Toplam = col D (for reference, we use gunduz/gece separately)
        kirim_list.append({
            "hat": hat_name,
            "gunduz": gunduz_kg,
            "gece": gece_kg,
            "status": "normal"
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
        gunduz_kg = v(ws, row_idx, 2)  # col B
        gece_kg   = v(ws, row_idx, 3)  # col C
        mikronize_list.append({
            "hat": hat_name,
            "gunduz": gunduz_kg,
            "gece": gece_kg,
            "status": "normal"
        })

    # ---- MIXER ----
    # Mixer 1: rows 10-17 (F=Makine, G=Reçete, H=Gündüz Şarj, I=Gece Şarj, J=Batch kg, K=Gündüz kg, L=Gece kg)
    # Mixer 2: rows 20-27
    # Mixer 3: rows 30-37
    mixer_list = []
    for makine, row_range in [
        ('Mixer 1', range(10, 18)),
        ('Mixer 2', range(20, 28)),
        ('Mixer 3', range(30, 38)),
    ]:
        for r in row_range:
            recipe  = s(ws, r, 7)   # col G
            if not recipe or 'TOPLAM' in recipe.upper() or 'MAKINE' in recipe.upper():
                continue
            batch_kg    = v(ws, r, 10)  # col J
            gunduz_sarj = v(ws, r, 8)   # col H
            gece_sarj   = v(ws, r, 9)   # col I
            gunduz_kg   = v(ws, r, 11)  # col K
            gece_kg     = v(ws, r, 12)  # col L
            toplam_kg   = gunduz_kg + gece_kg
            mixer_list.append({
                "makine": makine,
                "recipe": recipe,
                "batch_kg": round(batch_kg, 4),
                "gunduz_sarj": int(gunduz_sarj),
                "gece_sarj": int(gece_sarj),
                "gunduz_kg": round(gunduz_kg, 4),
                "gece_kg": round(gece_kg, 4),
                "toplam_kg": round(toplam_kg, 4),
            })

    # date format from sheet name: "01.08.2026" → "01.08.2026"
    date_str = sheet_name  # already "DD.MM.YYYY"

    return {
        "date": date_str,
        "gunduz": {
            "employees": emp_gunduz,
            "hours": 7.5,
            "operator": "",
            "extruders": [],
            "levha": []
        },
        "gece": {
            "employees": emp_gece,
            "hours": 7.5,
            "operator": "",
            "extruders": [],
            "levha": []
        },
        "downtimes": [],
        "kirim": kirim_list,
        "mikronize": mikronize_list,
        "mixer": mixer_list,
        "mixer_emp_gunduz": emp_gunduz,
        "mixer_emp_gece": emp_gece,
    }


def main():
    print(f"Excel okunuyor: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

    # Only parse date sheets (skip Aylık Özet, Reçeteler, Hammadde Özeti)
    date_sheets = [sn for sn in wb.sheetnames if sn[2] == '.' and sn[5] == '.']
    print(f"Toplam gün sayfası: {len(date_sheets)}")

    # Import app_backend to access data store
    import app_backend as backend
    data = backend.load_data()

    if "daily_data" not in data:
        data["daily_data"] = {}

    # Find current max key (numeric)
    existing_keys = [int(k) for k in data["daily_data"].keys() if k.isdigit()]
    next_key = max(existing_keys, default=0) + 1

    added = 0
    skipped = 0
    updated_keys = []

    for sheet_name in sorted(date_sheets):
        ws = wb[sheet_name]
        day_data = parse_sheet(ws, sheet_name)

        # Check if this date already exists
        date_str = day_data["date"]
        existing_key = None
        for k, v_day in data["daily_data"].items():
            if v_day.get("date") == date_str:
                existing_key = k
                break

        if existing_key:
            # Update existing entry — merge mixer/kirim/mikronize data
            data["daily_data"][existing_key].update({
                "kirim": day_data["kirim"],
                "mikronize": day_data["mikronize"],
                "mixer": day_data["mixer"],
                "mixer_emp_gunduz": day_data["mixer_emp_gunduz"],
                "mixer_emp_gece": day_data["mixer_emp_gece"],
            })
            print(f"  GÜNCELLENDI: {date_str} (key={existing_key})")
            updated_keys.append(existing_key)
            skipped += 1
        else:
            # Create new entry
            key = str(next_key)
            day_data["day"] = next_key
            data["daily_data"][key] = day_data
            updated_keys.append(key)
            next_key += 1
            added += 1
            print(f"  EKLENDI: {date_str} (key={key})")

    print(f"\nToplam eklenen: {added} | Güncellenen: {skipped}")

    # Save back via backend function (save_data already syncs to GitHub automatically)
    backend.save_data(data)
    print("Veriler kaydedildi ve GitHub senkronizasyonu tamamlandı!")

    # Quick summary
    total_mixer_kg = 0.0
    total_kirim_kg = 0.0
    total_mikronize_kg = 0.0
    days_count = 0
    for k in updated_keys:
        d = data["daily_data"].get(k, {})
        if not d:
            continue
        for m in (d.get("mixer") or []):
            total_mixer_kg += m.get("toplam_kg", 0)
        for kh in (d.get("kirim") or []):
            total_kirim_kg += kh.get("gunduz", 0) + kh.get("gece", 0)
        for mh in (d.get("mikronize") or []):
            total_mikronize_kg += mh.get("gunduz", 0) + mh.get("gece", 0)
        days_count += 1

    print(f"\n=== AĞUSTOS 2026 ÖZET ===")
    print(f"  İşlenen gün: {days_count}")
    print(f"  Toplam Mikser:    {total_mixer_kg/1000:.2f} Ton")
    print(f"  Toplam Kırım:     {total_kirim_kg/1000:.2f} Ton")
    print(f"  Toplam Mikronize: {total_mikronize_kg/1000:.2f} Ton")
    print(f"  Genel Toplam:     {(total_kirim_kg+total_mikronize_kg+total_mixer_kg)/1000:.2f} Ton")


if __name__ == "__main__":
    main()
