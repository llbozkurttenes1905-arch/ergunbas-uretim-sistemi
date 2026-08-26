from fastapi import FastAPI, HTTPException, Body, Header
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, Response
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import json
import os
import re
import math
import base64
import urllib.request
import urllib.error
from datetime import datetime
import openpyxl
from io import BytesIO

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(APP_DIR, "data.json")
USERS_FILE = os.path.join(APP_DIR, "users.json")

# ============================================================================
# GITHUB TABANLI KALICI DEPOLAMA
# ============================================================================
# Render'ın (ve benzer platformların) ücretsiz planlarında yerel disk KALICI
# DEĞİLDİR: servis her yeniden başladığında (uyku sonrası uyanma, redeploy vb.)
# yerel dosyalardaki değişiklikler kaybolur. Bunu önlemek için veriyi ayrıca
# GitHub reposuna (Contents API üzerinden) commit ediyoruz. Böylece:
#   - Her kayıt işlemi GitHub'da bir commit oluşturuyor (otomatik yedek/geçmiş)
#   - Servis yeniden başladığında, yerel dosya yerine GitHub'daki EN GÜNCEL
#     veriyi çekiyoruz, veri kaybı yaşanmıyor
#   - Üçüncü parti bir veritabanı şirketine ihtiyaç yok, tamamen kendi
#     GitHub reponuz kullanılıyor
#
# Devreye almak için Render'da şu ortam değişkenlerini tanımlayın:
#   GITHUB_TOKEN  -> "repo" yetkisine sahip bir GitHub Personal Access Token
#   GITHUB_REPO   -> "kullanici-adi/repo-adi" formatında (örn: llbozkurttenes1905-arch/ergunbas-uretim-sistemi)
#   GITHUB_BRANCH -> (opsiyonel, varsayılan "main")
#
# Bu değişkenler tanımlı değilse sistem otomatik olarak eski (yerel dosya)
# yöntemine döner, hiçbir şey bozulmaz.
# ============================================================================

GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO = os.environ.get("GITHUB_REPO", "")
GITHUB_BRANCH = os.environ.get("GITHUB_BRANCH", "main")
GITHUB_API_BASE = "https://api.github.com"

_data_cache = None
_users_cache = None


def _github_enabled():
    return bool(GITHUB_TOKEN and GITHUB_REPO)


def _github_request(method, url, payload=None):
    """GitHub API'sine düşük seviyeli istek atar. (sonuç_dict, http_status) döner."""
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"Bearer {GITHUB_TOKEN}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("User-Agent", "ergunbas-uretim-sistemi")
    if data:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            body = resp.read().decode("utf-8")
            return (json.loads(body) if body else {}), resp.status
    except urllib.error.HTTPError as e:
        try:
            err_body = json.loads(e.read().decode("utf-8"))
        except Exception:
            err_body = {}
        return err_body, e.code
    except Exception as e:
        return {"error": str(e)}, 0


def github_get_file(path):
    """GitHub reposundaki bir dosyanın içeriğini (JSON) ve sha'sını getirir.
    Dosya yoksa veya GitHub devre dışıysa (None, None) döner."""
    if not _github_enabled():
        return None, None
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_REPO}/contents/{path}?ref={GITHUB_BRANCH}"
    result, status = _github_request("GET", url)
    if status == 200 and "content" in result:
        try:
            content_str = base64.b64decode(result["content"]).decode("utf-8")
            return json.loads(content_str), result.get("sha")
        except Exception as e:
            print(f"[github_get_file] '{path}' ayrıştırılamadı: {e}")
            return None, None
    return None, None


def github_put_file(path, data_dict, message):
    """Bir JSON içeriği GitHub reposuna commit eder (varsa günceller, yoksa oluşturur).
    Başarılıysa True, değilse False döner. Asla exception fırlatmaz (best-effort)."""
    if not _github_enabled():
        return False
    try:
        _, sha = github_get_file(path)
        content_str = json.dumps(data_dict, ensure_ascii=False, indent=2)
        content_b64 = base64.b64encode(content_str.encode("utf-8")).decode("utf-8")
        url = f"{GITHUB_API_BASE}/repos/{GITHUB_REPO}/contents/{path}"
        payload = {"message": message, "content": content_b64, "branch": GITHUB_BRANCH}
        if sha:
            payload["sha"] = sha
        result, status = _github_request("PUT", url, payload)
        if status not in (200, 201):
            print(f"[github_put_file] '{path}' commit edilemedi (HTTP {status}): {result}")
            return False
        return True
    except Exception as e:
        print(f"[github_put_file] '{path}' commit edilirken hata: {e}")
        return False


def load_users():
    global _users_cache
    if _users_cache is not None:
        return _users_cache

    # Önce GitHub'daki (kalıcı) en güncel veriyi çekmeyi dene
    gh_data, _ = github_get_file("users.json")
    if gh_data is not None:
        _users_cache = gh_data
        return _users_cache

    # GitHub devre dışı/başarısızsa yerel dosyaya düş
    if not os.path.exists(USERS_FILE):
        default_users = {
            "users": [
                {"id": "u1", "username": "admin", "password": "ergunbas2026", "role": "admin", "name": "Sistem Yöneticisi"},
                {"id": "u2", "username": "vardiya1", "password": "vardiya1", "role": "operator", "name": "Gündüz Operatörü"},
                {"id": "u3", "username": "vardiya2", "password": "vardiya2", "role": "operator", "name": "Gece Operatörü"},
            ]
        }
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(default_users, f, ensure_ascii=False, indent=2)
        _users_cache = default_users
        return _users_cache

    with open(USERS_FILE, "r", encoding="utf-8") as f:
        _users_cache = json.load(f)
    return _users_cache


def save_users(users_data):
    global _users_cache
    _users_cache = users_data
    # Yerel dosyaya da yaz (aynı process içinde hızlı erişim + GitHub başarısız olursa yedek)
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users_data, f, ensure_ascii=False, indent=2)
    # Kalıcı depolama: GitHub'a commit et
    github_put_file("users.json", users_data, "Kullanıcı verisi güncellendi (otomatik)")

def require_editor(x_username: Optional[str]):
    """Sadece 'admin' veya 'operator' rolündeki kullanıcılar veri girişi/düzenleme yapabilir.
    'viewer' (görüntüleyici) rolündeki kullanıcılar salt okunur erişime sahiptir ve
    veri kaydetme/düzenleme isteklerinde reddedilir."""
    if not x_username:
        return  # header gönderilmediyse (eski istemci) engelleme, geriye dönük uyumluluk
    users_data = load_users()
    user = next((u for u in users_data.get("users", []) if u.get("username") == x_username), None)
    if user and user.get("role") == "viewer":
        raise HTTPException(status_code=403, detail="Salt okunur yetkiniz var, veri girişi/düzenleme yapamazsınız.")

def require_daily_operator(x_username: Optional[str]):
    """GÜNLÜK ÜRETİM VERİ GİRİŞİ sadece 'operator' rolündeki kullanıcı tarafından yapılabilir.
    'admin' rolü sistem yönetimi (kullanıcı/makine/ürün) için ayrılmıştır ve günlük veri
    girişi yapamaz. 'viewer' zaten hiçbir şekilde veri giremez."""
    if not x_username:
        return  # header gönderilmediyse (eski istemci) engelleme, geriye dönük uyumluluk
    users_data = load_users()
    user = next((u for u in users_data.get("users", []) if u.get("username") == x_username), None)
    if user and user.get("role") != "operator":
        raise HTTPException(
            status_code=403,
            detail="Günlük üretim veri girişi sadece yetkili operatör (Sorumlu Mühendis) tarafından yapılabilir."
        )

app = FastAPI(title="ERGUNBAS Group Ekstrüder ve Levha Üretim Yönetim Sistemi")

def load_data():
    global _data_cache
    if _data_cache is not None:
        return _data_cache

    # Önce GitHub'daki (kalıcı) en güncel veriyi çekmeyi dene
    gh_data, _ = github_get_file("data.json")
    if gh_data is not None:
        _data_cache = gh_data
        return _data_cache

    # GitHub devre dışı/başarısızsa yerel dosyaya düş
    if not os.path.exists(DATA_FILE):
        _data_cache = {"machines": [], "products": [], "daily_data": {}}
        return _data_cache
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        _data_cache = json.load(f)
    return _data_cache

def save_data(data):
    global _data_cache
    _data_cache = data
    # Yerel dosyaya da yaz (aynı process içinde hızlı erişim + GitHub başarısız olursa yedek)
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    # Kalıcı depolama: GitHub'a commit et
    github_put_file("data.json", data, "Üretim verisi güncellendi (otomatik)")

# Models
class MachineCreate(BaseModel):
    name: str
    type: str  # 'extruder' or 'levha'

class ProductCreate(BaseModel):
    name: str
    category: str
    door_ratio: float

class ExtruderEntry(BaseModel):
    row: Optional[int] = 0
    hat: str
    product: str
    length: Optional[float] = 0
    speed: Optional[float] = 0
    heads: Optional[float] = 1
    hours: Optional[float] = 0
    prod_kg: float = 0
    fire_kg: float = 0
    qty: int = 0
    sets: Optional[float] = 0

class LevhaEntry(BaseModel):
    row: Optional[int] = 0
    hat: str
    color: str
    width: Optional[float] = 93
    length: Optional[float] = 208
    m2_one: Optional[float] = 0
    qty: int = 0
    total_m2: Optional[float] = 0
    width_fire_cm: Optional[float] = 0
    dead_fire_m2: Optional[float] = 0
    dead_fire_kg: Optional[float] = 0
    total_kg: float = 0
    sets: Optional[float] = 0
    kg_per_m2: Optional[float] = 3.5
    kalip_cikis_eni: Optional[float] = 108.0

class DowntimeEntry(BaseModel):
    shift: str
    hat: str
    fire_reason: Optional[str] = ""
    fire_kg: Optional[float] = 0
    down_reason: Optional[str] = ""
    down_min: Optional[float] = 0
    desc: Optional[str] = ""

class ShiftData(BaseModel):
    employees: int = 10
    hours: float = 12
    operator: Optional[str] = ""
    extruders: List[ExtruderEntry] = []
    levha: List[LevhaEntry] = []

class DailyDataUpdate(BaseModel):
    gunduz: ShiftData
    gece: ShiftData
    downtimes: List[DowntimeEntry] = []

class AddDateRequest(BaseModel):
    date_str: str


def compute_door_capacity(db_data, filter_date_keys: Optional[List[str]] = None):
    """Calculates door capacity and component carryovers for all or specified date keys."""
    pervaz_prod = 0.0
    kasa_prod = 0.0
    seren_prod = 0.0
    levha_prod = 0.0

    days_to_process = filter_date_keys if filter_date_keys else list(db_data["daily_data"].keys())

    for d_str in days_to_process:
        if d_str not in db_data["daily_data"]:
            continue
        day_obj = db_data["daily_data"][d_str]

        for shift in ["gunduz", "gece"]:
            for item in day_obj.get(shift, {}).get("extruders", []):
                p_name = item.get("product", "").lower()
                qty = item.get("qty", 0)
                if "pervaz" in p_name:
                    pervaz_prod += qty
                elif "kasa" in p_name:
                    kasa_prod += qty
                elif "seren" in p_name:
                    seren_prod += qty

            for item in day_obj.get(shift, {}).get("levha", []):
                qty = item.get("qty", 0)
                levha_prod += qty

    pervaz_req = 5.0
    kasa_req = 2.5
    seren_req = 3.5
    levha_req = 2.0

    pervaz_eq = pervaz_prod / pervaz_req if pervaz_req else 0
    kasa_eq = kasa_prod / kasa_req if kasa_req else 0
    seren_eq = seren_prod / seren_req if seren_req else 0
    levha_eq = levha_prod / levha_req if levha_req else 0

    completable_doors = math.floor(min(pervaz_eq, kasa_eq, seren_eq, levha_eq)) if (pervaz_prod or kasa_prod or seren_prod or levha_prod) else 0

    pervaz_used = completable_doors * pervaz_req
    kasa_used = completable_doors * kasa_req
    seren_used = completable_doors * seren_req
    levha_used = completable_doors * levha_req

    pervaz_carry = pervaz_prod - pervaz_used
    kasa_carry = kasa_prod - kasa_used
    seren_carry = seren_prod - seren_used
    levha_carry = levha_prod - levha_used

    return {
        "completable_doors": completable_doors,
        "details": {
            "pervaz": {
                "produced": pervaz_prod,
                "req_per_door": pervaz_req,
                "door_eq": round(pervaz_eq, 2),
                "used": pervaz_used,
                "carryover": round(pervaz_carry, 2)
            },
            "kasa": {
                "produced": kasa_prod,
                "req_per_door": kasa_req,
                "door_eq": round(kasa_eq, 2),
                "used": kasa_used,
                "carryover": round(kasa_carry, 2)
            },
            "seren": {
                "produced": seren_prod,
                "req_per_door": seren_req,
                "door_eq": round(seren_eq, 2),
                "used": seren_used,
                "carryover": round(seren_carry, 2)
            },
            "levha": {
                "produced": levha_prod,
                "req_per_door": levha_req,
                "door_eq": round(levha_eq, 2),
                "used": levha_used,
                "carryover": round(levha_carry, 2)
            }
        }
    }


def parse_date_label(date_label):
    """DD.MM.YYYY formatındaki tarih etiketini datetime objesine çevirir. Ayrıştırılamazsa None döner."""
    try:
        return datetime.strptime(date_label, "%d.%m.%Y")
    except (ValueError, TypeError):
        return None


def get_sorted_day_keys(daily_data):
    """Günleri, kayıt anahtarına değil GERÇEK TAKVİM TARİHİNE göre kronolojik sıraya dizer.
    Bu sayede günler ay/yıl sınırı olmadan (Ağustos->Eylül->...) ve hangi sırayla
    eklenmiş olursa olsun her zaman doğru sıralanır."""
    def sort_key(k):
        dt = parse_date_label(daily_data.get(k, {}).get("date", ""))
        if dt:
            return (0, dt)
        return (1, int(k) if k.isdigit() else 999999)
    return sorted(daily_data.keys(), key=sort_key)


def get_day_category_qty(day_obj):
    """Bir günün ham üretim adetlerini kategori bazında döndürür (pervaz, kasa, seren, levha).
    Kapı devir zincirinde (gün gün kümülatif aktarım) kullanılır."""
    qty = {"pervaz": 0.0, "kasa": 0.0, "seren": 0.0, "levha": 0.0}
    for shift in ["gunduz", "gece"]:
        for item in day_obj.get(shift, {}).get("extruders", []):
            p_name = item.get("product", "").lower()
            q = item.get("qty", 0)
            if "pervaz" in p_name:
                qty["pervaz"] += q
            elif "kasa" in p_name:
                qty["kasa"] += q
            elif "seren" in p_name:
                qty["seren"] += q
        for item in day_obj.get(shift, {}).get("levha", []):
            qty["levha"] += item.get("qty", 0)
    return qty


@app.get("/api/dashboard")
def get_dashboard_summary():
    data = load_data()
    total_prod_kg = 0.0
    total_fire_kg = 0.0
    total_employees = 0
    total_hours = 0.0
    total_downtime_min = 0.0

    machine_totals = {}
    fire_reasons_summary = {}

    daily_chart = []
    sorted_keys = get_sorted_day_keys(data["daily_data"])

    # Kapı kapasitesi DEVİR ZİNCİRİ: bir günün fazlası, sıradaki güne (kronolojik
    # sırada) taşınır. Kategori bazında koşan (running) bakiye.
    door_req = {"pervaz": 5.0, "kasa": 2.5, "seren": 3.5, "levha": 2.0}
    running_carryover = {"pervaz": 0.0, "kasa": 0.0, "seren": 0.0, "levha": 0.0}

    for d_str in sorted_keys:
        day_obj = data["daily_data"][d_str]

        day_prod_kg = 0.0
        day_fire_kg = 0.0
        day_emp = 0
        day_hours = 0.0
        day_machine_totals = {}  # Bu güne özel makine/hat bazında kırılım
        day_shifts = {
            "gunduz": {"employees": 0, "hours": 0.0, "prod_kg": 0.0, "fire_kg": 0.0},
            "gece": {"employees": 0, "hours": 0.0, "prod_kg": 0.0, "fire_kg": 0.0}
        }

        for shift in ["gunduz", "gece"]:
            s_data = day_obj.get(shift, {})
            s_emp = s_data.get("employees", 10)
            s_hours = s_data.get("hours", 12)
            day_emp += s_emp
            day_hours += s_hours
            day_shifts[shift]["employees"] = s_emp
            day_shifts[shift]["hours"] = s_hours

            for ext in s_data.get("extruders", []):
                p_kg = ext.get("prod_kg", 0)
                f_kg = ext.get("fire_kg", 0)
                h_name = ext.get("hat", "Bilinmeyen Hat")
                h_product = ext.get("product", "")
                h_hours = ext.get("hours", 0)
                h_qty = ext.get("qty", 0)

                day_prod_kg += p_kg
                day_fire_kg += f_kg
                day_shifts[shift]["prod_kg"] += p_kg
                day_shifts[shift]["fire_kg"] += f_kg

                if h_name not in machine_totals:
                    machine_totals[h_name] = {"prod": 0.0, "fire": 0.0}
                machine_totals[h_name]["prod"] += p_kg
                machine_totals[h_name]["fire"] += f_kg

                dm_key = f"ext_{h_name}"
                if dm_key not in day_machine_totals:
                    day_machine_totals[dm_key] = {"hat": h_name, "type": "Ekstrüder", "products": {}, "prod_kg": 0.0, "fire_kg": 0.0, "hours": 0.0}
                if h_product:
                    if h_product not in day_machine_totals[dm_key]["products"]:
                        day_machine_totals[dm_key]["products"][h_product] = {"qty": 0, "prod_kg": 0.0, "fire_kg": 0.0}
                    day_machine_totals[dm_key]["products"][h_product]["qty"] += h_qty
                    day_machine_totals[dm_key]["products"][h_product]["prod_kg"] += p_kg
                    day_machine_totals[dm_key]["products"][h_product]["fire_kg"] += f_kg
                day_machine_totals[dm_key]["prod_kg"] += p_kg
                day_machine_totals[dm_key]["fire_kg"] += f_kg
                day_machine_totals[dm_key]["hours"] += h_hours

            for lev in s_data.get("levha", []):
                p_kg = lev.get("total_kg", 0)
                f_kg = lev.get("dead_fire_kg", 0)
                h_name = lev.get("hat", "Levha Hattı")
                h_product = lev.get("color", "")  # Levha'da ürün/varyant bilgisi 'color' (Renk/Model) alanında tutulur
                h_qty = lev.get("qty", 0)

                day_prod_kg += p_kg
                day_fire_kg += f_kg
                day_shifts[shift]["prod_kg"] += p_kg
                day_shifts[shift]["fire_kg"] += f_kg

                if h_name not in machine_totals:
                    machine_totals[h_name] = {"prod": 0.0, "fire": 0.0}
                machine_totals[h_name]["prod"] += p_kg
                machine_totals[h_name]["fire"] += f_kg

                dm_key = f"lev_{h_name}"
                if dm_key not in day_machine_totals:
                    # Levha satırlarında ayrı bir çalışma süresi alanı yok; verimlilik hesabında
                    # o vardiyanın toplam çalışma saati yaklaşık değer olarak kullanılır.
                    day_machine_totals[dm_key] = {"hat": h_name, "type": "Levha", "products": {}, "prod_kg": 0.0, "fire_kg": 0.0, "hours": 0.0}
                if h_product:
                    if h_product not in day_machine_totals[dm_key]["products"]:
                        day_machine_totals[dm_key]["products"][h_product] = {"qty": 0, "prod_kg": 0.0, "fire_kg": 0.0}
                    day_machine_totals[dm_key]["products"][h_product]["qty"] += h_qty
                    day_machine_totals[dm_key]["products"][h_product]["prod_kg"] += p_kg
                    day_machine_totals[dm_key]["products"][h_product]["fire_kg"] += f_kg
                day_machine_totals[dm_key]["prod_kg"] += p_kg
                day_machine_totals[dm_key]["fire_kg"] += f_kg
                day_machine_totals[dm_key]["hours"] = max(day_machine_totals[dm_key]["hours"], s_hours)

        day_downtime_min = 0.0
        day_fire_reasons = {}   # Bu güne özel fire/duruş sebepleri kırılımı
        day_downtimes_list = []  # Bu güne özel ham duruş kayıtları
        for dt in day_obj.get("downtimes", []):
            dt_min = dt.get("down_min", 0)
            r_reason = dt.get("fire_reason") or dt.get("down_reason") or "Diğer Sebepler"
            r_fire = dt.get("fire_kg", 0)

            day_downtime_min += dt_min
            total_downtime_min += dt_min

            if r_reason not in fire_reasons_summary:
                fire_reasons_summary[r_reason] = 0.0
            fire_reasons_summary[r_reason] += r_fire

            if r_reason not in day_fire_reasons:
                day_fire_reasons[r_reason] = {"fire_kg": 0.0, "down_min": 0.0}
            day_fire_reasons[r_reason]["fire_kg"] += r_fire
            day_fire_reasons[r_reason]["down_min"] += dt_min

            day_downtimes_list.append({
                "machine": dt.get("machine") or dt.get("hat") or "-",
                "reason": r_reason,
                "down_min": dt_min,
                "fire_kg": r_fire,
                "note": dt.get("note", "")
            })

        total_prod_kg += day_prod_kg
        total_fire_kg += day_fire_kg
        total_employees += day_emp
        total_hours += day_hours

        fire_ratio = (day_fire_kg / (day_prod_kg + day_fire_kg) * 100) if (day_prod_kg + day_fire_kg) > 0 else 0
        date_label = day_obj.get("date") or (f"{int(d_str):02d}.08.2026" if d_str.isdigit() else d_str)

        # Bu güne ait makine/hat bazında kırılım listesi (üretim çoktan aza sıralı)
        day_machines_list = []
        for dm in day_machine_totals.values():
            dm_prod = round(dm["prod_kg"], 2)
            dm_fire = round(dm["fire_kg"], 2)
            dm_fire_ratio = round((dm_fire / (dm_prod + dm_fire) * 100), 2) if (dm_prod + dm_fire) > 0 else 0
            dm_hours = round(dm.get("hours", 0), 2)
            dm_kg_per_hour_net = round((dm_prod / dm_hours), 2) if dm_hours > 0 else 0
            dm_kg_per_hour_gross = round(((dm_prod + dm_fire) / dm_hours), 2) if dm_hours > 0 else 0

            products_list = sorted(
                [
                    {
                        "name": p_name,
                        "qty": p_info["qty"],
                        "prod_kg": round(p_info["prod_kg"], 2),
                        "fire_kg": round(p_info["fire_kg"], 2)
                    }
                    for p_name, p_info in dm["products"].items()
                ],
                key=lambda x: x["prod_kg"], reverse=True
            )
            products_summary = ", ".join(f"{p['name']} ({p['qty']:g} adet)" for p in products_list) if products_list else "-"

            day_machines_list.append({
                "hat": dm["hat"],
                "type": dm["type"],
                "products": products_summary,
                "products_detail": products_list,
                "prod_kg": dm_prod,
                "fire_kg": dm_fire,
                "fire_ratio": dm_fire_ratio,
                "hours": dm_hours,
                "kg_per_hour": dm_kg_per_hour_net,
                "kg_per_hour_gross": dm_kg_per_hour_gross
            })
        day_machines_list.sort(key=lambda x: x["prod_kg"], reverse=True)

        # EKSTRÜDER ve LEVHA hatlarını ayrı ayrı grupla, her biri için kendi alt toplamını hesapla
        def build_type_summary(machines_of_type):
            t_prod = sum(m["prod_kg"] for m in machines_of_type)
            t_fire = sum(m["fire_kg"] for m in machines_of_type)
            t_hours = sum(m["hours"] for m in machines_of_type)
            t_fire_ratio = round((t_fire / (t_prod + t_fire) * 100), 2) if (t_prod + t_fire) > 0 else 0
            t_kg_per_hour_net = round((t_prod / t_hours), 2) if t_hours > 0 else 0
            t_kg_per_hour_gross = round(((t_prod + t_fire) / t_hours), 2) if t_hours > 0 else 0
            return {
                "machines": machines_of_type,
                "total_prod_kg": round(t_prod, 2),
                "total_fire_kg": round(t_fire, 2),
                "fire_ratio": t_fire_ratio,
                "total_hours": round(t_hours, 2),
                "kg_per_hour_net": t_kg_per_hour_net,
                "kg_per_hour_gross": t_kg_per_hour_gross
            }

        day_extruder_summary = build_type_summary([m for m in day_machines_list if m["type"] == "Ekstrüder"])
        day_levha_summary = build_type_summary([m for m in day_machines_list if m["type"] == "Levha"])

        # Genel toplam (Ekstrüder + Levha) — iki ayrı özetin birleşik sonucu
        day_combined_summary = {
            "total_prod_kg": round(day_extruder_summary["total_prod_kg"] + day_levha_summary["total_prod_kg"], 2),
            "total_fire_kg": round(day_extruder_summary["total_fire_kg"] + day_levha_summary["total_fire_kg"], 2),
        }
        _cp = day_combined_summary["total_prod_kg"]
        _cf = day_combined_summary["total_fire_kg"]
        day_combined_summary["fire_ratio"] = round((_cf / (_cp + _cf) * 100), 2) if (_cp + _cf) > 0 else 0
        _ch = round(day_extruder_summary["total_hours"] + day_levha_summary["total_hours"], 2)
        day_combined_summary["total_hours"] = _ch
        day_combined_summary["kg_per_hour_net"] = round((_cp / _ch), 2) if _ch > 0 else 0
        day_combined_summary["kg_per_hour_gross"] = round(((_cp + _cf) / _ch), 2) if _ch > 0 else 0

        # Bu güne ait fire/duruş sebepleri kırılımı (fire kg'ye göre çoktan aza)
        day_fire_reasons_list = sorted(
            [{"reason": k, "fire_kg": round(v["fire_kg"], 2), "down_min": round(v["down_min"], 1)} for k, v in day_fire_reasons.items()],
            key=lambda x: x["fire_kg"], reverse=True
        )

        # Bu güne özel kapı kapasitesi/reçete eşdeğeri hesabı — DEVİR ZİNCİRİ:
        # önceki günden gelen fazlalık (running_carryover) bugünün üretimine eklenir,
        # tamamlanan kapılar düşüldükten sonra kalan fazlalık bir sonraki güne aktarılır.
        today_qty = get_day_category_qty(day_obj)
        available = {cat: running_carryover[cat] + today_qty[cat] for cat in door_req}
        eq = {cat: (available[cat] / door_req[cat] if door_req[cat] else 0) for cat in door_req}
        has_any = any(available[cat] > 0 for cat in door_req)
        day_completable_doors = math.floor(min(eq.values())) if has_any else 0
        used = {cat: day_completable_doors * door_req[cat] for cat in door_req}
        carryover_out = {cat: available[cat] - used[cat] for cat in door_req}

        day_door_stats = {
            "completable_doors": day_completable_doors,
            "details": {
                cat: {
                    "produced": today_qty[cat],
                    "carryover_in": round(running_carryover[cat], 2),
                    "available": round(available[cat], 2),
                    "req_per_door": door_req[cat],
                    "door_eq": round(eq[cat], 2),
                    "used": used[cat],
                    "carryover": round(carryover_out[cat], 2)
                } for cat in door_req
            }
        }

        # Bir sonraki güne devret
        running_carryover = carryover_out

        # Bu güne özel verimlilik metrikleri
        day_kg_per_employee = round((day_prod_kg / day_emp), 2) if day_emp > 0 else 0
        day_kg_per_hour = round((day_prod_kg / day_hours), 2) if day_hours > 0 else 0

        # Vardiya bazında detaylı verimlilik (Kg/Çalışan, Net + Brüt kg/saat, Fire Oranı)
        def build_shift_detail(s):
            s_prod = s["prod_kg"]
            s_fire = s["fire_kg"]
            s_emp = s["employees"]
            s_hours = s["hours"]
            s_fire_ratio = round((s_fire / (s_prod + s_fire) * 100), 2) if (s_prod + s_fire) > 0 else 0
            s_kg_per_employee = round((s_prod / s_emp), 2) if s_emp > 0 else 0
            s_kg_per_hour_net = round((s_prod / s_hours), 2) if s_hours > 0 else 0
            s_kg_per_hour_gross = round(((s_prod + s_fire) / s_hours), 2) if s_hours > 0 else 0
            return {
                "employees": s_emp,
                "hours": round(s_hours, 2),
                "prod_kg": round(s_prod, 2),
                "fire_kg": round(s_fire, 2),
                "fire_ratio": s_fire_ratio,
                "kg_per_employee": s_kg_per_employee,
                "kg_per_hour_net": s_kg_per_hour_net,
                "kg_per_hour_gross": s_kg_per_hour_gross
            }

        daily_chart.append({
            "key": d_str,
            "date": date_label,
            "prod_kg": round(day_prod_kg, 2),
            "fire_kg": round(day_fire_kg, 2),
            "fire_ratio": round(fire_ratio, 2),
            "employees": day_emp,
            "hours": round(day_hours, 2),
            "downtime_min": round(day_downtime_min, 2),
            "kg_per_employee": day_kg_per_employee,
            "kg_per_hour": day_kg_per_hour,
            "machines": day_machines_list,
            "extruder_summary": day_extruder_summary,
            "levha_summary": day_levha_summary,
            "combined_summary": day_combined_summary,
            "shifts": {
                "gunduz": build_shift_detail(day_shifts["gunduz"]),
                "gece": build_shift_detail(day_shifts["gece"])
            },
            "door_stats": day_door_stats,
            "fire_reasons": day_fire_reasons_list,
            "downtimes": day_downtimes_list
        })

    # Toplam kayıtlı gün sayısı (varsayılan görüntülenecek gün = verisi olan en güncel gün)
    latest_day_key = sorted_keys[-1] if sorted_keys else None
    days_with_data = [d["key"] for d in daily_chart if (d["prod_kg"] > 0 or d["fire_kg"] > 0)]
    if days_with_data:
        latest_day_key = days_with_data[-1]

    # Weekly summary — AY/YIL SINIRI YOK: kronolojik sıradaki günler 7'şerli gruplara
    # ayrılır (takvim ayına göre sabit aralıklar yerine). Böylece Ağustos bittiğinde
    # Eylül (ve sonrası) günleri de sorunsuz şekilde haftalara dahil olur.
    weekly_summary = []
    chunk_size = 7
    for i in range(0, len(sorted_keys), chunk_size):
        w_keys = sorted_keys[i:i + chunk_size]
        if not w_keys:
            continue

        week_num = (i // chunk_size) + 1
        first_label = data["daily_data"][w_keys[0]].get("date", w_keys[0])
        last_label = data["daily_data"][w_keys[-1]].get("date", w_keys[-1])
        w_name = f"{week_num}. Hafta ({first_label} - {last_label})" if first_label != last_label else f"{week_num}. Hafta ({first_label})"

        w_prod = 0.0
        w_fire = 0.0
        w_emp = 0

        for k in w_keys:
            d_obj = data["daily_data"][k]
            for shift in ["gunduz", "gece"]:
                w_emp += d_obj.get(shift, {}).get("employees", 10)
                for ext in d_obj.get(shift, {}).get("extruders", []):
                    w_prod += ext.get("prod_kg", 0)
                    w_fire += ext.get("fire_kg", 0)
                for lev in d_obj.get(shift, {}).get("levha", []):
                    w_prod += lev.get("total_kg", 0)
                    w_fire += lev.get("dead_fire_kg", 0)

        w_fire_ratio = (w_fire / (w_prod + w_fire) * 100) if (w_prod + w_fire) > 0 else 0
        w_door_stats = compute_door_capacity(data, filter_date_keys=w_keys)

        weekly_summary.append({
            "name": w_name,
            "keys": w_keys,
            "prod_ton": round(w_prod / 1000.0, 2),
            "fire_ton": round(w_fire / 1000.0, 2),
            "fire_ratio": round(w_fire_ratio, 2),
            "employees": w_emp,
            "doors": w_door_stats["completable_doors"]
        })

    # Monthly Summary Details
    monthly_summary = {
        "title": "Ağustos 2026 Genel Aylık Özet",
        "total_days": len(sorted_keys),
        "total_prod_ton": round(total_prod_kg / 1000.0, 2),
        "total_fire_ton": round(total_fire_kg / 1000.0, 2),
        "fire_ratio": round((total_fire_kg / (total_prod_kg + total_fire_kg) * 100) if (total_prod_kg + total_fire_kg) > 0 else 0, 2),
        "total_employees": total_employees,
        "completable_doors": compute_door_capacity(data)["completable_doors"],
        "top_producing_machines": sorted([{"name": k, "prod_ton": round(v["prod"]/1000.0, 2), "fire_kg": round(v["fire"], 1)} for k, v in machine_totals.items()], key=lambda x: x["prod_ton"], reverse=True),
        "top_scrap_reasons": sorted([{"reason": k, "fire_kg": round(v, 1)} for k, v in fire_reasons_summary.items() if v > 0], key=lambda x: x["fire_kg"], reverse=True)
    }

    door_stats = compute_door_capacity(data)
    overall_fire_ratio = (total_fire_kg / (total_prod_kg + total_fire_kg) * 100) if (total_prod_kg + total_fire_kg) > 0 else 0
    kg_per_employee = (total_prod_kg / total_employees) if total_employees > 0 else 0
    kg_per_hour = (total_prod_kg / total_hours) if total_hours > 0 else 0

    return {
        "company": "ERGUNBAS Group",
        "total_prod_ton": round(total_prod_kg / 1000.0, 2),
        "total_fire_ton": round(total_fire_kg / 1000.0, 2),
        "overall_fire_ratio": round(overall_fire_ratio, 2),
        "total_employees": total_employees,
        "kg_per_employee": round(kg_per_employee, 2),
        "kg_per_hour": round(kg_per_hour, 2),
        "total_downtime_min": round(total_downtime_min, 1),
        "door_stats": door_stats,
        "daily_chart": daily_chart,
        "weekly_summary": weekly_summary,
        "monthly_summary": monthly_summary,
        "available_dates": [{"key": k, "date": data["daily_data"][k].get("date", k)} for k in sorted_keys],
        "latest_day_key": latest_day_key,
        "machines_count": len(data.get("machines", [])),
        "products_count": len(data.get("products", []))
    }

@app.get("/api/formulas")
def get_formulas():
    return {
        "extruder_formulas": {
            "fire_ratio": "Fire Oranı (%) = [ Fire (kg) / (Üretim kg + Fire kg) ] * 100",
            "set_count_pervaz": "Pervaz Takım Sayısı = Üretim Adedi / 5.0",
            "set_count_kasa": "Kasa Takım Sayısı = Üretim Adedi / 2.5",
            "set_count_seren": "Seren Takım Sayısı = Üretim Adedi / 3.5",
        },
        "levha_formulas": {
            "m2_one": "1 Levha m² = (En_cm / 100) * (Boy_cm / 100)",
            "total_m2": "Toplam m² = 1 Levha m² * Üretim Adedi",
            "width_fire_cm": "En Firesi (cm) = MAX(0, Kalıp Çıkış Eni - En_cm)",
            "dead_fire_m2": "Ölü Fire m² = (En Firesi / 100) * (Boy_cm / 100) * Üretim Adedi",
            "dead_fire_kg": "Ölü Fire (kg) = Ölü Fire m² * m² Ağırlığı (kg/m²)",
            "total_kg": "Toplam Üretim (kg) = Toplam m² * m² Ağırlığı (kg/m²)",
            "set_count_levha": "Levha Takım Sayısı = Üretim Adedi / 2.0"
        },
        "door_capacity_formulas": {
            "door_equivalent": "Kapı Eşdeğeri = Üretim Adedi / Reçete Katsayısı",
            "completable_doors": "Tamamlanabilir Kapı = FLOOR( MIN(Pervaz_Eşdeğeri, Kasa_Eşdeğeri, Seren_Eşdeğeri, Levha_Eşdeğeri) )",
            "used_qty": "Kullanılan Adet = Tamamlanan Kapı * Reçete Katsayısı",
            "carryover_stock": "Ertesi Güne Devir = Toplam Mevcut Adet - Kullanılan Adet"
        },
        "efficiency_formulas": {
            "kg_per_employee": "Kg / Çalışan = Toplam Üretim (kg) / Vardiyadaki Çalışan Sayısı",
            "kg_per_hour_net": "Net Verimlilik (kg/saat) = Sağlam Üretim (kg) / Çalışma Saati — Fire HARİÇ, satılabilir çıktıyı ölçer",
            "kg_per_hour_gross": "Brüt Verimlilik (kg/saat) = (Sağlam Üretim (kg) + Fire (kg)) / Çalışma Saati — makinenin toplam işlem hızını ölçer, fire DAHİL",
            "efficiency_note": "İkisi birlikte değerlendirilir: Net verimlilik düşükken Brüt verimlilik yüksekse sorun HIZ değil FIRE'dır (Fire Oranı sütununa bakın). İkisi de düşükse makine hattı gerçekten yavaş çalışıyordur."
        }
    }

@app.get("/api/machines")
def get_machines():
    data = load_data()
    return data.get("machines", [])

@app.post("/api/machines")
def add_machine(machine: MachineCreate, x_username: Optional[str] = Header(None)):
    require_editor(x_username)
    data = load_data()
    new_id = f"{machine.type[:3]}_{len(data['machines']) + 1}"
    new_machine = {
        "id": new_id,
        "name": machine.name,
        "type": machine.type
    }
    data["machines"].append(new_machine)
    save_data(data)
    return new_machine

@app.delete("/api/machines/{machine_id}")
def delete_machine(machine_id: str, x_username: Optional[str] = Header(None)):
    require_editor(x_username)
    data = load_data()
    data["machines"] = [m for m in data["machines"] if m["id"] != machine_id]
    save_data(data)
    return {"status": "success"}

@app.get("/api/products")
def get_products():
    data = load_data()
    return data.get("products", [])

@app.post("/api/products")
def add_product(product: ProductCreate, x_username: Optional[str] = Header(None)):
    require_editor(x_username)
    data = load_data()
    new_id = f"p_{len(data['products']) + 1}"
    new_prod = {
        "id": new_id,
        "name": product.name,
        "category": product.category,
        "door_ratio": product.door_ratio
    }
    data["products"].append(new_prod)
    save_data(data)
    return new_prod

# =====================================================================
# USER MANAGEMENT ENDPOINTS
# =====================================================================

class LoginRequest(BaseModel):
    username: str
    password: str

class UserCreate(BaseModel):
    username: str
    password: str
    role: str  # 'admin' or 'operator'
    name: str

class UserUpdate(BaseModel):
    password: Optional[str] = None
    role: Optional[str] = None
    name: Optional[str] = None

@app.post("/api/auth/login")
def login(req: LoginRequest):
    users_data = load_users()
    for u in users_data["users"]:
        if u["username"] == req.username and u["password"] == req.password:
            return {
                "status": "success",
                "user": {
                    "id": u["id"],
                    "username": u["username"],
                    "role": u["role"],
                    "name": u["name"]
                }
            }
    raise HTTPException(status_code=401, detail="Kullanıcı adı veya şifre hatalı!")

@app.get("/api/users")
def get_users():
    users_data = load_users()
    # Don't expose passwords in list
    return [{"id": u["id"], "username": u["username"], "role": u["role"], "name": u["name"]} for u in users_data["users"]]

@app.post("/api/users")
def create_user(user: UserCreate, x_username: Optional[str] = Header(None)):
    users_data = load_users()

    # Sadece 'admin' rolündeki kullanıcılar yeni kullanıcı ekleyebilir
    if x_username:
        acting_user = next((u for u in users_data["users"] if u.get("username") == x_username), None)
        if acting_user and acting_user.get("role") != "admin":
            raise HTTPException(status_code=403, detail="Kullanıcı eklemek için yönetici (admin) yetkisi gereklidir.")

    # Check if username already exists
    for u in users_data["users"]:
        if u["username"] == user.username:
            raise HTTPException(status_code=400, detail="Bu kullanıcı adı zaten mevcut!")

    # ID çakışmasını önlemek için mevcut en yüksek numaradan devam et (silme sonrası sayıya dayalı çakışma riskini ortadan kaldırır)
    existing_nums = []
    for u in users_data["users"]:
        m = re.match(r"^u(\d+)", u.get("id", ""))
        if m:
            existing_nums.append(int(m.group(1)))
    next_num = (max(existing_nums) + 1) if existing_nums else 1
    new_id = f"u{next_num}_{user.username[:3]}"

    new_user = {
        "id": new_id,
        "username": user.username,
        "password": user.password,
        "role": user.role,
        "name": user.name
    }
    users_data["users"].append(new_user)
    save_users(users_data)
    return {"status": "success", "id": new_id}

@app.delete("/api/users/{user_id}")
def delete_user(user_id: str, x_username: Optional[str] = Header(None)):
    users_data = load_users()

    if x_username:
        acting_user = next((u for u in users_data["users"] if u.get("username") == x_username), None)
        if acting_user and acting_user.get("role") != "admin":
            raise HTTPException(status_code=403, detail="Kullanıcı silmek için yönetici (admin) yetkisi gereklidir.")

    if user_id == "u1":
        raise HTTPException(status_code=400, detail="Admin kullanıcısı silinemez!")
    users_data["users"] = [u for u in users_data["users"] if u["id"] != user_id]
    save_users(users_data)
    return {"status": "success"}

@app.put("/api/users/{user_id}")
def update_user(user_id: str, update: UserUpdate, x_username: Optional[str] = Header(None)):
    users_data = load_users()

    if x_username:
        acting_user = next((u for u in users_data["users"] if u.get("username") == x_username), None)
        if acting_user and acting_user.get("role") != "admin":
            raise HTTPException(status_code=403, detail="Kullanıcı güncellemek için yönetici (admin) yetkisi gereklidir.")

    for u in users_data["users"]:
        if u["id"] == user_id:
            if update.password:
                u["password"] = update.password
            if update.role:
                u["role"] = update.role
            if update.name:
                u["name"] = update.name
            save_users(users_data)
            return {"status": "success"}
    raise HTTPException(status_code=404, detail="Kullanıcı bulunamadı!")

# =====================================================================
# LEVHA FIRE FORMUL API - Otomatik Hesaplama
# =====================================================================

class LevhaCalcRequest(BaseModel):
    width_cm: float         # Levha eni (cm)
    length_cm: float        # Levha boyu (cm)
    qty: int                # Adet
    kalip_cikis_eni: Optional[float] = None   # Kalıp çıkış eni (cm) - fire için
    kg_per_m2: Optional[float] = 3.5          # kg/m² ağırlık (varsayılan PVC levha)

@app.post("/api/calc/levha")
def calc_levha(req: LevhaCalcRequest):
    """Levha formüllerini otomatik hesapla:
    - 1 Levha m² = (En/100) * (Boy/100)
    - Toplam m² = 1_m2 * adet
    - En Firesi = MAX(0, kalip_cikis_eni - En)
    - Ölü Fire m² = (en_firesi/100) * (boy/100) * adet
    - Ölü Fire kg = olü_fire_m2 * kg_per_m2
    - Toplam kg = toplam_m2 * kg_per_m2
    """
    w = req.width_cm
    l = req.length_cm
    qty = req.qty
    kg_m2 = req.kg_per_m2 or 3.5

    m2_one = (w / 100.0) * (l / 100.0)
    total_m2 = round(m2_one * qty, 3)

    # En firesi
    en_firesi_cm = 0.0
    dead_fire_m2 = 0.0
    dead_fire_kg = 0.0
    if req.kalip_cikis_eni and req.kalip_cikis_eni > w:
        en_firesi_cm = req.kalip_cikis_eni - w
        dead_fire_m2 = round((en_firesi_cm / 100.0) * (l / 100.0) * qty, 3)
        dead_fire_kg = round(dead_fire_m2 * kg_m2, 2)

    total_kg = round(total_m2 * kg_m2, 2)
    sets = round(qty / 2.0, 1)

    return {
        "m2_one": round(m2_one, 4),
        "total_m2": total_m2,
        "en_firesi_cm": round(en_firesi_cm, 2),
        "dead_fire_m2": dead_fire_m2,
        "dead_fire_kg": dead_fire_kg,
        "total_kg": total_kg,
        "sets": sets,
        "kg_per_m2_used": kg_m2
    }

# =====================================================================
# EXTRUDER FIRE FORMUL API - Otomatik Hesaplama
# =====================================================================

class ExtruderCalcRequest(BaseModel):
    prod_kg: float
    fire_kg: float
    qty: int
    product_name: str = ""
    calisma_suresi_dk: Optional[float] = None  # Çalışma süresi (dakika)
    hiz_m_dk: Optional[float] = None           # Hız (m/dk)
    uzunluk_m: Optional[float] = None          # Profil uzunluğu (m)
    heads: Optional[float] = 1                 # Kafa sayısı (aynı anda çıkan ürün sayısı)

@app.post("/api/calc/extruder")
def calc_extruder(req: ExtruderCalcRequest):
    """Ekstrüder formüllerini otomatik hesapla"""
    prod_kg = req.prod_kg
    fire_kg = req.fire_kg
    qty = req.qty
    product_name = req.product_name.lower()

    total = prod_kg + fire_kg
    fire_ratio = round((fire_kg / total * 100), 2) if total > 0 else 0.0

    # Takım sayısı
    sets = 0.0
    if "pervaz" in product_name:
        sets = round(qty / 5.0, 1)
    elif "kasa" in product_name:
        sets = round(qty / 2.5, 1)
    elif "seren" in product_name:
        sets = round(qty / 3.5, 1)
    elif "levha" in product_name:
        sets = round(qty / 2.0, 1)

    # Teorik üretim (eğer hız ve süre verilmişse)
    heads = req.heads or 1
    teorik_adet = None
    if req.calisma_suresi_dk and req.hiz_m_dk and req.uzunluk_m and req.uzunluk_m > 0:
        teorik_uzunluk_m = req.hiz_m_dk * req.calisma_suresi_dk * heads
        teorik_adet = int(teorik_uzunluk_m / req.uzunluk_m)

    return {
        "fire_ratio": fire_ratio,
        "sets": sets,
        "teorik_adet": teorik_adet,
        "formula": {
            "fire_ratio": f"({fire_kg} / ({prod_kg} + {fire_kg})) * 100 = %{fire_ratio}",
            "sets": f"{qty} / oran = {sets} takım"
        }
    }


@app.get("/api/daily/{date_key}")
def get_daily_data(date_key: str):
    data = load_data()
    if date_key not in data["daily_data"]:
        date_label = f"{int(date_key):02d}.08.2026" if date_key.isdigit() else date_key
        data["daily_data"][date_key] = {
            "day": date_key,
            "date": date_label,
            "gunduz": {"employees": 10, "hours": 12, "extruders": [], "levha": []},
            "gece": {"employees": 10, "hours": 12, "extruders": [], "levha": []},
            "downtimes": []
        }
        save_data(data)

    return {
        "day_data": data["daily_data"][date_key],
        "door_stats": compute_door_capacity(data, filter_date_keys=[date_key])
    }

@app.post("/api/daily/add_date")
def add_new_date(req: AddDateRequest, x_username: Optional[str] = Header(None)):
    require_daily_operator(x_username)
    data = load_data()
    date_str = req.date_str.strip()

    # Aynı tarih zaten kayıtlıysa mevcut kaydı döndür (tekrar eklemeyi önle)
    existing_key = None
    for k, v in data["daily_data"].items():
        if v.get("date") == date_str:
            existing_key = k
            break

    if existing_key:
        key = existing_key
    else:
        # AY/YIL SINIRI YOK: her yeni gün, mevcut en yüksek gün numarasından devam eder.
        # Böylece Ağustos bittiğinde Eylül (ve sonrası) günleri sorunsuz sıraya eklenir,
        # kronolojik sıralama ve devir zinciri bozulmadan çalışmaya devam eder.
        existing_nums = [int(k) for k in data["daily_data"].keys() if k.isdigit()]
        key = str((max(existing_nums) + 1) if existing_nums else 1)

    if key not in data["daily_data"]:
        data["daily_data"][key] = {
            "day": key,
            "date": date_str,
            "gunduz": {"employees": 10, "hours": 12, "extruders": [], "levha": []},
            "gece": {"employees": 10, "hours": 12, "extruders": [], "levha": []},
            "downtimes": []
        }
        save_data(data)

    return {"status": "success", "key": key, "date": date_str}

@app.post("/api/daily/{date_key}")
def update_daily_data(date_key: str, update: DailyDataUpdate, x_username: Optional[str] = Header(None)):
    require_daily_operator(x_username)
    data = load_data()
    date_label = data["daily_data"].get(date_key, {}).get("date") or date_key

    data["daily_data"][date_key] = {
        "day": date_key,
        "date": date_label,
        "gunduz": update.gunduz.dict(),
        "gece": update.gece.dict(),
        "downtimes": [d.dict() for d in update.downtimes]
    }

    for shift in ["gunduz", "gece"]:
        for ext in data["daily_data"][date_key][shift]["extruders"]:
            p = ext.get("product", "").lower()
            q = ext.get("qty", 0)
            length = ext.get("length", 0) or 0
            speed = ext.get("speed", 0) or 0
            hours = ext.get("hours", 0) or 0
            hat_str = ext.get("hat", "")
            prod_kg = ext.get("prod_kg", 0) or 0

            # Otomatik Ekstrüder Fire Hesabı
            # Kafa Sayısı: bu hattan aynı anda kaç ürün çıkıyor (kullanıcı girer, varsayılan 1)
            heads = ext.get("heads") or 1
            if hours > 0 and speed > 0 and length > 0 and q > 0 and prod_kg > 0:
                teorik_m = hours * 60 * speed * heads
                net_m = q * length
                kg_m = prod_kg / net_m if net_m > 0 else 0
                if teorik_m > net_m and kg_m > 0:
                    ext["fire_kg"] = round((teorik_m - net_m) * kg_m, 2)

            if "pervaz" in p:
                ext["sets"] = round(q / 5.0, 1)
            elif "kasa" in p:
                ext["sets"] = round(q / 2.5, 1)
            elif "seren" in p:
                ext["sets"] = round(q / 3.5, 1)
            else:
                ext["sets"] = 0

        for lev in data["daily_data"][date_key][shift]["levha"]:
            q = lev.get("qty", 0)
            w = lev.get("width", 93)
            l = lev.get("length", 208)
            kg_m2 = lev.get("kg_per_m2", 3.5)
            kalip_eni = lev.get("kalip_cikis_eni", 0) or 0

            m2_one = (w / 100.0) * (l / 100.0)
            total_m2 = round(m2_one * q, 3)

            lev["m2_one"] = round(m2_one, 4)
            lev["total_m2"] = total_m2
            lev["total_kg"] = round(total_m2 * kg_m2, 2)
            lev["sets"] = round(q / 2.0, 1)
            lev["kg_per_m2"] = kg_m2
            lev["kalip_cikis_eni"] = kalip_eni

            # Ölü fire hesabı: kalıp çıkış eni > levha eni ise
            if kalip_eni > w:
                en_firesi = kalip_eni - w
                dead_m2 = round((en_firesi / 100.0) * (l / 100.0) * q, 3)
                lev["en_firesi_cm"] = round(en_firesi, 2)
                lev["dead_fire_m2"] = dead_m2
                lev["dead_fire_kg"] = round(dead_m2 * kg_m2, 2)
            else:
                lev["en_firesi_cm"] = 0.0
                lev["dead_fire_m2"] = 0.0
                lev["dead_fire_kg"] = 0.0

    save_data(data)
    return {"status": "success", "door_stats": compute_door_capacity(data, filter_date_keys=[date_key])}


@app.get("/api/export_excel")
def export_excel():
    data = load_data()
    wb = openpyxl.Workbook()
    
    ws_summary = wb.active
    ws_summary.title = "ERGUNBAS Yönetici Özeti"
    ws_summary.append(["ERGUNBAS GROUP - EKSTRÜDER VE LEVHA ÜRETİM RAPORU"])
    ws_summary.append([])
    ws_summary.append(["Tarih Key", "Tarih", "Gündüz Çalışan", "Gece Çalışan", "Toplam Üretim (kg)", "Toplam Fire (kg)", "Fire Oranı (%)"])

    sorted_keys = get_sorted_day_keys(data["daily_data"])
    
    for k in sorted_keys:
        d = data["daily_data"][k]
        g_emp = d.get("gunduz", {}).get("employees", 0)
        n_emp = d.get("gece", {}).get("employees", 0)

        tot_prod = 0.0
        tot_fire = 0.0

        for shift in ["gunduz", "gece"]:
            for e in d.get(shift, {}).get("extruders", []):
                tot_prod += e.get("prod_kg", 0)
                tot_fire += e.get("fire_kg", 0)
            for l in d.get(shift, {}).get("levha", []):
                tot_prod += l.get("total_kg", 0)
                tot_fire += l.get("dead_fire_kg", 0)

        fire_ratio = (tot_fire / (tot_prod + tot_fire) * 100) if (tot_prod + tot_fire) > 0 else 0
        ws_summary.append([k, d.get("date", k), g_emp, n_emp, round(tot_prod, 2), round(tot_fire, 2), round(fire_ratio, 2)])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    headers = {'Content-Disposition': 'attachment; filename="ERGUNBAS_Uretim_Raporu.xlsx"'}
    return Response(content=output.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)

# Serve Web Interface
@app.get("/", response_class=HTMLResponse)
def root():
    index_path = os.path.join(APP_DIR, "static", "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>ERGUNBAS Group Üretim Takip Sistemi</h1>"

# Mount static files
static_dir = os.path.join(APP_DIR, "static")
os.makedirs(static_dir, exist_ok=True)
app.mount("/static", StaticFiles(directory=static_dir), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
