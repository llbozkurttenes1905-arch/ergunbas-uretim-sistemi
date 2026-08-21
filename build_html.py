import base64
import os

APP_DIR = os.path.dirname(os.path.abspath(__file__))
logo_path = os.path.join(APP_DIR, "static", "logo.png")

with open(logo_path, "rb") as f:
    b64_logo = base64.b64encode(f.read()).decode("utf-8")

logo_data_uri = f"data:image/png;base64,{b64_logo}"

template = """<!DOCTYPE html>
<html lang="tr" class="h-full bg-slate-950 text-slate-100">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ERGUNBAS Group | Ekstrüder & Levha Canlı Üretim Yönetimi</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- Lucide Icons -->
    <script src="https://unpkg.com/lucide@latest"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        ergunbas: {
                            50: '#fff1f2',
                            100: '#ffe4e6',
                            500: '#f43f5e',
                            600: '#e11d48',
                            700: '#be123c',
                            800: '#9f1239',
                            900: '#881337',
                        }
                    }
                }
            }
        }
    </script>
    <style>
        .custom-scrollbar::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
            background: #0f172a;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
            background: #334155;
            border-radius: 3px;
        }
    </style>
</head>
<body class="h-full flex flex-col font-sans antialiased bg-slate-950 text-slate-100">

    <!-- LOGIN OVERLAY -->
    <div id="login-overlay" class="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-950/95 backdrop-blur-sm">
        <div class="bg-slate-900 border border-slate-700 rounded-2xl shadow-2xl p-8 w-full max-w-sm space-y-6">
            <div class="flex flex-col items-center gap-3">
                <div class="bg-white p-3 rounded-xl shadow-md">
                    <img src="{{LOGO_DATA_URI}}" alt="ERGUNBAS" class="h-14 w-auto object-contain">
                </div>
                <h2 class="text-2xl font-black text-white">ERGUNBAS Group</h2>
                <p class="text-xs text-slate-400 text-center">Üretim Takip Sistemine giriş yapın</p>
            </div>
            <form id="login-form" onsubmit="handleLogin(event)" class="space-y-4">
                <div>
                    <label class="block text-xs font-bold text-slate-400 mb-1.5 uppercase tracking-wider">Kullanıcı Adı</label>
                    <input type="text" id="login-username" required autocomplete="username"
                        class="w-full bg-slate-950 border border-slate-700 text-white rounded-xl px-4 py-3 text-sm focus:ring-2 focus:ring-ergunbas-500 focus:outline-none"
                        placeholder="Kullanıcı adınızı girin">
                </div>
                <div>
                    <label class="block text-xs font-bold text-slate-400 mb-1.5 uppercase tracking-wider">Şifre</label>
                    <input type="password" id="login-password" required autocomplete="current-password"
                        class="w-full bg-slate-950 border border-slate-700 text-white rounded-xl px-4 py-3 text-sm focus:ring-2 focus:ring-ergunbas-500 focus:outline-none"
                        placeholder="Şifrenizi girin">
                </div>
                <div id="login-error" class="hidden text-xs text-red-400 font-bold text-center bg-red-950/40 py-2 px-3 rounded-lg border border-red-800/50">
                    Kullanıcı adı veya şifre hatalı!
                </div>
                <button type="submit" class="w-full py-3 bg-ergunbas-600 hover:bg-ergunbas-500 text-white font-black rounded-xl transition shadow-lg shadow-ergunbas-600/30 flex items-center justify-center gap-2">
                    <i data-lucide="log-in" class="w-5 h-5"></i> Giriş Yap
                </button>
            </form>
        </div>
    </div>

    <!-- ERGUNBAS Corporate Header -->
    <header class="bg-slate-900 border-b border-slate-800 px-6 py-3.5 flex flex-wrap items-center justify-between gap-4 sticky top-0 z-50 shadow-xl">
        <div class="flex items-center space-x-4">
            <!-- Corporate Logo Embedded as Base64 -->
            <div class="bg-white p-2 rounded-xl shadow-md flex items-center justify-center border border-slate-200">
                <img src="{{LOGO_DATA_URI}}" alt="ERGUNBAS Group Logo" class="h-10 w-auto object-contain">
            </div>
            <div>
                <h1 class="text-xl font-black tracking-tight text-white flex items-center gap-2">
                    ERGUNBAS Group
                    <span class="text-xs font-bold uppercase tracking-wider px-2.5 py-0.5 rounded-md bg-ergunbas-600/20 text-ergunbas-500 border border-ergunbas-600/30">Üretim & Fire Yönetimi</span>
                </h1>
                <p class="text-xs text-slate-400 font-medium">Otomatik Hesaplama Motoru ve İnteraktif Yönetim Portalı</p>
            </div>
        </div>

        <!-- Navigation Tabs & Actions -->
        <div class="flex flex-wrap items-center gap-3">
            <nav class="flex space-x-1 bg-slate-950 p-1.5 rounded-xl border border-slate-800">
                <button onclick="switchTab('dashboard')" id="tab-dashboard" class="px-3.5 py-2 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5 bg-ergunbas-600 text-white shadow-lg shadow-ergunbas-600/30">
                    <i data-lucide="layout-dashboard" class="w-4 h-4"></i> Özetler (Aylık, Haftalık & Günlük)
                </button>
                <button onclick="switchTab('daily')" id="tab-daily" class="px-3.5 py-2 text-xs font-bold rounded-lg transition-all text-slate-400 hover:text-slate-200 flex items-center gap-1.5">
                    <i data-lucide="edit-3" class="w-4 h-4"></i> Günlük Veri Girişi
                </button>
                <button onclick="switchTab('doors')" id="tab-doors" class="px-3.5 py-2 text-xs font-bold rounded-lg transition-all text-slate-400 hover:text-slate-200 flex items-center gap-1.5">
                    <i data-lucide="door-open" class="w-4 h-4"></i> Kapı & Stok Takibi
                </button>
                <button onclick="switchTab('formulas')" id="tab-formulas" class="px-3.5 py-2 text-xs font-bold rounded-lg transition-all text-slate-400 hover:text-slate-200 flex items-center gap-1.5">
                    <i data-lucide="calculator" class="w-4 h-4 text-emerald-400"></i> Formüller & Mantık
                </button>
                <button onclick="switchTab('settings')" id="tab-settings" class="px-3.5 py-2 text-xs font-bold rounded-lg transition-all text-slate-400 hover:text-slate-200 flex items-center gap-1.5">
                    <i data-lucide="settings" class="w-4 h-4"></i> Makineler & Ürünler
                </button>
                <button onclick="switchTab('users')" id="tab-users" class="px-3.5 py-2 text-xs font-bold rounded-lg transition-all text-slate-400 hover:text-slate-200 flex items-center gap-1.5 admin-only" style="display:none">
                    <i data-lucide="users" class="w-4 h-4 text-violet-400"></i> Kullanıcılar
                </button>
            </nav>

            <!-- User info + Logout -->
            <div class="flex items-center gap-2">
                <div class="text-xs text-slate-400 font-medium hidden sm:block">
                    <span id="header-username" class="text-white font-bold"></span>
                    <span id="header-role" class="ml-1 text-slate-500"></span>
                </div>
                <button onclick="handleLogout()" class="px-3 py-2 bg-slate-800 hover:bg-red-900/50 text-slate-400 hover:text-red-400 font-bold text-xs rounded-xl transition flex items-center gap-1.5 border border-slate-700">
                    <i data-lucide="log-out" class="w-3.5 h-3.5"></i> Çıkış
                </button>
            </div>

            <a href="/api/export_excel" download class="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs rounded-xl transition flex items-center gap-1.5 shadow-md">
                <i data-lucide="download" class="w-4 h-4"></i> Excel İndir
            </a>
        </div>
    </header>

    <!-- Notification Toast -->
    <div id="toast" class="fixed bottom-6 right-6 z-50 bg-emerald-600 text-white px-5 py-3 rounded-xl shadow-2xl font-semibold text-sm flex items-center gap-2 hidden transition-all duration-300">
        <i data-lucide="check-circle" class="w-5 h-5"></i> <span id="toast-message">Veriler Başarıyla Kaydedildi!</span>
    </div>

    <!-- DETAIL MODAL -->
    <div id="modal-detail" class="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-md flex items-center justify-center p-4 hidden">
        <div class="bg-slate-900 border border-slate-800 rounded-3xl max-w-3xl w-full p-6 space-y-6 shadow-2xl max-h-[90vh] overflow-y-auto custom-scrollbar relative">
            <div class="flex justify-between items-start border-b border-slate-800 pb-4">
                <div>
                    <h3 class="text-xl font-black text-white flex items-center gap-2" id="modal-title">
                        <i data-lucide="info" class="w-6 h-6 text-ergunbas-500"></i> Detay Görünümü
                    </h3>
                    <p class="text-xs text-slate-400 mt-1" id="modal-subtitle">Seçilen dönemin detaylı üretimi, fireleri ve performans dökümü</p>
                </div>
                <button onclick="closeModal()" class="p-2 text-slate-400 hover:text-white rounded-xl hover:bg-slate-800 transition">
                    <i data-lucide="x" class="w-6 h-6"></i>
                </button>
            </div>
            <div id="modal-body" class="space-y-4 text-sm text-slate-300"></div>
            <div class="flex justify-end pt-4 border-t border-slate-800">
                <button onclick="closeModal()" class="px-5 py-2 bg-slate-800 hover:bg-slate-700 text-white font-bold text-xs rounded-xl transition">Kapat</button>
            </div>
        </div>
    </div>

    <!-- Main Content Area -->
    <main class="flex-1 p-6 overflow-y-auto custom-scrollbar space-y-6">

        <!-- TAB 1: DASHBOARD -->
        <div id="view-dashboard" class="space-y-6">

            <!-- AYLIK ÖZET BANNER CARD -->
            <div onclick="showMonthlyDetail()" class="bg-gradient-to-r from-ergunbas-900/60 via-slate-900 to-slate-900 border border-ergunbas-600/30 hover:border-ergunbas-500 rounded-2xl p-6 shadow-lg cursor-pointer transition flex flex-wrap items-center justify-between gap-6 group">
                <div class="space-y-1">
                    <span class="text-xs font-black uppercase tracking-wider text-ergunbas-500 flex items-center gap-1.5">
                        <i data-lucide="calendar" class="w-4 h-4"></i> Ağustos 2026 Genel Aylık Özet
                    </span>
                    <h2 class="text-2xl font-black text-white group-hover:text-ergunbas-400 transition flex items-center gap-2">
                        Aylık Üretim & Fire Raporu Detayı
                        <i data-lucide="chevron-right" class="w-5 h-5 text-slate-400 group-hover:translate-x-1 transition-transform"></i>
                    </h2>
                    <p class="text-xs text-slate-400">Tüm ekstrüder ve levha hatlarının 31 günlük toplam verilerini ve en çok üreten makineleri incelemek için tıklayın.</p>
                </div>
                <div class="flex items-center gap-4 bg-slate-950/60 px-5 py-3 rounded-xl border border-slate-800">
                    <div class="text-center px-3 border-r border-slate-800">
                        <span class="text-xs text-slate-400 block font-semibold">Toplam Üretim</span>
                        <span class="text-lg font-black text-white" id="monthly-banner-prod">Yükleniyor...</span>
                    </div>
                    <div class="text-center px-3 border-r border-slate-800">
                        <span class="text-xs text-slate-400 block font-semibold">Toplam Fire</span>
                        <span class="text-lg font-black text-amber-400" id="monthly-banner-fire">Yükleniyor...</span>
                    </div>
                    <div class="text-center px-3">
                        <span class="text-xs text-slate-400 block font-semibold">Tam Kapı</span>
                        <span class="text-lg font-black text-emerald-400" id="monthly-banner-doors">Yükleniyor...</span>
                    </div>
                </div>
            </div>

            <!-- Key Metric Cards -->
            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-5">
                <div onclick="showMonthlyDetail()" class="bg-slate-900 border border-slate-800 rounded-2xl p-5 relative overflow-hidden shadow-sm hover:border-ergunbas-600/40 cursor-pointer transition">
                    <div class="flex justify-between items-start">
                        <div>
                            <p class="text-xs font-bold text-slate-400 uppercase tracking-wider">Toplam Üretim Tonajı</p>
                            <h3 class="text-3xl font-black text-white mt-2" id="kpi-prod-ton">Yükleniyor...</h3>
                            <p class="text-xs text-ergunbas-500 mt-1 font-semibold flex items-center gap-1">
                                <i data-lucide="trending-up" class="w-3.5 h-3.5"></i> Ekstrüder + Levha Üretimi
                            </p>
                        </div>
                        <div class="p-3 bg-ergunbas-600/10 text-ergunbas-500 rounded-xl border border-ergunbas-600/20">
                            <i data-lucide="factory" class="w-6 h-6"></i>
                        </div>
                    </div>
                </div>

                <div onclick="showMonthlyDetail()" class="bg-slate-900 border border-slate-800 rounded-2xl p-5 relative overflow-hidden shadow-sm hover:border-amber-500/40 cursor-pointer transition">
                    <div class="flex justify-between items-start">
                        <div>
                            <p class="text-xs font-bold text-slate-400 uppercase tracking-wider">Toplam Fire Tonajı</p>
                            <h3 class="text-3xl font-black text-amber-400 mt-2" id="kpi-fire-ton">Yükleniyor...</h3>
                            <p class="text-xs text-slate-400 mt-1 font-semibold" id="kpi-fire-ratio">Fire Oranı: -</p>
                        </div>
                        <div class="p-3 bg-amber-500/10 text-amber-400 rounded-xl border border-amber-500/20">
                            <i data-lucide="flame" class="w-6 h-6"></i>
                        </div>
                    </div>
                </div>

                <div onclick="switchTab('doors')" class="bg-slate-900 border border-slate-800 rounded-2xl p-5 relative overflow-hidden shadow-sm hover:border-emerald-500/40 cursor-pointer transition">
                    <div class="flex justify-between items-start">
                        <div>
                            <p class="text-xs font-bold text-slate-400 uppercase tracking-wider">Tamamlanan Tam Kapı</p>
                            <h3 class="text-3xl font-black text-emerald-400 mt-2" id="kpi-doors">Yükleniyor...</h3>
                            <p class="text-xs text-emerald-400 mt-1 font-semibold">Stok & Reçete Eşdeğeri</p>
                        </div>
                        <div class="p-3 bg-emerald-500/10 text-emerald-400 rounded-xl border border-emerald-500/20">
                            <i data-lucide="door-open" class="w-6 h-6"></i>
                        </div>
                    </div>
                </div>

                <div onclick="showMonthlyDetail()" class="bg-slate-900 border border-slate-800 rounded-2xl p-5 relative overflow-hidden shadow-sm hover:border-purple-500/40 cursor-pointer transition">
                    <div class="flex justify-between items-start">
                        <div>
                            <p class="text-xs font-bold text-slate-400 uppercase tracking-wider">Çalışan Verimliliği</p>
                            <h3 class="text-3xl font-black text-purple-400 mt-2" id="kpi-efficiency">Yükleniyor...</h3>
                            <p class="text-xs text-slate-400 mt-1 font-medium" id="kpi-employees">Toplam Çalışan: -</p>
                        </div>
                        <div class="p-3 bg-purple-500/10 text-purple-400 rounded-xl border border-purple-500/20">
                            <i data-lucide="users" class="w-6 h-6"></i>
                        </div>
                    </div>
                </div>
            </div>

            <!-- HAFTALIK ÖZET TABLOSU -->
            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-sm">
                <div class="flex justify-between items-center mb-4 border-b border-slate-800 pb-3">
                    <h2 class="text-base font-bold text-white flex items-center gap-2">
                        <i data-lucide="calendar-range" class="w-5 h-5 text-ergunbas-500"></i> Haftalık Üretim ve Fire Özeti (Detay İçin Tıklayın)
                    </h2>
                    <span class="text-xs font-semibold text-slate-400">Ağustos 2026 Haftalık Kırılımları</span>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-5 gap-4" id="weekly-cards-container"></div>
            </div>

            <!-- Charts Section -->
            <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div class="lg:col-span-2 bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-sm">
                    <div class="flex justify-between items-center mb-4">
                        <h2 class="text-base font-bold text-white flex items-center gap-2">
                            <i data-lucide="line-chart" class="w-5 h-5 text-ergunbas-500"></i> Günlük Üretim Trendi (kg)
                        </h2>
                        <span class="text-xs text-slate-400">Barın Üzerine Tıklayıp Günlük Detayı Görün</span>
                    </div>
                    <div class="h-72">
                        <canvas id="chart-daily-prod"></canvas>
                    </div>
                </div>

                <div class="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-sm">
                    <div class="flex justify-between items-center mb-4">
                        <h2 class="text-base font-bold text-white flex items-center gap-2">
                            <i data-lucide="pie-chart" class="w-5 h-5 text-amber-400"></i> Fire Oranı Dağılımı (%)
                        </h2>
                    </div>
                    <div class="h-72">
                        <canvas id="chart-fire-ratio"></canvas>
                    </div>
                </div>
            </div>

            <!-- Door Component Equivalents Table -->
            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-sm">
                <h2 class="text-base font-bold text-white mb-4 flex items-center gap-2">
                    <i data-lucide="layers" class="w-5 h-5 text-emerald-400"></i> Kapı Üretim Kapasitesi ve Parça Devir Özeti
                </h2>
                <div class="overflow-x-auto">
                    <table class="w-full text-left text-sm text-slate-300">
                        <thead class="bg-slate-950 text-slate-400 uppercase text-xs">
                            <tr>
                                <th class="p-3">Ürün Grubu</th>
                                <th class="p-3">1 Kapı İhtiyacı (Reçete)</th>
                                <th class="p-3">Toplam Üretim (Adet)</th>
                                <th class="p-3">Kapı Eşdeğeri</th>
                                <th class="p-3">Kullanılan Miktar</th>
                                <th class="p-3">Ertesi Güne Devir</th>
                            </tr>
                        </thead>
                        <tbody id="tbl-door-summary" class="divide-y divide-slate-800"></tbody>
                    </table>
                </div>
            </div>
        </div>


        <!-- TAB 2: GÜNLÜK VERİ GİRİŞİ -->
        <div id="view-daily" class="space-y-6 hidden">
            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-4 flex flex-wrap items-center justify-between gap-4 shadow-md">
                <div class="flex flex-wrap items-center gap-3">
                    <label class="text-sm font-bold text-slate-300 flex items-center gap-2">
                        <i data-lucide="calendar" class="w-4 h-4 text-ergunbas-500"></i> Giriş Yapılacak Tarih:
                    </label>
                    <select id="select-day" onchange="loadDayData()" class="bg-slate-950 border border-slate-700 text-white rounded-xl px-4 py-2 text-sm font-bold focus:ring-2 focus:ring-ergunbas-500 focus:outline-none"></select>
                    <button onclick="promptAddNewDate()" class="px-3.5 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-bold rounded-xl transition border border-slate-700 flex items-center gap-1.5">
                        <i data-lucide="plus-circle" class="w-4 h-4 text-ergunbas-500"></i> Yeni Tarih / Gün Ekle
                    </button>
                </div>

                <div class="flex items-center space-x-3">
                    <span class="text-sm font-black text-slate-300" id="current-day-label">01.08.2026</span>
                    <button onclick="saveCurrentDay()" class="px-6 py-2.5 bg-ergunbas-600 hover:bg-ergunbas-500 text-white font-extrabold text-sm rounded-xl transition flex items-center gap-2 shadow-lg shadow-ergunbas-600/30">
                        <i data-lucide="save" class="w-4.5 h-4.5"></i> Tüm Verileri Kaydet
                    </button>
                </div>
            </div>

            <!-- GÜNDÜZ VARDIYASI -->
            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6 shadow-md">
                <div class="flex items-center justify-between border-b border-slate-800 pb-4">
                    <h3 class="text-lg font-black text-amber-400 flex items-center gap-2">
                        <i data-lucide="sun" class="w-5 h-5"></i> Gündüz Vardiyası Üretim Girişleri
                    </h3>
                    <div class="flex items-center space-x-4 text-sm">
                        <div class="flex items-center space-x-2">
                            <span class="text-slate-400 font-semibold">Sorumlu Operatör:</span>
                            <select id="gunduz-operator" class="bg-slate-950 border border-slate-700 rounded-lg px-3 py-1 text-white font-bold text-xs focus:ring-2 focus:ring-amber-500"></select>
                        </div>
                        <div class="flex items-center space-x-2">
                            <span class="text-slate-400 font-semibold">Çalışan Sayısı:</span>
                            <input type="number" id="gunduz-emp" value="10" class="w-20 bg-slate-950 border border-slate-700 rounded-lg px-3 py-1 text-white font-bold text-center focus:ring-2 focus:ring-amber-500">
                        </div>
                        <div class="flex items-center space-x-2">
                            <span class="text-slate-400 font-semibold">Süre (Saat):</span>
                            <input type="number" id="gunduz-hours" value="12" class="w-20 bg-slate-950 border border-slate-700 rounded-lg px-3 py-1 text-white font-bold text-center focus:ring-2 focus:ring-amber-500">
                        </div>
                    </div>
                </div>

                <div>
                    <div class="flex justify-between items-center mb-3">
                        <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Ekstrüder Hatları Üretimi</h4>
                        <button onclick="addExtruderRow('tbl-gunduz-ext')" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-xs font-bold text-sky-400 rounded-lg border border-slate-700 flex items-center gap-1">
                            <i data-lucide="plus" class="w-3.5 h-3.5"></i> Satır Ekle
                        </button>
                    </div>
                    <div class="overflow-x-auto custom-scrollbar">
                        <table class="w-full text-left text-sm text-slate-300">
                            <thead class="bg-slate-950 text-slate-400 uppercase text-xs">
                                <tr>
                                    <th class="p-2.5">Hat No</th>
                                    <th class="p-2.5">Ürün İsmi</th>
                                    <th class="p-2.5">Uzunluk (m)</th>
                                    <th class="p-2.5">Hız (m/dk)</th>
                                    <th class="p-2.5">Çalışma Süresi (sa)</th>
                                    <th class="p-2.5">Üretim (kg)</th>
                                    <th class="p-2.5">Fire (kg)</th>
                                    <th class="p-2.5">Üretim Adedi</th>
                                    <th class="p-2.5">Fire Oranı (%)</th>
                                    <th class="p-2.5">Takım Sayısı</th>
                                    <th class="p-2.5 text-center">Sil</th>
                                </tr>
                            </thead>
                            <tbody id="tbl-gunduz-ext" class="divide-y divide-slate-800"></tbody>
                        </table>
                    </div>
                </div>

                <div>
                    <div class="flex justify-between items-center mb-3">
                        <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Levha Hatları Üretimi</h4>
                        <button onclick="addLevhaRow('tbl-gunduz-lev')" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-xs font-bold text-amber-400 rounded-lg border border-slate-700 flex items-center gap-1">
                            <i data-lucide="plus" class="w-3.5 h-3.5"></i> Satır Ekle
                        </button>
                    </div>
                    <div class="overflow-x-auto custom-scrollbar">
                        <table class="w-full text-left text-sm text-slate-300">
                            <thead class="bg-slate-950 text-slate-400 uppercase text-xs">
                                <tr>
                                    <th class="p-2.5">Hat</th>
                                    <th class="p-2.5">Renk / Model</th>
                                    <th class="p-2.5">En (cm)</th>
                                    <th class="p-2.5">Boy (cm)</th>
                                    <th class="p-2.5">Kalıp Eni (cm)</th>
                                    <th class="p-2.5">kg/m²</th>
                                    <th class="p-2.5">Adet</th>
                                    <th class="p-2.5">Toplam m²</th>
                                    <th class="p-2.5">Ölü Fire (kg)</th>
                                    <th class="p-2.5">Toplam (kg)</th>
                                    <th class="p-2.5">Fire Oranı (%)</th>
                                    <th class="p-2.5">Takım Sayısı</th>
                                    <th class="p-2.5 text-center">Sil</th>
                                </tr>
                            </thead>
                            <tbody id="tbl-gunduz-lev" class="divide-y divide-slate-800"></tbody>
                        </table>
                    </div>
                </div>
            </div>


            <!-- GECE VARDIYASI -->
            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6 shadow-md">
                <div class="flex items-center justify-between border-b border-slate-800 pb-4">
                    <h3 class="text-lg font-black text-indigo-400 flex items-center gap-2">
                        <i data-lucide="moon" class="w-5 h-5"></i> Gece Vardiyası Üretim Girişleri
                    </h3>
                    <div class="flex items-center space-x-4 text-sm">
                        <div class="flex items-center space-x-2">
                            <span class="text-slate-400 font-semibold">Sorumlu Operatör:</span>
                            <select id="gece-operator" class="bg-slate-950 border border-slate-700 rounded-lg px-3 py-1 text-white font-bold text-xs focus:ring-2 focus:ring-indigo-500"></select>
                        </div>
                        <div class="flex items-center space-x-2">
                            <span class="text-slate-400 font-semibold">Çalışan Sayısı:</span>
                            <input type="number" id="gece-emp" value="10" class="w-20 bg-slate-950 border border-slate-700 rounded-lg px-3 py-1 text-white font-bold text-center focus:ring-2 focus:ring-indigo-500">
                        </div>
                        <div class="flex items-center space-x-2">
                            <span class="text-slate-400 font-semibold">Süre (Saat):</span>
                            <input type="number" id="gece-hours" value="12" class="w-20 bg-slate-950 border border-slate-700 rounded-lg px-3 py-1 text-white font-bold text-center focus:ring-2 focus:ring-indigo-500">
                        </div>
                    </div>
                </div>

                <div>
                    <div class="flex justify-between items-center mb-3">
                        <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Ekstrüder Hatları Üretimi</h4>
                        <button onclick="addExtruderRow('tbl-gece-ext')" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-xs font-bold text-sky-400 rounded-lg border border-slate-700 flex items-center gap-1">
                            <i data-lucide="plus" class="w-3.5 h-3.5"></i> Satır Ekle
                        </button>
                    </div>
                    <div class="overflow-x-auto custom-scrollbar">
                        <table class="w-full text-left text-sm text-slate-300">
                            <thead class="bg-slate-950 text-slate-400 uppercase text-xs">
                                <tr>
                                    <th class="p-2.5">Hat No</th>
                                    <th class="p-2.5">Ürün İsmi</th>
                                    <th class="p-2.5">Uzunluk (m)</th>
                                    <th class="p-2.5">Hız (m/dk)</th>
                                    <th class="p-2.5">Çalışma Süresi (sa)</th>
                                    <th class="p-2.5">Üretim (kg)</th>
                                    <th class="p-2.5">Fire (kg)</th>
                                    <th class="p-2.5">Üretim Adedi</th>
                                    <th class="p-2.5">Fire Oranı (%)</th>
                                    <th class="p-2.5">Takım Sayısı</th>
                                    <th class="p-2.5 text-center">Sil</th>
                                </tr>
                            </thead>
                            <tbody id="tbl-gece-ext" class="divide-y divide-slate-800"></tbody>
                        </table>
                    </div>
                </div>

                <div>
                    <div class="flex justify-between items-center mb-3">
                        <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Levha Hatları Üretimi</h4>
                        <button onclick="addLevhaRow('tbl-gece-lev')" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-xs font-bold text-amber-400 rounded-lg border border-slate-700 flex items-center gap-1">
                            <i data-lucide="plus" class="w-3.5 h-3.5"></i> Satır Ekle
                        </button>
                    </div>
                    <div class="overflow-x-auto custom-scrollbar">
                        <table class="w-full text-left text-sm text-slate-300">
                            <thead class="bg-slate-950 text-slate-400 uppercase text-xs">
                                <tr>
                                    <th class="p-2.5">Hat</th>
                                    <th class="p-2.5">Renk / Model</th>
                                    <th class="p-2.5">En (cm)</th>
                                    <th class="p-2.5">Boy (cm)</th>
                                    <th class="p-2.5">Kalıp Eni (cm)</th>
                                    <th class="p-2.5">kg/m²</th>
                                    <th class="p-2.5">Adet</th>
                                    <th class="p-2.5">Toplam m²</th>
                                    <th class="p-2.5">Ölü Fire (kg)</th>
                                    <th class="p-2.5">Toplam (kg)</th>
                                    <th class="p-2.5">Fire Oranı (%)</th>
                                    <th class="p-2.5">Takım Sayısı</th>
                                    <th class="p-2.5 text-center">Sil</th>
                                </tr>
                            </thead>
                            <tbody id="tbl-gece-lev" class="divide-y divide-slate-800"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- Fire ve Duruş Sebepleri Tablosu -->
            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-4 shadow-md">
                <div class="flex items-center justify-between">
                    <h3 class="text-base font-bold text-white flex items-center gap-2">
                        <i data-lucide="alert-triangle" class="w-5 h-5 text-amber-400"></i> Günlük Fire ve Duruş Sebepleri
                    </h3>
                    <button onclick="addDowntimeRow()" class="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-xs font-bold text-slate-200 rounded-lg transition flex items-center gap-1 border border-slate-700">
                        <i data-lucide="plus" class="w-3.5 h-3.5"></i> Kayıt Ekle
                    </button>
                </div>

                <div class="overflow-x-auto custom-scrollbar">
                    <table class="w-full text-left text-sm text-slate-300">
                        <thead class="bg-slate-950 text-slate-400 uppercase text-xs">
                            <tr>
                                <th class="p-2.5">Vardiya</th>
                                <th class="p-2.5">Hat No</th>
                                <th class="p-2.5">Fire Sebebi</th>
                                <th class="p-2.5">Fire (kg)</th>
                                <th class="p-2.5">Duruş Sebebi</th>
                                <th class="p-2.5">Duruş (dk)</th>
                                <th class="p-2.5">Açıklama</th>
                                <th class="p-2.5 text-center">İşlem</th>
                            </tr>
                        </thead>
                        <tbody id="tbl-downtimes" class="divide-y divide-slate-800"></tbody>
                    </table>
                </div>
            </div>
        </div>


        <!-- TAB 3: KAPI & STOK TAKİBİ -->
        <div id="view-doors" class="space-y-6 hidden">
            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6">
                <div class="flex justify-between items-center border-b border-slate-800 pb-4">
                    <div>
                        <h2 class="text-xl font-black text-white flex items-center gap-2">
                            <i data-lucide="door-open" class="w-6 h-6 text-emerald-400"></i> Kümülatif Kapı ve Parça Eşdeğerliği
                        </h2>
                        <p class="text-xs text-slate-400 mt-1">ERGUNBAS Group reçetesine göre ürün gruplarının tamamlayabildiği kapı miktarları ve stok devirleri</p>
                    </div>
                    <div class="bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 px-5 py-2.5 rounded-2xl text-center">
                        <span class="text-xs font-bold block uppercase tracking-wider">Tamamlanan Kapı</span>
                        <span class="text-3xl font-black" id="doors-total-val">0 Adet</span>
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-4 gap-5" id="door-cards-container"></div>
            </div>
        </div>


        <!-- TAB 4: FORMÜLLER -->
        <div id="view-formulas" class="space-y-6 hidden">
            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6">
                <div class="flex justify-between items-center border-b border-slate-800 pb-4">
                    <div>
                        <h2 class="text-xl font-black text-white flex items-center gap-2">
                            <i data-lucide="calculator" class="w-6 h-6 text-emerald-400"></i> Formüller ve Otomatik Hesaplama Mantığı
                        </h2>
                        <p class="text-xs text-slate-400 mt-1">Sistemde makineler ve ürünler eklendiğinde çalışan tüm resmi matematiksel formüller</p>
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div class="bg-slate-950 border border-slate-800 rounded-2xl p-5 space-y-3">
                        <h3 class="text-sm font-extrabold text-sky-400 flex items-center gap-2">
                            <i data-lucide="cpu" class="w-4 h-4"></i> Ekstrüder Hatları Formülleri
                        </h3>
                        <div class="space-y-2 text-xs font-mono text-slate-300">
                            <div class="p-2.5 bg-slate-900 rounded-xl border border-slate-800">
                                <span class="text-amber-400 font-bold block mb-1">Fire Oranı (%):</span>
                                <code>Fire Oranı (%) = [ Fire (kg) / (Üretim (kg) + Fire (kg)) ] * 100</code>
                            </div>
                            <div class="p-2.5 bg-slate-900 rounded-xl border border-slate-800">
                                <span class="text-emerald-400 font-bold block mb-1">Pervaz Takım Sayısı:</span>
                                <code>Takım Sayısı = Üretim Adedi / 5.0</code>
                            </div>
                            <div class="p-2.5 bg-slate-900 rounded-xl border border-slate-800">
                                <span class="text-emerald-400 font-bold block mb-1">Kasa Takım Sayısı:</span>
                                <code>Takım Sayısı = Üretim Adedi / 2.5</code>
                            </div>
                            <div class="p-2.5 bg-slate-900 rounded-xl border border-slate-800">
                                <span class="text-emerald-400 font-bold block mb-1">Seren Takım Sayısı:</span>
                                <code>Takım Sayısı = Üretim Adedi / 3.5</code>
                            </div>
                        </div>
                    </div>

                    <div class="bg-slate-950 border border-slate-800 rounded-2xl p-5 space-y-3">
                        <h3 class="text-sm font-extrabold text-amber-400 flex items-center gap-2">
                            <i data-lucide="layers" class="w-4 h-4"></i> Levha Hatları Formülleri
                        </h3>
                        <div class="space-y-2 text-xs font-mono text-slate-300">
                            <div class="p-2.5 bg-slate-900 rounded-xl border border-slate-800">
                                <span class="text-sky-400 font-bold block mb-1">1 Levha m² Hesabı:</span>
                                <code>1 Levha m² = (En_cm / 100) * (Boy_cm / 100)</code>
                            </div>
                            <div class="p-2.5 bg-slate-900 rounded-xl border border-slate-800">
                                <span class="text-sky-400 font-bold block mb-1">En Firesi (cm):</span>
                                <code>En Firesi (cm) = MAX(0, Kalıp Çıkış Eni - En_cm)</code>
                            </div>
                            <div class="p-2.5 bg-slate-900 rounded-xl border border-slate-800">
                                <span class="text-amber-400 font-bold block mb-1">Ölü Fire (kg):</span>
                                <code>Ölü Fire (kg) = (En Firesi/100) * (Boy/100) * Aded * (kg/m²)</code>
                            </div>
                            <div class="p-2.5 bg-slate-900 rounded-xl border border-slate-800">
                                <span class="text-emerald-400 font-bold block mb-1">Levha Takım Sayısı:</span>
                                <code>Levha Takım Sayısı = Üretim Adedi / 2.0</code>
                            </div>
                        </div>
                    </div>

                    <div class="bg-slate-950 border border-slate-800 rounded-2xl p-5 space-y-3">
                        <h3 class="text-sm font-extrabold text-emerald-400 flex items-center gap-2">
                            <i data-lucide="door-open" class="w-4 h-4"></i> Kapı Eşdeğeri ve Stok Devir Formülü
                        </h3>
                        <div class="space-y-2 text-xs font-mono text-slate-300">
                            <div class="p-2.5 bg-slate-900 rounded-xl border border-slate-800">
                                <span class="text-emerald-400 font-bold block mb-1">Tamamlanabilir Tam Kapı:</span>
                                <code>Tam Kapı = FLOOR( MIN(Pervaz_Eq, Kasa_Eq, Seren_Eq, Levha_Eq) )</code>
                            </div>
                            <div class="p-2.5 bg-slate-900 rounded-xl border border-slate-800">
                                <span class="text-amber-400 font-bold block mb-1">Ertesi Güne Kalan Devir Stok:</span>
                                <code>Ertesi Güne Devir = Toplam Mevcut Adet - (Tam Kapı * Reçete Oranı)</code>
                            </div>
                        </div>
                    </div>

                    <div class="bg-slate-950 border border-slate-800 rounded-2xl p-5 space-y-3">
                        <h3 class="text-sm font-extrabold text-purple-400 flex items-center gap-2">
                            <i data-lucide="trending-up" class="w-4 h-4"></i> Vardiya Performans Formülleri
                        </h3>
                        <div class="space-y-2 text-xs font-mono text-slate-300">
                            <div class="p-2.5 bg-slate-900 rounded-xl border border-slate-800">
                                <span class="text-purple-400 font-bold block mb-1">Kg / Çalışan Verimliliği:</span>
                                <code>Kg / Çalışan = Toplam Üretim (kg) / Vardiya Çalışan Sayısı</code>
                            </div>
                            <div class="p-2.5 bg-slate-900 rounded-xl border border-slate-800">
                                <span class="text-purple-400 font-bold block mb-1">Üretim Hızı (kg/saat):</span>
                                <code>Üretim Hızı = Toplam Üretim (kg) / Vardiya Çalışma Saati</code>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>


        <!-- TAB 5: DİNAMİK MAKINELER & ÜRÜNLER -->
        <div id="view-settings" class="space-y-6 hidden">
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div class="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6">
                    <div class="flex justify-between items-center border-b border-slate-800 pb-4">
                        <div>
                            <h3 class="text-lg font-bold text-white flex items-center gap-2">
                                <i data-lucide="cpu" class="w-5 h-5 text-ergunbas-500"></i> Makineler & Üretim Hatları
                            </h3>
                            <p class="text-xs text-slate-400">Fabrikaya yeni Ekstrüder veya Levha hattı ekleyebilirsiniz</p>
                        </div>
                    </div>
                    <form onsubmit="handleAddMachine(event)" class="flex gap-3 bg-slate-950 p-3 rounded-xl border border-slate-800">
                        <input type="text" id="new-machine-name" placeholder="Örn: Ekstrüder Hat 10" required class="flex-1 bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-ergunbas-500 focus:outline-none">
                        <select id="new-machine-type" class="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-ergunbas-500 focus:outline-none">
                            <option value="extruder">Ekstrüder Hat</option>
                            <option value="levha">Levha Hat</option>
                        </select>
                        <button type="submit" class="px-4 py-2 bg-ergunbas-600 hover:bg-ergunbas-500 text-white font-bold text-sm rounded-lg transition flex items-center gap-1">
                            <i data-lucide="plus" class="w-4 h-4"></i> Ekle
                        </button>
                    </form>
                    <div class="overflow-x-auto">
                        <table class="w-full text-left text-sm text-slate-300">
                            <thead class="bg-slate-950 text-slate-400 uppercase text-xs">
                                <tr>
                                    <th class="p-3">Hat İsmi</th>
                                    <th class="p-3">Tip</th>
                                    <th class="p-3 text-right">İşlem</th>
                                </tr>
                            </thead>
                            <tbody id="tbl-machines" class="divide-y divide-slate-800"></tbody>
                        </table>
                    </div>
                </div>

                <div class="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6">
                    <div class="flex justify-between items-center border-b border-slate-800 pb-4">
                        <div>
                            <h3 class="text-lg font-bold text-white flex items-center gap-2">
                                <i data-lucide="package" class="w-5 h-5 text-emerald-400"></i> Ürünler ve Reçete Oranları
                            </h3>
                            <p class="text-xs text-slate-400">Yeni ürün çeşidi ve 1 Kapı için reçete oranını tanımlayın</p>
                        </div>
                    </div>
                    <form onsubmit="handleAddProduct(event)" class="space-y-3 bg-slate-950 p-4 rounded-xl border border-slate-800">
                        <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
                            <input type="text" id="new-prod-name" placeholder="Örn: 120 mm Kasa" required class="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                            <select id="new-prod-cat" class="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                                <option value="pervaz">Pervaz (5 Adet/Kapı)</option>
                                <option value="kasa">Kasa (2.5 Adet/Kapı)</option>
                                <option value="seren">Seren (3.5 Adet/Kapı)</option>
                                <option value="levha">Levha (2 Adet/Kapı)</option>
                                <option value="diger">Diğer / Ekstra</option>
                            </select>
                            <input type="number" step="0.1" id="new-prod-ratio" placeholder="Reçete (Örn: 5.0)" required class="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                        </div>
                        <button type="submit" class="w-full py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-sm rounded-lg transition flex items-center justify-center gap-1">
                            <i data-lucide="plus" class="w-4 h-4"></i> Yeni Ürün & Reçete Kaydet
                        </button>
                    </form>
                    <div class="overflow-x-auto">
                        <table class="w-full text-left text-sm text-slate-300">
                            <thead class="bg-slate-950 text-slate-400 uppercase text-xs">
                                <tr>
                                    <th class="p-3">Ürün İsmi</th>
                                    <th class="p-3">Kategori</th>
                                    <th class="p-3">1 Kapı Katsayısı</th>
                                </tr>
                            </thead>
                            <tbody id="tbl-products" class="divide-y divide-slate-800"></tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 6: KULLANICILAR -->
        <div id="view-users" class="space-y-6 hidden">
            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6">
                <div class="flex justify-between items-center border-b border-slate-800 pb-4">
                    <div>
                        <h3 class="text-lg font-bold text-white flex items-center gap-2">
                            <i data-lucide="users" class="w-5 h-5 text-violet-400"></i> Kullanıcı Yönetimi
                        </h3>
                        <p class="text-xs text-slate-400">Sisteme yeni operatör veya yönetici ekleyebilir, mevcut kullanıcıları güncelleyebilir veya silebilirsiniz.</p>
                    </div>
                </div>
                
                <!-- Yeni Kullanıcı Ekleme Formu -->
                <form onsubmit="handleAddUser(event)" class="space-y-3 bg-slate-950 p-4 rounded-xl border border-slate-800">
                    <h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Yeni Kullanıcı Ekle</h4>
                    <div class="grid grid-cols-1 md:grid-cols-4 gap-3">
                        <input type="text" id="new-user-username" placeholder="Kullanıcı Adı" required class="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-violet-500 focus:outline-none">
                        <input type="text" id="new-user-name" placeholder="Ad Soyad / Vardiya İsmi" required class="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-violet-500 focus:outline-none">
                        <input type="password" id="new-user-password" placeholder="Şifre" required class="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-violet-500 focus:outline-none">
                        <select id="new-user-role" class="bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-violet-500 focus:outline-none">
                            <option value="operator">Operatör</option>
                            <option value="admin">Yönetici (Admin)</option>
                        </select>
                    </div>
                    <button type="submit" class="w-full py-2 bg-violet-600 hover:bg-violet-500 text-white font-bold text-sm rounded-lg transition flex items-center justify-center gap-1">
                        <i data-lucide="user-plus" class="w-4 h-4"></i> Kullanıcı Kaydet
                    </button>
                </form>

                <!-- Kullanıcı Listesi Tablosu -->
                <div class="overflow-x-auto">
                    <table class="w-full text-left text-sm text-slate-300">
                        <thead class="bg-slate-950 text-slate-400 uppercase text-xs">
                            <tr>
                                <th class="p-3">Ad Soyad</th>
                                <th class="p-3">Kullanıcı Adı</th>
                                <th class="p-3">Yetki / Rol</th>
                                <th class="p-3 text-right">İşlem</th>
                            </tr>
                        </thead>
                        <tbody id="tbl-users" class="divide-y divide-slate-800"></tbody>
                    </table>
                </div>
            </div>
        </div>

    </main>

    <!-- App Logic JavaScript -->
    <script>
        let currentDateKey = '1';
        let dayData = null;
        let dashboardData = null;
        let machinesList = [];
        let productsList = [];
        let prodChart = null;
        let fireChart = null;

        function safeCreateIcons() {
            try {
                if (typeof lucide !== 'undefined' && lucide.createIcons) {
                    lucide.createIcons();
                }
            } catch (e) {
                console.warn('Lucide icon warning:', e);
            }
        }

        document.addEventListener('DOMContentLoaded', () => {
            safeCreateIcons();
            
            // Check if user session exists
            const savedUser = localStorage.getItem('currentUser');
            if (savedUser) {
                currentUser = JSON.parse(savedUser);
                
                // Hide overlay
                document.getElementById('login-overlay').classList.add('hidden');
                
                // Update header
                document.getElementById('header-username').textContent = currentUser.name;
                document.getElementById('header-role').textContent = `(${currentUser.role === 'admin' ? 'Yönetici' : 'Operatör'})`;
                
                // Show users tab if admin
                const userTab = document.getElementById('tab-users');
                if (userTab) {
                    if (currentUser.role === 'admin') {
                        userTab.style.display = 'flex';
                    } else {
                        userTab.style.display = 'none';
                    }
                }
                
                // Load data
                fetchDashboardAndDates();
                loadDashboard();
                loadUsers(); // Operatör dropdown'larını doldur
            } else {
                // Keep overlay visible
                document.getElementById('login-overlay').classList.remove('hidden');
            }
        });

        function showToast(msg) {
            const toast = document.getElementById('toast');
            document.getElementById('toast-message').textContent = msg;
            toast.classList.remove('hidden');
            setTimeout(() => { toast.classList.add('hidden'); }, 3000);
        }

        function closeModal() {
            document.getElementById('modal-detail').classList.add('hidden');
        }

        async function fetchDashboardAndDates() {
            await fetchMachinesAndProducts();
            const res = await fetch('/api/dashboard');
            dashboardData = await res.json();

            const select = document.getElementById('select-day');
            if (select) {
                select.innerHTML = '';
                dashboardData.available_dates.forEach(d => {
                    const opt = document.createElement('option');
                    opt.value = d.key;
                    opt.textContent = d.date;
                    select.appendChild(opt);
                });
            }
        }

        async function promptAddNewDate() {
            const dateStr = prompt("Yeni giriş yapmak istediğiniz tarihi giriniz (Örn: 01.09.2026):", "01.09.2026");
            if (dateStr) {
                const res = await fetch('/api/daily/add_date', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ date_str: dateStr })
                });
                if (res.ok) {
                    const result = await res.json();
                    showToast(`${dateStr} günü sisteme eklendi!`);
                    await fetchDashboardAndDates();
                    document.getElementById('select-day').value = result.key;
                    loadDayData();
                }
            }
        }

        function switchTab(tabId) {
            ['dashboard', 'daily', 'doors', 'formulas', 'settings', 'users'].forEach(t => {
                const viewEl = document.getElementById(`view-${t}`);
                if (viewEl) viewEl.classList.add('hidden');
                const btn = document.getElementById(`tab-${t}`);
                if (btn) btn.className = "px-3.5 py-2 text-xs font-bold rounded-lg transition-all text-slate-400 hover:text-slate-200 flex items-center gap-1.5";
            });

            const targetView = document.getElementById(`view-${tabId}`);
            if (targetView) targetView.classList.remove('hidden');
            const activeBtn = document.getElementById(`tab-${tabId}`);
            if (activeBtn) activeBtn.className = "px-3.5 py-2 text-xs font-bold rounded-lg transition-all bg-ergunbas-600 text-white shadow-lg shadow-ergunbas-600/30 flex items-center gap-1.5";

            if (tabId === 'dashboard') loadDashboard();
            if (tabId === 'daily') loadDayData();
            if (tabId === 'doors') loadDoorStats();
            if (tabId === 'settings') loadSettings();
            if (tabId === 'users') loadUsers();
        }

        async function fetchMachinesAndProducts() {
            try {
                const [mRes, pRes] = await Promise.all([
                    fetch('/api/machines'),
                    fetch('/api/products')
                ]);
                machinesList = await mRes.json();
                productsList = await pRes.json();
            } catch(e) {
                console.error('Fetch machines/products failed:', e);
            }
        }

        async function loadDashboard() {
            try {
                const res = await fetch('/api/dashboard');
                dashboardData = await res.json();

                document.getElementById('kpi-prod-ton').textContent = `${dashboardData.total_prod_ton} Ton`;
                document.getElementById('kpi-fire-ton').textContent = `${dashboardData.total_fire_ton} Ton`;
                document.getElementById('kpi-fire-ratio').textContent = `Fire Oranı: %${dashboardData.overall_fire_ratio}`;
                document.getElementById('kpi-doors').textContent = `${dashboardData.door_stats.completable_doors} Adet`;
                document.getElementById('kpi-efficiency').textContent = `${dashboardData.kg_per_employee} kg/kişi`;
                document.getElementById('kpi-employees').textContent = `Toplam Vardiya Çalışanı: ${dashboardData.total_employees}`;

                document.getElementById('monthly-banner-prod').textContent = `${dashboardData.total_prod_ton} Ton`;
                document.getElementById('monthly-banner-fire').textContent = `${dashboardData.total_fire_ton} Ton`;
                document.getElementById('monthly-banner-doors').textContent = `${dashboardData.door_stats.completable_doors} Adet`;

                const wContainer = document.getElementById('weekly-cards-container');
                if (wContainer && dashboardData.weekly_summary) {
                    wContainer.innerHTML = dashboardData.weekly_summary.map((w, idx) => `
                        <div onclick="showWeeklyDetail(${idx})" class="bg-slate-950 border border-slate-800 hover:border-ergunbas-500 rounded-xl p-4 space-y-2 cursor-pointer transition group">
                            <div class="flex justify-between items-center">
                                <h4 class="text-xs font-black text-ergunbas-500 uppercase tracking-wider group-hover:text-ergunbas-400 transition">${w.name}</h4>
                                <i data-lucide="info" class="w-3.5 h-3.5 text-slate-500 group-hover:text-ergunbas-400"></i>
                            </div>
                            <div class="space-y-1 text-xs">
                                <div class="flex justify-between text-slate-400">
                                    <span>Üretim:</span>
                                    <span class="font-bold text-white">${w.prod_ton} Ton</span>
                                </div>
                                <div class="flex justify-between text-slate-400">
                                    <span>Fire:</span>
                                    <span class="font-bold text-amber-400">${w.fire_ton} Ton (%${w.fire_ratio})</span>
                                </div>
                                <div class="flex justify-between border-t border-slate-800 pt-1 text-slate-300 font-bold">
                                    <span>Tam Kapı:</span>
                                    <span class="text-emerald-400">${w.doors} Adet</span>
                                </div>
                            </div>
                        </div>
                    `).join('');
                }

                const doorDetails = dashboardData.door_stats.details;
                const tblSummary = document.getElementById('tbl-door-summary');
                if (tblSummary && doorDetails) {
                    tblSummary.innerHTML = `
                        <tr onclick="showDoorComponentDetail('pervaz')" class="hover:bg-slate-900/50 cursor-pointer transition">
                            <td class="p-3 font-bold text-white flex items-center gap-2"><span class="w-2 h-2 rounded-full bg-ergunbas-500"></span> Pervaz</td>
                            <td class="p-3 text-ergunbas-500 font-bold">5.0 Adet</td>
                            <td class="p-3">${doorDetails.pervaz.produced}</td>
                            <td class="p-3 font-bold text-emerald-400">${doorDetails.pervaz.door_eq} Kapı</td>
                            <td class="p-3 text-slate-400">${doorDetails.pervaz.used} Adet</td>
                            <td class="p-3 font-bold text-amber-400">${doorDetails.pervaz.carryover} Adet</td>
                        </tr>
                        <tr onclick="showDoorComponentDetail('kasa')" class="hover:bg-slate-900/50 cursor-pointer transition">
                            <td class="p-3 font-bold text-white flex items-center gap-2"><span class="w-2 h-2 rounded-full bg-emerald-500"></span> Kasa</td>
                            <td class="p-3 text-ergunbas-500 font-bold">2.5 Adet</td>
                            <td class="p-3">${doorDetails.kasa.produced}</td>
                            <td class="p-3 font-bold text-emerald-400">${doorDetails.kasa.door_eq} Kapı</td>
                            <td class="p-3 text-slate-400">${doorDetails.kasa.used} Adet</td>
                            <td class="p-3 font-bold text-amber-400">${doorDetails.kasa.carryover} Adet</td>
                        </tr>
                        <tr onclick="showDoorComponentDetail('seren')" class="hover:bg-slate-900/50 cursor-pointer transition">
                            <td class="p-3 font-bold text-white flex items-center gap-2"><span class="w-2 h-2 rounded-full bg-purple-500"></span> Seren</td>
                            <td class="p-3 text-ergunbas-500 font-bold">3.5 Adet</td>
                            <td class="p-3">${doorDetails.seren.produced}</td>
                            <td class="p-3 font-bold text-emerald-400">${doorDetails.seren.door_eq} Kapı</td>
                            <td class="p-3 text-slate-400">${doorDetails.seren.used} Adet</td>
                            <td class="p-3 font-bold text-amber-400">${doorDetails.seren.carryover} Adet</td>
                        </tr>
                        <tr onclick="showDoorComponentDetail('levha')" class="hover:bg-slate-900/50 cursor-pointer transition">
                            <td class="p-3 font-bold text-white flex items-center gap-2"><span class="w-2 h-2 rounded-full bg-amber-500"></span> Levha</td>
                            <td class="p-3 text-ergunbas-500 font-bold">2.0 Adet</td>
                            <td class="p-3">${doorDetails.levha.produced}</td>
                            <td class="p-3 font-bold text-emerald-400">${doorDetails.levha.door_eq} Kapı</td>
                            <td class="p-3 text-slate-400">${doorDetails.levha.used} Adet</td>
                            <td class="p-3 font-bold text-amber-400">${doorDetails.levha.carryover} Adet</td>
                        </tr>
                    `;
                }

                if (dashboardData.daily_chart) {
                    renderCharts(dashboardData.daily_chart);
                }
                safeCreateIcons();
            } catch(e) {
                console.error('loadDashboard failed:', e);
            }
        }

        function showMonthlyDetail() {
            if (!dashboardData || !dashboardData.monthly_summary) return;
            const ms = dashboardData.monthly_summary;

            document.getElementById('modal-title').innerHTML = `<i data-lucide="calendar" class="w-6 h-6 text-ergunbas-500"></i> ${ms.title}`;
            document.getElementById('modal-subtitle').textContent = `Toplam ${ms.total_days} Günlük Üretim, Makine Performansları ve Fire Sebepleri`;

            const topMachinesHtml = ms.top_producing_machines.map(m => `
                <div class="flex justify-between items-center p-2.5 bg-slate-950 rounded-xl border border-slate-800">
                    <span class="font-bold text-white">${m.name}</span>
                    <span class="text-xs"><strong class="text-ergunbas-500">${m.prod_ton} Ton</strong> Üretim | <span class="text-amber-400">${m.fire_kg} kg Fire</span></span>
                </div>
            `).join('');

            const topScrapHtml = ms.top_scrap_reasons.map(s => `
                <div class="flex justify-between items-center p-2.5 bg-slate-950 rounded-xl border border-slate-800">
                    <span class="font-semibold text-slate-300">${s.reason}</span>
                    <span class="font-bold text-amber-400 text-xs">${s.fire_kg} kg Fire</span>
                </div>
            `).join('');

            document.getElementById('modal-body').innerHTML = `
                <div class="grid grid-cols-2 md:grid-cols-4 gap-3 text-center mb-4">
                    <div class="p-3 bg-slate-950 rounded-xl border border-slate-800">
                        <span class="text-xs text-slate-400 block font-semibold">Toplam Üretim</span>
                        <span class="text-xl font-black text-white">${ms.total_prod_ton} Ton</span>
                    </div>
                    <div class="p-3 bg-slate-950 rounded-xl border border-slate-800">
                        <span class="text-xs text-slate-400 block font-semibold">Toplam Fire</span>
                        <span class="text-xl font-black text-amber-400">${ms.total_fire_ton} Ton</span>
                    </div>
                    <div class="p-3 bg-slate-950 rounded-xl border border-slate-800">
                        <span class="text-xs text-slate-400 block font-semibold">Fire Oranı</span>
                        <span class="text-xl font-black text-amber-400">%${ms.fire_ratio}</span>
                    </div>
                    <div class="p-3 bg-slate-950 rounded-xl border border-slate-800">
                        <span class="text-xs text-slate-400 block font-semibold">Tam Kapı</span>
                        <span class="text-xl font-black text-emerald-400">${ms.completable_doors} Adet</span>
                    </div>
                </div>

                <div class="space-y-3">
                    <h4 class="text-xs font-bold uppercase tracking-wider text-slate-400">En Çok Üretim Yapan Makineler</h4>
                    <div class="space-y-1.5 max-h-48 overflow-y-auto custom-scrollbar">
                        ${topMachinesHtml}
                    </div>
                </div>

                <div class="space-y-3 pt-3 border-t border-slate-800">
                    <h4 class="text-xs font-bold uppercase tracking-wider text-slate-400">Başlıca Fire ve Duruş Sebepleri</h4>
                    <div class="space-y-1.5 max-h-40 overflow-y-auto custom-scrollbar">
                        ${topScrapHtml}
                    </div>
                </div>
            `;

            document.getElementById('modal-detail').classList.remove('hidden');
            safeCreateIcons();
        }

        function showWeeklyDetail(wIndex) {
            if (!dashboardData || !dashboardData.weekly_summary) return;
            const w = dashboardData.weekly_summary[wIndex];

            document.getElementById('modal-title').innerHTML = `<i data-lucide="calendar-range" class="w-6 h-6 text-ergunbas-500"></i> ${w.name} Detayı`;
            document.getElementById('modal-subtitle').textContent = `Seçilen haftaya ait günlük üretimler ve tamamlanan kapı miktarları`;

            const daysHtml = w.keys.map(k => {
                const dayObj = dashboardData.daily_chart.find(d => String(d.key) === String(k));
                if (!dayObj) return '';
                return `
                    <tr onclick="openDayFromModal('${k}')" class="hover:bg-slate-900/80 cursor-pointer transition">
                        <td class="p-2.5 font-bold text-white">${dayObj.date}</td>
                        <td class="p-2.5 text-ergunbas-500 font-extrabold">${dayObj.prod_kg} kg</td>
                        <td class="p-2.5 text-amber-400 font-bold">${dayObj.fire_kg} kg (%${dayObj.fire_ratio})</td>
                        <td class="p-2.5 text-right font-semibold text-slate-400">Detay Veri Girişine Git &rarr;</td>
                    </tr>
                `;
            }).join('');

            document.getElementById('modal-body').innerHTML = `
                <div class="grid grid-cols-3 gap-3 text-center mb-4">
                    <div class="p-3 bg-slate-950 rounded-xl border border-slate-800">
                        <span class="text-xs text-slate-400 block font-semibold">Haftalık Üretim</span>
                        <span class="text-xl font-black text-white">${w.prod_ton} Ton</span>
                    </div>
                    <div class="p-3 bg-slate-950 rounded-xl border border-slate-800">
                        <span class="text-xs text-slate-400 block font-semibold">Haftalık Fire</span>
                        <span class="text-xl font-black text-amber-400">${w.fire_ton} Ton (%${w.fire_ratio})</span>
                    </div>
                    <div class="p-3 bg-slate-950 rounded-xl border border-slate-800">
                        <span class="text-xs text-slate-400 block font-semibold">Haftalık Tam Kapı</span>
                        <span class="text-xl font-black text-emerald-400">${w.doors} Adet</span>
                    </div>
                </div>

                <div class="overflow-x-auto">
                    <table class="w-full text-left text-sm text-slate-300">
                        <thead class="bg-slate-950 text-slate-400 uppercase text-xs">
                            <tr>
                                <th class="p-2.5">Tarih</th>
                                <th class="p-2.5">Üretim (kg)</th>
                                <th class="p-2.5">Fire (kg)</th>
                                <th class="p-2.5 text-right">İşlem</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-800">
                            ${daysHtml}
                        </tbody>
                    </table>
                </div>
            `;

            document.getElementById('modal-detail').classList.remove('hidden');
            safeCreateIcons();
        }

        function openDayFromModal(key) {
            closeModal();
            const selectEl = document.getElementById('select-day');
            if (selectEl) selectEl.value = key;
            switchTab('daily');
        }

        function showDoorComponentDetail(cat) {
            if (!dashboardData || !dashboardData.door_stats) return;
            const item = dashboardData.door_stats.details[cat];
            const titleMap = { 'pervaz': 'Pervaz (5 ad/kapı)', 'kasa': 'Kasa (2.5 ad/kapı)', 'seren': 'Seren (3.5 ad/kapı)', 'levha': 'Levha (2 ad/kapı)' };

            document.getElementById('modal-title').innerHTML = `<i data-lucide="layers" class="w-6 h-6 text-emerald-400"></i> ${titleMap[cat]} Ürün Detayı`;
            document.getElementById('modal-subtitle').textContent = `1 Kapı için gerekli reçete katsayısı ve devir stok durumu`;

            document.getElementById('modal-body').innerHTML = `
                <div class="space-y-4">
                    <div class="grid grid-cols-2 md:grid-cols-4 gap-3 text-center">
                        <div class="p-3 bg-slate-950 rounded-xl border border-slate-800">
                            <span class="text-xs text-slate-400 block">Üretilen Miktar</span>
                            <span class="text-xl font-bold text-white">${item.produced} Adet</span>
                        </div>
                        <div class="p-3 bg-slate-950 rounded-xl border border-slate-800">
                            <span class="text-xs text-slate-400 block">Kapı Eşdeğeri</span>
                            <span class="text-xl font-bold text-emerald-400">${item.door_eq} Kapı</span>
                        </div>
                        <div class="p-3 bg-slate-950 rounded-xl border border-slate-800">
                            <span class="text-xs text-slate-400 block">Kullanılan</span>
                            <span class="text-xl font-bold text-slate-300">${item.used} Adet</span>
                        </div>
                        <div class="p-3 bg-slate-950 rounded-xl border border-slate-800">
                            <span class="text-xs text-slate-400 block">Kalan Devir</span>
                            <span class="text-xl font-bold text-amber-400">${item.carryover} Adet</span>
                        </div>
                    </div>
                    <div class="p-4 bg-slate-950 rounded-xl border border-slate-800 text-xs text-slate-400 space-y-1">
                        <p class="font-bold text-white">💡 Hesaplama Mantığı:</p>
                        <p>1 adet tam kapı için tam olarak <strong>${item.req_per_door} adet ${cat}</strong> gereklidir. Toplam üretilen miktar bu orana bölündüğünde kapı eşdeğeri bulunur.</p>
                    </div>
                </div>
            `;

            document.getElementById('modal-detail').classList.remove('hidden');
            safeCreateIcons();
        }

        function renderCharts(dailyData) {
            if (typeof Chart === 'undefined') return;
            const labels = dailyData.map(d => d.date);
            const prodValues = dailyData.map(d => d.prod_kg);
            const fireRatios = dailyData.map(d => d.fire_ratio);

            if (prodChart) prodChart.destroy();
            const canvas1 = document.getElementById('chart-daily-prod');
            if (canvas1) {
                const ctx1 = canvas1.getContext('2d');
                prodChart = new Chart(ctx1, {
                    type: 'bar',
                    data: {
                        labels: labels,
                        datasets: [{
                            label: 'Üretim (kg)',
                            data: prodValues,
                            backgroundColor: '#e11d48',
                            borderRadius: 6
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        onClick: (e, elements) => {
                            if (elements.length > 0) {
                                const idx = elements[0].index;
                                const key = dailyData[idx].key;
                                openDayFromModal(key);
                            }
                        },
                        plugins: { legend: { display: false } },
                        scales: {
                            x: { grid: { display: false }, ticks: { color: '#64748b' } },
                            y: { grid: { color: '#1e293b' }, ticks: { color: '#64748b' } }
                        }
                    }
                });
            }

            if (fireChart) fireChart.destroy();
            const canvas2 = document.getElementById('chart-fire-ratio');
            if (canvas2) {
                const ctx2 = canvas2.getContext('2d');
                fireChart = new Chart(ctx2, {
                    type: 'line',
                    data: {
                        labels: labels,
                        datasets: [{
                            label: 'Fire Oranı (%)',
                            data: fireRatios,
                            borderColor: '#f59e0b',
                            backgroundColor: 'rgba(245, 158, 11, 0.1)',
                            fill: true,
                            tension: 0.3
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: { legend: { display: false } },
                        scales: {
                            x: { grid: { display: false }, ticks: { color: '#64748b' } },
                            y: { grid: { color: '#1e293b' }, ticks: { color: '#64748b' } }
                        }
                    }
                });
            }
        }

        async function loadDayData() {
            const daySelect = document.getElementById('select-day');
            if (!daySelect) return;
            currentDateKey = daySelect.value;
            const selectedText = daySelect.options[daySelect.selectedIndex] ? daySelect.options[daySelect.selectedIndex].text : currentDateKey;
            document.getElementById('current-day-label').textContent = selectedText;

            const res = await fetch(`/api/daily/${currentDateKey}`);
            const data = await res.json();
            dayData = data.day_data;

            document.getElementById('gunduz-emp').value = dayData.gunduz.employees || 10;
            document.getElementById('gunduz-hours').value = dayData.gunduz.hours || 12;
            const gOp = document.getElementById('gunduz-operator');
            if (gOp && dayData.gunduz.operator) gOp.value = dayData.gunduz.operator;
            renderExtruderTable('tbl-gunduz-ext', dayData.gunduz.extruders);
            renderLevhaTable('tbl-gunduz-lev', dayData.gunduz.levha);

            document.getElementById('gece-emp').value = dayData.gece.employees || 10;
            document.getElementById('gece-hours').value = dayData.gece.hours || 12;
            const geOp = document.getElementById('gece-operator');
            if (geOp && dayData.gece.operator) geOp.value = dayData.gece.operator;
            renderExtruderTable('tbl-gece-ext', dayData.gece.extruders);
            renderLevhaTable('tbl-gece-lev', dayData.gece.levha);

            renderDowntimesTable(dayData.downtimes || []);
        }

        function calcLiveRow(rowEl, type) {
            if (type === 'extruder') {
                const prodKg = parseFloat(rowEl.querySelector('.input-prod-kg').value) || 0;
                const fireKg = parseFloat(rowEl.querySelector('.input-fire-kg').value) || 0;
                const qty = parseInt(rowEl.querySelector('.input-qty').value) || 0;
                const prodName = rowEl.querySelector('.input-prod-name').value.toLowerCase();

                const total = prodKg + fireKg;
                const ratio = total > 0 ? ((fireKg / total) * 100).toFixed(1) : '0.0';
                rowEl.querySelector('.val-fire-ratio').textContent = `%${ratio}`;

                let sets = 0;
                if (prodName.includes('pervaz')) sets = qty / 5.0;
                else if (prodName.includes('kasa')) sets = qty / 2.5;
                else if (prodName.includes('seren')) sets = qty / 3.5;
                rowEl.querySelector('.val-sets').textContent = sets ? sets.toFixed(1) : '-';
            } else if (type === 'levha') {
                const width = parseFloat(rowEl.querySelector('.input-width').value) || 0;
                const length = parseFloat(rowEl.querySelector('.input-length').value) || 0;
                const kalipEni = parseFloat(rowEl.querySelector('.input-kalip-eni').value) || 0;
                const kgM2 = parseFloat(rowEl.querySelector('.input-kg-m2').value) || 0;
                const qty = parseInt(rowEl.querySelector('.input-qty').value) || 0;

                const m2One = (width / 100) * (length / 100);
                const totalM2 = m2One * qty;
                rowEl.querySelector('.val-total-m2').textContent = totalM2 > 0 ? totalM2.toFixed(1) : '-';

                let enFiresi = Math.max(0, kalipEni - width);
                let deadFireM2 = (enFiresi / 100) * (length / 100) * qty;
                let deadFireKg = deadFireM2 * kgM2;
                rowEl.querySelector('.val-dead-fire-kg').textContent = deadFireKg > 0 ? deadFireKg.toFixed(1) : '-';

                let totalKg = totalM2 * kgM2;
                rowEl.querySelector('.val-total-kg').textContent = totalKg > 0 ? totalKg.toFixed(1) : '-';

                const total = totalKg + deadFireKg;
                const ratio = total > 0 ? ((deadFireKg / total) * 100).toFixed(1) : '0.0';
                rowEl.querySelector('.val-fire-ratio').textContent = `%${ratio}`;
                rowEl.querySelector('.val-sets').textContent = (qty / 2.0).toFixed(1);
            }
        }

        function calcExtruderRow(tr) {
            const inputs = tr.querySelectorAll('input[type="number"]');
            // Columns: length, speed, hours, prod_kg, fire_kg(span), qty
            const length  = parseFloat(tr.querySelector('.input-ext-length')?.value) || 0;
            const speed   = parseFloat(tr.querySelector('.input-ext-speed')?.value)  || 0;
            const hours   = parseFloat(tr.querySelector('.input-ext-hours')?.value)  || 0;
            const prodKg  = parseFloat(tr.querySelector('.input-prod-kg')?.value)    || 0;
            const qty     = parseInt(tr.querySelector('.input-qty')?.value)          || 0;
            const hatName = tr.querySelector('.input-hat-ext')?.value || '';

            // Hat 3,4,5,9 => çift çıkışlı (2x hız)
            const hatNum = hatName.replace(/[^0-9]/g, '');
            const isDouble = ['3','4','5','9'].includes(hatNum);
            const mult = isDouble ? 2 : 1;

            let fireKg = 0;
            if (hours > 0 && speed > 0 && length > 0 && qty > 0 && prodKg > 0) {
                const teorikM = hours * 60 * speed * mult;
                const netM    = qty * length;
                const kgPerM  = netM > 0 ? prodKg / netM : 0;
                if (teorikM > netM) {
                    fireKg = Math.round((teorikM - netM) * kgPerM * 100) / 100;
                }
            }

            const fireSpan = tr.querySelector('.val-fire-kg');
            if (fireSpan) fireSpan.textContent = fireKg.toFixed(1);
            const fireInput = tr.querySelector('.input-fire-kg');
            if (fireInput) fireInput.value = fireKg;

            const total = prodKg + fireKg;
            const ratio = total > 0 ? ((fireKg / total) * 100).toFixed(1) : '0.0';
            const ratioEl = tr.querySelector('.val-fire-ratio');
            if (ratioEl) ratioEl.textContent = `%${ratio}`;

            // Takım sayısı
            const prodName = tr.querySelector('.input-prod-name')?.value?.toLowerCase() || '';
            let sets = 0;
            if (prodName.includes('pervaz'))     sets = Math.round(qty / 5.0 * 10) / 10;
            else if (prodName.includes('kasa'))  sets = Math.round(qty / 2.5 * 10) / 10;
            else if (prodName.includes('seren')) sets = Math.round(qty / 3.5 * 10) / 10;
            const setsEl = tr.querySelector('.val-sets');
            if (setsEl) setsEl.textContent = sets || '-';
        }

        function createExtruderRowHtml(hatName = '', item = {}) {
            const prodKg  = item.prod_kg || 0;
            const fireKg  = item.fire_kg || 0;
            const hours   = item.hours   || '';
            const ratio   = (prodKg + fireKg) > 0 ? ((fireKg / (prodKg + fireKg)) * 100).toFixed(1) : '0.0';

            return `
                <tr class="hover:bg-slate-900/50" oninput="calcExtruderRow(this)">
                    <td class="p-2.5 font-semibold text-white">
                        <input type="text" value="${hatName || item.hat || 'Hat 1'}" class="input-hat-ext bg-slate-950 border border-slate-800 text-white text-xs font-bold rounded px-2 py-1 w-28">
                    </td>
                    <td class="p-2.5">
                        <input type="text" value="${item.product || ''}" class="input-prod-name bg-slate-950 border border-slate-800 text-white text-xs rounded px-2 py-1 w-full font-medium" placeholder="Örn: 80X80 Pervaz">
                    </td>
                    <td class="p-2.5">
                        <input type="number" step="0.01" value="${item.length || ''}" class="input-ext-length bg-slate-950 border border-slate-800 text-white text-xs rounded px-2 py-1 w-20 text-center">
                    </td>
                    <td class="p-2.5">
                        <input type="number" step="0.01" value="${item.speed || ''}" class="input-ext-speed bg-slate-950 border border-slate-800 text-white text-xs rounded px-2 py-1 w-20 text-center">
                    </td>
                    <td class="p-2.5">
                        <input type="number" step="0.1" value="${hours}" placeholder="sa" class="input-ext-hours bg-slate-950 border border-amber-700/40 text-amber-300 text-xs rounded px-2 py-1 w-16 text-center" title="Bu hattın çalışma süresi (saat). Boş bırakılırsa fire hesaplanmaz.">
                    </td>
                    <td class="p-2.5">
                        <input type="number" step="1" value="${prodKg}" class="input-prod-kg bg-slate-950 border border-slate-800 text-ergunbas-500 font-extrabold text-xs rounded px-2 py-1 w-24 text-center">
                    </td>
                    <td class="p-2.5 text-center">
                        <input type="hidden" class="input-fire-kg" value="${fireKg}">
                        <span class="val-fire-kg font-extrabold text-amber-400 text-xs">${fireKg.toFixed(1)}</span>
                    </td>
                    <td class="p-2.5">
                        <input type="number" step="1" value="${item.qty || 0}" class="input-qty bg-slate-950 border border-slate-800 text-white text-xs rounded px-2 py-1 w-20 text-center">
                    </td>
                    <td class="p-2.5 font-bold text-amber-400 text-xs text-center val-fire-ratio">%${ratio}</td>
                    <td class="p-2.5 font-bold text-emerald-400 text-xs text-center val-sets">${item.sets ? Math.round(item.sets*10)/10 : '-'}</td>
                    <td class="p-2.5 text-center">
                        <button onclick="this.closest('tr').remove()" class="p-1 text-red-400 hover:text-red-300">
                            <i data-lucide="trash-2" class="w-4 h-4"></i>
                        </button>
                    </td>
                </tr>
            `;
        }


        function renderExtruderTable(elementId, items) {
            const tbody = document.getElementById(elementId);
            if (!tbody) return;
            tbody.innerHTML = '';
            const extruderMachines = machinesList.filter(m => m.type === 'extruder');

            if (items && items.length > 0) {
                items.forEach(item => {
                    tbody.insertAdjacentHTML('beforeend', createExtruderRowHtml(item.hat, item));
                });
            } else {
                extruderMachines.forEach(m => {
                    tbody.insertAdjacentHTML('beforeend', createExtruderRowHtml(m.name, {}));
                });
            }
            safeCreateIcons();
        }

        function addExtruderRow(elementId) {
            const tbody = document.getElementById(elementId);
            if (tbody) tbody.insertAdjacentHTML('beforeend', createExtruderRowHtml('Ekstrüder Hat', {}));
            safeCreateIcons();
        }

        function createLevhaRowHtml(hatName = '', item = {}) {
            const kalipEni = item.kalip_cikis_eni || (item.width_fire_cm && item.width ? (item.width + item.width_fire_cm) : 108);
            const kgM2Raw = item.kg_per_m2 || (item.total_kg && item.total_m2 ? (item.total_kg / item.total_m2) : 2.45);
            const kgM2 = Math.round(kgM2Raw * 100) / 100;

            const totalKg = item.total_kg || 0;
            const deadFireKg = item.dead_fire_kg || 0;
            const ratio = (totalKg + deadFireKg) > 0 ? ((deadFireKg / (totalKg + deadFireKg)) * 100).toFixed(1) : '0.0';

            return `
                <tr class="hover:bg-slate-900/50" oninput="calcLiveRow(this, 'levha')">
                    <td class="p-2.5 font-semibold text-white">
                        <input type="text" value="${hatName || item.hat || 'Levha 1'}" class="input-hat bg-slate-950 border border-slate-800 text-white text-xs font-bold rounded px-2 py-1 w-24">
                    </td>
                    <td class="p-2.5">
                        <input type="text" value="${item.color || ''}" class="input-color bg-slate-950 border border-slate-800 text-white text-xs rounded px-2 py-1 w-full" placeholder="Model/Renk">
                    </td>
                    <td class="p-2.5">
                        <input type="number" value="${item.width || 93}" class="input-width bg-slate-950 border border-slate-800 text-white text-xs rounded px-2 py-1 w-16 text-center">
                    </td>
                    <td class="p-2.5">
                        <input type="number" value="${item.length || 208}" class="input-length bg-slate-950 border border-slate-800 text-white text-xs rounded px-2 py-1 w-16 text-center">
                    </td>
                    <td class="p-2.5">
                        <input type="number" value="${kalipEni}" class="input-kalip-eni bg-slate-950 border border-slate-800 text-slate-400 text-xs rounded px-2 py-1 w-16 text-center">
                    </td>
                    <td class="p-2.5">
                        <input type="number" step="0.01" value="${kgM2}" class="input-kg-m2 bg-slate-950 border border-slate-800 text-slate-400 text-xs rounded px-2 py-1 w-16 text-center">
                    </td>
                    <td class="p-2.5">
                        <input type="number" value="${item.qty || 0}" class="input-qty bg-slate-950 border border-slate-800 text-white text-xs rounded px-2 py-1 w-20 text-center">
                    </td>
                    <td class="p-2.5 text-xs text-slate-400 text-center val-total-m2">${item.total_m2 ? item.total_m2.toFixed(1) : '-'}</td>
                    <td class="p-2.5 font-bold text-amber-400 text-xs text-center val-dead-fire-kg">${deadFireKg ? deadFireKg.toFixed(1) : '-'}</td>
                    <td class="p-2.5 font-bold text-ergunbas-500 text-xs text-center val-total-kg">${totalKg ? totalKg.toFixed(1) : '-'}</td>
                    <td class="p-2.5 font-bold text-amber-400 text-xs text-center val-fire-ratio">%${ratio}</td>
                    <td class="p-2.5 font-bold text-emerald-400 text-xs text-center val-sets">${item.sets ? item.sets.toFixed(1) : '-'}</td>
                    <td class="p-2.5 text-center">
                        <button onclick="this.closest('tr').remove()" class="p-1 text-red-400 hover:text-red-300">
                            <i data-lucide="trash-2" class="w-4 h-4"></i>
                        </button>
                    </td>
                </tr>
            `;
        }

        function renderLevhaTable(elementId, items) {
            const tbody = document.getElementById(elementId);
            if (!tbody) return;
            tbody.innerHTML = '';
            const levhaMachines = machinesList.filter(m => m.type === 'levha');

            if (items && items.length > 0) {
                items.forEach(item => {
                    tbody.insertAdjacentHTML('beforeend', createLevhaRowHtml(item.hat, item));
                });
            } else {
                levhaMachines.forEach(m => {
                    tbody.insertAdjacentHTML('beforeend', createLevhaRowHtml(m.name, {}));
                });
            }
            safeCreateIcons();
        }

        function addLevhaRow(elementId) {
            const tbody = document.getElementById(elementId);
            if (tbody) tbody.insertAdjacentHTML('beforeend', createLevhaRowHtml('Levha 1', {}));
            safeCreateIcons();
        }

        function renderDowntimesTable(downtimes) {
            const tbody = document.getElementById('tbl-downtimes');
            if (!tbody) return;
            tbody.innerHTML = '';

            downtimes.forEach((dt, idx) => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-900/50';
                tr.innerHTML = `
                    <td class="p-2.5">
                        <select class="bg-slate-950 border border-slate-800 text-white text-xs font-semibold rounded px-2 py-1">
                            <option value="Gündüz" ${dt.shift === 'Gündüz' ? 'selected' : ''}>Gündüz</option>
                            <option value="Gece" ${dt.shift === 'Gece' ? 'selected' : ''}>Gece</option>
                        </select>
                    </td>
                    <td class="p-2.5">
                        <input type="text" value="${dt.hat || ''}" class="bg-slate-950 border border-slate-800 text-white text-xs rounded px-2 py-1 w-24">
                    </td>
                    <td class="p-2.5">
                        <input type="text" value="${dt.fire_reason || ''}" class="bg-slate-950 border border-slate-800 text-white text-xs rounded px-2 py-1 w-full" placeholder="Fire Sebebi">
                    </td>
                    <td class="p-2.5">
                        <input type="number" value="${dt.fire_kg || 0}" class="bg-slate-950 border border-slate-800 text-amber-400 font-bold text-xs rounded px-2 py-1 w-20 text-center">
                    </td>
                    <td class="p-2.5">
                        <input type="text" value="${dt.down_reason || ''}" class="bg-slate-950 border border-slate-800 text-white text-xs rounded px-2 py-1 w-full" placeholder="Duruş Sebebi">
                    </td>
                    <td class="p-2.5">
                        <input type="number" value="${dt.down_min || 0}" class="bg-slate-950 border border-slate-800 text-red-400 font-bold text-xs rounded px-2 py-1 w-20 text-center">
                    </td>
                    <td class="p-2.5">
                        <input type="text" value="${dt.desc || ''}" class="bg-slate-950 border border-slate-800 text-slate-400 text-xs rounded px-2 py-1 w-full" placeholder="Açıklama">
                    </td>
                    <td class="p-2.5 text-center">
                        <button onclick="this.closest('tr').remove()" class="p-1 text-red-400 hover:text-red-300">
                            <i data-lucide="trash-2" class="w-4 h-4"></i>
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
            safeCreateIcons();
        }

        function addDowntimeRow() {
            renderDowntimesTable([
                ...(getDowntimesFromUI()),
                { shift: 'Gündüz', hat: '1', fire_reason: '', fire_kg: 0, down_reason: '', down_min: 0, desc: '' }
            ]);
        }

        function getDowntimesFromUI() {
            const rows = document.querySelectorAll('#tbl-downtimes tr');
            const list = [];
            rows.forEach(r => {
                const selects = r.querySelectorAll('select');
                const inputs = r.querySelectorAll('input');
                if (selects.length && inputs.length >= 6) {
                    list.push({
                        shift: selects[0].value,
                        hat: inputs[0].value,
                        fire_reason: inputs[1].value,
                        fire_kg: parseFloat(inputs[2].value) || 0,
                        down_reason: inputs[3].value,
                        down_min: parseFloat(inputs[4].value) || 0,
                        desc: inputs[5].value
                    });
                }
            });
            return list;
        }

        async function saveCurrentDay() {
            const payload = {
                gunduz: {
                    employees: parseInt(document.getElementById('gunduz-emp').value) || 10,
                    hours: parseFloat(document.getElementById('gunduz-hours').value) || 12,
                    operator: document.getElementById('gunduz-operator')?.value || '',
                    extruders: getExtrudersFromUI('tbl-gunduz-ext'),
                    levha: getLevhaFromUI('tbl-gunduz-lev')
                },
                gece: {
                    employees: parseInt(document.getElementById('gece-emp').value) || 10,
                    hours: parseFloat(document.getElementById('gece-hours').value) || 12,
                    operator: document.getElementById('gece-operator')?.value || '',
                    extruders: getExtrudersFromUI('tbl-gece-ext'),
                    levha: getLevhaFromUI('tbl-gece-lev')
                },
                downtimes: getDowntimesFromUI()
            };

            const res = await fetch(`/api/daily/${currentDateKey}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (res.ok) {
                showToast("Günlük veriler ERGUNBAS veritabanına başarıyla kaydedildi!");
                loadDashboard();
            }
        }

        function getExtrudersFromUI(elementId) {
            const rows = document.querySelectorAll(`#${elementId} tr`);
            const list = [];
            rows.forEach(r => {
                const hat     = r.querySelector('.input-hat-ext')?.value  || '';
                const product = r.querySelector('.input-prod-name')?.value || '';
                const length  = parseFloat(r.querySelector('.input-ext-length')?.value) || 0;
                const speed   = parseFloat(r.querySelector('.input-ext-speed')?.value)  || 0;
                const hours   = parseFloat(r.querySelector('.input-ext-hours')?.value)  || 0;
                const prodKg  = parseFloat(r.querySelector('.input-prod-kg')?.value)    || 0;
                const fireKg  = parseFloat(r.querySelector('.input-fire-kg')?.value)    || 0;
                const qty     = parseInt(r.querySelector('.input-qty')?.value)          || 0;
                if (hat || product) {
                    list.push({ hat, product, length, speed, hours, prod_kg: prodKg, fire_kg: fireKg, qty, sets: 0 });
                }
            });
            return list;
        }

        function getLevhaFromUI(elementId) {
            const rows = document.querySelectorAll(`#${elementId} tr`);
            const list = [];
            rows.forEach(r => {
                const hatEl = r.querySelector('.input-hat');
                if (hatEl) {
                    list.push({
                        hat: hatEl.value,
                        color: r.querySelector('.input-color')?.value || '',
                        width: parseFloat(r.querySelector('.input-width')?.value) || 0,
                        length: parseFloat(r.querySelector('.input-length')?.value) || 0,
                        kalip_cikis_eni: parseFloat(r.querySelector('.input-kalip-eni')?.value) || 0,
                        kg_per_m2: parseFloat(r.querySelector('.input-kg-m2')?.value) || 0,
                        qty: parseInt(r.querySelector('.input-qty')?.value) || 0
                    });
                }
            });
            return list;
        }

        async function loadDoorStats() {
            const res = await fetch('/api/dashboard');
            const data = await res.json();

            document.getElementById('doors-total-val').textContent = `${data.door_stats.completable_doors} Adet`;

            const container = document.getElementById('door-cards-container');
            const details = data.door_stats.details;

            const cards = [
                { cat: 'pervaz', title: 'Pervaz (5 Adet/Kapı)', color: 'ergunbas', data: details.pervaz },
                { cat: 'kasa', title: 'Kasa (2.5 Adet/Kapı)', color: 'emerald', data: details.kasa },
                { cat: 'seren', title: 'Seren (3.5 Adet/Kapı)', color: 'purple', data: details.seren },
                { cat: 'levha', title: 'Levha (2.0 Adet/Kapı)', color: 'amber', data: details.levha },
            ];

            container.innerHTML = cards.map(c => `
                <div onclick="showDoorComponentDetail('${c.cat}')" class="bg-slate-950 border border-slate-800 hover:border-emerald-500 rounded-2xl p-5 space-y-3 shadow-md cursor-pointer transition group">
                    <div class="flex justify-between items-center">
                        <h4 class="text-sm font-black text-${c.color === 'ergunbas' ? 'ergunbas-500' : c.color + '-400'}">${c.title}</h4>
                        <i data-lucide="info" class="w-4 h-4 text-slate-500 group-hover:text-emerald-400"></i>
                    </div>
                    <div class="space-y-1.5 text-xs">
                        <div class="flex justify-between text-slate-400">
                            <span>Üretilen:</span>
                            <span class="font-bold text-white">${c.data.produced} Adet</span>
                        </div>
                        <div class="flex justify-between text-slate-400">
                            <span>Kapı Eşdeğeri:</span>
                            <span class="font-bold text-emerald-400">${c.data.door_eq} Kapı</span>
                        </div>
                        <div class="flex justify-between text-slate-400">
                            <span>Kullanılan:</span>
                            <span class="font-medium text-slate-300">${c.data.used} Adet</span>
                        </div>
                        <div class="flex justify-between border-t border-slate-800 pt-2 text-slate-300 font-bold">
                            <span>Devir Stok:</span>
                            <span class="text-amber-400">${c.data.carryover} Adet</span>
                        </div>
                    </div>
                </div>
            `).join('');

            safeCreateIcons();
        }

        async function loadSettings() {
            await fetchMachinesAndProducts();

            const tblM = document.getElementById('tbl-machines');
            if (tblM) {
                tblM.innerHTML = machinesList.map(m => `
                    <tr class="hover:bg-slate-900/50">
                        <td class="p-3 font-semibold text-white">${m.name}</td>
                        <td class="p-3">
                            <span class="px-2.5 py-0.5 rounded text-xs font-bold ${m.type === 'extruder' ? 'bg-ergunbas-600/20 text-ergunbas-500 border border-ergunbas-600/30' : 'bg-amber-500/10 text-amber-400 border border-amber-500/20'}">
                                ${m.type === 'extruder' ? 'Ekstrüder' : 'Levha'}
                            </span>
                        </td>
                        <td class="p-3 text-right">
                            <button onclick="deleteMachine('${m.id}')" class="p-1 text-red-400 hover:text-red-300">
                                <i data-lucide="trash-2" class="w-4 h-4"></i>
                            </button>
                        </td>
                    </tr>
                `).join('');
            }

            const tblP = document.getElementById('tbl-products');
            if (tblP) {
                tblP.innerHTML = productsList.map(p => `
                    <tr class="hover:bg-slate-900/50">
                        <td class="p-3 font-semibold text-white">${p.name}</td>
                        <td class="p-3 uppercase text-xs font-bold text-slate-400">${p.category}</td>
                        <td class="p-3 font-bold text-emerald-400">${p.door_ratio} Adet/Kapı</td>
                    </tr>
                `).join('');
            }

            safeCreateIcons();
        }

        async function handleAddMachine(e) {
            e.preventDefault();
            const name = document.getElementById('new-machine-name').value;
            const type = document.getElementById('new-machine-type').value;

            await fetch('/api/machines', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ name, type })
            });

            document.getElementById('new-machine-name').value = '';
            showToast(`Yeni makine (${name}) eklendi ve hesaplamalara dahil edildi!`);
            loadSettings();
        }

        async function deleteMachine(id) {
            if (confirm('Bu makineyi silmek istediğinizden emin misiniz?')) {
                await fetch(`/api/machines/${id}`, { method: 'DELETE' });
                showToast("Makine silindi!");
                loadSettings();
            }
        }

        async function handleAddProduct(e) {
            e.preventDefault();
            const name = document.getElementById('new-prod-name').value;
            const category = document.getElementById('new-prod-cat').value;
            const door_ratio = parseFloat(document.getElementById('new-prod-ratio').value) || 0;

            await fetch('/api/products', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ name, category, door_ratio })
            });

            document.getElementById('new-prod-name').value = '';
            document.getElementById('new-prod-ratio').value = '';
            showToast(`Yeni ürün (${name}) reçeteye kaydedildi!`);
            loadSettings();
        }

        let currentUser = null;

        async function handleLogin(e) {
            if (e) e.preventDefault();
            const username = document.getElementById('login-username').value;
            const password = document.getElementById('login-password').value;
            const errorEl = document.getElementById('login-error');

            try {
                const res = await fetch('/api/auth/login', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ username, password })
                });

                if (res.ok) {
                    const data = await res.json();
                    currentUser = data.user;
                    localStorage.setItem('currentUser', JSON.stringify(currentUser));
                    
                    document.getElementById('login-overlay').classList.add('hidden');
                    errorEl.classList.add('hidden');
                    
                    document.getElementById('header-username').textContent = currentUser.name;
                    document.getElementById('header-role').textContent = `(${currentUser.role === 'admin' ? 'Yönetici' : 'Operatör'})`;
                    
                    const userTab = document.getElementById('tab-users');
                    if (userTab) {
                        if (currentUser.role === 'admin') {
                            userTab.style.display = 'flex';
                        } else {
                            userTab.style.display = 'none';
                        }
                    }
                    
                    await fetchDashboardAndDates();
                    loadDashboard();
                    
                    showToast(`Hoş geldiniz, ${currentUser.name}!`);
                } else {
                    const err = await res.json();
                    errorEl.textContent = err.detail || 'Kullanıcı adı veya şifre hatalı!';
                    errorEl.classList.remove('hidden');
                }
            } catch(err) {
                console.error(err);
                errorEl.textContent = 'Sunucu bağlantı hatası!';
                errorEl.classList.remove('hidden');
            }
        }

        function handleLogout() {
            currentUser = null;
            localStorage.removeItem('currentUser');
            document.getElementById('login-overlay').classList.remove('hidden');
            document.getElementById('login-username').value = '';
            document.getElementById('login-password').value = '';
            document.getElementById('header-username').textContent = '';
            document.getElementById('header-role').textContent = '';
            
            const userTab = document.getElementById('tab-users');
            if (userTab) userTab.style.display = 'none';
            
            switchTab('dashboard');
        }

        async function loadUsers() {
            try {
                const res = await fetch('/api/users');
                if (!res.ok) throw new Error('Kullanıcı listesi alınamadı');
                const users = await res.json();
                
                const tbl = document.getElementById('tbl-users');
                if (tbl) {
                    tbl.innerHTML = users.map(u => {
                        const isMainAdmin = u.id === 'u1';
                        return `
                            <tr class="hover:bg-slate-900/50">
                                <td class="p-3 font-semibold text-white">${u.name}</td>
                                <td class="p-3 font-mono text-xs text-slate-400">${u.username}</td>
                                <td class="p-3">
                                    <span class="px-2.5 py-0.5 rounded text-xs font-bold ${u.role === 'admin' ? 'bg-violet-600/20 text-violet-400 border border-violet-600/30' : 'bg-slate-500/10 text-slate-400 border border-slate-500/20'}">
                                        ${u.role === 'admin' ? 'Yönetici' : 'Operatör'}
                                    </span>
                                </td>
                                <td class="p-3 text-right space-x-2">
                                    <button onclick="promptUpdateUser('${u.id}', '${u.name}', '${u.role}')" class="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-bold rounded-lg border border-slate-700 transition">
                                        Düzenle
                                    </button>
                                    ${isMainAdmin ? '' : `
                                    <button onclick="deleteUser('${u.id}')" class="p-1 text-red-400 hover:text-red-300 inline-block align-middle">
                                        <i data-lucide="trash-2" class="w-4 h-4"></i>
                                    </button>
                                    `}
                                </td>
                            </tr>
                        `;
                    }).join('');
                }

                // Vardiya operatör dropdown'larını doldur
                ['gunduz-operator', 'gece-operator'].forEach(selId => {
                    const sel = document.getElementById(selId);
                    if (!sel) return;
                    const currentVal = sel.value;
                    sel.innerHTML = '<option value="">— Seçiniz —</option>' +
                        users.map(u => `<option value="${u.username}">${u.name}</option>`).join('');
                    if (currentVal) sel.value = currentVal;
                });

                safeCreateIcons();
            } catch(e) {
                console.error(e);
            }
        }

        async function handleAddUser(e) {
            e.preventDefault();
            const username = document.getElementById('new-user-username').value;
            const name = document.getElementById('new-user-name').value;
            const password = document.getElementById('new-user-password').value;
            const role = document.getElementById('new-user-role').value;

            const res = await fetch('/api/users', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ username, name, password, role })
            });

            if (res.ok) {
                document.getElementById('new-user-username').value = '';
                document.getElementById('new-user-name').value = '';
                document.getElementById('new-user-password').value = '';
                showToast(`Yeni kullanıcı (${name}) sisteme kaydedildi!`);
                loadUsers();
            } else {
                const err = await res.json();
                alert(err.detail || 'Kullanıcı eklenemedi!');
            }
        }

        async function deleteUser(id) {
            if (confirm('Bu kullanıcıyı silmek istediğinizden emin misiniz?')) {
                const res = await fetch(`/api/users/${id}`, { method: 'DELETE' });
                if (res.ok) {
                    showToast("Kullanıcı silindi!");
                    loadUsers();
                } else {
                    const err = await res.json();
                    alert(err.detail || 'Kullanıcı silinemedi!');
                }
            }
        }

        async function promptUpdateUser(id, currentName, currentRole) {
            const newName = prompt("Kullanıcı adı/soyadı güncelleyin:", currentName);
            if (newName === null) return;
            
            const newRole = prompt("Rolü güncelleyin ('admin' veya 'operator'):", currentRole);
            if (newRole === null) return;
            
            if (newRole !== 'admin' && newRole !== 'operator') {
                alert("Geçersiz rol! Sadece 'admin' veya 'operator' olmalıdır.");
                return;
            }
            
            const newPassword = prompt("Yeni şifre (Değiştirmek istemiyorsanız boş bırakın):", "");
            
            const payload = {
                name: newName,
                role: newRole
            };
            if (newPassword) {
                payload.password = newPassword;
            }

            const res = await fetch(`/api/users/${id}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (res.ok) {
                showToast("Kullanıcı bilgileri güncellendi!");
                loadUsers();
            } else {
                const err = await res.json();
                alert(err.detail || 'Kullanıcı güncellenemedi!');
            }
        }
    </script>
</body>
</html>
"""

final_html = template.replace("{{LOGO_DATA_URI}}", logo_data_uri)

with open(os.path.join(APP_DIR, "static", "index.html"), "w", encoding="utf-8") as f:
    f.write(final_html)

print("SUCCESS: static/index.html generated cleanly!")
