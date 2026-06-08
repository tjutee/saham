import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import inspect
import os
import re
import json
from pathlib import Path
from io import StringIO
from urllib.request import Request, urlopen


DATA_FILE = "Ringkasan.xlsx"
HISTORY_CACHE_DIR = Path("history_cache")
BLOCKING_PROXY_VALUES = {"http://127.0.0.1:9", "https://127.0.0.1:9", "127.0.0.1:9"}
IDX_COMPANY_PROFILE_URLS = [
    "https://www.idx.co.id/primary/ListedCompany/GetCompanyProfiles?emitenType=s",
    "https://www.idx.co.id/umbraco/Surface/ListedCompany/GetCompanyProfiles?emitenType=s",
]
IDX_STOCK_LIST_URL = "https://www.idx.id/primary/StockData/GetSecuritiesStock?start=0&length=9999&code=&sector=&board=&language=en-us"
TRADINGVIEW_SCAN_URL = "https://scanner.tradingview.com/indonesia/scan"
STOCKANALYSIS_IDX_URL = "https://stockanalysis.com/list/indonesia-stock-exchange/"
ONLINE_LOAD_PERIOD = "1y"
ONLINE_REFRESH_TTL = 6 * 60 * 60
TRADINGVIEW_FUNDAMENTAL_COLUMNS = [
    "name",
    "description",
    "sector",
    "industry",
    "market_cap_basic",
    "total_revenue",
    "price_earnings_ttm",
    "price_book_fq",
    "return_on_equity_fq",
    "return_on_assets_fq",
    "debt_to_equity_fq",
    "net_margin_ttm",
]

NUMERIC_COLUMNS = [
    "Penutupan",
    "PER",
    "PBV",
    "ROE",
    "ROA",
    "DER",
    "NPM",
    "Sebelumnya",
    "%Change",
    "Open",
    "High",
    "Low",
    "Volume",
    "NIM",
    "CAR",
    "LDR",
    "NPL",
    "BOPO",
    "CIR",
    "LAR",
    "Index_Count_Raw",
    "Index_Count_Sigma",
    "Index_Count_Metrik",
    "Mkt Cap",
    "Total Rev",
    "Market_Cap",
    "Revenue",
    "Sales_Multiple",
]

BASE_WEIGHTS = {
    "valuation": 25,
    "quality": 30,
    "risk": 15,
    "liquidity": 15,
    "momentum": 10,
    "index_strength": 5,
}

HISTORY_COLUMNS = {
    "4-wk %Pr. Chg.": "Return_4W",
    "13-wk %Pr. Chg.": "Return_13W",
    "26-wk %Pr. Chg.": "Return_26W",
    "52-wk %Pr. Chg.": "Return_52W",
    "MTD": "Return_MTD",
    "YTD": "Return_YTD",
}

ONLINE_RETURN_WINDOWS = {
    "Return_4W": 20,
    "Return_13W": 65,
    "Return_26W": 130,
    "Return_52W": 260,
}

NONBANK_THRESHOLDS = {
    "PER": ("<=", 15.0),
    "PBV": ("<=", 3.0),
    "ROE": (">=", 12.0),
    "ROA": (">=", 7.0),
    "DER": ("<=", 1.5),
    "NPM": (">=", 7.0),
}

BANKING_THRESHOLDS = {
    "PER": ("<=", 15.0),
    "PBV": ("<=", 3.0),
    "ROE": (">=", 12.0),
    "ROA": (">=", 1.0),
    "DER": ("<=", 1.5),
    "NPM": (">=", 7.0),
    "NIM": (">=", 3.5),
    "CAR": (">=", 15.0),
    "LDR": (">=", 75.0),
    "NPL": ("<=", 3.5),
    "BOPO": ("<=", 80.0),
    "CIR": ("<=", 65.0),
    "LAR": ("<=", 11.0),
}

PROFILE_WEIGHTS = {
    "Balanced": BASE_WEIGHTS,
    "Defensive": {
        "valuation": 22,
        "quality": 32,
        "risk": 22,
        "liquidity": 14,
        "momentum": 5,
        "index_strength": 5,
    },
    "Growth": {
        "valuation": 18,
        "quality": 34,
        "risk": 10,
        "liquidity": 14,
        "momentum": 18,
        "index_strength": 6,
    },
    "Value": {
        "valuation": 38,
        "quality": 25,
        "risk": 14,
        "liquidity": 12,
        "momentum": 6,
        "index_strength": 5,
    },
}

HELP_TEXT = {
    "profile": "Profil scoring hanya mengatur bobot faktor Score. Balanced seimbang, Defensive menekankan kualitas/risiko, Growth menekankan pertumbuhan dan momentum, Value menekankan valuasi.",
    "filter_preset": "Preset filter hanya mengatur ketat/longgarnya penyaringan dan tidak mengubah bobot Score. Konservatif Aman mengetatkan harga minimal 50, volume minimal 10 juta, PER 0.1-25, PBV <= 3.5, ROE >= 8%, NPM >= 3%, DER <= 1.5, Score >= 60, Threshold >= 65%, wajib inti valuasi/profit, dan Data Bersih saja.",
    "valuation": "Valuation_Score dihitung dari PER dan PBV yang valid. Skor tinggi berarti valuasi relatif lebih murah di antara kandidat, bukan berarti harga pasti akan naik.",
    "quality": "Quality_Score dihitung dari ROE, ROA, NPM, dan metrik bank bila tersedia. Skor tinggi berarti profitabilitas dan efisiensi historis lebih baik.",
    "risk": "Risk_Score adalah skor risiko relatif. Non-bank memakai DER dan volatilitas/pergerakan harga; bank memakai NIM, CAR, LDR, NPL, BOPO, serta CIR/LAR bila ada. Skor tinggi berarti risiko model lebih rendah, bukan bebas risiko.",
    "liquidity": "Liquidity_Score dihitung dari volume dan turnover. Skor tinggi berarti saham lebih aktif diperdagangkan, tetapi tetap tidak menjamin order besar bisa dieksekusi tanpa slippage.",
    "momentum": "Momentum_Score memakai %Change dan return historis 4W, 13W, 26W, 52W, serta YTD bila tersedia. Ini membaca tren historis, bukan prediksi harga.",
    "index_strength": "Index_Score memakai jumlah kemunculan saham pada indeks/sumber data. Bila tersedia, nilai utama diambil dari kolom Sigma i >= 7 di Excel sebagai sinyal coverage indeks; ini bukan jaminan kualitas perusahaan.",
    "sector": "Menyaring berdasarkan sektor bisnis dari daftar resmi BEI/IDX bila tersedia, lalu Excel fallback. Tidak mengubah rumus score, hanya membatasi saham yang dianalisis.",
    "industry": "Menyaring berdasarkan industri yang lebih spesifik di dalam sektor. Tidak mengubah score dasar.",
    "price": "Penutupan diprioritaskan dari yfinance online, lalu diisi Excel bila online kosong. Filter ini membatasi rentang harga nominal, bukan valuasi murah/mahal.",
    "volume": "Volume diprioritaskan dari yfinance online, lalu diisi Excel bila online kosong. Volume minimum membantu membuang saham sangat tidak likuid.",
    "per": "PER = harga saham / laba per saham. PER rendah sering terlihat murah jika laba positif; PER nol, negatif, atau ekstrem diberi penalti/dianggap tidak sehat.",
    "pbv": "PBV = harga saham / nilai buku per saham. PBV rendah berarti harga lebih dekat ke nilai buku, tetapi perlu dibaca bersama ROE, kualitas aset, dan sektor.",
    "roe": "ROE = laba bersih / ekuitas. Semakin tinggi semakin efisien modal pemegang saham menghasilkan laba, selama tidak didorong leverage berlebihan.",
    "npm": "NPM = laba bersih / pendapatan. Semakin tinggi berarti margin laba bersih lebih kuat; nilai negatif menunjukkan rugi bersih.",
    "der": "DER = total utang / ekuitas. Untuk non-bank, makin rendah umumnya lebih konservatif. Untuk bank, DER tidak otomatis buruk sehingga default tidak diterapkan ke Banking.",
    "score": "Score akhir = rata-rata tertimbang Valuasi, Kualitas, Risiko, Likuiditas, Momentum, dan Indeks, lalu dikurangi Penalty dan dibatasi 0-100.",
    "threshold_source": "Sumber aturan threshold. Auto memakai Banking untuk saham bank dan NonBank untuk saham selain bank; pilihan manual memaksa semua saham memakai satu set aturan.",
    "threshold_ratio": "Threshold_Pass_Ratio = Threshold_Pass_Count / Threshold_Applicable x 100. Rasio yang kolomnya ada tetapi nilainya kosong/tidak memenuhi batas dihitung tidak lolos.",
    "core_thresholds": "Jika aktif, saham wajib memenuhi inti konservatif: PER <= 15, PBV <= 3, ROE >= 12%, dan NPM >= 7%. Ini tambahan di luar slider umum.",
    "der_banking": "Jika aktif, filter DER maksimum juga diterapkan ke saham Banking. Default mati karena struktur neraca bank berbeda dari non-bank.",
    "history_source": "Online yfinance memakai ticker KODE.JK sebagai sumber histori utama. Excel Metrik tetap tersedia sebagai cadangan/pembanding bila data online kosong.",
    "fundamental_source": "Fundamental diprioritaskan dari online TradingView scanner bila tersedia, lalu Excel mengisi rasio/metadata yang kosong. Sumber BEI/IDX tetap utama untuk universe kode saham.",
    "history_scope": "Saham pilihan memakai kode yang dipilih manual. All/top N memakai saham teratas dari hasil filter saat ini, biasanya berdasarkan ranking Score setelah filter.",
    "history_top_n": "Jumlah kode dari hasil filter/ranking yang dimasukkan ke grafik All/top N. Makin besar makin lengkap, tetapi grafik online bisa lebih lambat.",
    "history_codes": "Masukkan kode IDX tanpa akhiran .JK, misalnya BBCA atau BBRI. Dashboard otomatis memanggil format online BBCA.JK.",
    "history_period": "Rentang data online: 1 minggu = sekitar 5 hari bursa terakhir, 1/3/6 bulan, 1/2/5/10 tahun, atau All sepanjang data tersedia dari sumber.",
    "recommendation": "Recommendation murni dari Score: Strong Buy >= 78, Buy >= 68, Watchlist >= 55, Speculative >= 42, selain itu Avoid. Ini hasil screener, bukan instruksi beli.",
    "risk_level": "Risk_Level adalah kategori risiko relatif dari model berdasarkan rasio, volatilitas, likuiditas, dan penalti. Tetap perlu validasi berita dan laporan keuangan.",
    "turnover": "Turnover = Penutupan x Volume. Ini estimasi nilai transaksi kasar dari data online atau Excel fallback.",
    "return": "Return = harga akhir / harga awal - 1. Nilai ditampilkan dalam persen dan hanya menjelaskan performa periode historis.",
    "reco_sort": "Metrik untuk mengurutkan grafik dan tabel rekomendasi. Mengubah urutan tampilan saja, bukan rumus Score.",
    "reco_limit": "Jumlah saham yang ditampilkan setelah seluruh filter sidebar, label rekomendasi, dan sort diterapkan.",
    "reco_ascending": "Aktifkan untuk mengurutkan dari nilai terendah ke tertinggi pada metrik pilihan. Berguna untuk audit saham lemah atau rasio rendah.",
    "table_columns": "Pilih kolom tabel yang ingin ditampilkan. Ini hanya mengatur tampilan, tidak mengubah data, filter, atau scoring.",
    "explorer_axis": "Pilih rasio atau skor untuk sumbu grafik eksplorasi. Pilihan sumbu hanya mengubah visual scatter, bukan rekomendasi.",
    "explore_color": "Warna titik menunjukkan dimensi tambahan seperti Score, risiko, rekomendasi, atau sektor agar pola lebih mudah dibaca.",
    "explore_size": "Ukuran bubble memakai metrik seperti Volume, Turnover, Score, Liquidity_Score, atau Index_Count. Ini hanya encoding visual.",
    "explore_limit": "Membatasi jumlah titik scatter agar grafik tetap cepat dan mudah dibaca. Data rekomendasi utama tidak dipotong.",
    "histogram": "Pilih sampai tiga metrik untuk melihat distribusi nilai. Berguna untuk mendeteksi outlier dan sebaran rasio.",
    "history_chart_type": "Line cocok untuk perbandingan presisi banyak saham. Area lebih enak untuk satu atau sedikit saham karena area bisa menumpuk secara visual.",
    "history_table": "Menampilkan ringkasan angka histori seperti tanggal terakhir, harga awal/akhir, dan return total di bawah grafik.",
    "sector_group": "Pilih agregasi berdasarkan sektor, subsektor, industri, subindustri, atau industry fallback. Ringkasan dihitung dari seluruh scored_df.",
    "sector_min": "Sembunyikan kelompok dengan jumlah saham terlalu sedikit agar median dan ranking kelompok tidak mudah bias.",
    "sector_sort": "Metrik untuk ranking kelompok sektor/industri, misalnya median score, jumlah Strong Buy, market cap, revenue, turnover, rata-rata ROE, atau jumlah saham.",
    "sector_chart": "Bar untuk ranking, Treemap untuk komposisi market cap, Scatter untuk membaca hubungan score dan ukuran kelompok.",
    "factor_inspect": "Pilih faktor untuk melihat distribusi dan contoh saham teratas. Ini alat audit metodologi, bukan filter baru.",
    "factor_top_n": "Jumlah contoh saham teratas yang ditampilkan untuk faktor yang sedang diinspeksi.",
    "quality_issue": "Pilih jenis masalah data untuk melihat contoh saham yang perlu direview. Detail ini membantu membersihkan sumber data sebelum memakai rekomendasi.",
    "audit_code": "Pilih satu atau beberapa kode saham untuk melihat alasan lolos/gagal pada filter aktif dan preset pembanding.",
    "audit_scope": "Cakupan audit filter. Semua saham mengecek seluruh universe final, Hasil filter aktif hanya saham yang lolos filter sidebar, Kode pilihan untuk investigasi manual.",
    "universe_audit": "Universe kode saham diprioritaskan dari daftar resmi BEI/IDX. Setelah itu dashboard melengkapi data dari sumber online seperti yfinance dan TradingView scanner, lalu Excel hanya sebagai fallback/ide algoritme.",
    "refresh_period": "Periode histori online yang akan diambil saat memperbarui cache. Pilih lebih panjang untuk analisis historis, lebih pendek untuk refresh cepat.",
    "refresh_top_n": "Jumlah saham teratas berdasarkan Index_Count yang cache historinya akan diperbarui dari sumber online.",
    "clean_data": "Jika aktif, hanya tampil saham Clean_Data=True: kode valid, harga > 0, volume >= 10 juta, PER 0.1-35, PBV 0.05-8, ROE >= 5, ROA ada, NPM >= 0, threshold >= 55%, Risk_Level bukan High, Penalty <= 10, metrik bank lengkap, dan DER non-bank <= 2.5.",
}

ANALYSIS_COLUMNS = [
    "Score",
    "Valuation_Score",
    "Quality_Score",
    "Risk_Score",
    "Liquidity_Score",
    "Momentum_Score",
    "History_Momentum_Score",
    "Threshold_Pass_Ratio",
    "Penutupan",
    "PER",
    "PBV",
    "ROE",
    "ROA",
    "DER",
    "NPM",
    "%Change",
    "Return_4W",
    "Return_13W",
    "Return_26W",
    "Return_52W",
    "Return_YTD",
    "Volume",
    "Turnover",
    "Index_Count",
    "Market_Cap",
    "Revenue",
    "Sales_Multiple",
]


st.set_page_config(
    page_title="Dashboard Rekomendasi Saham IDX",
    page_icon=":chart_with_upwards_trend:",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.25rem; }
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 14px 16px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }
    .small-note {
        color: #64748b;
        font-size: 0.86rem;
        line-height: 1.35;
    }
    .recommendation-card {
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 14px 16px;
        background: #ffffff;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def clean_text(value, default="-"):
    if pd.isna(value):
        return default
    value = str(value).strip()
    return value if value else default


def format_number(value, digits=1):
    if pd.isna(value):
        return "-"
    return f"{value:,.{digits}f}"


def format_rupiah(value):
    if pd.isna(value):
        return "-"
    return f"Rp {value:,.0f}"


def format_large_rupiah(value):
    if pd.isna(value):
        return "-"
    value = float(value)
    abs_value = abs(value)
    if abs_value >= 1_000_000_000_000:
        return f"Rp {value / 1_000_000_000_000:,.1f}T"
    if abs_value >= 1_000_000_000:
        return f"Rp {value / 1_000_000_000:,.1f}B"
    if abs_value >= 1_000_000:
        return f"Rp {value / 1_000_000:,.1f}M"
    return f"Rp {value:,.0f}"


def stretch_kwargs(func):
    try:
        params = inspect.signature(func).parameters
    except (TypeError, ValueError):
        return {"use_container_width": True}
    if "width" in params:
        return {"width": "stretch"}
    return {"use_container_width": True}


DATAFRAME_STRETCH = stretch_kwargs(st.dataframe)
PLOTLY_STRETCH = stretch_kwargs(st.plotly_chart)


def show_table(data=None, *args, **kwargs):
    kwargs.pop("width", None)
    kwargs.pop("use_container_width", None)
    return st.dataframe(data, *args, **DATAFRAME_STRETCH, **kwargs)


def show_chart(fig, *args, **kwargs):
    kwargs.pop("width", None)
    kwargs.pop("use_container_width", None)
    return st.plotly_chart(fig, *args, **PLOTLY_STRETCH, **kwargs)


def prepare_chart_frame(data, metric, limit=None):
    chart = data.copy()
    chart["Kode"] = chart["Kode"].astype(str).str.strip().str.upper()
    chart = chart[chart["Kode"].notna() & ~chart["Kode"].isin(["", "-", "NAN", "NONE"])]
    if metric in chart.columns:
        chart[metric] = pd.to_numeric(chart[metric], errors="coerce")
        chart = chart[chart[metric].notna() & np.isfinite(chart[metric])]
    if limit is not None:
        chart = chart.head(limit)
    chart["Chart_Label"] = chart["Kode"]
    return chart


def build_completeness_report(data):
    groups = {
        "Identitas": ["Kode", "Nama Perusahaan", "Sektor", "Industry", "ListingBoard"],
        "Hierarki Industri": ["Sektor_Metrik", "Subsektor", "Industri", "Subindustri"],
        "Harga & Likuiditas": ["Penutupan", "Volume", "Turnover", "%Change"],
        "Ukuran Emiten": ["Market_Cap", "Revenue", "Sales_Multiple"],
        "Valuasi": ["PER", "PBV"],
        "Profitabilitas": ["ROE", "ROA", "NPM"],
        "Banking": ["NIM", "CAR", "LDR", "NPL", "BOPO", "CIR", "LAR"],
        "Histori": ["Return_4W", "Return_13W", "Return_26W", "Return_52W", "Return_YTD"],
        "Scoring": ["Score", "Valuation_Score", "Quality_Score", "Risk_Score", "Liquidity_Score", "Momentum_Score", "Index_Score"],
        "Sumber Data": ["Price_Source", "Volume_Source", "Fundamental_Source", "Universe_Source", "Universe_Diff_Status"],
    }
    rows = []
    for group, columns in groups.items():
        available_columns = [column for column in columns if column in data.columns]
        if not available_columns:
            continue
        scoped_data = data
        if group == "Banking" and "Threshold_Mode" in data.columns:
            scoped_data = data[data["Threshold_Mode"].eq("Banking")]
        total_rows = max(len(scoped_data), 1)
        for column in available_columns:
            available = int(scoped_data[column].notna().sum())
            rows.append(
                {
                    "Grup": group,
                    "Kolom": column,
                    "Terisi": available,
                    "Kosong": int(len(scoped_data) - available),
                    "Coverage": available / total_rows * 100,
                }
            )
    return pd.DataFrame(rows)


def build_source_mix(data):
    source_columns = [column for column in ["Price_Source", "Volume_Source", "Fundamental_Source", "Data_Source", "Universe_Source", "Universe_Diff_Status"] if column in data.columns]
    rows = []
    for column in source_columns:
        counts = data[column].fillna("Tidak diketahui").astype(str).value_counts().reset_index()
        counts.columns = ["Nilai", "Jumlah"]
        counts["Area"] = column
        rows.append(counts[["Area", "Nilai", "Jumlah"]])
    if not rows:
        return pd.DataFrame(columns=["Area", "Nilai", "Jumlah"])
    return pd.concat(rows, ignore_index=True)


def get_file_status(path):
    file_path = Path(path)
    if not file_path.exists():
        return {"File": str(path), "Status": "Tidak ditemukan", "Ukuran": "-", "Last Modified": "-"}
    stat = file_path.stat()
    modified = pd.Timestamp.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    return {
        "File": str(path),
        "Status": "OK",
        "Ukuran": f"{stat.st_size / 1024 / 1024:.2f} MB",
        "Last Modified": modified,
    }


def get_data_update_label(raw, file_status):
    source = raw.attrs.get("data_source") if hasattr(raw, "attrs") else None
    online_update = raw.attrs.get("online_update") if hasattr(raw, "attrs") else None
    if online_update:
        return f"{source or 'online'}, {online_update}"

    last_update_column = next((column for column in raw.columns if str(column).lower().startswith("last update")), None)
    if last_update_column:
        label = str(last_update_column).replace("Last Update:", "", 1).strip()
        match = re.search(r"(\d{8})\s+(\d{1,2}:\d{2})", label)
        if match:
            prefix = label[: match.start()].strip(" ,")
            date_text = pd.to_datetime(match.group(1), format="%Y%m%d", errors="coerce")
            if pd.notna(date_text):
                formatted = f"{date_text:%Y-%m-%d} {match.group(2)}"
                return f"{prefix}, {formatted}" if prefix else formatted
        if label:
            return label
    modified = file_status.get("Last Modified", "-")
    return f"file modified {modified}" if modified != "-" else "belum tersedia"


def clean_stock_code(value):
    code = str(value).strip().upper()
    code = re.sub(r"\.JK$", "", code)
    code = re.sub(r"[^A-Z0-9-]", "", code)
    if code in {"", "-", "NAN", "NONE", "NULL"}:
        return ""
    return code


def count_index_memberships(value):
    if pd.isna(value):
        return 0
    text = str(value).strip()
    if not text or text.upper() in {"-", "NAN", "NONE"}:
        return 0
    parts = [part.strip() for part in re.split(r"[,;|]", text) if part.strip()]
    return len(set(parts))


def first_existing_column(dataframe, candidates):
    normalized_lookup = {str(column).strip().lower(): column for column in dataframe.columns}
    for candidate in candidates:
        found = normalized_lookup.get(candidate.strip().lower())
        if found is not None:
            return found
    return None


def find_index_sigma_column(columns):
    for column in columns:
        text = str(column).strip().lower()
        if "index" in text or "update" in text or "expired" in text:
            continue
        if "∑" in text or "sigma" in text or text.startswith("sum"):
            return column
    for column in columns:
        text = str(column).strip().lower()
        if "≥" in text and "7" in text and "index" not in text:
            return column
    return None


def normalize_universe_frame(dataframe, source):
    if dataframe is None or dataframe.empty:
        return pd.DataFrame(columns=["Kode", "Nama Perusahaan", "Sektor", "Industry", "ListingDate", "Shares", "ListingBoard", "Universe_Source"])

    code_column = first_existing_column(
        dataframe,
        ["Kode", "Code", "Stock Code", "StockCode", "Ticker", "Symbol", "EmitenCode"],
    )
    if code_column is None:
        return pd.DataFrame(columns=["Kode", "Nama Perusahaan", "Sektor", "Industry", "ListingDate", "Shares", "ListingBoard", "Universe_Source"])

    name_column = first_existing_column(
        dataframe,
        ["Nama Perusahaan", "Company Name", "Name", "Company", "EmitenName", "SecurityName"],
    )
    sector_column = first_existing_column(dataframe, ["Sektor", "Sector"])
    industry_column = first_existing_column(dataframe, ["Industry", "Subsector", "Sub Sector", "SubSektor"])
    listing_date_column = first_existing_column(dataframe, ["ListingDate", "Listing Date", "Tanggal Pencatatan"])
    shares_column = first_existing_column(dataframe, ["Shares", "Share", "Saham"])
    listing_board_column = first_existing_column(dataframe, ["ListingBoard", "Listing Board", "Papan Pencatatan", "Board"])

    universe = pd.DataFrame()
    universe["Kode"] = dataframe[code_column].map(clean_stock_code)
    universe["Nama Perusahaan"] = dataframe[name_column].map(clean_text) if name_column else "-"
    universe["Sektor"] = dataframe[sector_column].map(lambda value: clean_text(value, "No Sector")) if sector_column else "No Sector"
    universe["Industry"] = dataframe[industry_column].map(lambda value: clean_text(value, "No Industry")) if industry_column else "No Industry"
    universe["ListingDate"] = pd.to_datetime(dataframe[listing_date_column], errors="coerce") if listing_date_column else pd.NaT
    universe["Shares"] = pd.to_numeric(dataframe[shares_column], errors="coerce") if shares_column else np.nan
    universe["ListingBoard"] = dataframe[listing_board_column].map(lambda value: clean_text(value, "-")) if listing_board_column else "-"
    universe["Universe_Source"] = source
    universe = universe[universe["Kode"].ne("")]
    universe = universe.drop_duplicates("Kode")
    return universe


def read_url_text(url):
    request = Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "text/html,application/json;q=0.9,*/*;q=0.8",
        },
    )
    removed_proxy = remove_blocking_proxy_env()
    try:
        with urlopen(request, timeout=25) as response:
            charset = response.headers.get_content_charset() or "utf-8"
            return response.read().decode(charset, errors="replace")
    finally:
        restore_proxy_env(removed_proxy)


def post_json(url, payload):
    request = Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "application/json,text/plain,*/*",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    removed_proxy = remove_blocking_proxy_env()
    try:
        with urlopen(request, timeout=25) as response:
            charset = response.headers.get_content_charset() or "utf-8"
            return json.loads(response.read().decode(charset, errors="replace"))
    finally:
        restore_proxy_env(removed_proxy)


def load_tradingview_universe():
    payload = {
        "columns": ["name", "description", "sector", "industry"],
        "filter": [{"left": "exchange", "operation": "equal", "right": "IDX"}],
        "options": {"lang": "en"},
        "range": [0, 1200],
        "sort": {"sortBy": "name", "sortOrder": "asc"},
        "markets": ["indonesia"],
        "symbols": {"query": {"types": []}, "tickers": []},
    }
    response = post_json(TRADINGVIEW_SCAN_URL, payload)
    rows = []
    for item in response.get("data", []):
        values = item.get("d", [])
        rows.append(
            {
                "Kode": values[0] if len(values) > 0 else item.get("s", "").split(":")[-1],
                "Nama Perusahaan": values[1] if len(values) > 1 else "-",
                "Sektor": values[2] if len(values) > 2 else "No Sector",
                "Industry": values[3] if len(values) > 3 else "No Industry",
            }
        )
    return normalize_universe_frame(pd.DataFrame(rows), "TradingView online")


@st.cache_data(ttl=ONLINE_REFRESH_TTL, show_spinner=False)
def load_online_fundamentals():
    payload = {
        "columns": TRADINGVIEW_FUNDAMENTAL_COLUMNS,
        "filter": [{"left": "exchange", "operation": "equal", "right": "IDX"}],
        "options": {"lang": "en"},
        "range": [0, 1500],
        "sort": {"sortBy": "name", "sortOrder": "asc"},
        "markets": ["indonesia"],
        "symbols": {"query": {"types": []}, "tickers": []},
    }
    try:
        response = post_json(TRADINGVIEW_SCAN_URL, payload)
    except Exception as exc:
        output = pd.DataFrame(columns=["Kode"])
        output.attrs["fundamental_source"] = "empty"
        output.attrs["fundamental_error"] = f"TradingView fundamental gagal: {exc}"
        return output

    rows = []
    for item in response.get("data", []):
        values = item.get("d", [])
        row = dict(zip(TRADINGVIEW_FUNDAMENTAL_COLUMNS, values))
        code = clean_stock_code(row.get("name", item.get("s", "").split(":")[-1]))
        if not code:
            continue
        rows.append(
            {
                "Kode": code,
                "Nama_Perusahaan_Online": clean_text(row.get("description"), "-"),
                "Sektor_Online": clean_text(row.get("sector"), "No Sector"),
                "Industry_Online": clean_text(row.get("industry"), "No Industry"),
                "Market_Cap_Online": row.get("market_cap_basic"),
                "Revenue_Online": row.get("total_revenue"),
                "PER_Online": row.get("price_earnings_ttm"),
                "PBV_Online": row.get("price_book_fq"),
                "ROE_Online": row.get("return_on_equity_fq"),
                "ROA_Online": row.get("return_on_assets_fq"),
                "DER_Online": row.get("debt_to_equity_fq"),
                "NPM_Online": row.get("net_margin_ttm"),
                "Fundamental_Source": "TradingView scanner",
            }
        )

    output = pd.DataFrame(rows).drop_duplicates("Kode") if rows else pd.DataFrame(columns=["Kode"])
    for column in ["Market_Cap_Online", "Revenue_Online", "PER_Online", "PBV_Online", "ROE_Online", "ROA_Online", "DER_Online", "NPM_Online"]:
        if column in output.columns:
            output[column] = pd.to_numeric(output[column], errors="coerce")
    output.attrs["fundamental_source"] = "TradingView scanner" if not output.empty else "empty"
    output.attrs["fundamental_error"] = None if not output.empty else "TradingView scanner tidak mengembalikan data fundamental."
    return output


def load_official_idx_universe():
    payload = json.loads(read_url_text(IDX_STOCK_LIST_URL))
    records = payload.get("data", payload if isinstance(payload, list) else [])
    universe = normalize_universe_frame(pd.DataFrame(records), "BEI/IDX official")
    universe.attrs["records_total"] = payload.get("recordsTotal") if isinstance(payload, dict) else len(universe)
    return universe


@st.cache_data(ttl=ONLINE_REFRESH_TTL, show_spinner=False)
def load_idx_universe_online():
    errors = []

    try:
        universe = load_official_idx_universe()
        if not universe.empty:
            universe.attrs["universe_error"] = None
            return universe
    except Exception as exc:
        errors.append(f"BEI/IDX: {exc}")

    try:
        universe = load_tradingview_universe()
        if not universe.empty:
            universe.attrs["universe_error"] = None
            return universe
    except Exception as exc:
        errors.append(f"TradingView: {exc}")

    for url in IDX_COMPANY_PROFILE_URLS:
        try:
            payload = pd.read_json(StringIO(read_url_text(url)))
            if isinstance(payload, pd.DataFrame) and not payload.empty:
                if "data" in payload.columns and payload["data"].apply(lambda value: isinstance(value, dict)).any():
                    payload = pd.DataFrame(payload["data"].dropna().tolist())
                universe = normalize_universe_frame(payload, "IDX online")
                if not universe.empty:
                    universe.attrs["universe_error"] = None
                    return universe
        except Exception as exc:
            errors.append(f"IDX: {exc}")

    try:
        tables = pd.read_html(StringIO(read_url_text(STOCKANALYSIS_IDX_URL)))
        for table in tables:
            universe = normalize_universe_frame(table, "StockAnalysis online")
            if not universe.empty:
                universe.attrs["universe_error"] = "; ".join(errors) if errors else None
                return universe
    except Exception as exc:
        errors.append(f"StockAnalysis: {exc}")

    fallback = pd.DataFrame(columns=["Kode", "Nama Perusahaan", "Sektor", "Industry", "Universe_Source"])
    fallback.attrs["universe_error"] = "; ".join(errors) if errors else "Sumber universe online kosong."
    return fallback


def ensure_expected_columns(dataframe, columns):
    output = dataframe.copy()
    for column in columns:
        if column not in output.columns:
            output[column] = np.nan
    return output


def get_history_cache_status():
    if not HISTORY_CACHE_DIR.exists():
        return pd.DataFrame(columns=["Cache File", "Kode", "Period", "Rows", "Last Date", "Modified"])
    rows = []
    for cache_file in sorted(HISTORY_CACHE_DIR.glob("*.csv")):
        parts = cache_file.stem.rsplit("_", 1)
        code = parts[0] if parts else cache_file.stem
        period = parts[1] if len(parts) > 1 else "-"
        try:
            cached = pd.read_csv(cache_file)
            last_date = pd.to_datetime(cached.get("Date"), errors="coerce").max()
            rows_count = len(cached)
            last_date_text = last_date.strftime("%Y-%m-%d") if pd.notna(last_date) else "-"
        except Exception:
            rows_count = 0
            last_date_text = "Rusak/tidak terbaca"
        modified = pd.Timestamp.fromtimestamp(cache_file.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        rows.append(
            {
                "Cache File": cache_file.name,
                "Kode": code,
                "Period": period,
                "Rows": rows_count,
                "Last Date": last_date_text,
                "Modified": modified,
            }
        )
    return pd.DataFrame(rows)


def load_latest_history_cache_snapshot():
    if not HISTORY_CACHE_DIR.exists():
        return pd.DataFrame(columns=["Kode", "Online_Last_Date", "Close_Online", "Volume_Online_Latest"])
    frames = []
    for cache_file in HISTORY_CACHE_DIR.glob("*.csv"):
        try:
            cached = pd.read_csv(cache_file)
        except Exception:
            continue
        if not {"Kode", "Date", "Close", "Volume_Online"}.issubset(cached.columns):
            continue
        cached = cached[["Kode", "Date", "Close", "Volume_Online"]].copy()
        cached["Date"] = pd.to_datetime(cached["Date"], errors="coerce")
        cached["Close"] = pd.to_numeric(cached["Close"], errors="coerce")
        cached["Volume_Online"] = pd.to_numeric(cached["Volume_Online"], errors="coerce")
        cached["Kode"] = cached["Kode"].astype(str).str.strip().str.upper()
        cached = cached.dropna(subset=["Kode", "Date", "Close"])
        if not cached.empty:
            frames.append(cached)
    if not frames:
        return pd.DataFrame(columns=["Kode", "Online_Last_Date", "Close_Online", "Volume_Online_Latest"])
    snapshot = pd.concat(frames, ignore_index=True).sort_values(["Kode", "Date"])
    snapshot = snapshot.groupby("Kode", as_index=False).tail(1)
    return snapshot.rename(
        columns={"Date": "Online_Last_Date", "Close": "Close_Online", "Volume_Online": "Volume_Online_Latest"}
    )[["Kode", "Online_Last_Date", "Close_Online", "Volume_Online_Latest"]]


def apply_online_volume_fallback(df):
    snapshot = load_latest_history_cache_snapshot()
    output = df.copy()
    output["Volume_Original"] = output["Volume"]
    output["Volume_Source"] = "Excel"
    if snapshot.empty:
        output["Close_Online"] = np.nan
        output["Volume_Online_Latest"] = np.nan
        output["Online_Last_Date"] = pd.NaT
        return output

    output = output.merge(snapshot, on="Kode", how="left")
    close_match = (
        output["Penutupan"].gt(0)
        & output["Close_Online"].gt(0)
        & ((output["Close_Online"] - output["Penutupan"]).abs() / output["Penutupan"]).le(0.05)
    )
    use_online_volume = (
        output["Volume"].fillna(0).lt(1_000_000)
        & output["Volume_Online_Latest"].fillna(0).ge(1_000_000)
        & close_match
    )
    output.loc[use_online_volume, "Volume"] = output.loc[use_online_volume, "Volume_Online_Latest"]
    output.loc[use_online_volume, "Volume_Source"] = "Online cache"
    return output


def build_data_quality_report(scored, raw):
    invalid_code = scored["Kode"].isna() | scored["Kode"].astype(str).str.strip().str.upper().isin(["", "-", "NAN", "NONE"])
    invalid_price = scored["Penutupan"].isna() | scored["Penutupan"].le(0)
    invalid_volume = scored["Volume"].isna() | scored["Volume"].le(0)
    invalid_per = scored["PER"].isna() | scored["PER"].le(0)
    invalid_pbv = scored["PBV"].isna() | scored["PBV"].le(0)
    missing_profit = scored[["ROE", "ROA", "NPM"]].isna().any(axis=1)
    missing_size_data = scored.get("Market_Cap", pd.Series(index=scored.index)).isna() | scored.get("Revenue", pd.Series(index=scored.index)).isna()
    bank_metric_columns = [column for column in ["NIM", "CAR", "LDR", "NPL", "BOPO"] if column in scored.columns]
    bank_missing_metrics = scored["Threshold_Mode"].eq("Banking") & scored[bank_metric_columns].isna().any(axis=1) if bank_metric_columns else pd.Series(False, index=scored.index)
    low_threshold = scored["Threshold_Pass_Ratio"].lt(40)
    missing_history = scored["Return_52W"].isna()
    checks = [
        {
            "Area": "Kode",
            "Check": "Kode kosong/tidak valid",
            "Rows": int(invalid_code.sum()),
            "Severity": "High",
            "Action": "Perbaiki kode saham di sumber data sebelum scoring.",
        },
        {
            "Area": "Kode",
            "Check": "Kode duplikat setelah deduplikasi",
            "Rows": int(scored["Kode"].duplicated().sum()),
            "Severity": "Medium",
            "Action": "Cek proses deduplikasi jika nilai tidak nol.",
        },
        {
            "Area": "Harga",
            "Check": "Harga penutupan kosong atau <= 0",
            "Rows": int(invalid_price.sum()),
            "Severity": "High",
            "Action": "Update harga penutupan agar chart dan turnover valid.",
        },
        {
            "Area": "Likuiditas",
            "Check": "Volume kosong atau <= 0",
            "Rows": int(invalid_volume.sum()),
            "Severity": "High",
            "Action": "Update volume karena memengaruhi likuiditas dan penalti.",
        },
        {
            "Area": "Valuasi",
            "Check": "PER negatif atau kosong",
            "Rows": int(invalid_per.sum()),
            "Severity": "Medium",
            "Action": "PER negatif bisa berarti rugi; tetap tampil tapi diberi penalti.",
        },
        {
            "Area": "Valuasi",
            "Check": "PBV negatif atau kosong",
            "Rows": int(invalid_pbv.sum()),
            "Severity": "Medium",
            "Action": "Cek nilai buku atau sumber rasio.",
        },
        {
            "Area": "Profit",
            "Check": "ROE/ROA/NPM ada yang kosong",
            "Rows": int(missing_profit.sum()),
            "Severity": "Medium",
            "Action": "Lengkapi rasio profit agar kualitas score lebih akurat.",
        },
        {
            "Area": "Ukuran Emiten",
            "Check": "Market cap atau revenue kosong",
            "Rows": int(missing_size_data.sum()),
            "Severity": "Low",
            "Action": "Lengkapi sheet Metrik agar konteks ukuran emiten dan MCap/Revenue lebih akurat.",
        },
        {
            "Area": "Threshold",
            "Check": "Threshold lolos < 40 persen",
            "Rows": int(low_threshold.sum()),
            "Severity": "Medium",
            "Action": "Cek saham yang banyak gagal threshold; jangan jadi prioritas tanpa alasan kuat.",
        },
        {
            "Area": "Banking",
            "Check": "Metrik khusus bank ada yang kosong",
            "Rows": int(bank_missing_metrics.sum()),
            "Severity": "Medium",
            "Action": "Lengkapi NIM/CAR/LDR/NPL/BOPO untuk menilai bank lebih adil.",
        },
        {
            "Area": "Histori",
            "Check": "Return 52 minggu kosong",
            "Rows": int(missing_history.sum()),
            "Severity": "Low",
            "Action": "Gunakan tab Histori Harga online untuk melengkapi konteks tren.",
        },
        {
            "Area": "Sumber",
            "Check": "Baris sumber duplikat per kode",
            "Rows": int(len(raw) - raw["Kode"].astype(str).str.strip().str.upper().nunique()) if "Kode" in raw.columns else 0,
            "Severity": "Info",
            "Action": "Normal jika saham muncul di beberapa indeks; dashboard melakukan deduplikasi.",
        },
    ]
    report = pd.DataFrame(checks)
    report["Status"] = np.where(report["Rows"].eq(0), "OK", "Review")
    return report[["Status", "Severity", "Area", "Check", "Rows", "Action"]]


def get_quality_detail(scored, issue_key):
    if issue_key == "Harga penutupan kosong atau <= 0":
        mask = scored["Penutupan"].isna() | scored["Penutupan"].fillna(0).le(0)
    elif issue_key == "Volume kosong atau <= 0":
        mask = scored["Volume"].isna() | scored["Volume"].fillna(0).le(0)
    elif issue_key == "PER negatif atau kosong":
        mask = scored["PER"].isna() | scored["PER"].fillna(0).le(0)
    elif issue_key == "PBV negatif atau kosong":
        mask = scored["PBV"].isna() | scored["PBV"].fillna(0).le(0)
    elif issue_key == "ROE/ROA/NPM ada yang kosong":
        mask = scored[["ROE", "ROA", "NPM"]].isna().any(axis=1)
    elif issue_key == "Market cap atau revenue kosong":
        mask = scored.get("Market_Cap", pd.Series(index=scored.index)).isna() | scored.get("Revenue", pd.Series(index=scored.index)).isna()
    elif issue_key == "Threshold lolos < 40 persen":
        mask = scored["Threshold_Pass_Ratio"].lt(40)
    elif issue_key == "Metrik khusus bank ada yang kosong":
        bank_metric_columns = [column for column in ["NIM", "CAR", "LDR", "NPL", "BOPO"] if column in scored.columns]
        mask = scored["Threshold_Mode"].eq("Banking") & scored[bank_metric_columns].isna().any(axis=1) if bank_metric_columns else pd.Series(False, index=scored.index)
    elif issue_key == "Return 52 minggu kosong":
        mask = scored["Return_52W"].isna()
    else:
        mask = pd.Series(False, index=scored.index)
    columns = ["Kode", "Nama Perusahaan", "Sektor", "Subsektor", "Industri", "Subindustri", "Industry", "Score", "Recommendation", "Penutupan", "Volume", "Market_Cap", "Revenue", "Sales_Multiple", "PER", "PBV", "ROE", "ROA", "NPM", "NIM", "CAR", "LDR", "NPL", "BOPO", "Threshold_Pass_Ratio", "Return_52W"]
    return scored.loc[mask, [column for column in columns if column in scored.columns]].copy()


def add_safety_flags(scored):
    output = scored.copy()
    banking_mask = output["Threshold_Mode"].eq("Banking")
    valid_core = (
        output["Kode"].notna()
        & output["Kode"].astype(str).str.strip().ne("")
        & output["Penutupan"].gt(0)
        & output["Volume"].ge(10_000_000)
        & output["PER"].between(0.1, 35, inclusive="both")
        & output["PBV"].between(0.05, 8, inclusive="both")
        & output["ROE"].ge(5)
        & output["ROA"].notna()
        & output["NPM"].ge(0)
        & output["Threshold_Pass_Ratio"].ge(55)
        & output["Risk_Level"].ne("High")
        & output["Penalty"].le(10)
    )
    bank_metrics = [column for column in ["NIM", "CAR", "LDR", "NPL", "BOPO"] if column in output.columns]
    bank_ok = pd.Series(True, index=output.index)
    if bank_metrics:
        bank_ok = ~banking_mask | output[bank_metrics].notna().all(axis=1)
    nonbank_ok = banking_mask | output["DER"].fillna(np.inf).le(2.5)
    output["Clean_Data"] = valid_core & bank_ok & nonbank_ok
    output["Safety_Recommendation"] = np.where(
        output["Clean_Data"] & output["Score"].ge(78),
        "Bersih - Strong",
        np.where(output["Clean_Data"] & output["Score"].ge(68), "Bersih - Buy", np.where(output["Clean_Data"], "Bersih - Watch", "Review Data/Risiko")),
    )

    reasons = []
    for _, row in output.iterrows():
        row_reasons = []
        if pd.isna(row.get("Penutupan")) or row.get("Penutupan", 0) <= 0:
            row_reasons.append("harga tidak valid")
        if pd.isna(row.get("Volume")) or row.get("Volume", 0) < 10_000_000:
            row_reasons.append("volume rendah/kosong")
        if pd.isna(row.get("PER")) or row.get("PER", 0) <= 0:
            row_reasons.append("PER tidak sehat")
        if pd.isna(row.get("PBV")) or row.get("PBV", 0) <= 0:
            row_reasons.append("PBV tidak sehat")
        if pd.isna(row.get("ROE")) or row.get("ROE", 0) < 5:
            row_reasons.append("ROE rendah/kosong")
        if pd.isna(row.get("NPM")) or row.get("NPM", 0) < 0:
            row_reasons.append("NPM negatif/kosong")
        if row.get("Threshold_Pass_Ratio", 0) < 55:
            row_reasons.append("threshold rendah")
        if row.get("Risk_Level") == "High":
            row_reasons.append("risiko High")
        if row.get("Threshold_Mode") == "Banking":
            missing_bank = [metric for metric in bank_metrics if pd.isna(row.get(metric))]
            if missing_bank:
                row_reasons.append("metrik bank kosong")
        elif row.get("DER", np.inf) > 2.5:
            row_reasons.append("DER tinggi")
        reasons.append(", ".join(row_reasons) if row_reasons else "OK")
    output["Safety_Notes"] = reasons
    return output


def make_filter_criteria(
    name,
    price_range,
    min_volume,
    per_range,
    pbv_max,
    roe_min,
    npm_min,
    der_max,
    apply_der_to_banking,
    min_score,
    min_threshold_ratio,
    require_core_thresholds,
    clean_data_only,
    sector_filter="Semua Sektor",
    industry_filter="Semua Industri",
):
    return {
        "name": name,
        "price_range": price_range,
        "min_volume": min_volume,
        "per_range": per_range,
        "pbv_max": pbv_max,
        "roe_min": roe_min,
        "npm_min": npm_min,
        "der_max": der_max,
        "apply_der_to_banking": apply_der_to_banking,
        "min_score": min_score,
        "min_threshold_ratio": min_threshold_ratio,
        "require_core_thresholds": require_core_thresholds,
        "clean_data_only": clean_data_only,
        "sector_filter": sector_filter,
        "industry_filter": industry_filter,
    }


def default_filter_criteria(name, conservative, price_min, price_max):
    if conservative:
        return make_filter_criteria(
            name=name,
            price_range=(max(price_min, 50), min(price_max, 50_000)),
            min_volume=10_000_000,
            per_range=(0.1, 25.0),
            pbv_max=3.5,
            roe_min=8.0,
            npm_min=3.0,
            der_max=1.5,
            apply_der_to_banking=False,
            min_score=60,
            min_threshold_ratio=65,
            require_core_thresholds=True,
            clean_data_only=True,
        )
    return make_filter_criteria(
        name=name,
        price_range=(price_min, min(price_max, 50_000)),
        min_volume=5_000_000,
        per_range=(0.0, 35.0),
        pbv_max=5.0,
        roe_min=5.0,
        npm_min=0.0,
        der_max=2.5,
        apply_der_to_banking=False,
        min_score=45,
        min_threshold_ratio=50,
        require_core_thresholds=False,
        clean_data_only=False,
    )


def audit_row_against_criteria(row, criteria):
    checks = []

    def add_check(area, ok, actual, required):
        checks.append(
            {
                "Area": area,
                "Status": "Lolos" if bool(ok) else "Gagal",
                "Actual": actual,
                "Required": required,
            }
        )

    if criteria["sector_filter"] != "Semua Sektor":
        add_check("Sektor", row.get("Sektor") == criteria["sector_filter"], row.get("Sektor"), criteria["sector_filter"])
    if criteria["industry_filter"] != "Semua Industri":
        add_check("Industri", row.get("Industry") == criteria["industry_filter"], row.get("Industry"), criteria["industry_filter"])

    price_low, price_high = criteria["price_range"]
    add_check("Harga", price_low <= row.get("Penutupan", np.nan) <= price_high, format_rupiah(row.get("Penutupan")), f"{format_rupiah(price_low)} - {format_rupiah(price_high)}")
    add_check("Volume", row.get("Volume", 0) >= criteria["min_volume"], format_volume(row.get("Volume")), f">= {format_volume(criteria['min_volume'])}")

    per_low, per_high = criteria["per_range"]
    add_check("PER", per_low <= row.get("PER", np.nan) <= per_high, format_number(row.get("PER"), 2), f"{per_low:g} - {per_high:g}")
    add_check("PBV", row.get("PBV", np.inf) <= criteria["pbv_max"], format_number(row.get("PBV"), 2), f"<= {criteria['pbv_max']:g}")
    add_check("ROE", row.get("ROE", -np.inf) >= criteria["roe_min"], f"{format_number(row.get('ROE'))}%", f">= {criteria['roe_min']:g}%")
    add_check("NPM", row.get("NPM", -np.inf) >= criteria["npm_min"], f"{format_number(row.get('NPM'))}%", f">= {criteria['npm_min']:g}%")

    is_banking = row.get("Threshold_Mode") == "Banking"
    der_ok = row.get("DER", np.inf) <= criteria["der_max"] or (is_banking and not criteria["apply_der_to_banking"])
    der_required = "Tidak diterapkan ke Banking" if is_banking and not criteria["apply_der_to_banking"] else f"<= {criteria['der_max']:g}"
    add_check("DER", der_ok, format_number(row.get("DER"), 2), der_required)

    add_check("Score", row.get("Score", 0) >= criteria["min_score"], format_number(row.get("Score")), f">= {criteria['min_score']:g}")
    add_check("Threshold", row.get("Threshold_Pass_Ratio", 0) >= criteria["min_threshold_ratio"], f"{format_number(row.get('Threshold_Pass_Ratio'), 0)}%", f">= {criteria['min_threshold_ratio']:g}%")

    if criteria["require_core_thresholds"]:
        core_ok = (
            row.get("PER", np.inf) <= 15
            and row.get("PBV", np.inf) <= 3
            and row.get("ROE", -np.inf) >= 12
            and row.get("NPM", -np.inf) >= 7
        )
        add_check("Core valuasi/profit", core_ok, f"PER {format_number(row.get('PER'), 2)}, PBV {format_number(row.get('PBV'), 2)}, ROE {format_number(row.get('ROE'))}%, NPM {format_number(row.get('NPM'))}%", "PER<=15, PBV<=3, ROE>=12%, NPM>=7%")

    if criteria["clean_data_only"]:
        add_check("Clean_Data", bool(row.get("Clean_Data")), row.get("Safety_Notes", "-"), "Clean_Data=True")

    check_df = pd.DataFrame(checks)
    failed = check_df[check_df["Status"] == "Gagal"]["Area"].tolist()
    return {
        "Filter": criteria["name"],
        "Status": "Lolos" if not failed else "Gagal",
        "Gagal": len(failed),
        "Alasan Utama": ", ".join(failed[:5]) if failed else "OK",
        "Detail": check_df,
    }


def build_filter_audit(data, selected_codes, criteria_list):
    rows = []
    details = []
    for code in selected_codes:
        matches = data[data["Kode"].eq(code)]
        if matches.empty:
            rows.append({"Kode": code, "Filter": "-", "Status": "Tidak ditemukan", "Gagal": "-", "Alasan Utama": "Kode tidak ada di data"})
            continue
        row = matches.iloc[0]
        for criteria in criteria_list:
            audit = audit_row_against_criteria(row, criteria)
            rows.append(
                {
                    "Kode": code,
                    "Nama Perusahaan": row.get("Nama Perusahaan"),
                    "Filter": audit["Filter"],
                    "Status": audit["Status"],
                    "Gagal": audit["Gagal"],
                    "Alasan Utama": audit["Alasan Utama"],
                    "Score": row.get("Score"),
                    "Recommendation": row.get("Recommendation"),
                    "Clean_Data": row.get("Clean_Data"),
                    "Safety_Notes": row.get("Safety_Notes"),
                    "Volume_Source": row.get("Volume_Source", "Excel"),
                }
            )
            detail = audit["Detail"].copy()
            detail.insert(0, "Kode", code)
            detail.insert(1, "Filter", audit["Filter"])
            details.append(detail)
    detail_df = pd.concat(details, ignore_index=True) if details else pd.DataFrame()
    return pd.DataFrame(rows), detail_df


def format_volume(value):
    if pd.isna(value):
        return "-"
    if value >= 1_000_000_000:
        return f"{value / 1_000_000_000:.2f}B"
    if value >= 1_000_000:
        return f"{value / 1_000_000:.1f}M"
    if value >= 1_000:
        return f"{value / 1_000:.1f}K"
    return f"{value:,.0f}"


def cache_file_for_history(code, period):
    safe_code = str(code).strip().upper().replace("/", "_").replace("\\", "_")
    safe_period = str(period).strip().lower().replace("/", "_").replace("\\", "_")
    return HISTORY_CACHE_DIR / f"{safe_code}_{safe_period}.csv"


def read_history_cache(codes, period):
    frames = []
    for code in codes:
        cache_file = cache_file_for_history(code, period)
        if cache_file.exists():
            try:
                cached = pd.read_csv(cache_file)
                cached["Kode"] = str(code).strip().upper()
                frames.append(cached)
            except Exception:
                continue
    if not frames:
        return pd.DataFrame()
    return normalize_history_frame(pd.concat(frames, ignore_index=True))


def write_history_cache(history, period):
    if history.empty:
        return
    HISTORY_CACHE_DIR.mkdir(exist_ok=True)
    cache_columns = [column for column in ["Date", "Kode", "Open", "High", "Low", "Close", "Volume_Online"] if column in history.columns]
    for code, group in history.groupby("Kode"):
        cache_file = cache_file_for_history(code, period)
        group[cache_columns].to_csv(cache_file, index=False)


def yahoo_period_to_dates(period):
    end = pd.Timestamp.today().normalize()
    days_by_period = {
        "5d": 7,
        "1mo": 31,
        "3mo": 92,
        "6mo": 183,
        "1y": 365,
        "2y": 365 * 2,
        "5y": 365 * 5,
        "10y": 365 * 10,
    }
    if period == "max":
        return None, None
    return end - pd.Timedelta(days=days_by_period.get(period, 365)), end


def remove_blocking_proxy_env():
    removed = {}
    for key in ["HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"]:
        value = os.environ.get(key)
        if value and value.strip().lower() in BLOCKING_PROXY_VALUES:
            removed[key] = value
            os.environ.pop(key, None)
    return removed


def restore_proxy_env(removed):
    for key, value in removed.items():
        os.environ[key] = value


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_yahoo_history(codes, period="max"):
    cleaned_codes = [str(code).strip().upper() for code in codes if str(code).strip()]
    if not cleaned_codes:
        return pd.DataFrame(), "Pilih minimal satu kode saham.", "empty"

    try:
        import yfinance as yf
    except Exception as exc:
        cached = read_history_cache(cleaned_codes, period)
        if not cached.empty:
            return cached, f"Library yfinance belum tersedia, memakai cache lokal. Detail: {exc}", "cache"
        return pd.DataFrame(), f"Library yfinance belum tersedia: {exc}", "empty"

    tickers = [f"{code}.JK" for code in cleaned_codes]
    error_messages = []
    removed_proxy = remove_blocking_proxy_env()
    try:
        downloaded = yf.download(
            tickers=tickers,
            period=period,
            interval="1d",
            auto_adjust=True,
            progress=False,
            threads=True,
            group_by="ticker",
        )
    except Exception as exc:
        downloaded = pd.DataFrame()
        error_messages.append(f"yfinance gagal: {exc}")
    finally:
        restore_proxy_env(removed_proxy)

    if downloaded.empty:
        fallback_history, fallback_error = fetch_history_pandas_datareader(cleaned_codes, period)
        if not fallback_history.empty:
            write_history_cache(fallback_history, period)
            return fallback_history, None, "pandas-datareader"
        if fallback_error:
            error_messages.append(fallback_error)

        cached = read_history_cache(cleaned_codes, period)
        if not cached.empty:
            return cached, "Data live kosong/gagal, memakai cache lokal terakhir.", "cache"
        return pd.DataFrame(), "Data online kosong. Cek koneksi atau ticker IDX. " + " ".join(error_messages), "empty"

    records = []
    if isinstance(downloaded.columns, pd.MultiIndex):
        for code, ticker in zip(cleaned_codes, tickers):
            if ticker not in downloaded.columns.get_level_values(0):
                continue
            ticker_data = downloaded[ticker].reset_index()
            if "Close" not in ticker_data.columns:
                continue
            for _, row in ticker_data.iterrows():
                close = row.get("Close")
                if pd.isna(close):
                    continue
                records.append(
                    {
                        "Date": row.get("Date"),
                        "Kode": code,
                        "Open": row.get("Open", np.nan),
                        "High": row.get("High", np.nan),
                        "Low": row.get("Low", np.nan),
                        "Close": close,
                        "Volume_Online": row.get("Volume", np.nan),
                    }
                )
    else:
        ticker_data = downloaded.reset_index()
        code = cleaned_codes[0]
        for _, row in ticker_data.iterrows():
            close = row.get("Close")
            if pd.isna(close):
                continue
            records.append(
                {
                "Date": row.get("Date"),
                "Kode": code,
                "Open": row.get("Open", np.nan),
                "High": row.get("High", np.nan),
                "Low": row.get("Low", np.nan),
                "Close": close,
                "Volume_Online": row.get("Volume", np.nan),
            }
            )

    history = pd.DataFrame(records)
    if history.empty:
        cached = read_history_cache(cleaned_codes, period)
        if not cached.empty:
            return cached, "Data live tidak menemukan harga penutupan, memakai cache lokal terakhir.", "cache"
        return history, "Data online tidak menemukan harga penutupan.", "empty"

    history = normalize_history_frame(history)
    write_history_cache(history, period)
    return history, None, "yfinance"


def fetch_history_pandas_datareader(codes, period="max"):
    try:
        from pandas_datareader import data as pdr
    except Exception as exc:
        return pd.DataFrame(), f"pandas-datareader belum tersedia: {exc}"

    start, end = yahoo_period_to_dates(period)
    if start is None:
        start = pd.Timestamp("1990-01-01")
        end = pd.Timestamp.today().normalize()

    frames = []
    errors = []
    removed_proxy = remove_blocking_proxy_env()
    try:
        for code in codes:
            ticker = f"{code}.JK"
            try:
                data = pdr.DataReader(ticker, "yahoo", start, end).reset_index()
            except Exception as exc:
                errors.append(f"{code}: {exc}")
                continue
            if data.empty or "Close" not in data.columns:
                continue
            frame = pd.DataFrame(
                {
                    "Date": data["Date"],
                    "Kode": code,
                    "Open": data["Open"] if "Open" in data.columns else np.nan,
                    "High": data["High"] if "High" in data.columns else np.nan,
                    "Low": data["Low"] if "Low" in data.columns else np.nan,
                    "Close": data["Close"],
                    "Volume_Online": data["Volume"] if "Volume" in data.columns else np.nan,
                }
            )
            frames.append(frame)
    finally:
        restore_proxy_env(removed_proxy)

    if not frames:
        return pd.DataFrame(), "Fallback pandas-datareader kosong. " + "; ".join(errors[:3])
    return normalize_history_frame(pd.concat(frames, ignore_index=True)), None


def normalize_history_frame(history):
    if history.empty:
        return history
    history = history.copy()
    history["Date"] = pd.to_datetime(history["Date"], errors="coerce")
    history["Close"] = pd.to_numeric(history["Close"], errors="coerce")
    for column in ["Open", "High", "Low"]:
        if column in history.columns:
            history[column] = pd.to_numeric(history[column], errors="coerce")
    if "Volume_Online" in history.columns:
        history["Volume_Online"] = pd.to_numeric(history["Volume_Online"], errors="coerce")
    history["Kode"] = history["Kode"].astype(str).str.strip().str.upper()
    history = history.dropna(subset=["Date", "Close"]).sort_values(["Kode", "Date"])
    history["Normalized"] = history.groupby("Kode")["Close"].transform(
        lambda series: series / series.iloc[0] * 100 if len(series) else series
    )
    return history


def calculate_window_return(group, days):
    if group.empty:
        return np.nan
    group = group.sort_values("Date")
    latest = group.iloc[-1]
    target_date = latest["Date"] - pd.Timedelta(days=int(days * 1.45))
    candidates = group[group["Date"].le(target_date)]
    start_row = candidates.iloc[-1] if not candidates.empty else group.iloc[0]
    start_close = start_row.get("Close")
    latest_close = latest.get("Close")
    if pd.isna(start_close) or pd.isna(latest_close) or start_close <= 0:
        return np.nan
    return (latest_close / start_close - 1) * 100


def calculate_ytd_return(group):
    if group.empty:
        return np.nan
    group = group.sort_values("Date")
    latest = group.iloc[-1]
    year_start = pd.Timestamp(year=latest["Date"].year, month=1, day=1)
    candidates = group[group["Date"].ge(year_start)]
    start_row = candidates.iloc[0] if not candidates.empty else group.iloc[0]
    start_close = start_row.get("Close")
    latest_close = latest.get("Close")
    if pd.isna(start_close) or pd.isna(latest_close) or start_close <= 0:
        return np.nan
    return (latest_close / start_close - 1) * 100


def build_online_market_frame(codes, period=ONLINE_LOAD_PERIOD):
    history, error, source = fetch_yahoo_history(codes, period=period)
    if history.empty:
        output = pd.DataFrame(columns=["Kode"])
        output.attrs["market_source"] = source
        output.attrs["market_error"] = error
        return output

    history = normalize_history_frame(history)
    rows = []
    for code, group in history.groupby("Kode"):
        group = group.sort_values("Date")
        latest = group.iloc[-1]
        previous = group.iloc[-2] if len(group) > 1 else latest
        close = latest.get("Close")
        previous_close = previous.get("Close")
        daily_change = np.nan
        if pd.notna(close) and pd.notna(previous_close) and previous_close > 0:
            daily_change = (close / previous_close - 1) * 100
        row = {
            "Kode": code,
            "Penutupan": close,
            "Sebelumnya": previous_close,
            "%Change": daily_change,
            "Open": latest.get("Open", np.nan),
            "High": latest.get("High", np.nan),
            "Low": latest.get("Low", np.nan),
            "Volume": latest.get("Volume_Online", np.nan),
            "Online_Last_Date": latest.get("Date"),
            "Close_Online": close,
            "Volume_Online_Latest": latest.get("Volume_Online", np.nan),
            "Price_Source": source,
            "Volume_Source": source,
        }
        for column, days in ONLINE_RETURN_WINDOWS.items():
            row[column] = calculate_window_return(group, days)
        row["Return_YTD"] = calculate_ytd_return(group)
        rows.append(row)

    output = pd.DataFrame(rows)
    output.attrs["market_source"] = source
    output.attrs["market_error"] = error
    return output


def score_percentile(series, higher_is_better=True, valid_mask=None):
    values = pd.to_numeric(series, errors="coerce")
    if valid_mask is None:
        valid_mask = values.notna()
    valid_mask = valid_mask & values.notna()

    output = pd.Series(0.0, index=series.index)
    valid_values = values[valid_mask]
    if valid_values.empty:
        return output

    lower = valid_values.quantile(0.03)
    upper = valid_values.quantile(0.97)
    clipped = values.clip(lower, upper)
    ranks = clipped[valid_mask].rank(pct=True, method="average")
    if not higher_is_better:
        ranks = 1 - ranks
    output.loc[valid_mask] = ranks * 100
    return output.clip(0, 100)


def score_target_range(series, low, high, center=None):
    values = pd.to_numeric(series, errors="coerce")
    if center is None:
        center = (low + high) / 2
    width = max(high - low, 1)
    distance = (values - center).abs()
    score = 100 - (distance / width * 100)
    return score.clip(0, 100).fillna(0)


def read_sheet(sheet_name):
    try:
        df_sheet = pd.read_excel(DATA_FILE, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()
    df_sheet.columns = df_sheet.columns.str.strip()
    return df_sheet


def normalize_ratio_columns(df_sheet):
    rename_map = {}
    for column in df_sheet.columns:
        normalized = column.strip().upper()
        for metric in ["PER", "PBV", "ROE", "ROA", "DER", "NPM", "NIM", "CAR", "LDR", "NPL", "BOPO", "CIR", "LAR"]:
            if normalized.startswith(metric):
                rename_map[column] = metric
                break
    return df_sheet.rename(columns=rename_map)


def aggregate_threshold_sheet(sheet_name):
    threshold_raw = normalize_ratio_columns(read_sheet(sheet_name))
    if threshold_raw.empty or "Kode" not in threshold_raw.columns:
        return pd.DataFrame(columns=["Kode"])
    for column in NUMERIC_COLUMNS:
        if column in threshold_raw.columns:
            threshold_raw[column] = pd.to_numeric(threshold_raw[column], errors="coerce")
    threshold_raw["Kode"] = threshold_raw["Kode"].astype(str).str.strip().str.upper()
    ratio_columns = [
        column
        for column in ["PER", "PBV", "ROE", "ROA", "DER", "NPM", "NIM", "CAR", "LDR", "NPL", "BOPO", "CIR", "LAR"]
        if column in threshold_raw.columns
    ]
    aggregations = {column: "median" for column in ratio_columns}
    aggregations.update(
        {
            "Nama Perusahaan": lambda x: clean_text(x.dropna().iloc[0]) if x.dropna().size else "-",
            "Sektor": lambda x: clean_text(x.dropna().iloc[0], "No Sector") if x.dropna().size else "No Sector",
            "Industry": lambda x: clean_text(x.dropna().iloc[0], "No Industry") if x.dropna().size else "No Industry",
        }
    )
    return threshold_raw.groupby("Kode", as_index=False).agg(aggregations)


def load_history_metrics():
    history = read_sheet("Metrik")
    if history.empty or "Kode Saham" not in history.columns:
        return pd.DataFrame(columns=["Kode"])
    history["Kode"] = history["Kode Saham"].astype(str).str.strip().str.upper()
    keep_columns = [
        "Kode",
        "Sektor",
        "Subsektor",
        "Industri",
        "Subindustri",
        "Index",
        "Mkt Cap",
        "Total Rev",
    ] + list(HISTORY_COLUMNS)
    keep_columns = [column for column in keep_columns if column in history.columns]
    history = history[keep_columns].rename(
        columns={
            **HISTORY_COLUMNS,
            "Sektor": "Sektor_Metrik",
            "Index": "Index_Metrik",
        }
    )
    if "Index_Metrik" in history.columns:
        history["Index_Count_Metrik"] = history["Index_Metrik"].apply(count_index_memberships)
    for column in [
        "Mkt Cap",
        "Total Rev",
        "Return_4W",
        "Return_13W",
        "Return_26W",
        "Return_52W",
        "Return_MTD",
        "Return_YTD",
        "Index_Count_Metrik",
    ]:
        if column in history.columns:
            history[column] = pd.to_numeric(history[column], errors="coerce")
    return history.drop_duplicates("Kode")


def build_excel_fallback_summary():
    raw = read_sheet("Ringkasan")
    if raw.empty or "Kode" not in raw.columns:
        return pd.DataFrame(columns=["Kode"]), raw

    index_count_column = find_index_sigma_column(raw.columns)
    raw["Index_Count_Raw"] = raw[index_count_column] if index_count_column else np.nan
    raw["Index_Count_Sigma"] = raw["Index_Count_Raw"]

    for column in NUMERIC_COLUMNS:
        if column in raw.columns:
            raw[column] = pd.to_numeric(raw[column], errors="coerce")

    raw["Kode"] = raw["Kode"].astype(str).str.strip().str.upper()
    raw = raw[raw["Kode"].notna() & (raw["Kode"] != "") & (raw["Kode"] != "NAN")]

    aggregations = {
        "Nama Perusahaan": lambda x: clean_text(x.dropna().iloc[0]) if x.dropna().size else "-",
        "Sektor": lambda x: clean_text(x.dropna().iloc[0], "No Sector") if x.dropna().size else "No Sector",
        "Industry": lambda x: clean_text(x.dropna().iloc[0], "No Industry") if x.dropna().size else "No Industry",
        "Index": lambda x: ", ".join(sorted({clean_text(v) for v in x.dropna()})),
        "Penutupan": "median",
        "PER": "median",
        "PBV": "median",
        "ROE": "median",
        "ROA": "median",
        "DER": "median",
        "NPM": "median",
        "NIM": "median",
        "CAR": "median",
        "LDR": "median",
        "NPL": "median",
        "BOPO": "median",
        "CIR": "median",
        "LAR": "median",
        "Sebelumnya": "median",
        "%Change": "median",
        "Open": "median",
        "High": "median",
        "Low": "median",
        "Volume": "median",
        "Index_Count_Raw": "max",
        "Index_Count_Sigma": "max",
    }

    existing = {key: value for key, value in aggregations.items() if key in raw.columns}
    summary = raw.groupby("Kode", as_index=False).agg(existing)
    if "Index" in summary.columns:
        summary["Index_Count"] = summary["Index"].apply(lambda text: 0 if not text else len(text.split(", ")))
    else:
        summary["Index"] = ""
        summary["Index_Count"] = 0
    index_signal_columns = [column for column in ["Index_Count", "Index_Count_Raw", "Index_Count_Sigma"] if column in summary.columns]
    if index_signal_columns:
        summary["Index_Count"] = summary[index_signal_columns].max(axis=1).fillna(0)
    return summary, raw


def merge_universe_with_excel_fallback(universe, excel_summary):
    if excel_summary.empty or "Kode" not in excel_summary.columns:
        return universe

    fallback_columns = [column for column in ["Kode", "Nama Perusahaan", "Sektor", "Industry"] if column in excel_summary.columns]
    excel_universe = excel_summary[fallback_columns].copy()
    for column, default in [("Nama Perusahaan", "-"), ("Sektor", "No Sector"), ("Industry", "No Industry")]:
        if column not in excel_universe.columns:
            excel_universe[column] = default
    excel_universe["ListingDate"] = pd.NaT
    excel_universe["Shares"] = np.nan
    excel_universe["ListingBoard"] = "-"
    excel_universe["Universe_Source"] = "Excel fallback"
    excel_universe["In_IDX_Official"] = False
    excel_universe["Universe_Diff_Status"] = "Excel fallback only"

    if universe.empty:
        return excel_universe.drop_duplicates("Kode")

    online_lookup = universe.drop_duplicates("Kode").set_index("Kode")
    output = excel_universe.drop_duplicates("Kode").copy()
    online_codes = output["Kode"].isin(online_lookup.index)
    output.loc[online_codes, "Universe_Source"] = output.loc[online_codes, "Kode"].map(online_lookup["Universe_Source"])
    output["In_IDX_Official"] = online_codes & output["Universe_Source"].eq("BEI/IDX official")
    output["Universe_Diff_Status"] = np.where(
        output["In_IDX_Official"],
        "Match BEI/IDX official",
        np.where(online_codes, "Match non-IDX online fallback", "Excel fallback only"),
    )

    for column in ["Nama Perusahaan", "Sektor", "Industry", "ListingDate", "Shares", "ListingBoard"]:
        online_values = output["Kode"].map(online_lookup[column]) if column in online_lookup.columns else pd.Series(np.nan, index=output.index)
        missing = output[column].isna() | output[column].astype(str).str.strip().isin(["", "-", "No Sector", "No Industry", "nan", "NaT"])
        output.loc[missing & online_values.notna(), column] = online_values[missing & online_values.notna()]

    return output.reset_index(drop=True)


def summarize_universe_source(universe):
    if universe.empty or "Kode" not in universe.columns:
        return "Universe kosong"
    total = universe["Kode"].nunique()
    idx_count = int(universe.get("In_IDX_Official", pd.Series(False, index=universe.index)).fillna(False).sum())
    fallback_count = total - idx_count
    return f"Universe {total:,} kode ({idx_count:,} BEI/IDX official, {fallback_count:,} fallback)"


def is_banking_row(row):
    industry_text = f"{row.get('Industry', '')} {row.get('Industri', '')} {row.get('Subindustri', '')}".lower()
    name_text = str(row.get("Nama Perusahaan", "")).lower()
    return "bank" in industry_text or name_text.startswith("bank ") or " bank " in f" {name_text} "


def threshold_pass(value, operator, threshold):
    if pd.isna(value):
        return False
    if operator == "<=":
        return value <= threshold
    if operator == ">=":
        return value >= threshold
    return False


def apply_threshold_profile(df_input, forced_mode=None):
    output = df_input.copy()
    modes = []
    pass_counts = []
    applicable_counts = []

    for _, row in output.iterrows():
        is_bank = is_banking_row(row) if forced_mode is None else forced_mode == "Banking"
        thresholds = BANKING_THRESHOLDS if is_bank else NONBANK_THRESHOLDS
        passed = 0
        applicable = 0
        for metric, (operator, threshold) in thresholds.items():
            if metric not in output.columns:
                continue
            applicable += 1
            if threshold_pass(row.get(metric), operator, threshold):
                passed += 1
        modes.append("Banking" if is_bank else "NonBank")
        pass_counts.append(passed)
        applicable_counts.append(applicable)

    output["Threshold_Mode"] = modes
    output["Threshold_Pass_Count"] = pass_counts
    output["Threshold_Applicable"] = applicable_counts
    output["Threshold_Pass_Ratio"] = np.where(
        output["Threshold_Applicable"] > 0,
        output["Threshold_Pass_Count"] / output["Threshold_Applicable"] * 100,
        0,
    )
    return output


@st.cache_data(show_spinner=False)
def load_data():
    excel_summary, raw = build_excel_fallback_summary()
    universe = load_idx_universe_online()
    universe_error = universe.attrs.get("universe_error")

    if universe.empty:
        universe = excel_summary[["Kode", "Nama Perusahaan", "Sektor", "Industry"]].copy()
        universe["ListingDate"] = pd.NaT
        universe["Shares"] = np.nan
        universe["ListingBoard"] = "-"
        universe["Universe_Source"] = "Excel fallback"
        universe["In_IDX_Official"] = False
        universe["Universe_Diff_Status"] = "Excel fallback only"
    else:
        universe = merge_universe_with_excel_fallback(universe, excel_summary)

    online_market = build_online_market_frame(universe["Kode"].tolist(), period=ONLINE_LOAD_PERIOD)
    market_source = online_market.attrs.get("market_source", "empty")
    market_error = online_market.attrs.get("market_error")
    online_fundamentals = load_online_fundamentals()
    fundamental_source = online_fundamentals.attrs.get("fundamental_source", "empty")
    fundamental_error = online_fundamentals.attrs.get("fundamental_error")

    df = universe.merge(online_market, on="Kode", how="left")
    if not online_fundamentals.empty:
        df = df.merge(online_fundamentals, on="Kode", how="left")
    else:
        df["Fundamental_Source"] = np.nan

    text_online_columns = {
        "Nama Perusahaan": "Nama_Perusahaan_Online",
        "Sektor": "Sektor_Online",
        "Industry": "Industry_Online",
    }
    for column, online_column in text_online_columns.items():
        if online_column in df.columns:
            missing = df[column].isna() | df[column].astype(str).str.strip().isin(["", "-", "No Sector", "No Industry", "nan"])
            df.loc[missing, column] = df.loc[missing, online_column]

    numeric_online_columns = {
        "PER": "PER_Online",
        "PBV": "PBV_Online",
        "ROE": "ROE_Online",
        "ROA": "ROA_Online",
        "DER": "DER_Online",
        "NPM": "NPM_Online",
        "Mkt Cap": "Market_Cap_Online",
        "Total Rev": "Revenue_Online",
    }
    for column, online_column in numeric_online_columns.items():
        if online_column in df.columns:
            if column not in df.columns:
                df[column] = np.nan
            df[column] = pd.to_numeric(df[column], errors="coerce").fillna(pd.to_numeric(df[online_column], errors="coerce"))

    excel_columns = [
        "Kode",
        "Nama Perusahaan",
        "Sektor",
        "Industry",
        "ListingDate",
        "Shares",
        "ListingBoard",
        "Universe_Source",
        "In_IDX_Official",
        "Universe_Diff_Status",
        "Index",
        "Index_Count",
        "Index_Count_Raw",
        "Index_Count_Sigma",
        "Penutupan",
        "PER",
        "PBV",
        "ROE",
        "ROA",
        "DER",
        "NPM",
        "NIM",
        "CAR",
        "LDR",
        "NPL",
        "BOPO",
        "CIR",
        "LAR",
        "Sebelumnya",
        "%Change",
        "Open",
        "High",
        "Low",
        "Volume",
    ]
    excel_merge = excel_summary[[column for column in excel_columns if column in excel_summary.columns]].copy()
    df = df.merge(excel_merge, on="Kode", how="left", suffixes=("", "_Excel"))

    text_columns = ["Nama Perusahaan", "Sektor", "Industry", "Index"]
    numeric_fallback_columns = [
        "Penutupan",
        "Sebelumnya",
        "%Change",
        "Open",
        "High",
        "Low",
        "Volume",
        "PER",
        "PBV",
        "ROE",
        "ROA",
        "DER",
        "NPM",
        "NIM",
        "CAR",
        "LDR",
        "NPL",
        "BOPO",
        "CIR",
        "LAR",
        "Index_Count",
        "Index_Count_Raw",
        "Index_Count_Sigma",
    ]
    for column in text_columns:
        excel_column = f"{column}_Excel"
        if column not in df.columns:
            df[column] = np.nan
        if excel_column in df.columns:
            missing = df[column].isna() | df[column].astype(str).str.strip().isin(["", "-", "No Sector", "No Industry", "nan"])
            df.loc[missing, column] = df.loc[missing, excel_column]
            df = df.drop(columns=[excel_column])
    for column in numeric_fallback_columns:
        excel_column = f"{column}_Excel"
        if column not in df.columns:
            df[column] = np.nan
        df[column] = pd.to_numeric(df[column], errors="coerce")
        if excel_column in df.columns:
            excel_values = pd.to_numeric(df[excel_column], errors="coerce")
            fill_from_excel = df[column].isna() & excel_values.notna()
            if column == "Penutupan":
                if "Price_Source" not in df.columns:
                    df["Price_Source"] = np.nan
                df.loc[fill_from_excel, "Price_Source"] = "Excel fallback"
            if column == "Volume":
                df["Volume_Original"] = excel_values
                if "Volume_Source" not in df.columns:
                    df["Volume_Source"] = np.nan
                df.loc[fill_from_excel, "Volume_Source"] = "Excel fallback"
            if column in ["PER", "PBV", "ROE", "ROA", "DER", "NPM", "NIM", "CAR", "LDR", "NPL", "BOPO", "CIR", "LAR"]:
                if "Fundamental_Source" not in df.columns:
                    df["Fundamental_Source"] = np.nan
                df.loc[fill_from_excel & df["Fundamental_Source"].isna(), "Fundamental_Source"] = "Excel fallback"
            df[column] = df[column].fillna(excel_values)
            df = df.drop(columns=[excel_column])

    index_missing = df["Index"].isna() | df["Index"].astype(str).str.strip().eq("")
    df.loc[index_missing, "Index"] = df.loc[index_missing, "Universe_Source"]
    df["Index_Count"] = pd.to_numeric(df["Index_Count"], errors="coerce").fillna(0)
    df["Index_Count"] = np.where(df["Index_Count"].le(0), 1, df["Index_Count"])
    if "Volume_Original" not in df.columns:
        df["Volume_Original"] = df["Volume"]
    df["Price_Source"] = df.get("Price_Source", pd.Series(index=df.index)).fillna("Excel fallback")
    df["Volume_Source"] = df.get("Volume_Source", pd.Series(index=df.index)).fillna("Excel fallback")
    df["Fundamental_Source"] = df.get("Fundamental_Source", pd.Series(index=df.index)).fillna("Excel fallback")
    df["Data_Source"] = np.where(
        df["Price_Source"].isin(["yfinance", "pandas-datareader", "cache"]) | df["Fundamental_Source"].eq("TradingView scanner"),
        "Online-first mixed",
        "Excel fallback",
    )
    df["Turnover"] = df["Penutupan"].fillna(0) * df["Volume"].fillna(0)
    df["Intraday_Range_%"] = np.where(
        df["Penutupan"] > 0,
        (df["High"].fillna(df["Penutupan"]) - df["Low"].fillna(df["Penutupan"]))
        / df["Penutupan"]
        * 100,
        np.nan,
    )
    df["Valid_Data"] = (
        (df["Penutupan"] > 0)
        & (df["Volume"] > 0)
        & df[["PER", "PBV", "ROE", "ROA", "DER", "NPM"]].notna().any(axis=1)
    )
    if "Kode" in raw.columns:
        df["Source_Rows"] = raw.groupby("Kode").size().reindex(df["Kode"]).fillna(0).to_numpy()
    else:
        df["Source_Rows"] = 0
    history = load_history_metrics()
    df = df.merge(history, on="Kode", how="left", suffixes=("", "_ExcelHistory"))
    for column in [
        "Return_4W",
        "Return_13W",
        "Return_26W",
        "Return_52W",
        "Return_MTD",
        "Return_YTD",
        "Mkt Cap",
        "Total Rev",
        "Index_Count_Metrik",
    ]:
        excel_column = f"{column}_ExcelHistory"
        if excel_column in df.columns:
            df[column] = pd.to_numeric(df.get(column), errors="coerce").fillna(pd.to_numeric(df[excel_column], errors="coerce"))
            df = df.drop(columns=[excel_column])
    for column in ["Sektor_Metrik", "Subsektor", "Industri", "Subindustri", "Index_Metrik"]:
        excel_column = f"{column}_ExcelHistory"
        if excel_column in df.columns:
            if column not in df.columns:
                df[column] = np.nan
            missing = df[column].isna() | df[column].astype(str).str.strip().isin(["", "-", "nan"])
            df.loc[missing, column] = df.loc[missing, excel_column]
            df = df.drop(columns=[excel_column])
    if "Sektor_Metrik" in df.columns:
        missing_sector = df["Sektor"].isna() | df["Sektor"].astype(str).str.strip().isin(["", "-", "No Sector", "nan"])
        df.loc[missing_sector, "Sektor"] = df.loc[missing_sector, "Sektor_Metrik"]
    if "Subindustri" in df.columns:
        missing_industry = df["Industry"].isna() | df["Industry"].astype(str).str.strip().isin(["", "-", "No Industry", "nan"])
        subindustry_values = df.loc[missing_industry, "Subindustri"]
        industry_values = df.loc[missing_industry, "Industri"] if "Industri" in df.columns else pd.Series(np.nan, index=df.index[missing_industry])
        df.loc[missing_industry, "Industry"] = subindustry_values.where(subindustry_values.notna(), industry_values)
    if "Index_Metrik" in df.columns:
        index_missing = df["Index"].isna() | df["Index"].astype(str).str.strip().isin(["", "-", "nan"])
        df.loc[index_missing, "Index"] = df.loc[index_missing, "Index_Metrik"]
    if "Index_Count_Metrik" in df.columns:
        df["Index_Count"] = pd.concat(
            [
                pd.to_numeric(df.get("Index_Count"), errors="coerce"),
                pd.to_numeric(df["Index_Count_Metrik"], errors="coerce"),
            ],
            axis=1,
        ).max(axis=1).fillna(0)
    if "Mkt Cap" in df.columns:
        df["Market_Cap"] = pd.to_numeric(df["Mkt Cap"], errors="coerce")
    else:
        df["Market_Cap"] = np.nan
    if "Total Rev" in df.columns:
        df["Revenue"] = pd.to_numeric(df["Total Rev"], errors="coerce")
    else:
        df["Revenue"] = np.nan
    df["Sales_Multiple"] = np.where(df["Revenue"].fillna(0) > 0, df["Market_Cap"] / df["Revenue"], np.nan)
    threshold_bank = aggregate_threshold_sheet("Banking")
    threshold_nonbank = aggregate_threshold_sheet("NonBank")
    threshold_values = pd.concat([threshold_bank, threshold_nonbank], ignore_index=True)
    threshold_values = threshold_values.groupby("Kode", as_index=False).median(numeric_only=True) if "Kode" in threshold_values.columns else pd.DataFrame(columns=["Kode"])
    fill_columns = ["NIM", "CAR", "LDR", "NPL", "BOPO", "CIR", "LAR"]
    merge_columns = ["Kode"] + [column for column in fill_columns if column in threshold_values.columns]
    df = df.merge(threshold_values[merge_columns], on="Kode", how="left", suffixes=("", "_Threshold"))
    for column in fill_columns:
        threshold_column = f"{column}_Threshold"
        if threshold_column in df.columns:
            df[column] = df[column].fillna(df[threshold_column])
            df = df.drop(columns=[threshold_column])
    df = apply_threshold_profile(df)
    raw.attrs["data_source"] = f"{summarize_universe_source(universe)}; Market {market_source}; Fundamental {fundamental_source}"
    raw.attrs["online_update"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    raw.attrs["universe_error"] = universe_error
    raw.attrs["market_error"] = market_error
    raw.attrs["fundamental_error"] = fundamental_error
    return df, raw


def calculate_scores(df, weights):
    scored = df.copy()

    positive_per = scored["PER"].between(0.1, scored["PER"].quantile(0.95), inclusive="both")
    positive_pbv = scored["PBV"].between(0.05, scored["PBV"].quantile(0.95), inclusive="both")
    positive_der = scored["DER"].between(0, scored["DER"].quantile(0.97), inclusive="both")

    scored["PER_Score"] = score_percentile(scored["PER"], higher_is_better=False, valid_mask=positive_per)
    scored["PBV_Score"] = score_percentile(scored["PBV"], higher_is_better=False, valid_mask=positive_pbv)
    scored["Valuation_Score"] = (scored["PER_Score"] * 0.58) + (scored["PBV_Score"] * 0.42)

    scored["ROE_Score"] = score_percentile(scored["ROE"], higher_is_better=True, valid_mask=scored["ROE"] > 0)
    scored["ROA_Score"] = score_percentile(scored["ROA"], higher_is_better=True, valid_mask=scored["ROA"] > 0)
    scored["NPM_Score"] = score_percentile(scored["NPM"], higher_is_better=True, valid_mask=scored["NPM"] > 0)
    scored["Quality_Score"] = (
        scored["ROE_Score"] * 0.45 + scored["ROA_Score"] * 0.30 + scored["NPM_Score"] * 0.25
    )

    scored["DER_Score"] = score_percentile(scored["DER"], higher_is_better=False, valid_mask=positive_der)
    scored["Volatility_Score"] = score_percentile(
        scored["Intraday_Range_%"], higher_is_better=False, valid_mask=scored["Intraday_Range_%"] >= 0
    )
    scored["Risk_Score"] = scored["DER_Score"] * 0.72 + scored["Volatility_Score"] * 0.28
    banking_mask = scored.get("Threshold_Mode", pd.Series("", index=scored.index)).eq("Banking")
    scored["CAR_Score"] = score_percentile(scored.get("CAR", pd.Series(index=scored.index)), higher_is_better=True, valid_mask=scored.get("CAR", pd.Series(index=scored.index)).gt(0))
    scored["NPL_Score"] = score_percentile(scored.get("NPL", pd.Series(index=scored.index)), higher_is_better=False, valid_mask=scored.get("NPL", pd.Series(index=scored.index)).ge(0))
    scored["BOPO_Score"] = score_percentile(scored.get("BOPO", pd.Series(index=scored.index)), higher_is_better=False, valid_mask=scored.get("BOPO", pd.Series(index=scored.index)).gt(0))
    scored["LDR_Score"] = score_target_range(scored.get("LDR", pd.Series(index=scored.index)), low=70, high=100, center=85)
    scored["Banking_Risk_Score"] = (
        scored["CAR_Score"] * 0.30
        + scored["NPL_Score"] * 0.30
        + scored["BOPO_Score"] * 0.25
        + scored["LDR_Score"] * 0.15
    )
    scored["Risk_Score"] = np.where(
        banking_mask,
        scored["Banking_Risk_Score"].fillna(scored["Risk_Score"]),
        scored["Risk_Score"],
    )

    scored["Volume_Score"] = score_percentile(scored["Volume"], higher_is_better=True, valid_mask=scored["Volume"] > 0)
    scored["Turnover_Score"] = score_percentile(scored["Turnover"], higher_is_better=True, valid_mask=scored["Turnover"] > 0)
    scored["Liquidity_Score"] = scored["Volume_Score"] * 0.55 + scored["Turnover_Score"] * 0.45

    scored["Trend_Score"] = score_target_range(scored["%Change"], low=-7, high=9, center=2)
    scored["Return_4W_Score"] = score_target_range(scored.get("Return_4W", pd.Series(index=scored.index)), low=-20, high=35, center=8)
    scored["Return_13W_Score"] = score_target_range(scored.get("Return_13W", pd.Series(index=scored.index)), low=-25, high=45, center=12)
    scored["Return_26W_Score"] = score_target_range(scored.get("Return_26W", pd.Series(index=scored.index)), low=-35, high=70, center=18)
    scored["Return_52W_Score"] = score_target_range(scored.get("Return_52W", pd.Series(index=scored.index)), low=-45, high=100, center=25)
    scored["History_Momentum_Score"] = (
        scored["Return_4W_Score"] * 0.25
        + scored["Return_13W_Score"] * 0.30
        + scored["Return_26W_Score"] * 0.25
        + scored["Return_52W_Score"] * 0.20
    )
    scored["Momentum_Score"] = (
        scored["History_Momentum_Score"] * 0.60
        + scored["Trend_Score"] * 0.20
        + score_percentile(scored["%Change"], higher_is_better=True, valid_mask=scored["%Change"].between(-15, 20))
        * 0.20
    )

    scored["Index_Score"] = score_percentile(
        scored["Index_Count"], higher_is_better=True, valid_mask=scored["Index_Count"] > 0
    )

    weighted_score = (
        scored["Valuation_Score"] * weights["valuation"]
        + scored["Quality_Score"] * weights["quality"]
        + scored["Risk_Score"] * weights["risk"]
        + scored["Liquidity_Score"] * weights["liquidity"]
        + scored["Momentum_Score"] * weights["momentum"]
        + scored["Index_Score"] * weights["index_strength"]
    ) / sum(weights.values())

    penalty = np.zeros(len(scored))
    penalty += np.where(scored["PER"].le(0) | scored["PBV"].le(0), 10, 0)
    penalty += np.where(scored["ROE"].le(0) | scored["ROA"].le(0), 9, 0)
    penalty += np.where(scored["NPM"].le(0), 6, 0)
    penalty += np.where(scored["Volume"].lt(1_000_000), 6, 0)
    penalty += np.where(scored["Penutupan"].le(0), 25, 0)
    penalty += np.where(scored["%Change"].abs().gt(25), 6, 0)
    if "Threshold_Pass_Ratio" in scored.columns:
        penalty += np.where(scored["Threshold_Pass_Ratio"].lt(40), 5, 0)

    scored["Penalty"] = penalty
    scored["Score"] = (weighted_score - scored["Penalty"]).clip(0, 100)

    conditions = [
        scored["Score"] >= 78,
        scored["Score"] >= 68,
        scored["Score"] >= 55,
        scored["Score"] >= 42,
    ]
    labels = ["Strong Buy", "Buy", "Watchlist", "Speculative"]
    scored["Recommendation"] = np.select(conditions, labels, default="Avoid")

    banking_high_risk = banking_mask & (
        (scored.get("NPL", pd.Series(index=scored.index)) > 5)
        | (scored.get("BOPO", pd.Series(index=scored.index)) > 90)
        | (scored.get("CAR", pd.Series(index=scored.index)) < 12)
        | (scored["ROE"] < 0)
        | (scored["NPM"] < 0)
        | (scored["Volume"] < 1_000_000)
    )
    banking_medium_risk = banking_mask & (
        (scored.get("NPL", pd.Series(index=scored.index)) > 3.5)
        | (scored.get("BOPO", pd.Series(index=scored.index)) > 80)
        | (scored.get("LDR", pd.Series(index=scored.index)) > 100)
        | (scored["Volume"] < 10_000_000)
        | (scored["%Change"].abs() > 10)
    )
    nonbank_mask = ~banking_mask
    high_risk = banking_high_risk | (
        nonbank_mask
        & (
            (scored["DER"] > 2.5)
            | (scored["ROE"] < 0)
            | (scored["NPM"] < 0)
            | (scored["Volume"] < 1_000_000)
            | (scored["Intraday_Range_%"] > 12)
        )
    )
    medium_risk = banking_medium_risk | (
        nonbank_mask
        & (
            (scored["DER"] > 1.2)
            | (scored["Volume"] < 10_000_000)
            | (scored["Intraday_Range_%"] > 7)
            | (scored["%Change"].abs() > 10)
        )
    )
    scored["Risk_Level"] = np.select([high_risk, medium_risk], ["High", "Medium"], default="Low")
    return scored.sort_values("Score", ascending=False)


df, raw_df = load_data()
data_file_status = get_file_status(DATA_FILE)
data_update_label = get_data_update_label(raw_df, data_file_status)

st.title("Dashboard Rekomendasi Saham IDX")
st.caption(
    f"Data online-first, update {data_update_label}. Universe kode diprioritaskan dari BEI/IDX; yfinance mengisi harga/histori; TradingView scanner mengisi fundamental online; {DATA_FILE} menjadi fallback dan acuan metodologi. Sistem scoring multi-factor untuk screening awal, bukan nasihat investasi."
)
if raw_df.attrs.get("universe_error"):
    st.warning(f"Daftar kode online memakai fallback. Detail: {raw_df.attrs.get('universe_error')}")
if raw_df.attrs.get("market_error"):
    st.warning(f"Sebagian data pasar online memakai fallback/cache. Detail: {raw_df.attrs.get('market_error')}")
if raw_df.attrs.get("fundamental_error"):
    st.warning(f"Sebagian data fundamental online memakai fallback Excel. Detail: {raw_df.attrs.get('fundamental_error')}")

with st.expander("Panduan dashboard, istilah, dan cara membaca hasil", expanded=False):
    st.markdown(
        """
        Arahkan kursor ke ikon bantuan pada menu/filter untuk melihat penjelasan singkat langsung di tempatnya.

        **Menu utama**
        - **Ringkasan**: snapshot eksekutif berisi kondisi universe, sumber data, distribusi rekomendasi, top kandidat, dan matriks faktor.
        - **Rekomendasi**: ranking saham berdasarkan score multi-factor, filter sidebar, label rekomendasi, dan sort aktif.
        - **Explorer**: grafik sebar untuk melihat hubungan valuasi, profitabilitas, risiko, likuiditas, sektor, dan outlier.
        - **Histori Harga**: grafik return dari yfinance online dengan format `KODE.JK`, serta mode Excel Metrik sebagai pembanding/cadangan.
        - **Sektor**: ringkasan score, jumlah saham, Strong Buy, ROE, dan turnover per sektor/industri.
        - **Data Quality**: audit data, cache histori, kelengkapan rasio, dan catatan kualitas data.
        - **Metodologi**: bobot aktif, threshold NonBank/Banking, rumus scoring, penalti, dan distribusi faktor.

        **Istilah penting**
        - **PER**: Price to Earnings Ratio. Lebih rendah umumnya lebih murah, selama laba positif.
        - **PBV**: Price to Book Value. Lebih rendah berarti harga lebih dekat ke nilai buku.
        - **ROE / ROA**: kemampuan menghasilkan laba dari ekuitas/aset. Lebih tinggi lebih baik.
        - **DER**: Debt to Equity Ratio. Dipakai untuk non-bank; untuk bank tidak menjadi filter utama karena struktur bisnis bank memang leverage tinggi.
        - **NPM**: Net Profit Margin. Margin laba bersih terhadap pendapatan.
        - **NIM, CAR, LDR, NPL, BOPO**: metrik khusus bank. CAR tinggi, NPL rendah, BOPO rendah, dan LDR sehat lebih baik.
        - **CIR / LAR**: metrik efisiensi dan risiko aset bank bila tersedia pada sheet Banking.
        - **Volume**: jumlah saham yang diperdagangkan.
        - **Turnover**: estimasi nilai transaksi dari harga penutupan dikali volume.
        - **Momentum**: sinyal tren dari `%Change`, return 4W, 13W, 26W, 52W, dan YTD bila tersedia.
        - **Index Count / Kekuatan indeks**: jumlah kemunculan saham pada indeks/sumber data. Jika tersedia, dashboard memakai nilai kolom Sigma i >= 7 dari Excel sebagai coverage indeks utama; ini sinyal visibilitas, bukan jaminan kualitas.
        - **Threshold**: persentase rasio yang lolos batas dari sheet `NonBank` atau `Banking`.
        - **Threshold Mode**: sumber batas rasio yang dipakai, yaitu `NonBank` atau `Banking`.
        - **Threshold Pass Count / Applicable**: jumlah rasio yang lolos dibanding jumlah rasio yang bisa dinilai.
        - **Score**: nilai akhir 0-100 dari bobot faktor dikurangi penalti. Semakin tinggi semakin menarik sebagai kandidat screening.
        - **Recommendation**: label dari score akhir: Strong Buy, Buy, Watchlist, Speculative, atau Avoid.
        - **Clean_Data**: penanda bahwa data dan rasio utama lolos filter kebersihan minimum.
        - **Safety_Recommendation**: ringkasan kelayakan data seperti `Bersih - Strong`; di kartu utama ditampilkan sebagai `Data`, bukan jaminan aman investasi.
        - **Safety_Notes**: alasan saham perlu direview, misalnya volume rendah, rasio kosong, threshold rendah, atau risiko tinggi.
        - **Risk Level**: estimasi risiko relatif berdasarkan rasio, volatilitas, penalti, dan likuiditas, bukan jaminan keamanan.
        - **KODE.JK**: format ticker saham Indonesia di Yahoo Finance/yfinance, misalnya `BBCA.JK`.
        - **All / top N**: grafik histori untuk saham teratas dari hasil filter/ranking saat ini.

        **Rumus ringkas**
        - `Score = weighted average(Valuation, Quality, Risk, Liquidity, Momentum, Index) - Penalty`, lalu dibatasi 0-100.
        - `Threshold_Pass_Ratio = Threshold_Pass_Count / Threshold_Applicable * 100`.
        - `Turnover = Penutupan * Volume`.
        - `Return Total = Harga Akhir / Harga Awal - 1`.
        - Jika `Wajib lolos valuasi & profit inti` aktif: `PER <= 15`, `PBV <= 3`, `ROE >= 12%`, dan `NPM >= 7%`.

        **Profil scoring vs preset filter**
        - **Profil scoring** mengubah bobot Score, misalnya lebih condong ke valuasi, kualitas, risiko, likuiditas, atau momentum.
        - **Preset filter** mengubah batas minimum/maksimum yang dipakai untuk menyaring hasil. Preset `Konservatif Aman` tidak mengubah bobot Score.
        - Kombinasi yang rapi: pilih profil `Defensive` bila ingin bobot lebih hati-hati, lalu pilih preset filter `Konservatif Aman` bila ingin hasil yang lolos data dan rasio lebih ketat.

        Gunakan hasil ini sebagai screener awal. Keputusan investasi tetap perlu cek berita, laporan keuangan, aksi korporasi, dan diversifikasi portofolio.

        **Prioritas sumber data**
        1. BEI/IDX resmi untuk universe kode saham dan metadata listing.
        2. Sumber online pelengkap: yfinance untuk harga/histori dan TradingView scanner untuk fundamental massal.
        3. Excel dipakai terakhir sebagai fallback, audit pembanding, dan acuan ide algoritme sampai seluruh kolom penting punya sumber online yang stabil.
        """
    )

with st.sidebar:
    st.header("Filter & Strategi")
    profile = st.selectbox("Profil scoring", list(PROFILE_WEIGHTS), index=0, help=HELP_TEXT["profile"])
    weights = PROFILE_WEIGHTS[profile].copy()
    filter_preset = st.selectbox("Preset filter", ["Normal", "Konservatif Aman"], index=0, help=HELP_TEXT["filter_preset"])
    safe_preset = filter_preset == "Konservatif Aman"
    if safe_preset:
        st.info("Preset filter konservatif aktif: filter dibuat lebih ketat. Bobot Score tetap mengikuti Profil scoring yang dipilih.")

    with st.expander("Sesuaikan bobot", expanded=False):
        weights["valuation"] = st.slider("Valuasi", 0, 50, weights["valuation"], help=HELP_TEXT["valuation"])
        weights["quality"] = st.slider("Kualitas profit", 0, 50, weights["quality"], help=HELP_TEXT["quality"])
        weights["risk"] = st.slider("Risiko relatif", 0, 40, weights["risk"], help=HELP_TEXT["risk"])
        weights["liquidity"] = st.slider("Likuiditas", 0, 40, weights["liquidity"], help=HELP_TEXT["liquidity"])
        weights["momentum"] = st.slider("Momentum", 0, 40, weights["momentum"], help=HELP_TEXT["momentum"])
        weights["index_strength"] = st.slider("Kekuatan indeks", 0, 25, weights["index_strength"], help=HELP_TEXT["index_strength"])
        if sum(weights.values()) == 0:
            st.warning("Minimal satu bobot harus lebih dari 0.")
            weights = BASE_WEIGHTS.copy()

    sectors = ["Semua Sektor"] + sorted(df["Sektor"].dropna().unique().tolist())
    industries = ["Semua Industri"] + sorted(df["Industry"].dropna().unique().tolist())
    sector_filter = st.selectbox("Sektor", sectors, help=HELP_TEXT["sector"])
    industry_filter = st.selectbox("Industri", industries, help=HELP_TEXT["industry"])

    price_min, price_max = int(df["Penutupan"].min()), int(df["Penutupan"].max())
    default_price_range = (max(price_min, 50), min(price_max, 50_000)) if safe_preset else (price_min, min(price_max, 50_000))
    price_range = st.slider("Harga penutupan", price_min, price_max, default_price_range, help=HELP_TEXT["price"])
    min_volume = st.select_slider(
        "Minimum volume",
        options=[0, 1_000_000, 5_000_000, 10_000_000, 50_000_000, 100_000_000, 500_000_000],
        value=10_000_000 if safe_preset else 5_000_000,
        format_func=format_volume,
        help=HELP_TEXT["volume"],
    )

    per_range = st.slider("PER", 0.0, 80.0, (0.1, 25.0) if safe_preset else (0.0, 35.0), step=0.5, help=HELP_TEXT["per"])
    pbv_max = st.slider("PBV maksimum", 0.0, 15.0, 3.5 if safe_preset else 5.0, step=0.1, help=HELP_TEXT["pbv"])
    roe_min = st.slider("ROE minimum (%)", -50.0, 100.0, 8.0 if safe_preset else 5.0, step=0.5, help=HELP_TEXT["roe"])
    npm_min = st.slider("NPM minimum (%)", -50.0, 100.0, 3.0 if safe_preset else 0.0, step=0.5, help=HELP_TEXT["npm"])
    der_max = st.slider("DER maksimum", 0.0, 8.0, 1.5 if safe_preset else 2.5, step=0.1, help=HELP_TEXT["der"])
    apply_der_to_banking = st.checkbox("Terapkan DER juga ke Banking", value=False, help=HELP_TEXT["der_banking"])
    min_score = st.slider("Score minimum", 0, 100, 60 if safe_preset else 45, help=HELP_TEXT["score"])
    st.divider()
    st.subheader("Threshold Sheet")
    threshold_source = st.selectbox("Sumber threshold", ["Auto: Banking untuk bank, NonBank untuk lainnya", "NonBank", "Banking"], help=HELP_TEXT["threshold_source"])
    min_threshold_ratio = st.slider("Minimum lolos threshold (%)", 0, 100, 65 if safe_preset else 50, step=5, help=HELP_TEXT["threshold_ratio"])
    require_core_thresholds = st.checkbox("Wajib lolos valuasi & profit inti", value=safe_preset, help=HELP_TEXT["core_thresholds"])
    clean_data_only = st.checkbox("Data bersih saja", value=safe_preset, help=HELP_TEXT["clean_data"])

    st.divider()
    with st.expander("Workflow Update", expanded=False):
        file_status = data_file_status
        st.caption(f"Sumber aktif: {data_update_label}")
        st.caption(f"Excel fallback: {file_status['Status']} | Modified: {file_status['Last Modified']} | Size: {file_status['Ukuran']}")
        cache_status_sidebar = get_history_cache_status()
        st.caption(f"Cache histori: {len(cache_status_sidebar):,} file di `{HISTORY_CACHE_DIR}`")
        refresh_period = st.selectbox(
            "Periode refresh cache",
            ["5d", "1mo", "3mo", "6mo", "1y", "2y", "5y", "10y", "max"],
            index=4,
            format_func=lambda value: {
                "5d": "1 minggu",
                "1mo": "1 bulan",
                "3mo": "3 bulan",
                "6mo": "6 bulan",
                "1y": "1 tahun",
                "2y": "2 tahun",
                "5y": "5 tahun",
                "10y": "10 tahun",
                "max": "All / sepanjang masa",
            }.get(value, value),
            help=HELP_TEXT["refresh_period"],
        )
        refresh_top_n = st.slider("Jumlah top saham untuk refresh", 5, 50, 10, step=5, help=HELP_TEXT["refresh_top_n"])
        if st.button("Refresh cache histori top saham"):
            with st.spinner("Mengambil data histori online..."):
                fetch_yahoo_history.clear()
                top_codes_for_refresh = df.sort_values("Index_Count", ascending=False)["Kode"].head(refresh_top_n).tolist()
                refreshed_history, refresh_error, refresh_source = fetch_yahoo_history(top_codes_for_refresh, period=refresh_period)
            if refresh_error and refreshed_history.empty:
                st.error(refresh_error)
            else:
                st.success(f"Cache diperbarui dari {refresh_source}: {len(refreshed_history):,} baris histori.")
        if st.button("Clear cache aplikasi"):
            st.cache_data.clear()
            st.success("Cache aplikasi dibersihkan. Dashboard akan dimuat ulang.")
            st.rerun()

active_filter_criteria = make_filter_criteria(
    name="Filter aktif",
    price_range=price_range,
    min_volume=min_volume,
    per_range=per_range,
    pbv_max=pbv_max,
    roe_min=roe_min,
    npm_min=npm_min,
    der_max=der_max,
    apply_der_to_banking=apply_der_to_banking,
    min_score=min_score,
    min_threshold_ratio=min_threshold_ratio,
    require_core_thresholds=require_core_thresholds,
    clean_data_only=clean_data_only,
    sector_filter=sector_filter,
    industry_filter=industry_filter,
)
preset_filter_criteria = [
    active_filter_criteria,
    default_filter_criteria("Preset Normal", conservative=False, price_min=price_min, price_max=price_max),
    default_filter_criteria("Preset Konservatif Aman", conservative=True, price_min=price_min, price_max=price_max),
]

forced_threshold_mode = None if threshold_source.startswith("Auto") else threshold_source
threshold_df = apply_threshold_profile(df, forced_mode=forced_threshold_mode)
scored_df = calculate_scores(threshold_df, weights)
scored_df = add_safety_flags(scored_df)
filtered = scored_df.copy()

if sector_filter != "Semua Sektor":
    filtered = filtered[filtered["Sektor"] == sector_filter]
if industry_filter != "Semua Industri":
    filtered = filtered[filtered["Industry"] == industry_filter]

der_filter = filtered["DER"].fillna(np.inf) <= der_max
if not apply_der_to_banking:
    der_filter = der_filter | filtered["Threshold_Mode"].eq("Banking")

filtered = filtered[
    (filtered["Penutupan"].between(price_range[0], price_range[1]))
    & (filtered["Volume"].fillna(0) >= min_volume)
    & (filtered["PER"].between(per_range[0], per_range[1]))
    & (filtered["PBV"].fillna(np.inf) <= pbv_max)
    & (filtered["ROE"].fillna(-np.inf) >= roe_min)
    & (filtered["NPM"].fillna(-np.inf) >= npm_min)
    & der_filter
    & (filtered["Score"] >= min_score)
    & (filtered["Threshold_Pass_Ratio"] >= min_threshold_ratio)
]

if require_core_thresholds:
    filtered = filtered[
        (filtered["PER"].fillna(np.inf) <= 15)
        & (filtered["PBV"].fillna(np.inf) <= 3)
        & (filtered["ROE"].fillna(-np.inf) >= 12)
        & (filtered["NPM"].fillna(-np.inf) >= 7)
    ]

if clean_data_only:
    filtered = filtered[filtered["Clean_Data"]]

status_cols = st.columns(4)
status_cols[0].metric("Universe", f"{len(df):,}", f"{len(raw_df):,} baris sumber")
status_cols[1].metric("Lolos filter", f"{len(filtered):,}")
status_cols[2].metric("Top score", f"{filtered['Score'].max():.1f}" if len(filtered) else "-")
status_cols[3].metric("Data bersih", f"{filtered['Clean_Data'].sum():,}" if len(filtered) else "0")

tab_summary, tab_reco, tab_explore, tab_history, tab_sector, tab_quality, tab_method = st.tabs(
    ["Ringkasan", "Rekomendasi", "Explorer", "Histori Harga", "Sektor", "Data Quality", "Metodologi"]
)

with tab_summary:
    st.subheader("Ringkasan eksekutif")
    summary_scope = st.radio(
        "Cakupan ringkasan",
        ["Hasil filter aktif", "Semua universe"],
        horizontal=True,
        help="Hasil filter aktif mengikuti seluruh filter sidebar. Semua universe memakai seluruh saham yang sudah di-score.",
    )
    summary_data = filtered.copy() if summary_scope == "Hasil filter aktif" else scored_df.copy()
    if summary_data.empty:
        st.warning("Tidak ada data pada cakupan ini. Longgarkan filter atau pilih Semua universe.")
    else:
        idx_match = int(summary_data.get("In_IDX_Official", pd.Series(False, index=summary_data.index)).fillna(False).sum())
        fallback_only = int((~summary_data.get("In_IDX_Official", pd.Series(False, index=summary_data.index)).fillna(False)).sum())
        online_price = int(summary_data.get("Price_Source", pd.Series("", index=summary_data.index)).astype(str).str.contains("Online|yfinance|cache", case=False, na=False).sum())
        clean_ratio = summary_data["Clean_Data"].mean() * 100 if "Clean_Data" in summary_data.columns and len(summary_data) else 0
        total_market_cap = summary_data.get("Market_Cap", pd.Series(index=summary_data.index)).sum(skipna=True)

        summary_cols = st.columns(6)
        summary_cols[0].metric("Saham dianalisis", f"{summary_data['Kode'].nunique():,}")
        summary_cols[1].metric("Match BEI/IDX", f"{idx_match:,}")
        summary_cols[2].metric("Fallback kode", f"{fallback_only:,}")
        summary_cols[3].metric("Harga online/cache", f"{online_price:,}")
        summary_cols[4].metric("Market cap", format_large_rupiah(total_market_cap) if total_market_cap else "-")
        summary_cols[5].metric("Clean data", f"{clean_ratio:.0f}%")

        chart_cols = st.columns([1, 1, 1])
        with chart_cols[0]:
            reco_counts = summary_data["Recommendation"].value_counts().reindex(["Strong Buy", "Buy", "Watchlist", "Speculative", "Avoid"]).dropna().reset_index()
            reco_counts.columns = ["Recommendation", "Jumlah"]
            fig = px.bar(
                reco_counts,
                x="Recommendation",
                y="Jumlah",
                color="Recommendation",
                title="Distribusi rekomendasi",
                color_discrete_map={
                    "Strong Buy": "#15803d",
                    "Buy": "#65a30d",
                    "Watchlist": "#ca8a04",
                    "Speculative": "#ea580c",
                    "Avoid": "#dc2626",
                },
            )
            fig.update_layout(height=330, showlegend=False, margin=dict(l=20, r=20, t=60, b=40))
            show_chart(fig)
        with chart_cols[1]:
            risk_counts = summary_data["Risk_Level"].value_counts().reindex(["Low", "Medium", "High"]).dropna().reset_index()
            risk_counts.columns = ["Risk_Level", "Jumlah"]
            fig = px.pie(
                risk_counts,
                names="Risk_Level",
                values="Jumlah",
                hole=0.45,
                title="Komposisi risiko",
                color="Risk_Level",
                color_discrete_map={"Low": "#16a34a", "Medium": "#ca8a04", "High": "#dc2626"},
            )
            fig.update_layout(height=330, margin=dict(l=20, r=20, t=60, b=40))
            show_chart(fig)
        with chart_cols[2]:
            source_mix = build_source_mix(summary_data)
            source_view = source_mix[source_mix["Area"].isin(["Price_Source", "Fundamental_Source", "Universe_Diff_Status"])].copy()
            if source_view.empty:
                st.info("Ringkasan sumber data belum tersedia.")
            else:
                fig = px.bar(
                    source_view,
                    x="Jumlah",
                    y="Nilai",
                    color="Area",
                    orientation="h",
                    title="Sumber harga & status kode",
                )
                fig.update_layout(height=330, yaxis_title="", margin=dict(l=20, r=20, t=60, b=40))
                show_chart(fig)

        overview_cols = st.columns([1.2, 1])
        with overview_cols[0]:
            st.write("Top kandidat profesional")
            top_summary_columns = [
                "Kode",
                "Nama Perusahaan",
                "Recommendation",
                "Risk_Level",
                "Clean_Data",
                "Score",
                "Threshold_Pass_Ratio",
                "Penutupan",
                "PER",
                "PBV",
                "ROE",
                "NPM",
                "Return_52W",
                "Volume",
                "Market_Cap",
                "Revenue",
                "Sales_Multiple",
                "Index_Count",
                "Subsektor",
                "Subindustri",
                "Price_Source",
                "Fundamental_Source",
                "Universe_Diff_Status",
            ]
            top_summary = summary_data.sort_values(["Score", "Threshold_Pass_Ratio", "Liquidity_Score"], ascending=False).head(15)
            show_table(
                top_summary[[column for column in top_summary_columns if column in top_summary.columns]],                hide_index=True,
                column_config={
                    "Score": st.column_config.ProgressColumn("Score", min_value=0, max_value=100, format="%.1f", help=HELP_TEXT["score"]),
                    "Threshold_Pass_Ratio": st.column_config.ProgressColumn("Threshold", min_value=0, max_value=100, format="%.0f%%", help=HELP_TEXT["threshold_ratio"]),
                    "Penutupan": st.column_config.NumberColumn("Harga", format="Rp %.0f", help=HELP_TEXT["price"]),
                    "Volume": st.column_config.NumberColumn("Volume", format="%.0f", help=HELP_TEXT["volume"]),
                    "Market_Cap": st.column_config.NumberColumn("Market Cap", format="Rp %.0f", help="Kapitalisasi pasar dari sheet Metrik bila tersedia."),
                    "Revenue": st.column_config.NumberColumn("Revenue", format="Rp %.0f", help="Total revenue dari sheet Metrik bila tersedia."),
                    "Sales_Multiple": st.column_config.NumberColumn("MCap/Revenue", format="%.2f", help="Market cap dibagi total revenue. Dipakai sebagai konteks tambahan, bukan rumus utama score."),
                    "Return_52W": st.column_config.NumberColumn("52W", format="%.1f%%", help=HELP_TEXT["return"]),
                    "Clean_Data": st.column_config.CheckboxColumn("Clean Data", help=HELP_TEXT["clean_data"]),
                },
            )
        with overview_cols[1]:
            st.write("Matriks faktor top score")
            factor_columns = ["Valuation_Score", "Quality_Score", "Risk_Score", "Liquidity_Score", "Momentum_Score", "Index_Score"]
            factor_matrix = summary_data.sort_values("Score", ascending=False).head(12).set_index("Kode")[
                [column for column in factor_columns if column in summary_data.columns]
            ]
            if factor_matrix.empty:
                st.info("Matriks faktor belum tersedia.")
            else:
                fig = px.imshow(
                    factor_matrix,
                    aspect="auto",
                    text_auto=".0f",
                    color_continuous_scale="RdYlGn",
                    zmin=0,
                    zmax=100,
                    title="Skor komponen 0-100",
                )
                fig.update_layout(height=460, xaxis_title="", yaxis_title="", margin=dict(l=20, r=20, t=60, b=40))
                show_chart(fig)

with tab_reco:
    if filtered.empty:
        st.warning("Tidak ada saham yang sesuai filter. Longgarkan kriteria di sidebar.")
    else:
        reco_controls = st.columns([1, 1, 1, 1])
        with reco_controls[0]:
            reco_limit = st.slider("Jumlah tampil", 5, 100, 25, step=5, help=HELP_TEXT["reco_limit"])
        with reco_controls[1]:
            reco_sort = st.selectbox(
                "Urutkan berdasarkan",
                ["Score", "Threshold_Pass_Ratio", "Valuation_Score", "Quality_Score", "Risk_Score", "Liquidity_Score", "Momentum_Score", "Return_52W", "Volume", "Turnover"],
                help=HELP_TEXT["reco_sort"],
            )
        with reco_controls[2]:
            reco_ascending = st.toggle("Urut naik", value=False, help=HELP_TEXT["reco_ascending"])
        with reco_controls[3]:
            reco_labels = st.multiselect(
                "Label rekomendasi",
                ["Strong Buy", "Buy", "Watchlist", "Speculative", "Avoid"],
                default=["Strong Buy", "Buy", "Watchlist"],
                help=HELP_TEXT["recommendation"],
            )

        reco_view = filtered[filtered["Recommendation"].isin(reco_labels)].copy()
        if reco_view.empty:
            st.info("Tidak ada saham pada label rekomendasi yang dipilih.")
            reco_view = filtered.copy()
        reco_view = reco_view.sort_values(reco_sort, ascending=reco_ascending, na_position="last").head(reco_limit)

        left, right = st.columns([1.25, 1])
        with left:
            chart_data = prepare_chart_frame(reco_view.sort_values(reco_sort), reco_sort)
            if chart_data.empty:
                st.warning(f"Tidak ada data valid untuk grafik {reco_sort}. Pilih metrik sort lain.")
            else:
                fig = px.bar(
                    chart_data,
                    x=reco_sort,
                    y="Chart_Label",
                    text="Kode",
                    orientation="h",
                    color="Recommendation",
                    hover_name="Kode",
                    hover_data={
                        "Chart_Label": False,
                        "Nama Perusahaan": True,
                        "Sektor": True,
                        "PER": ":.2f",
                        "PBV": ":.2f",
                        "ROE": ":.1f",
                        "DER": ":.2f",
                        "Threshold_Pass_Ratio": ":.0f",
                        "Return_52W": ":.1f",
                        "Volume": ":,.0f",
                    },
                    title=f"Top rekomendasi berdasarkan {reco_sort}",
                    color_discrete_map={
                        "Strong Buy": "#15803d",
                        "Buy": "#65a30d",
                        "Watchlist": "#ca8a04",
                        "Speculative": "#ea580c",
                        "Avoid": "#dc2626",
                    },
                )
                fig.update_traces(textposition="outside", cliponaxis=False)
                fig.update_layout(height=520, xaxis_title=reco_sort, yaxis_title="", margin=dict(l=20, r=80, t=70, b=40))
                show_chart(fig)

        with right:
            component_cols = [
                "Valuation_Score",
                "Quality_Score",
                "Risk_Score",
                "Liquidity_Score",
                "Momentum_Score",
                "Index_Score",
            ]
            radar_base = reco_view.head(5)
            fig = go.Figure()
            for _, row in radar_base.iterrows():
                values = [row[col] for col in component_cols]
                fig.add_trace(
                    go.Scatterpolar(
                        r=values + [values[0]],
                        theta=[
                            "Valuasi",
                            "Kualitas",
                            "Risiko",
                            "Likuiditas",
                            "Momentum",
                            "Indeks",
                            "Valuasi",
                        ],
                        fill="toself",
                        name=row["Kode"],
                    )
                )
            fig.update_layout(
                title="Profil faktor top 5",
                polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
                height=520,
                margin=dict(l=30, r=30, t=60, b=30),
            )
            show_chart(fig)

        display_columns = [
            "Kode",
            "Nama Perusahaan",
            "Recommendation",
            "Safety_Recommendation",
            "Risk_Level",
            "Clean_Data",
            "Safety_Notes",
            "Score",
            "Valuation_Score",
            "Quality_Score",
            "Risk_Score",
            "Liquidity_Score",
            "Momentum_Score",
            "History_Momentum_Score",
            "Threshold_Mode",
            "Threshold_Pass_Count",
            "Threshold_Applicable",
            "Threshold_Pass_Ratio",
            "Penutupan",
            "PER",
            "PBV",
            "ROE",
            "ROA",
            "DER",
            "NPM",
            "%Change",
            "Return_4W",
            "Return_13W",
            "Return_26W",
            "Return_52W",
            "Return_YTD",
            "Volume",
            "Market_Cap",
            "Revenue",
            "Sales_Multiple",
            "Index_Count",
            "Index_Count_Sigma",
            "Index_Count_Metrik",
            "Price_Source",
            "Volume_Source",
            "Fundamental_Source",
            "Data_Source",
            "Universe_Source",
            "Universe_Diff_Status",
            "ListingBoard",
            "ListingDate",
            "Subsektor",
            "Industri",
            "Subindustri",
            "Volume_Original",
            "Volume_Online_Latest",
            "Online_Last_Date",
            "Sektor",
            "Industry",
            "Index",
        ]
        default_table_columns = [
            "Kode",
            "Nama Perusahaan",
            "Recommendation",
            "Safety_Recommendation",
            "Risk_Level",
            "Clean_Data",
            "Score",
            "Threshold_Pass_Ratio",
            "Penutupan",
            "PER",
            "PBV",
            "ROE",
            "DER",
            "Return_52W",
            "Volume",
            "Market_Cap",
            "Index_Count",
            "Price_Source",
            "Fundamental_Source",
            "Volume_Source",
            "Universe_Diff_Status",
            "Safety_Notes",
            "Sektor",
            "Industry",
        ]
        selected_table_columns = st.multiselect(
            "Kolom tabel",
            display_columns,
            default=default_table_columns,
            help=HELP_TEXT["table_columns"],
        )
        if not selected_table_columns:
            selected_table_columns = default_table_columns
        table = reco_view[selected_table_columns].copy()
        show_table(
            table,            hide_index=True,
            column_config={
                "Score": st.column_config.ProgressColumn("Score", min_value=0, max_value=100, format="%.1f", help=HELP_TEXT["score"]),
                "Valuation_Score": st.column_config.NumberColumn("Valuasi", format="%.1f", help=HELP_TEXT["valuation"]),
                "Quality_Score": st.column_config.NumberColumn("Kualitas", format="%.1f", help=HELP_TEXT["quality"]),
                "Risk_Score": st.column_config.NumberColumn("Risiko", format="%.1f", help=HELP_TEXT["risk"]),
                "Liquidity_Score": st.column_config.NumberColumn("Likuiditas", format="%.1f", help=HELP_TEXT["liquidity"]),
                "Momentum_Score": st.column_config.NumberColumn("Momentum", format="%.1f", help=HELP_TEXT["momentum"]),
                "History_Momentum_Score": st.column_config.NumberColumn("Histori", format="%.1f", help=HELP_TEXT["momentum"]),
                "Threshold_Pass_Ratio": st.column_config.ProgressColumn("Threshold", min_value=0, max_value=100, format="%.0f%%", help=HELP_TEXT["threshold_ratio"]),
                "Penutupan": st.column_config.NumberColumn("Harga", format="Rp %.0f", help=HELP_TEXT["price"]),
                "Volume": st.column_config.NumberColumn("Volume", format="%.0f", help=HELP_TEXT["volume"]),
                "Market_Cap": st.column_config.NumberColumn("Market Cap", format="Rp %.0f", help="Kapitalisasi pasar dari sheet Metrik bila tersedia."),
                "Revenue": st.column_config.NumberColumn("Revenue", format="Rp %.0f", help="Total revenue dari sheet Metrik bila tersedia."),
                "Sales_Multiple": st.column_config.NumberColumn("MCap/Revenue", format="%.2f", help="Market cap dibagi revenue. Konteks tambahan, bukan rumus utama score."),
                "Index_Count": st.column_config.NumberColumn("Index Count", format="%.0f", help=HELP_TEXT["index_strength"]),
                "Index_Count_Sigma": st.column_config.NumberColumn("Sigma i", format="%.0f", help="Nilai coverage indeks dari kolom Sigma i >= 7 di Excel fallback bila tersedia."),
                "Index_Count_Metrik": st.column_config.NumberColumn("Index Count Metrik", format="%.0f", help="Jumlah indeks dari daftar gabungan pada sheet Metrik."),
                "Price_Source": st.column_config.TextColumn("Sumber Harga", help="Menunjukkan apakah harga berasal dari yfinance/cache atau Excel fallback."),
                "Fundamental_Source": st.column_config.TextColumn("Sumber Fundamental", help=HELP_TEXT["fundamental_source"]),
                "Volume_Original": st.column_config.NumberColumn("Volume Excel", format="%.0f", help="Volume dari Excel fallback bila tersedia."),
                "Volume_Online_Latest": st.column_config.NumberColumn("Volume Online", format="%.0f", help="Volume terakhir dari yfinance/cache."),
                "Volume_Source": st.column_config.TextColumn("Sumber Volume", help="Menunjukkan apakah volume berasal dari yfinance/cache atau Excel fallback."),
                "Data_Source": st.column_config.TextColumn("Sumber Data", help="Ringkasan sumber data pasar utama untuk baris ini."),
                "Universe_Source": st.column_config.TextColumn("Sumber Kode", help="Sumber universe kode saham: BEI/IDX official atau fallback."),
                "Universe_Diff_Status": st.column_config.TextColumn("Status Kode", help="Menunjukkan apakah kode match dengan daftar resmi BEI/IDX atau hanya ada di fallback."),
                "ListingBoard": st.column_config.TextColumn("Papan", help="Papan pencatatan dari daftar resmi BEI/IDX bila tersedia."),
                "ListingDate": st.column_config.DateColumn("Listing", help="Tanggal pencatatan dari daftar resmi BEI/IDX bila tersedia."),
                "Online_Last_Date": st.column_config.DateColumn("Tanggal Online", help="Tanggal data online terakhir dari yfinance/cache."),
                "Return_4W": st.column_config.NumberColumn("4W", format="%.1f%%", help=HELP_TEXT["return"]),
                "Return_13W": st.column_config.NumberColumn("13W", format="%.1f%%", help=HELP_TEXT["return"]),
                "Return_26W": st.column_config.NumberColumn("26W", format="%.1f%%", help=HELP_TEXT["return"]),
                "Return_52W": st.column_config.NumberColumn("52W", format="%.1f%%", help=HELP_TEXT["return"]),
                "Return_YTD": st.column_config.NumberColumn("YTD", format="%.1f%%", help=HELP_TEXT["return"]),
                "Recommendation": st.column_config.TextColumn("Recommendation", help=HELP_TEXT["recommendation"]),
                "Safety_Recommendation": st.column_config.TextColumn("Data Check", help="Ringkasan Clean_Data dan Score. Label Bersih berarti lolos filter data minimum, bukan jaminan aman investasi."),
                "Risk_Level": st.column_config.TextColumn("Risk Level", help=HELP_TEXT["risk_level"]),
                "Clean_Data": st.column_config.CheckboxColumn("Clean Data", help=HELP_TEXT["clean_data"]),
                "Safety_Notes": st.column_config.TextColumn("Catatan Data", help="Alasan saham perlu direview jika belum lolos Clean_Data."),
            },
        )

        csv = reco_view[display_columns].to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download hasil rekomendasi CSV",
            data=csv,
            file_name="rekomendasi_saham_multi_factor.csv",
            mime="text/csv",
        )

with tab_explore:
    explorer_data = filtered if not filtered.empty else scored_df
    explore_controls = st.columns([1, 1, 1, 1])
    with explore_controls[0]:
        explore_x = st.selectbox("Sumbu X", ANALYSIS_COLUMNS, index=ANALYSIS_COLUMNS.index("PER"), help=HELP_TEXT["explorer_axis"])
    with explore_controls[1]:
        explore_y = st.selectbox("Sumbu Y", ANALYSIS_COLUMNS, index=ANALYSIS_COLUMNS.index("ROE"), help=HELP_TEXT["explorer_axis"])
    with explore_controls[2]:
        explore_color = st.selectbox("Warna", ["Score", "Quality_Score", "Risk_Score", "Threshold_Pass_Ratio", "Recommendation", "Risk_Level", "Sektor"], help=HELP_TEXT["explore_color"])
    with explore_controls[3]:
        explore_size = st.selectbox("Ukuran bubble", ["Volume", "Turnover", "Score", "Liquidity_Score", "Index_Count"], help=HELP_TEXT["explore_size"])

    explore_max = max(1, min(500, len(explorer_data)))
    explore_min = min(50, explore_max)
    explore_default = min(250, explore_max)
    explore_step = 25 if explore_max >= 50 else 1
    explore_limit = st.slider("Jumlah titik Explorer", explore_min, explore_max, explore_default, step=explore_step, help=HELP_TEXT["explore_limit"])
    explore_plot = explorer_data.sort_values("Score", ascending=False).head(explore_limit)

    left, right = st.columns(2)
    with left:
        fig = px.scatter(
            explore_plot,
            x=explore_x,
            y=explore_y,
            size=explore_size,
            color=explore_color,
            hover_name="Kode",
            hover_data=["Nama Perusahaan", "Sektor", "PBV", "DER", "NPM", "Recommendation"],
            title=f"{explore_x} vs {explore_y}",
            color_continuous_scale="RdYlGn",
        )
        fig.update_layout(height=460)
        show_chart(fig)

    with right:
        pair_x = st.selectbox("Pembanding X", ["PBV", "PER", "DER", "Threshold_Pass_Ratio", "Return_52W"], index=0)
        pair_y = st.selectbox("Pembanding Y", ["DER", "ROE", "NPM", "Score", "Risk_Score"], index=0)
        fig = px.scatter(
            explore_plot,
            x=pair_x,
            y=pair_y,
            color="Quality_Score",
            size="Turnover",
            hover_name="Kode",
            hover_data=["Nama Perusahaan", "Sektor", "Score", "Risk_Level"],
            title=f"{pair_x} vs {pair_y}, warna = kualitas profit",
            color_continuous_scale="Viridis",
        )
        fig.update_layout(height=460)
        show_chart(fig)

    hist_cols = st.columns(3)
    histogram_columns = st.multiselect(
        "Histogram",
        ANALYSIS_COLUMNS,
        default=["Score", "ROE", "PER"],
        help=HELP_TEXT["histogram"],
    )[:3]
    histogram_colors = ["#2563eb", "#16a34a", "#9333ea"]
    for index, (column, container) in enumerate(zip(histogram_columns, hist_cols)):
        with container:
            fig = px.histogram(
                explore_plot,
                x=column,
                nbins=35,
                title=f"Distribusi {column}",
                color_discrete_sequence=[histogram_colors[index]],
            )
            fig.update_layout(height=340)
            show_chart(fig)

with tab_history:
    history_source = filtered if not filtered.empty else scored_df
    history_mode = st.radio(
        "Sumber grafik histori",
        ["Excel Metrik 4W-52W", "Online yfinance KODE.JK"],
        horizontal=True,
        help=HELP_TEXT["history_source"],
    )
    chart_scope = st.radio(
        "Cakupan grafik",
        ["Saham pilihan", "All/top N hasil filter"],
        horizontal=True,
        help=HELP_TEXT["history_scope"],
    )
    top_n_history = st.slider("Jumlah saham untuk grafik all/top N", 5, 100, 25, step=5, help=HELP_TEXT["history_top_n"])
    history_chart_type = st.segmented_control(
        "Tipe grafik histori",
        ["Line", "Area"],
        default="Line",
        help=HELP_TEXT["history_chart_type"],
    )
    show_history_table = st.toggle("Tampilkan tabel ringkasan histori", value=True, help=HELP_TEXT["history_table"])

    if chart_scope == "All/top N hasil filter":
        selected_codes = history_source.head(top_n_history)["Kode"].tolist()
        st.caption(f"Mode all memakai top {len(selected_codes)} saham dari hasil filter/ranking saat ini.")
    else:
        default_codes = history_source.head(5)["Kode"].tolist()
        selected_codes = st.multiselect(
            "Pilih saham untuk grafik histori",
            options=scored_df["Kode"].tolist(),
            default=default_codes,
            help=HELP_TEXT["history_codes"],
        )

    if history_mode == "Online yfinance KODE.JK":
        period = st.selectbox(
            "Rentang data online",
            ["5d", "1mo", "3mo", "6mo", "1y", "2y", "5y", "10y", "max"],
            index=4,
            format_func=lambda value: {
                "5d": "1 minggu",
                "1mo": "1 bulan",
                "3mo": "3 bulan",
                "6mo": "6 bulan",
                "1y": "1 tahun",
                "2y": "2 tahun",
                "5y": "5 tahun",
                "10y": "10 tahun",
                "max": "All / sepanjang masa",
            }.get(value, value),
            help=HELP_TEXT["history_period"],
        )
        online_history, online_error, online_source = fetch_yahoo_history(selected_codes, period=period)
        if online_error:
            st.warning(online_error)
            st.caption("Dashboard memakai cache atau Excel fallback untuk bagian data yang tidak tersedia online.")
        elif online_history.empty:
            st.warning("Data online kosong.")
        else:
            st.caption(f"Sumber histori aktif: {online_source}")
            last_dates = online_history.groupby("Kode")["Date"].max().reset_index()
            last_dates["Last Update Online"] = last_dates["Date"].dt.strftime("%Y-%m-%d")
            if show_history_table:
                show_table(last_dates[["Kode", "Last Update Online"]], hide_index=True)

            chart_func = px.area if history_chart_type == "Area" else px.line
            fig = chart_func(online_history, x="Date", y="Close", color="Kode", title="Harga penutupan historis online", labels={"Close": "Harga penutupan", "Date": "Tanggal"})
            fig.update_layout(height=480)
            show_chart(fig)

            fig = chart_func(online_history, x="Date", y="Normalized", color="Kode", title="Perbandingan performa, indeks awal = 100", labels={"Normalized": "Indeks performa", "Date": "Tanggal"})
            fig.add_hline(y=100, line_dash="dash", line_color="#64748b")
            fig.update_layout(height=440)
            show_chart(fig)

            summary = (
                online_history.groupby("Kode")
                .agg(
                    Start=("Date", "min"),
                    End=("Date", "max"),
                    Start_Close=("Close", "first"),
                    Last_Close=("Close", "last"),
                    Days=("Date", "count"),
                )
                .reset_index()
            )
            summary["Return_Total_%"] = (summary["Last_Close"] / summary["Start_Close"] - 1) * 100
            if show_history_table:
                show_table(
                    summary.sort_values("Return_Total_%", ascending=False),                    hide_index=True,
                    column_config={
                        "Start": st.column_config.DateColumn("Awal"),
                        "End": st.column_config.DateColumn("Akhir"),
                        "Start_Close": st.column_config.NumberColumn("Harga Awal", format="%.0f"),
                        "Last_Close": st.column_config.NumberColumn("Harga Terakhir", format="%.0f"),
                        "Return_Total_%": st.column_config.NumberColumn("Return Total", format="%.1f%%", help=HELP_TEXT["return"]),
                    },
                )

        st.caption("Sumber online memakai ticker IDX format KODE.JK, misalnya BBCA.JK. Jika data live gagal, dashboard mencoba fallback dan cache lokal.")
    else:
        history_columns = ["Return_4W", "Return_13W", "Return_26W", "Return_52W"]
        available_history = [column for column in history_columns if column in history_source.columns]

        if not available_history:
            st.warning("Kolom histori belum tersedia di sheet Metrik.")
        else:
            selected_history = scored_df[scored_df["Kode"].isin(selected_codes)].copy()

            if selected_history.empty:
                st.info("Pilih minimal satu saham untuk menampilkan grafik histori.")
            else:
                long_history = selected_history.melt(
                    id_vars=["Kode", "Nama Perusahaan", "Score", "Threshold_Pass_Ratio", "Return_YTD"],
                    value_vars=available_history,
                    var_name="Periode",
                    value_name="Return",
                )
                period_order = {
                    "Return_4W": "4 minggu",
                    "Return_13W": "13 minggu",
                    "Return_26W": "26 minggu",
                    "Return_52W": "52 minggu",
                }
                long_history["Periode"] = long_history["Periode"].map(period_order)

                chart_func = px.area if history_chart_type == "Area" else px.line
                fig = chart_func(
                    long_history,
                    x="Periode",
                    y="Return",
                    color="Kode",
                    hover_data=["Nama Perusahaan", "Score", "Threshold_Pass_Ratio"],
                    title="Histori return sampai 1 tahun",
                    category_orders={"Periode": ["4 minggu", "13 minggu", "26 minggu", "52 minggu"]},
                )
                fig.add_hline(y=0, line_dash="dash", line_color="#64748b")
                fig.update_layout(height=480, yaxis_title="Return (%)", xaxis_title="")
                show_chart(fig)

                compare = selected_history[
                    [
                        "Kode",
                        "Nama Perusahaan",
                        "Score",
                        "Threshold_Pass_Ratio",
                        "Return_4W",
                        "Return_13W",
                        "Return_26W",
                        "Return_52W",
                        "Return_YTD",
                    ]
                ].sort_values("Return_52W", ascending=False)
                if show_history_table:
                    show_table(
                        compare,                        hide_index=True,
                        column_config={
                            "Score": st.column_config.NumberColumn("Score", format="%.1f", help=HELP_TEXT["score"]),
                            "Threshold_Pass_Ratio": st.column_config.NumberColumn("Threshold", format="%.0f%%", help=HELP_TEXT["threshold_ratio"]),
                            "Return_4W": st.column_config.NumberColumn("4W", format="%.1f%%", help=HELP_TEXT["return"]),
                            "Return_13W": st.column_config.NumberColumn("13W", format="%.1f%%", help=HELP_TEXT["return"]),
                            "Return_26W": st.column_config.NumberColumn("26W", format="%.1f%%", help=HELP_TEXT["return"]),
                            "Return_52W": st.column_config.NumberColumn("52W", format="%.1f%%", help=HELP_TEXT["return"]),
                            "Return_YTD": st.column_config.NumberColumn("YTD", format="%.1f%%", help=HELP_TEXT["return"]),
                        },
                    )

            top_history = history_source.dropna(subset=["Return_52W"]).head(25)
            if not top_history.empty:
                fig = px.scatter(
                    top_history,
                    x="Return_52W",
                    y="Score",
                    size="Volume",
                    color="Threshold_Pass_Ratio",
                    hover_name="Kode",
                    hover_data=["Nama Perusahaan", "Sektor", "Return_26W", "Return_YTD"],
                    title="Score vs return 52 minggu pada saham terfilter",
                    color_continuous_scale="RdYlGn",
                )
                fig.add_vline(x=0, line_dash="dash", line_color="#64748b")
                fig.update_layout(height=440, xaxis_title="Return 52 minggu (%)", yaxis_title="Score")
                show_chart(fig)

with tab_sector:
    sector_controls = st.columns([1, 1, 1, 1])
    with sector_controls[0]:
        sector_group_options = [column for column in ["Sektor", "Subsektor", "Industri", "Subindustri", "Industry"] if column in scored_df.columns]
        sector_group = st.selectbox("Kelompok", sector_group_options, help=HELP_TEXT["sector_group"])
    with sector_controls[1]:
        sector_min_count = st.slider("Minimum saham per kelompok", 1, 25, 3, help=HELP_TEXT["sector_min"])
    with sector_controls[2]:
        sector_sort = st.selectbox("Urutkan sektor", ["Median_Score", "Strong_Buy", "Total_Market_Cap", "Total_Revenue", "Total_Turnover", "Avg_ROE", "Saham"], help=HELP_TEXT["sector_sort"])
    with sector_controls[3]:
        sector_chart = st.selectbox("Visual utama", ["Bar", "Treemap", "Scatter"], help=HELP_TEXT["sector_chart"])

    sector_summary = (
        scored_df.groupby(sector_group, dropna=False)
        .agg(
            Saham=("Kode", "count"),
            Median_Score=("Score", "median"),
            Avg_ROE=("ROE", "mean"),
            Median_PER=("PER", "median"),
            Median_PBV=("PBV", "median"),
            Total_Turnover=("Turnover", "sum"),
            Total_Market_Cap=("Market_Cap", "sum"),
            Total_Revenue=("Revenue", "sum"),
            Median_Sales_Multiple=("Sales_Multiple", "median"),
            Strong_Buy=("Recommendation", lambda x: (x == "Strong Buy").sum()),
        )
        .reset_index()
    )
    sector_summary_all = sector_summary.sort_values(sector_sort, ascending=False)
    sector_summary = sector_summary_all[sector_summary_all["Saham"] >= sector_min_count]
    if sector_summary.empty:
        st.warning("Tidak ada kelompok yang memenuhi minimum saham. Menampilkan semua kelompok sementara.")
        sector_summary = sector_summary_all

    left, right = st.columns([1.1, 1])
    with left:
        if sector_chart == "Scatter":
            fig = px.scatter(
                sector_summary,
                x="Total_Market_Cap",
                y="Median_Score",
                size="Saham",
                color="Strong_Buy",
                hover_name=sector_group,
                title=f"{sector_group}: market cap vs median score",
                color_continuous_scale="Greens",
            )
        elif sector_chart == "Treemap":
            fig = px.treemap(
                sector_summary,
                path=[sector_group],
                values="Total_Market_Cap",
                color="Median_Score",
                title=f"Peta market cap dan score {sector_group.lower()}",
                color_continuous_scale="RdYlGn",
            )
        else:
            fig = px.bar(
                sector_summary,
                x=sector_sort,
                y=sector_group,
                orientation="h",
                color="Strong_Buy",
                title=f"Ranking {sector_group.lower()} berdasarkan {sector_sort}",
                color_continuous_scale="Greens",
            )
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
        fig.update_layout(height=520)
        show_chart(fig)

    with right:
        market_cap_view = sector_summary.sort_values("Total_Market_Cap", ascending=True).tail(15)
        fig = px.bar(
            market_cap_view,
            x="Total_Market_Cap",
            y=sector_group,
            orientation="h",
            color="Median_Score",
            title=f"Kontribusi market cap {sector_group.lower()}",
            color_continuous_scale="RdYlGn",
        )
        fig.update_layout(xaxis_title="Total market cap", yaxis_title="")
        fig.update_layout(height=520)
        show_chart(fig)

    show_table(
        sector_summary,        hide_index=True,
        column_config={
            "Median_Score": st.column_config.NumberColumn("Median Score", format="%.1f"),
            "Avg_ROE": st.column_config.NumberColumn("Avg ROE", format="%.1f%%"),
            "Median_PER": st.column_config.NumberColumn("Median PER", format="%.1f"),
            "Median_PBV": st.column_config.NumberColumn("Median PBV", format="%.1f"),
            "Total_Turnover": st.column_config.NumberColumn("Total Turnover", format="Rp %.0f"),
            "Total_Market_Cap": st.column_config.NumberColumn("Total Market Cap", format="Rp %.0f"),
            "Total_Revenue": st.column_config.NumberColumn("Total Revenue", format="Rp %.0f"),
            "Median_Sales_Multiple": st.column_config.NumberColumn("Median MCap/Revenue", format="%.2f"),
        },
    )

with tab_quality:
    st.subheader("Data Quality & Workflow Update")
    quality_report = build_data_quality_report(scored_df, raw_df)
    review_count = int((quality_report["Rows"].gt(0) & ~quality_report["Severity"].eq("Info")).sum())
    high_count = int(((quality_report["Rows"] > 0) & quality_report["Severity"].eq("High")).sum())
    quality_cols = st.columns(4)
    quality_cols[0].metric("Status check", f"{len(quality_report) - review_count}/{len(quality_report)} OK")
    quality_cols[1].metric("Perlu review", f"{review_count}")
    quality_cols[2].metric("High severity", f"{high_count}")
    quality_cols[3].metric("Lolos data bersih", f"{scored_df['Clean_Data'].sum():,}")

    universe_summary = (
        scored_df.groupby(["Universe_Diff_Status", "Universe_Source"], dropna=False)
        .agg(Jumlah=("Kode", "nunique"))
        .reset_index()
        .rename(columns={"Universe_Diff_Status": "Status Kode", "Universe_Source": "Sumber Kode"})
        .sort_values(["Status Kode", "Sumber Kode"])
    )
    universe_cols = st.columns(4)
    universe_cols[0].metric("Kode universe", f"{scored_df['Kode'].nunique():,}")
    universe_cols[1].metric("Match BEI/IDX", f"{scored_df['In_IDX_Official'].sum():,}")
    universe_cols[2].metric("Fallback non-BEI", f"{(~scored_df['In_IDX_Official']).sum():,}")
    universe_cols[3].metric("Sumber aktif", f"{scored_df['Universe_Source'].nunique():,}")
    with st.expander("Audit sumber kode saham", expanded=False):
        st.caption(HELP_TEXT["universe_audit"])
        show_table(
            universe_summary,            hide_index=True,
            column_config={"Jumlah": st.column_config.NumberColumn("Jumlah", format="%d")},
        )
        diff_codes = scored_df[~scored_df["In_IDX_Official"]][
            ["Kode", "Nama Perusahaan", "Universe_Source", "Universe_Diff_Status", "Sektor", "Industry", "ListingBoard"]
        ].sort_values("Kode")
        if diff_codes.empty:
            st.success("Semua kode di universe match dengan daftar resmi BEI/IDX.")
        else:
            st.warning(f"Ada {len(diff_codes):,} kode yang tidak match dengan daftar resmi BEI/IDX dan tetap dipertahankan dari fallback.")
            show_table(diff_codes, hide_index=True)

    with st.expander("Kelengkapan kolom & sumber data", expanded=True):
        completeness_report = build_completeness_report(scored_df)
        completeness_metric = completeness_report.groupby("Grup", as_index=False)["Coverage"].mean().sort_values("Coverage")
        source_mix = build_source_mix(scored_df)
        data_health_cols = st.columns([1, 1])
        with data_health_cols[0]:
            fig = px.bar(
                completeness_metric,
                x="Coverage",
                y="Grup",
                orientation="h",
                text="Coverage",
                title="Coverage rata-rata per grup data",
                color="Coverage",
                color_continuous_scale="RdYlGn",
                range_color=[0, 100],
            )
            fig.update_traces(texttemplate="%{text:.0f}%", textposition="outside", cliponaxis=False)
            fig.update_layout(height=360, xaxis_title="Coverage", yaxis_title="", margin=dict(l=20, r=70, t=60, b=40))
            show_chart(fig)
        with data_health_cols[1]:
            if source_mix.empty:
                st.info("Ringkasan sumber data belum tersedia.")
            else:
                source_focus = source_mix[source_mix["Area"].isin(["Price_Source", "Volume_Source", "Fundamental_Source", "Universe_Diff_Status"])]
                fig = px.bar(
                    source_focus,
                    x="Jumlah",
                    y="Nilai",
                    color="Area",
                    orientation="h",
                    title="Campuran sumber data utama",
                )
                fig.update_layout(height=360, yaxis_title="", margin=dict(l=20, r=20, t=60, b=40))
                show_chart(fig)

        show_table(
            completeness_report.sort_values(["Coverage", "Grup", "Kolom"]),            hide_index=True,
            column_config={
                "Coverage": st.column_config.ProgressColumn("Coverage", min_value=0, max_value=100, format="%.0f%%"),
                "Terisi": st.column_config.NumberColumn("Terisi", format="%d"),
                "Kosong": st.column_config.NumberColumn("Kosong", format="%d"),
            },
        )

    with st.expander("Audit Kode Saham: alasan lolos/gagal filter", expanded=True):
        audit_scope = st.radio(
            "Cakupan audit",
            ["Semua saham", "Hasil filter aktif", "Kode pilihan"],
            horizontal=True,
            help=HELP_TEXT["audit_scope"],
        )
        if audit_scope == "Semua saham":
            audit_codes = scored_df["Kode"].tolist()
            st.caption(f"Mengaudit seluruh universe: {len(audit_codes):,} saham x {len(preset_filter_criteria)} skenario filter.")
        elif audit_scope == "Hasil filter aktif":
            audit_codes = filtered["Kode"].tolist()
            st.caption(f"Mengaudit saham yang lolos filter sidebar saat ini: {len(audit_codes):,} saham.")
        else:
            default_audit_codes = [code for code in ["BBCA", "BBRI", "BMRI", "BBNI"] if code in scored_df["Kode"].values]
            audit_codes = st.multiselect(
                "Pilih kode untuk diaudit",
                options=scored_df["Kode"].tolist(),
                default=default_audit_codes,
                help=HELP_TEXT["audit_code"],
            )
        if audit_codes:
            audit_summary, audit_detail = build_filter_audit(scored_df, audit_codes, preset_filter_criteria)
            audit_metric_cols = st.columns(4)
            audit_metric_cols[0].metric("Kode diaudit", f"{len(audit_codes):,}")
            audit_metric_cols[1].metric("Baris audit", f"{len(audit_summary):,}")
            audit_metric_cols[2].metric("Lolos", f"{audit_summary['Status'].eq('Lolos').sum():,}")
            audit_metric_cols[3].metric("Gagal", f"{audit_summary['Status'].eq('Gagal').sum():,}")

            audit_filter_cols = st.columns([1, 1, 2])
            with audit_filter_cols[0]:
                audit_filter_view = st.selectbox(
                    "Tampilkan skenario",
                    ["Semua"] + [criteria["name"] for criteria in preset_filter_criteria],
                    help="Batasi ringkasan audit ke satu skenario filter.",
                )
            with audit_filter_cols[1]:
                audit_status_view = st.selectbox(
                    "Tampilkan status",
                    ["Semua", "Lolos", "Gagal"],
                    help="Batasi ringkasan audit berdasarkan status lolos/gagal.",
                )
            with audit_filter_cols[2]:
                audit_search = st.text_input(
                    "Cari kode/nama/alasan",
                    value="",
                    help="Cari kode, nama perusahaan, atau alasan gagal pada ringkasan audit.",
                ).strip().upper()

            audit_view = audit_summary.copy()
            if audit_filter_view != "Semua":
                audit_view = audit_view[audit_view["Filter"].eq(audit_filter_view)]
            if audit_status_view != "Semua":
                audit_view = audit_view[audit_view["Status"].eq(audit_status_view)]
            if audit_search:
                search_text = (
                    audit_view["Kode"].astype(str)
                    + " "
                    + audit_view["Nama Perusahaan"].astype(str)
                    + " "
                    + audit_view["Alasan Utama"].astype(str)
                ).str.upper()
                audit_view = audit_view[search_text.str.contains(audit_search, na=False)]

            show_table(
                audit_view,                hide_index=True,
                column_config={
                    "Score": st.column_config.NumberColumn("Score", format="%.1f", help=HELP_TEXT["score"]),
                    "Clean_Data": st.column_config.CheckboxColumn("Clean Data", help=HELP_TEXT["clean_data"]),
                    "Volume_Source": st.column_config.TextColumn("Sumber Volume"),
                },
            )
            if audit_view.empty:
                st.info("Tidak ada baris audit yang cocok dengan filter tampilan.")
            detail_cols = st.columns([1, 1])
            with detail_cols[0]:
                detail_options = audit_view["Kode"].drop_duplicates().tolist() if not audit_view.empty else audit_codes
                detail_code = st.selectbox("Detail kode", detail_options, help="Pilih kode untuk melihat checklist filter satu per satu.")
            with detail_cols[1]:
                detail_filter = st.selectbox(
                    "Detail filter",
                    [criteria["name"] for criteria in preset_filter_criteria],
                    help="Pilih skenario filter yang ingin dibedah.",
                )
            selected_detail = audit_detail[
                audit_detail["Kode"].eq(detail_code) & audit_detail["Filter"].eq(detail_filter)
            ]
            show_table(
                selected_detail,                hide_index=True,
                column_config={
                    "Status": st.column_config.TextColumn("Status"),
                    "Actual": st.column_config.TextColumn("Nilai Saat Ini"),
                    "Required": st.column_config.TextColumn("Kriteria"),
                },
            )
        else:
            st.info("Pilih minimal satu kode saham untuk audit.")

    show_table(
        quality_report,        hide_index=True,
        column_config={
            "Rows": st.column_config.NumberColumn("Rows", format="%d"),
        },
    )

    issue_options = quality_report.loc[quality_report["Rows"].gt(0), "Check"].tolist()
    if issue_options:
        selected_issue = st.selectbox("Lihat detail masalah", issue_options, help=HELP_TEXT["quality_issue"])
        detail = get_quality_detail(scored_df, selected_issue)
        show_table(
            detail.head(200),            hide_index=True,
            column_config={
                "Score": st.column_config.NumberColumn("Score", format="%.1f", help=HELP_TEXT["score"]),
                "Penutupan": st.column_config.NumberColumn("Harga", format="Rp %.0f", help=HELP_TEXT["price"]),
                "Volume": st.column_config.NumberColumn("Volume", format="%.0f", help=HELP_TEXT["volume"]),
                "Market_Cap": st.column_config.NumberColumn("Market Cap", format="Rp %.0f"),
                "Revenue": st.column_config.NumberColumn("Revenue", format="Rp %.0f"),
                "Sales_Multiple": st.column_config.NumberColumn("MCap/Revenue", format="%.2f"),
                "PER": st.column_config.NumberColumn("PER", format="%.2f", help=HELP_TEXT["per"]),
                "PBV": st.column_config.NumberColumn("PBV", format="%.2f", help=HELP_TEXT["pbv"]),
                "ROE": st.column_config.NumberColumn("ROE", format="%.1f%%", help=HELP_TEXT["roe"]),
                "ROA": st.column_config.NumberColumn("ROA", format="%.1f%%", help=HELP_TEXT["quality"]),
                "NPM": st.column_config.NumberColumn("NPM", format="%.1f%%", help=HELP_TEXT["npm"]),
                "Threshold_Pass_Ratio": st.column_config.NumberColumn("Threshold", format="%.0f%%", help=HELP_TEXT["threshold_ratio"]),
                "Return_52W": st.column_config.NumberColumn("52W", format="%.1f%%", help=HELP_TEXT["return"]),
            },
        )
    else:
        st.success("Tidak ada issue data quality pada check utama.")

    status_left, status_right = st.columns([1, 1])
    with status_left:
        st.write("Status Excel fallback")
        show_table(pd.DataFrame([get_file_status(DATA_FILE)]), hide_index=True)
    with status_right:
        st.write("Status cache histori")
        cache_status = get_history_cache_status()
        if cache_status.empty:
            st.info("Belum ada cache histori online.")
        else:
            show_table(cache_status.sort_values("Modified", ascending=False).head(50), hide_index=True)

    with st.expander("Workflow rutin yang disarankan", expanded=True):
        st.markdown(
            """
            1. Ambil universe kode dari daftar resmi BEI/IDX sebagai sumber utama.
            2. Lengkapi harga/histori dari yfinance dan fundamental massal dari TradingView scanner.
            3. Pakai `Ringkasan.xlsx` hanya untuk mengisi kolom yang belum tersedia online dan sebagai pembanding metodologi.
            4. Buka `Audit sumber kode saham` dan `Kelengkapan kolom & sumber data` untuk memastikan fallback terlihat jelas.
            5. Jalankan `Refresh cache histori top saham` di sidebar untuk memperbarui cache online.
            6. Update `Ringkasan.xlsx` hanya bila ada data offline yang lebih baik atau ide algoritme baru yang perlu diuji.
            7. Simpan/export hasil rekomendasi hanya setelah check High severity terkendali.
            """
        )

with tab_method:
    st.subheader("Formula scoring multi-factor")
    st.markdown(
        """
        Score akhir memakai normalisasi percentile yang dipotong di persentil 3 dan 97 agar outlier ekstrem tidak mendominasi.

        Faktor yang dihitung:
        - Valuasi: PER rendah dan PBV rendah, hanya untuk nilai positif yang masuk akal.
        - Kualitas profit: ROE, ROA, dan NPM positif.
        - Risiko: non-bank memakai DER rendah dan intraday range rendah; Banking memakai CAR, NPL, BOPO, dan LDR bila tersedia.
        - Likuiditas: volume dan turnover harga x volume.
        - Momentum: kombinasi histori online 4, 13, 26, 52 minggu dan perubahan harga harian yang tidak ekstrem, dengan Excel Metrik sebagai fallback.
        - Kekuatan indeks: nilai kolom Sigma i >= 7 dari Excel bila tersedia, atau jumlah indeks/tempat kemunculan dari fallback; bila hanya universe BEI/IDX tersedia, minimal dihitung sebagai saham listed.
        - Threshold sheet: rasio dibandingkan dengan batas dari sheet NonBank atau Banking sebagai cadangan metodologi fundamental.

        Penalti diterapkan untuk PER/PBV negatif, profitabilitas negatif, NPM negatif, volume rendah, harga nol, pergerakan harian ekstrem, dan kelulusan threshold yang terlalu rendah.
        """
    )

    st.info(f"Profil scoring aktif: {profile}. Preset filter aktif: {filter_preset}. Profil scoring mengubah bobot, preset filter mengubah batas penyaringan.")

    method_cols = st.columns(5)
    method_cols[0].metric("Baris sumber", f"{len(raw_df):,}")
    method_cols[1].metric("Saham unik setelah deduplikasi", f"{len(df):,}")
    method_cols[2].metric("Duplikasi indeks dibersihkan", f"{len(raw_df) - len(df):,}")
    method_cols[3].metric("Histori 52W tersedia", f"{df['Return_52W'].notna().sum():,}")
    method_cols[4].metric("Update data", data_update_label)

    st.write("Bobot aktif:")
    weight_df = pd.DataFrame(
        [{"Faktor": key.replace("_", " ").title(), "Bobot": value} for key, value in weights.items()]
    )
    show_table(weight_df, hide_index=True)

    method_view_cols = st.columns([1, 1])
    with method_view_cols[0]:
        factor_to_inspect = st.selectbox(
            "Inspeksi faktor",
            ["Score", "Valuation_Score", "Quality_Score", "Risk_Score", "Liquidity_Score", "Momentum_Score", "Index_Score", "Penalty"],
            help=HELP_TEXT["factor_inspect"],
        )
    with method_view_cols[1]:
        factor_top_n = st.slider("Jumlah contoh faktor", 5, 50, 15, step=5, help=HELP_TEXT["factor_top_n"])

    factor_examples = scored_df.sort_values(factor_to_inspect, ascending=False).head(factor_top_n)
    fig = px.histogram(scored_df, x=factor_to_inspect, nbins=40, title=f"Distribusi {factor_to_inspect}", color_discrete_sequence=["#2563eb"])
    fig.update_layout(height=320)
    show_chart(fig)
    factor_example_columns = list(
        dict.fromkeys(
            ["Kode", "Nama Perusahaan", "Recommendation", "Risk_Level", factor_to_inspect, "Score", "Threshold_Pass_Ratio", "Sektor"]
        )
    )
    show_table(
        factor_examples[factor_example_columns],        hide_index=True,
    )

    threshold_info = pd.DataFrame(
        [
            {"Mode": "NonBank", "Metric": metric, "Rule": f"{operator} {threshold:g}"}
            for metric, (operator, threshold) in NONBANK_THRESHOLDS.items()
        ]
        + [
            {"Mode": "Banking", "Metric": metric, "Rule": f"{operator} {threshold:g}"}
            for metric, (operator, threshold) in BANKING_THRESHOLDS.items()
        ]
    )
    threshold_mode_view = st.segmented_control("Tampilkan threshold", ["Semua", "NonBank", "Banking"], default="Semua")
    if threshold_mode_view != "Semua":
        threshold_info = threshold_info[threshold_info["Mode"] == threshold_mode_view]
    st.write("Threshold aktif dari sheet:")
    show_table(threshold_info, hide_index=True)

    penalty_info = pd.DataFrame(
        [
            {"Penalti": "PER atau PBV negatif/tidak sehat", "Dampak": "-10"},
            {"Penalti": "ROE atau ROA negatif", "Dampak": "-9"},
            {"Penalti": "NPM negatif", "Dampak": "-6"},
            {"Penalti": "Volume di bawah 1 juta", "Dampak": "-6"},
            {"Penalti": "Harga penutupan nol/tidak valid", "Dampak": "-25"},
            {"Penalti": "Perubahan harian ekstrem", "Dampak": "-6"},
            {"Penalti": "Threshold lolos di bawah 40 persen", "Dampak": "-5"},
        ]
    )
    with st.expander("Daftar penalti scoring"):
        show_table(penalty_info, hide_index=True)

    with st.expander("Catatan risiko"):
        st.markdown(
            """
            Dashboard ini adalah alat screening kuantitatif awal. Hasil terbaik tetap perlu dicek manual:
            laporan keuangan terbaru, aksi korporasi, kualitas manajemen, valuasi historis, berita material,
            dan rencana alokasi portofolio. Satu hari data harga juga tidak cukup untuk menyimpulkan tren jangka panjang.
            """
        )

st.caption("Built with Streamlit + Plotly. Gunakan sebagai screener, bukan pengganti keputusan investasi.")
