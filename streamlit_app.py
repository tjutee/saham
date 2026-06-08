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
YFINANCE_CACHE_DIR = HISTORY_CACHE_DIR / ".yfinance"
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
APP_BUILD = os.environ.get("STREAMLIT_GIT_COMMIT", "technical-visible-ui")
ONLINE_PERIOD_OPTIONS = ["5d", "2wk", "1mo", "3mo", "6mo", "1y", "2y", "5y", "10y", "max"]
ONLINE_PERIOD_LABELS = {
    "5d": "1 minggu",
    "2wk": "2 minggu",
    "1mo": "1 bulan",
    "3mo": "3 bulan",
    "6mo": "6 bulan",
    "1y": "1 tahun",
    "2y": "2 tahun",
    "5y": "5 tahun",
    "10y": "10 tahun",
    "max": "All / sepanjang masa",
}
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
    "history_source": "Online yfinance memakai ticker KODE.JK dan menjadi sumber utama grafik histori. Excel Metrik hanya mode pembanding/cadangan bila data online kosong.",
    "fundamental_source": "Fundamental diprioritaskan dari online TradingView scanner bila tersedia, lalu Excel mengisi rasio/metadata yang kosong. Sumber BEI/IDX tetap utama untuk universe kode saham.",
    "history_scope": "Saham pilihan memakai kode yang dipilih manual. All/top N memakai saham teratas dari hasil filter saat ini, biasanya berdasarkan ranking Score setelah filter.",
    "history_top_n": "Jumlah kode dari hasil filter/ranking yang dimasukkan ke grafik All/top N. Makin besar makin lengkap, tetapi grafik online bisa lebih lambat.",
    "history_codes": "Masukkan kode IDX tanpa akhiran .JK, misalnya BBCA atau BBRI. Dashboard otomatis memanggil format online BBCA.JK.",
    "history_period": "Rentang data online: 1 minggu, 2 minggu untuk short swing, 1/3/6 bulan, 1/2/5/10 tahun, atau All sepanjang data tersedia dari sumber.",
    "recommendation": "Recommendation murni dari Score: Strong Buy >= 78, Buy >= 68, Watchlist >= 55, Speculative >= 42, selain itu Avoid. Ini hasil screener, bukan instruksi beli.",
    "risk_level": "Risk_Level adalah kategori risiko relatif dari model berdasarkan rasio, volatilitas, likuiditas, dan penalti. Tetap perlu validasi berita dan laporan keuangan.",
    "turnover": "Turnover = Penutupan x Volume. Grafik utama memakai harga/volume yfinance/cache bila tersedia; Excel hanya fallback saat online kosong.",
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
    "technical_period": "Rentang OHLCV online, sama seperti bagian histori pada tab Harga & Teknikal. Periode pendek cocok untuk RSI/MACD cepat; minimal 1-2 tahun disarankan agar MA200 dan 52W lebih stabil.",
    "technical_code": "Pilih satu kode saham untuk candlestick dan indikator detail. Data diambil dari yfinance/cache memakai format KODE.JK.",
    "technical_score": "Technical_Score adalah konfirmasi timing berbasis trend, RSI, MACD, volume, dan volatilitas. Ini tidak mengganti Score fundamental utama.",
    "technical_filter": "Filter sinyal teknikal untuk melihat kandidat dengan kondisi trend/momentum tertentu dari hasil filter aktif. Kosongkan pilihan untuk menampilkan semua sinyal.",
    "entry_action": "Entry_Action menggabungkan fundamental dan teknikal: fundamental memilih saham layak, teknikal menentukan timing entry/tunggu/tahan/take profit. Kosongkan filter untuk menampilkan semua aksi entry.",
    "position_action": "Position_Action adalah arahan umum untuk saham yang sudah dimiliki, tanpa memakai harga beli pribadi. Kosongkan filter untuk menampilkan semua aksi posisi.",
    "sector_relative": "Sector_Relative_Score membandingkan valuasi dan kualitas saham terhadap sektor yang sama. Ini membantu mengurangi bias karena PER/PBV/ROE wajar berbeda antar sektor.",
    "explainability": "Decision_Summary, Top_Strengths, Top_Risks, dan Action_Checklist menjelaskan faktor utama di balik ranking agar keputusan tidak menjadi black box.",
    "atr_stop": "ATR_Stop_2x adalah zona risiko teknikal berbasis dua kali ATR dari harga terakhir. Ini bukan instruksi order otomatis dan tetap perlu disesuaikan dengan profil risiko pribadi.",
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

RECOMMENDATION_COLORS = {
    "Strong Buy": "#15803d",
    "Buy": "#65a30d",
    "Watchlist": "#ca8a04",
    "Speculative": "#ea580c",
    "Avoid": "#dc2626",
}
RISK_COLORS = {"Low": "#15803d", "Medium": "#ca8a04", "High": "#dc2626"}
TECHNICAL_SIGNAL_COLORS = {
    "Bullish": "#15803d",
    "Constructive": "#65a30d",
    "Neutral": "#2563eb",
    "Overbought": "#ea580c",
    "Weak": "#dc2626",
}
ENTRY_ACTION_COLORS = {
    "Buy Candidate": "#15803d",
    "Wait Pullback": "#ca8a04",
    "Wait Confirmation": "#2563eb",
    "Hold/Monitor": "#0f766e",
    "Take Profit / Tight Stop": "#ea580c",
    "Avoid Entry": "#dc2626",
}
ENTRY_ACTION_ORDER = [
    "Buy Candidate",
    "Wait Confirmation",
    "Hold/Monitor",
    "Wait Pullback",
    "Take Profit / Tight Stop",
    "Avoid Entry",
]
POSITION_ACTION_COLORS = {
    "Hold": "#15803d",
    "Add on Pullback": "#65a30d",
    "Tight Stop": "#ca8a04",
    "Take Profit": "#ea580c",
    "Reduce": "#ea580c",
    "Exit / Sell": "#dc2626",
    "Review Position": "#2563eb",
}
POSITION_ACTION_ORDER = [
    "Hold",
    "Add on Pullback",
    "Review Position",
    "Tight Stop",
    "Take Profit",
    "Reduce",
    "Exit / Sell",
]
SOURCE_COLORS = {
    "Price_Source": "#2563eb",
    "Volume_Source": "#0891b2",
    "Fundamental_Source": "#7c3aed",
    "Data_Source": "#475569",
    "Universe_Source": "#0f766e",
    "Universe_Diff_Status": "#ea580c",
}
FACTOR_COLORS = {
    "Score": "#2563eb",
    "Valuation_Score": "#0f766e",
    "Quality_Score": "#15803d",
    "Risk_Score": "#ca8a04",
    "Liquidity_Score": "#0891b2",
    "Momentum_Score": "#7c3aed",
    "History_Momentum_Score": "#9333ea",
    "Threshold_Pass_Ratio": "#475569",
    "PER": "#0f766e",
    "PBV": "#14b8a6",
    "ROE": "#15803d",
    "DER": "#ca8a04",
    "Return_52W": "#7c3aed",
}
STOCK_LINE_COLORS = [
    "#2563eb",
    "#0f766e",
    "#7c3aed",
    "#0891b2",
    "#ca8a04",
    "#db2777",
    "#475569",
    "#65a30d",
]
SCORE_SCALE = "RdYlGn"
COUNT_SCALE = "Blues"
CHART_AXIS_COLOR = "#64748b"


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


def format_percent(value, digits=1):
    if pd.isna(value):
        return "-"
    return f"{value:,.{digits}f}%"


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
    fig.update_layout(
        font=dict(color="#0f172a"),
        hoverlabel=dict(bgcolor="#ffffff", font_size=12),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return st.plotly_chart(fig, *args, **PLOTLY_STRETCH, **kwargs)


def chart_color_kwargs(color_field):
    if color_field == "Recommendation":
        return {"color_discrete_map": RECOMMENDATION_COLORS}
    if color_field == "Risk_Level":
        return {"color_discrete_map": RISK_COLORS}
    if color_field == "Technical_Signal":
        return {"color_discrete_map": TECHNICAL_SIGNAL_COLORS}
    if color_field == "Entry_Action":
        return {"color_discrete_map": ENTRY_ACTION_COLORS}
    if color_field == "Position_Action":
        return {"color_discrete_map": POSITION_ACTION_COLORS}
    if color_field in ["Sektor", "Industry", "Kode"]:
        return {"color_discrete_sequence": STOCK_LINE_COLORS}
    if color_field in ["Score", "Quality_Score", "Risk_Score", "Threshold_Pass_Ratio", "Valuation_Score", "Liquidity_Score", "Momentum_Score", "History_Momentum_Score"]:
        return {"color_continuous_scale": SCORE_SCALE, "range_color": [0, 100]}
    return {"color_discrete_sequence": STOCK_LINE_COLORS}


def safe_slider(label, min_value, max_value, value, *, step=None, help=None, format=None):
    if pd.isna(min_value) or pd.isna(max_value):
        st.caption(f"{label}: data belum tersedia.")
        return value
    if max_value <= min_value:
        st.caption(f"{label}: {min_value}")
        return (min_value, min_value) if isinstance(value, tuple) else min_value
    if isinstance(value, tuple):
        low, high = value
        low = max(min_value, min(max_value, low))
        high = max(min_value, min(max_value, high))
        if low > high:
            low, high = high, low
        value = (low, high)
    else:
        value = max(min_value, min(max_value, value))
    kwargs = {"help": help}
    if step is not None:
        kwargs["step"] = step
    if format is not None:
        kwargs["format"] = format
    return st.slider(label, min_value, max_value, value, **kwargs)


def has_online_market(data):
    if "Price_Source" not in data.columns:
        return pd.Series(False, index=data.index)
    return data["Price_Source"].astype(str).str.contains("yfinance|pandas-datareader|cache", case=False, na=False)


def chart_market_frame(data, label="grafik"):
    online_mask = has_online_market(data)
    if online_mask.any():
        chart_data = data[online_mask].copy()
        fallback_count = int((~online_mask).sum())
        if fallback_count:
            st.caption(f"{label}: memakai {len(chart_data):,} saham dengan harga yfinance/cache. {fallback_count:,} saham Excel fallback tidak dimasukkan ke grafik utama.")
        else:
            st.caption(f"{label}: seluruh saham memakai harga yfinance/cache.")
        return chart_data
    st.caption(f"{label}: data yfinance/cache belum tersedia, memakai Excel fallback sementara.")
    return data.copy()


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
            "Action": "Gunakan tab Harga & Teknikal untuk melengkapi konteks tren.",
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


def grouped_percentile_score(dataframe, group_column, value_column, higher_is_better=True, min_group_size=5):
    values = pd.to_numeric(dataframe.get(value_column, pd.Series(np.nan, index=dataframe.index)), errors="coerce")
    groups = dataframe.get(group_column, pd.Series("All", index=dataframe.index)).fillna("All").astype(str)
    output = pd.Series(50.0, index=dataframe.index)

    for _, index in groups.groupby(groups).groups.items():
        group_values = values.loc[index]
        valid_values = group_values.dropna()
        if len(valid_values) < min_group_size:
            continue
        ranks = valid_values.rank(method="average", ascending=True)
        if len(valid_values) == 1:
            scores = pd.Series(50.0, index=valid_values.index)
        elif higher_is_better:
            scores = (ranks - 1) / (len(valid_values) - 1) * 100
        else:
            scores = (len(valid_values) - ranks) / (len(valid_values) - 1) * 100
        output.loc[scores.index] = scores.clip(0, 100)
    return output


def add_sector_relative_scores(scored):
    output = scored.copy()
    group_column = "Sektor" if "Sektor" in output.columns else "Industry"

    output["Sector_PER_Score"] = grouped_percentile_score(output, group_column, "PER", higher_is_better=False)
    output["Sector_PBV_Score"] = grouped_percentile_score(output, group_column, "PBV", higher_is_better=False)
    output["Sector_ROE_Score"] = grouped_percentile_score(output, group_column, "ROE", higher_is_better=True)
    output["Sector_ROA_Score"] = grouped_percentile_score(output, group_column, "ROA", higher_is_better=True)
    output["Sector_NPM_Score"] = grouped_percentile_score(output, group_column, "NPM", higher_is_better=True)
    output["Sector_Valuation_Score"] = output["Sector_PER_Score"] * 0.58 + output["Sector_PBV_Score"] * 0.42
    output["Sector_Quality_Score"] = (
        output["Sector_ROE_Score"] * 0.45
        + output["Sector_ROA_Score"] * 0.30
        + output["Sector_NPM_Score"] * 0.25
    )
    output["Sector_Relative_Score"] = (
        output["Sector_Valuation_Score"] * 0.45
        + output["Sector_Quality_Score"] * 0.40
        + output.get("Liquidity_Score", pd.Series(50, index=output.index)) * 0.15
    ).clip(0, 100).round(1)
    output["Sector_Relative_Label"] = np.select(
        [
            output["Sector_Relative_Score"].ge(75),
            output["Sector_Relative_Score"].ge(60),
            output["Sector_Relative_Score"].lt(40),
        ],
        ["Unggul sektor", "Kompetitif", "Lemah sektor"],
        default="Netral sektor",
    )
    return output


def add_decision_explainability(scored):
    output = scored.copy()
    factor_labels = {
        "Valuation_Score": "valuasi",
        "Quality_Score": "kualitas profit",
        "Risk_Score": "risiko relatif",
        "Liquidity_Score": "likuiditas",
        "Momentum_Score": "momentum",
        "Index_Score": "coverage indeks",
        "Sector_Relative_Score": "relatif sektor",
    }

    summaries = []
    strengths = []
    risks = []
    checklists = []
    for _, row in output.iterrows():
        factor_values = {
            label: pd.to_numeric(row.get(column), errors="coerce")
            for column, label in factor_labels.items()
        }
        valid_factors = {label: value for label, value in factor_values.items() if pd.notna(value)}
        top_strengths = sorted(valid_factors.items(), key=lambda item: item[1], reverse=True)[:3]
        weak_factors = sorted(valid_factors.items(), key=lambda item: item[1])[:3]

        strengths.append(", ".join(f"{label} {value:.0f}" for label, value in top_strengths) if top_strengths else "-")
        risks.append(", ".join(f"{label} {value:.0f}" for label, value in weak_factors) if weak_factors else "-")

        checklist = []
        if not bool(row.get("Clean_Data", False)):
            checklist.append("validasi data")
        if clean_text(row.get("Risk_Level")) == "High":
            checklist.append("kurangi risiko")
        if pd.to_numeric(row.get("Sector_Relative_Score"), errors="coerce") < 45:
            checklist.append("bandingkan ulang sektor")
        if pd.to_numeric(row.get("Threshold_Pass_Ratio"), errors="coerce") < 65:
            checklist.append("cek threshold")
        if pd.to_numeric(row.get("Momentum_Score"), errors="coerce") < 45:
            checklist.append("tunggu momentum")
        if not checklist:
            checklist.append("konfirmasi teknikal")
        checklists.append(", ".join(checklist))

        recommendation = clean_text(row.get("Recommendation"), "Avoid")
        sector_label = clean_text(row.get("Sector_Relative_Label"), "Netral sektor")
        risk_level = clean_text(row.get("Risk_Level"), "High")
        summaries.append(f"{recommendation}; {sector_label}; risiko {risk_level}")

    output["Top_Strengths"] = strengths
    output["Top_Risks"] = risks
    output["Action_Checklist"] = checklists
    output["Decision_Summary"] = summaries
    return output


def build_data_freshness(scored):
    if scored.empty:
        return {
            "Freshness_Label": "Unknown",
            "Latest_Online_Date": None,
            "Online_Data_Lag_Days": np.nan,
            "Online_Price_Coverage_%": 0.0,
            "Online_Fundamental_Coverage_%": 0.0,
            "Excel_Fundamental_Coverage_%": 0.0,
            "Stale_Price_Rows": 0,
        }

    today = pd.Timestamp.today().normalize()
    online_dates = pd.to_datetime(scored.get("Online_Last_Date", pd.Series(pd.NaT, index=scored.index)), errors="coerce")
    latest_online_date = online_dates.max()
    lag_days = (today - latest_online_date.normalize()).days if pd.notna(latest_online_date) else np.nan
    stale_rows = int((online_dates.notna() & ((today - online_dates.dt.normalize()).dt.days > 5)).sum())
    online_price_coverage = (
        scored.get("Price_Source", pd.Series("", index=scored.index))
        .astype(str)
        .str.contains("yfinance|cache|pandas-datareader", case=False, na=False)
        .mean()
        * 100
    )
    online_fundamental_count = pd.to_numeric(
        scored.get("Online_Fundamental_Field_Count", pd.Series(0, index=scored.index)), errors="coerce"
    ).fillna(0)
    excel_fundamental_count = pd.to_numeric(
        scored.get("Excel_Fundamental_Field_Count", pd.Series(0, index=scored.index)), errors="coerce"
    ).fillna(0)
    fundamental_total = online_fundamental_count + excel_fundamental_count
    online_fundamental_coverage = (
        online_fundamental_count.sum() / fundamental_total.sum() * 100 if fundamental_total.sum() > 0 else 0.0
    )
    excel_fundamental_coverage = (
        excel_fundamental_count.sum() / fundamental_total.sum() * 100 if fundamental_total.sum() > 0 else 0.0
    )

    if pd.isna(lag_days):
        label = "Unknown"
    elif lag_days <= 3 and online_price_coverage >= 70:
        label = "Fresh"
    elif lag_days <= 7 and online_price_coverage >= 40:
        label = "Stale"
    else:
        label = "Needs Refresh"

    return {
        "Freshness_Label": label,
        "Latest_Online_Date": latest_online_date,
        "Online_Data_Lag_Days": lag_days,
        "Online_Price_Coverage_%": round(float(online_price_coverage), 1),
        "Online_Fundamental_Coverage_%": round(float(online_fundamental_coverage), 1),
        "Excel_Fundamental_Coverage_%": round(float(excel_fundamental_coverage), 1),
        "Stale_Price_Rows": stale_rows,
    }


@st.cache_data(ttl=3600, show_spinner=False)
def build_market_regime(period="2y"):
    history, error, source = fetch_yahoo_symbol_history(["^JKSE"], period=period)
    if history.empty:
        return {
            "Market_Regime": "Unknown",
            "IHSG_Close": np.nan,
            "IHSG_MA50": np.nan,
            "IHSG_MA200": np.nan,
            "IHSG_Return_20D_%": np.nan,
            "IHSG_Return_60D_%": np.nan,
            "IHSG_Last_Date": pd.NaT,
            "Regime_Reason": error or "Data IHSG kosong.",
            "Market_Source": source,
            "Market_Error": error,
        }

    ihsg = history[history["Symbol"].eq("^JKSE")].sort_values("Date").copy()
    ihsg["MA50"] = ihsg["Close"].rolling(50, min_periods=20).mean()
    ihsg["MA200"] = ihsg["Close"].rolling(200, min_periods=80).mean()
    ihsg["Return_20D_%"] = ihsg["Close"].pct_change(20) * 100
    ihsg["Return_60D_%"] = ihsg["Close"].pct_change(60) * 100
    latest = ihsg.iloc[-1]
    close = latest.get("Close")
    ma50 = latest.get("MA50")
    ma200 = latest.get("MA200")
    ret20 = latest.get("Return_20D_%")
    ret60 = latest.get("Return_60D_%")
    above50 = pd.notna(close) and pd.notna(ma50) and close >= ma50
    above200 = pd.notna(close) and pd.notna(ma200) and close >= ma200
    positive20 = pd.notna(ret20) and ret20 >= 0
    positive60 = pd.notna(ret60) and ret60 >= 0

    if above50 and above200 and positive20 and positive60:
        regime = "Risk-On"
        reason = "IHSG di atas MA50/MA200 dan momentum 20D/60D positif."
    elif (pd.notna(close) and pd.notna(ma200) and close < ma200) or ((not above50) and pd.notna(ret20) and ret20 < 0):
        regime = "Risk-Off"
        reason = "IHSG di bawah MA200 atau trend pendek melemah."
    else:
        regime = "Neutral"
        reason = "Sinyal IHSG campuran; gunakan konfirmasi tambahan."

    return {
        "Market_Regime": regime,
        "IHSG_Close": close,
        "IHSG_MA50": ma50,
        "IHSG_MA200": ma200,
        "IHSG_Return_20D_%": ret20,
        "IHSG_Return_60D_%": ret60,
        "IHSG_Last_Date": latest.get("Date"),
        "Regime_Reason": reason,
        "Market_Source": source,
        "Market_Error": error,
    }


@st.cache_data(ttl=3600, show_spinner=False)
def build_market_breadth(codes, period="1y", limit=50):
    selected_codes = [str(code).strip().upper() for code in codes if str(code).strip()][:limit]
    if not selected_codes:
        return {
            "Breadth_Count": 0,
            "Above_MA50_%": np.nan,
            "Above_MA200_%": np.nan,
            "Breadth_Label": "Unknown",
            "Breadth_Source": "empty",
            "Breadth_Error": "Kode breadth kosong.",
        }

    history, error, source = fetch_yahoo_history(selected_codes, period=period)
    if history.empty:
        return {
            "Breadth_Count": 0,
            "Above_MA50_%": np.nan,
            "Above_MA200_%": np.nan,
            "Breadth_Label": "Unknown",
            "Breadth_Source": source,
            "Breadth_Error": error,
        }

    indicators = build_technical_indicators(history)
    latest = indicators.sort_values(["Kode", "Date"]).groupby("Kode", as_index=False).tail(1)
    close = pd.to_numeric(latest.get("Close", pd.Series(index=latest.index)), errors="coerce")
    ma50 = pd.to_numeric(latest.get("MA50", pd.Series(index=latest.index)), errors="coerce")
    ma200 = pd.to_numeric(latest.get("MA200", pd.Series(index=latest.index)), errors="coerce")
    valid50 = close.notna() & ma50.notna()
    valid200 = close.notna() & ma200.notna()
    above50 = (close[valid50] >= ma50[valid50]).mean() * 100 if valid50.any() else np.nan
    above200 = (close[valid200] >= ma200[valid200]).mean() * 100 if valid200.any() else np.nan

    if pd.notna(above50) and pd.notna(above200) and above50 >= 60 and above200 >= 55:
        label = "Healthy"
    elif pd.notna(above50) and pd.notna(above200) and (above50 < 40 or above200 < 40):
        label = "Weak"
    else:
        label = "Mixed"

    return {
        "Breadth_Count": int(latest["Kode"].nunique()),
        "Above_MA50_%": round(float(above50), 1) if pd.notna(above50) else np.nan,
        "Above_MA200_%": round(float(above200), 1) if pd.notna(above200) else np.nan,
        "Breadth_Label": label,
        "Breadth_Source": source,
        "Breadth_Error": error,
    }


def add_market_context_to_explainability(scored, market_context):
    output = scored.copy()
    regime = clean_text(market_context.get("Market_Regime"), "Unknown")
    breadth = clean_text(market_context.get("Breadth_Label"), "Unknown")
    if regime == "Risk-Off" or breadth == "Weak":
        append_note = "market risk-off"
    elif regime == "Neutral" or breadth == "Mixed":
        append_note = "konfirmasi market"
    else:
        append_note = ""
    if append_note:
        checklist = output.get("Action_Checklist", pd.Series("", index=output.index)).astype(str)
        output["Action_Checklist"] = np.where(
            checklist.str.strip().eq(""),
            append_note,
            checklist + ", " + append_note,
        )
    output["Market_Regime"] = regime
    output["Market_Breadth"] = breadth
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


def provider_period(period):
    return "1mo" if period == "2wk" else period


def trim_history_to_period(history, period):
    if history.empty or period in ["max", None]:
        return history
    start, _ = yahoo_period_to_dates(period)
    if start is None:
        return history
    output = history.copy()
    output["Date"] = pd.to_datetime(output["Date"], errors="coerce")
    return output[output["Date"].ge(start)].copy()


def yahoo_period_to_dates(period):
    end = pd.Timestamp.today().normalize()
    days_by_period = {
        "5d": 7,
        "2wk": 14,
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
        YFINANCE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        if hasattr(yf, "set_tz_cache_location"):
            yf.set_tz_cache_location(str(YFINANCE_CACHE_DIR))
    except Exception as exc:
        cached = read_history_cache(cleaned_codes, period)
        if not cached.empty:
            return cached, f"Library yfinance belum tersedia, memakai cache lokal. Detail: {exc}", "cache"
        return pd.DataFrame(), f"Library yfinance belum tersedia: {exc}", "empty"

    tickers = [f"{code}.JK" for code in cleaned_codes]
    error_messages = []
    removed_proxy = remove_blocking_proxy_env()
    download_period = provider_period(period)
    try:
        downloaded = yf.download(
            tickers=tickers,
            period=download_period,
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
    history = trim_history_to_period(history, period)
    write_history_cache(history, period)
    return history, None, "yfinance"


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_yahoo_symbol_history(symbols, period="2y"):
    cleaned_symbols = [str(symbol).strip().upper() for symbol in symbols if str(symbol).strip()]
    if not cleaned_symbols:
        return pd.DataFrame(), "Pilih minimal satu simbol.", "empty"

    try:
        import yfinance as yf
        YFINANCE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        if hasattr(yf, "set_tz_cache_location"):
            yf.set_tz_cache_location(str(YFINANCE_CACHE_DIR))
    except Exception as exc:
        return pd.DataFrame(), f"Library yfinance belum tersedia: {exc}", "empty"

    removed_proxy = remove_blocking_proxy_env()
    try:
        downloaded = yf.download(
            tickers=cleaned_symbols,
            period=provider_period(period),
            interval="1d",
            auto_adjust=False,
            progress=False,
            threads=True,
            group_by="ticker",
        )
    except Exception as exc:
        downloaded = pd.DataFrame()
        error = f"yfinance simbol gagal: {exc}"
    else:
        error = None
    finally:
        restore_proxy_env(removed_proxy)

    if downloaded.empty:
        return pd.DataFrame(), error or "Data simbol online kosong.", "empty"

    records = []
    if isinstance(downloaded.columns, pd.MultiIndex):
        for symbol in cleaned_symbols:
            if symbol not in downloaded.columns.get_level_values(0):
                continue
            symbol_data = downloaded[symbol].reset_index()
            if "Close" not in symbol_data.columns:
                continue
            for _, row in symbol_data.iterrows():
                close = row.get("Close")
                if pd.isna(close):
                    continue
                records.append(
                    {
                        "Date": row.get("Date"),
                        "Symbol": symbol,
                        "Open": row.get("Open", np.nan),
                        "High": row.get("High", np.nan),
                        "Low": row.get("Low", np.nan),
                        "Close": close,
                        "Volume": row.get("Volume", np.nan),
                    }
                )
    else:
        symbol_data = downloaded.reset_index()
        symbol = cleaned_symbols[0]
        for _, row in symbol_data.iterrows():
            close = row.get("Close")
            if pd.isna(close):
                continue
            records.append(
                {
                    "Date": row.get("Date"),
                    "Symbol": symbol,
                    "Open": row.get("Open", np.nan),
                    "High": row.get("High", np.nan),
                    "Low": row.get("Low", np.nan),
                    "Close": close,
                    "Volume": row.get("Volume", np.nan),
                }
            )

    history = pd.DataFrame(records)
    if history.empty:
        return history, "Data simbol online tidak menemukan harga penutupan.", "empty"
    history["Date"] = pd.to_datetime(history["Date"], errors="coerce")
    for column in ["Open", "High", "Low", "Close", "Volume"]:
        history[column] = pd.to_numeric(history[column], errors="coerce")
    history = history.dropna(subset=["Date", "Close"]).sort_values(["Symbol", "Date"])
    history = trim_history_to_period(history, period)
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


def compute_rsi(close, window=14):
    delta = close.diff()
    gain = delta.clip(lower=0).ewm(alpha=1 / window, adjust=False, min_periods=window).mean()
    loss = (-delta.clip(upper=0)).ewm(alpha=1 / window, adjust=False, min_periods=window).mean()
    rs = gain / loss.replace(0, np.nan)
    rsi = 100 - (100 / (1 + rs))
    rsi = rsi.mask(loss.eq(0) & gain.gt(0), 100)
    rsi = rsi.mask(loss.eq(0) & gain.eq(0), 50)
    return rsi.clip(0, 100)


def build_technical_indicators(history):
    if history.empty:
        return pd.DataFrame()

    frames = []
    for code, group in normalize_history_frame(history).groupby("Kode"):
        tech = group.sort_values("Date").copy()
        close = tech["Close"]
        high = pd.to_numeric(tech.get("High", close), errors="coerce").fillna(close)
        low = pd.to_numeric(tech.get("Low", close), errors="coerce").fillna(close)
        volume = pd.to_numeric(tech.get("Volume_Online", pd.Series(np.nan, index=tech.index)), errors="coerce")

        tech["MA20"] = close.rolling(20, min_periods=5).mean()
        tech["MA50"] = close.rolling(50, min_periods=10).mean()
        tech["MA200"] = close.rolling(200, min_periods=40).mean()
        tech["RSI14"] = compute_rsi(close)
        ema12 = close.ewm(span=12, adjust=False, min_periods=12).mean()
        ema26 = close.ewm(span=26, adjust=False, min_periods=26).mean()
        tech["MACD"] = ema12 - ema26
        tech["MACD_Signal"] = tech["MACD"].ewm(span=9, adjust=False, min_periods=9).mean()
        tech["BB_Mid"] = tech["MA20"]
        bb_std = close.rolling(20, min_periods=10).std()
        tech["BB_Upper"] = tech["BB_Mid"] + 2 * bb_std
        tech["BB_Lower"] = tech["BB_Mid"] - 2 * bb_std
        prev_close = close.shift(1)
        true_range = pd.concat([(high - low).abs(), (high - prev_close).abs(), (low - prev_close).abs()], axis=1).max(axis=1)
        tech["ATR14"] = true_range.rolling(14, min_periods=5).mean()
        tech["Volume_MA20"] = volume.rolling(20, min_periods=5).mean()
        tech["Volume_Ratio"] = volume / tech["Volume_MA20"].replace(0, np.nan)
        tech["High_52W"] = high.rolling(252, min_periods=40).max()
        tech["Low_52W"] = low.rolling(252, min_periods=40).min()
        tech["Distance_52W_High_%"] = (close / tech["High_52W"] - 1) * 100
        tech["Distance_52W_Low_%"] = (close / tech["Low_52W"] - 1) * 100
        tech["ATR_%"] = tech["ATR14"] / close.replace(0, np.nan) * 100
        tech["Return_20D_%"] = close.pct_change(20) * 100
        tech["Return_60D_%"] = close.pct_change(60) * 100
        tech["Trend_Bullish"] = (close > tech["MA50"]) & (tech["MA50"] > tech["MA200"])
        tech["MACD_Bullish"] = tech["MACD"] > tech["MACD_Signal"]
        tech["Volume_Confirm"] = tech["Volume_Ratio"] >= 1
        frames.append(tech)

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def summarize_technical_indicators(technical_history):
    if technical_history.empty:
        return pd.DataFrame()

    latest = technical_history.sort_values(["Kode", "Date"]).groupby("Kode", as_index=False).tail(1).copy()
    latest["Trend_Score"] = np.select(
        [
            latest["Close"].gt(latest["MA50"]) & latest["MA50"].gt(latest["MA200"]),
            latest["Close"].gt(latest["MA50"]),
            latest["Close"].lt(latest["MA200"]),
        ],
        [100, 70, 25],
        default=50,
    )
    latest["RSI_Score"] = score_target_range(latest["RSI14"], low=35, high=75, center=55)
    latest["MACD_Score"] = np.where(latest["MACD_Bullish"], 80, 35)
    latest["Volume_Score"] = score_target_range(latest["Volume_Ratio"], low=0.5, high=2.0, center=1.2)
    latest["Volatility_Score"] = (100 - (latest["ATR_%"].fillna(6) / 12 * 100)).clip(0, 100)
    latest["Technical_Score"] = (
        latest["Trend_Score"] * 0.30
        + latest["RSI_Score"] * 0.20
        + latest["MACD_Score"] * 0.20
        + latest["Volume_Score"] * 0.15
        + latest["Volatility_Score"] * 0.15
    ).round(1)
    latest["Technical_Signal"] = np.select(
        [
            latest["RSI14"].ge(75),
            latest["Technical_Score"].ge(75) & latest["Trend_Bullish"],
            latest["Technical_Score"].ge(60),
            latest["Technical_Score"].lt(45),
        ],
        ["Overbought", "Bullish", "Constructive", "Weak"],
        default="Neutral",
    )
    latest["Technical_Notes"] = latest.apply(
        lambda row: ", ".join(
            note
            for note, ok in [
                ("trend bullish", bool(row.get("Trend_Bullish"))),
                ("MACD bullish", bool(row.get("MACD_Bullish"))),
                ("volume confirm", bool(row.get("Volume_Confirm"))),
                ("RSI tinggi", pd.notna(row.get("RSI14")) and row.get("RSI14") >= 70),
                ("volatilitas tinggi", pd.notna(row.get("ATR_%")) and row.get("ATR_%") >= 8),
            ]
            if ok
        )
        or "netral",
        axis=1,
    )
    return latest


def build_entry_decision(row):
    score = pd.to_numeric(row.get("Score"), errors="coerce")
    threshold = pd.to_numeric(row.get("Threshold_Pass_Ratio"), errors="coerce")
    technical_score = pd.to_numeric(row.get("Technical_Score"), errors="coerce")
    signal = clean_text(row.get("Technical_Signal"), "Neutral")
    recommendation = clean_text(row.get("Recommendation"), "Avoid")
    risk_level = clean_text(row.get("Risk_Level"), "High")
    clean_data = bool(row.get("Clean_Data", False))

    fundamental_ok = (
        pd.notna(score)
        and score >= 68
        and recommendation in ["Strong Buy", "Buy"]
        and risk_level != "High"
        and clean_data
        and (pd.isna(threshold) or threshold >= 55)
    )
    fundamental_watch = pd.notna(score) and score >= 55 and recommendation in ["Strong Buy", "Buy", "Watchlist"] and risk_level != "High"

    if not fundamental_watch:
        action = "Avoid Entry"
        combined = "Fundamental belum cukup kuat"
        reason = "Prioritas rendah: score/rekomendasi/risiko/kualitas data belum memenuhi batas watchlist."
    elif not fundamental_ok:
        action = "Hold/Monitor"
        combined = "Fundamental watchlist"
        reason = "Fundamental cukup untuk dipantau, tetapi belum memenuhi kriteria kandidat entry yang bersih."
    elif signal == "Bullish" and technical_score >= 75:
        action = "Buy Candidate"
        combined = "Fundamental kuat + teknikal bullish"
        reason = "Fundamental lolos dan trend teknikal mendukung entry bertahap."
    elif signal == "Constructive" and technical_score >= 60:
        action = "Wait Confirmation"
        combined = "Fundamental kuat + teknikal membaik"
        reason = "Fundamental kuat, tetapi butuh konfirmasi lanjutan dari breakout, volume, atau MACD."
    elif signal == "Overbought":
        action = "Take Profit / Tight Stop"
        combined = "Fundamental kuat + harga panas"
        reason = "Saham layak, tetapi RSI/kenaikan teknikal sudah panas. Hindari FOMO, tunggu pullback atau gunakan stop ketat."
    elif signal == "Weak":
        action = "Wait Pullback"
        combined = "Fundamental kuat + teknikal lemah"
        reason = "Fundamental baik, tetapi timing belum mendukung. Tunggu reversal atau harga kembali di atas MA kunci."
    else:
        action = "Wait Confirmation"
        combined = "Fundamental kuat + teknikal netral"
        reason = "Fundamental lolos, tetapi sinyal teknikal belum cukup tegas."

    return pd.Series({"Entry_Action": action, "Combined_View": combined, "Timing_Reason": reason})


def build_position_decision(row):
    score = pd.to_numeric(row.get("Score"), errors="coerce")
    technical_score = pd.to_numeric(row.get("Technical_Score"), errors="coerce")
    rsi = pd.to_numeric(row.get("RSI14"), errors="coerce")
    atr_pct = pd.to_numeric(row.get("ATR_%"), errors="coerce")
    close = pd.to_numeric(row.get("Close"), errors="coerce")
    ma50 = pd.to_numeric(row.get("MA50"), errors="coerce")
    ma200 = pd.to_numeric(row.get("MA200"), errors="coerce")
    signal = clean_text(row.get("Technical_Signal"), "Neutral")
    recommendation = clean_text(row.get("Recommendation"), "Avoid")
    risk_level = clean_text(row.get("Risk_Level"), "High")
    clean_data = bool(row.get("Clean_Data", False))
    trend_bullish = bool(row.get("Trend_Bullish", False))
    macd_bullish = bool(row.get("MACD_Bullish", False))

    fundamental_strong = pd.notna(score) and score >= 68 and recommendation in ["Strong Buy", "Buy"] and risk_level != "High" and clean_data
    fundamental_weak = pd.isna(score) or score < 55 or recommendation in ["Speculative", "Avoid"] or risk_level == "High"
    breakdown_ma50 = pd.notna(close) and pd.notna(ma50) and close < ma50
    breakdown_ma200 = pd.notna(close) and pd.notna(ma200) and close < ma200
    high_volatility = pd.notna(atr_pct) and atr_pct >= 8

    if fundamental_weak and (signal == "Weak" or breakdown_ma200):
        action = "Exit / Sell"
        exit_risk = "High"
        reason = "Fundamental tidak mendukung dan teknikal lemah/breakdown. Prioritaskan keluar atau hindari mempertahankan posisi."
    elif fundamental_weak:
        action = "Reduce"
        exit_risk = "High"
        reason = "Fundamental lemah untuk posisi inti. Kurangi eksposur, terutama bila teknikal tidak segera membaik."
    elif signal == "Overbought" and pd.notna(rsi) and rsi >= 80:
        action = "Take Profit"
        exit_risk = "Medium"
        reason = "Fundamental masih layak, tetapi RSI sangat panas. Amankan sebagian profit dan tunggu pullback."
    elif signal == "Overbought":
        action = "Tight Stop"
        exit_risk = "Medium"
        reason = "Momentum kuat tetapi mulai panas. Pertahankan hanya dengan trailing stop atau batas risiko ketat."
    elif fundamental_strong and trend_bullish and macd_bullish and technical_score >= 70:
        action = "Hold"
        exit_risk = "Low"
        reason = "Fundamental kuat dan trend teknikal masih mendukung. Pertahankan posisi selama tidak breakdown."
    elif fundamental_strong and signal == "Weak":
        action = "Review Position"
        exit_risk = "Medium"
        reason = "Fundamental baik tetapi teknikal melemah. Evaluasi posisi dan tunggu pemulihan di atas MA kunci."
    elif fundamental_strong and breakdown_ma50:
        action = "Tight Stop"
        exit_risk = "Medium"
        reason = "Fundamental baik, tetapi harga di bawah MA50. Perketat stop sampai trend pulih."
    elif fundamental_strong and signal == "Constructive":
        action = "Add on Pullback"
        exit_risk = "Low"
        reason = "Fundamental baik dan teknikal membaik. Tambahan posisi lebih ideal saat pullback sehat, bukan mengejar harga."
    elif high_volatility:
        action = "Review Position"
        exit_risk = "Medium"
        reason = "Volatilitas tinggi. Tinjau ukuran posisi meskipun sinyal utama belum bearish."
    else:
        action = "Review Position"
        exit_risk = "Medium"
        reason = "Sinyal campuran. Pertahankan secara selektif sambil menunggu konfirmasi trend."

    return pd.Series({"Position_Action": action, "Exit_Risk": exit_risk, "Position_Reason": reason})


def add_entry_decision(technical_summary, scored):
    if technical_summary.empty:
        return technical_summary
    fundamental_columns = [
        "Kode",
        "Nama Perusahaan",
        "Sektor",
        "Score",
        "Recommendation",
        "Risk_Level",
        "Clean_Data",
        "Threshold_Pass_Ratio",
    ]
    available_columns = [column for column in fundamental_columns if column in scored.columns]
    output = technical_summary.merge(scored[available_columns], on="Kode", how="left")
    entry_decision = output.apply(build_entry_decision, axis=1)
    position_decision = output.apply(build_position_decision, axis=1)
    output = pd.concat([output, entry_decision, position_decision], axis=1)
    return add_trade_risk_levels(output)


def add_trade_risk_levels(technical_decision):
    if technical_decision.empty:
        return technical_decision
    output = technical_decision.copy()
    close = pd.to_numeric(output.get("Close", pd.Series(np.nan, index=output.index)), errors="coerce")
    atr = pd.to_numeric(output.get("ATR14", pd.Series(np.nan, index=output.index)), errors="coerce")
    atr_pct = pd.to_numeric(output.get("ATR_%", pd.Series(np.nan, index=output.index)), errors="coerce")
    output["ATR_Stop_2x"] = (close - 2 * atr).where(close.gt(0) & atr.gt(0))
    output["ATR_Stop_Distance_%"] = (2 * atr / close.replace(0, np.nan) * 100).where(close.gt(0) & atr.gt(0))
    output["Position_Risk_Bucket"] = np.select(
        [
            atr_pct.ge(8),
            atr_pct.ge(5),
            atr_pct.lt(3),
        ],
        ["High volatility", "Medium volatility", "Low volatility"],
        default="Normal volatility",
    )
    output["Risk_Note"] = np.where(
        output["ATR_Stop_2x"].notna(),
        "2x ATR stop sebagai zona risiko teknikal, bukan instruksi order otomatis.",
        "ATR belum cukup untuk menghitung zona risiko.",
    )
    return output


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
            online_values = pd.to_numeric(df[online_column], errors="coerce")
            has_online_value = online_values.notna()
            df[column] = pd.to_numeric(df[column], errors="coerce")
            df.loc[has_online_value, column] = online_values[has_online_value]

    online_fundamental_fields = [column for column in numeric_online_columns.values() if column in df.columns]
    if online_fundamental_fields:
        df["Online_Fundamental_Field_Count"] = df[online_fundamental_fields].notna().sum(axis=1)
    else:
        df["Online_Fundamental_Field_Count"] = 0

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
    excel_fundamental_fill_count = pd.Series(0, index=df.index, dtype="int64")
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
                excel_fundamental_fill_count = excel_fundamental_fill_count + fill_from_excel.astype("int64")
            df[column] = df[column].fillna(excel_values)
            df = df.drop(columns=[excel_column])
    fundamental_metric_columns = [
        column
        for column in ["PER", "PBV", "ROE", "ROA", "DER", "NPM", "NIM", "CAR", "LDR", "NPL", "BOPO", "CIR", "LAR"]
        if column in df.columns
    ]
    final_fundamental_field_count = df[fundamental_metric_columns].notna().sum(axis=1) if fundamental_metric_columns else pd.Series(0, index=df.index)
    df["Excel_Fundamental_Field_Count"] = (
        final_fundamental_field_count
        - pd.to_numeric(df.get("Online_Fundamental_Field_Count", pd.Series(0, index=df.index)), errors="coerce").fillna(0)
    ).clip(lower=0)

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
st.caption(f"Build UI: `{APP_BUILD}`. Histori dan teknikal digabung di tab **Harga & Teknikal**.")
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
        - **Harga & Teknikal**: grafik return dari yfinance online, mode Excel Metrik sebagai pembanding/cadangan, ringkasan teknikal top kandidat, candlestick/line, MA20/50/200, RSI, MACD, ATR, technical score, entry action, dan position action dari OHLCV yfinance/cache.
        - **Sektor**: ringkasan score, jumlah saham, Strong Buy, ROE, dan turnover per sektor/industri.
        - **Kualitas Data**: audit data, cache histori, kelengkapan rasio, dan catatan kualitas data.
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
        - **Technical Score / Signal**: konfirmasi timing berbasis MA, RSI, MACD, volume, ATR, dan tren harga. Ini tidak mengubah Score fundamental.
        - **Entry Action**: arahan untuk calon pembelian. Fundamental menjadi gerbang awal, teknikal menentukan timing entry/tunggu/hindari.
        - **Position Action**: arahan umum untuk saham yang sudah dimiliki tanpa memakai harga beli pribadi, misalnya Hold, Take Profit, Reduce, atau Exit / Sell.
        - **Exit Risk**: risiko keluar/pengetatan posisi berdasarkan kombinasi fundamental dan teknikal, bukan perhitungan profit pribadi.
        - **Sector Relative Score**: perbandingan valuasi dan kualitas terhadap saham lain dalam sektor yang sama.
        - **Decision Summary / Top Strengths / Top Risks**: ringkasan alasan kuantitatif agar ranking mudah diaudit.
        - **ATR Stop 2x**: zona risiko teknikal berbasis volatilitas ATR, bukan instruksi order otomatis.
        - **Clean_Data**: penanda bahwa data dan rasio utama lolos filter kebersihan minimum.
        - **Safety_Recommendation**: ringkasan kelayakan data seperti `Bersih - Strong`; di kartu utama ditampilkan sebagai `Data`, bukan jaminan aman investasi.
        - **Safety_Notes**: alasan saham perlu direview, misalnya volume rendah, rasio kosong, threshold rendah, atau risiko tinggi.
        - **Risk Level**: estimasi risiko relatif berdasarkan rasio, volatilitas, penalti, dan likuiditas, bukan jaminan keamanan.
        - **KODE.JK**: format ticker saham Indonesia di Yahoo Finance/yfinance, misalnya `BBCA.JK`.
        - **All / top N**: grafik histori untuk saham teratas dari hasil filter/ranking saat ini.

        **Rumus ringkas**
        - `Score = weighted average(Valuation, Quality, Risk, Liquidity, Momentum, Index) - Penalty`, lalu dibatasi 0-100.
        - `Technical_Score = Trend 30% + RSI 20% + MACD 20% + Volume 15% + Volatilitas 15%`.
        - `Sector_Relative_Score = Valuasi relatif sektor 45% + Kualitas relatif sektor 40% + Likuiditas 15%`.
        - `Entry_Action`: fundamental memilih saham layak, teknikal menentukan timing untuk calon pembelian.
        - `Position_Action`: teknikal dan fundamental memberi arahan hold/reduce/take profit/exit untuk saham yang sudah dimiliki.
        - `ATR_Stop_2x = Close - 2 x ATR14`.
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

        **Makna warna**
        - Hijau: kondisi lebih kuat atau risiko rendah.
        - Kuning/oranye: area watchlist, medium, atau perlu dicermati.
        - Merah: avoid, risiko tinggi, atau kondisi paling lemah.
        - Biru/ungu: sumber data, seri saham, likuiditas, momentum, atau dimensi netral/non-keputusan.
        - Grafik jumlah/coverage memakai skala netral; grafik score memakai skala 0-100 agar tidak tercampur dengan warna kategori.
        """
    )

with st.sidebar:
    st.header("Filter & Strategi")
    ui_mode = st.radio(
        "Mode filter",
        ["Cepat", "Lengkap"],
        horizontal=True,
        help="Cepat menampilkan kontrol utama saja. Lengkap membuka rasio, threshold, dan bobot untuk analisis detail.",
    )
    advanced_mode = ui_mode == "Lengkap"
    profile = st.selectbox("Profil scoring", list(PROFILE_WEIGHTS), index=0, help=HELP_TEXT["profile"])
    weights = PROFILE_WEIGHTS[profile].copy()
    filter_preset = st.selectbox("Preset filter", ["Normal", "Konservatif Aman"], index=0, help=HELP_TEXT["filter_preset"])
    safe_preset = filter_preset == "Konservatif Aman"
    if safe_preset:
        st.info("Preset filter konservatif aktif: filter dibuat lebih ketat. Bobot Score tetap mengikuti Profil scoring yang dipilih.")

    with st.expander("Sesuaikan bobot", expanded=advanced_mode):
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
    price_range = safe_slider("Harga penutupan", price_min, price_max, default_price_range, help=HELP_TEXT["price"])
    min_volume = st.select_slider(
        "Minimum volume",
        options=[0, 1_000_000, 5_000_000, 10_000_000, 50_000_000, 100_000_000, 500_000_000],
        value=10_000_000 if safe_preset else 5_000_000,
        format_func=format_volume,
        help=HELP_TEXT["volume"],
    )

    min_score = st.slider("Score minimum", 0, 100, 60 if safe_preset else 45, help=HELP_TEXT["score"])
    clean_data_only = st.checkbox("Data bersih saja", value=safe_preset, help=HELP_TEXT["clean_data"])

    per_range = (0.1, 25.0) if safe_preset else (0.0, 35.0)
    pbv_max = 3.5 if safe_preset else 5.0
    roe_min = 8.0 if safe_preset else 5.0
    npm_min = 3.0 if safe_preset else 0.0
    der_max = 1.5 if safe_preset else 2.5
    apply_der_to_banking = False
    threshold_source = "Auto: Banking untuk bank, NonBank untuk lainnya"
    min_threshold_ratio = 65 if safe_preset else 50
    require_core_thresholds = safe_preset

    with st.expander("Filter rasio & threshold", expanded=advanced_mode):
        per_range = st.slider("PER", 0.0, 80.0, per_range, step=0.5, help=HELP_TEXT["per"])
        pbv_max = st.slider("PBV maksimum", 0.0, 15.0, pbv_max, step=0.1, help=HELP_TEXT["pbv"])
        roe_min = st.slider("ROE minimum (%)", -50.0, 100.0, roe_min, step=0.5, help=HELP_TEXT["roe"])
        npm_min = st.slider("NPM minimum (%)", -50.0, 100.0, npm_min, step=0.5, help=HELP_TEXT["npm"])
        der_max = st.slider("DER maksimum", 0.0, 8.0, der_max, step=0.1, help=HELP_TEXT["der"])
        apply_der_to_banking = st.checkbox("Terapkan DER juga ke Banking", value=False, help=HELP_TEXT["der_banking"])
        st.divider()
        threshold_source = st.selectbox("Sumber threshold", ["Auto: Banking untuk bank, NonBank untuk lainnya", "NonBank", "Banking"], help=HELP_TEXT["threshold_source"])
        min_threshold_ratio = st.slider("Minimum lolos threshold (%)", 0, 100, min_threshold_ratio, step=5, help=HELP_TEXT["threshold_ratio"])
        require_core_thresholds = st.checkbox("Wajib lolos valuasi & profit inti", value=require_core_thresholds, help=HELP_TEXT["core_thresholds"])

    st.divider()
    with st.expander("Workflow Update", expanded=False):
        file_status = data_file_status
        st.caption(f"Sumber aktif: {data_update_label}")
        st.caption(f"Excel fallback: {file_status['Status']} | Modified: {file_status['Last Modified']} | Size: {file_status['Ukuran']}")
        cache_status_sidebar = get_history_cache_status()
        st.caption(f"Cache histori: {len(cache_status_sidebar):,} file di `{HISTORY_CACHE_DIR}`")
        refresh_period = st.selectbox(
            "Periode refresh cache",
            ONLINE_PERIOD_OPTIONS,
            index=ONLINE_PERIOD_OPTIONS.index("1y"),
            format_func=lambda value: ONLINE_PERIOD_LABELS.get(value, value),
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
scored_df = add_sector_relative_scores(scored_df)
scored_df = add_decision_explainability(scored_df)
market_regime = build_market_regime(period="2y")
breadth_codes = (
    scored_df.sort_values(["Index_Count", "Volume"], ascending=False)["Kode"]
    .dropna()
    .astype(str)
    .head(50)
    .tolist()
)
market_breadth = build_market_breadth(tuple(breadth_codes), period="1y", limit=50)
market_context = {**market_regime, **market_breadth}
scored_df = add_market_context_to_explainability(scored_df, market_context)
data_freshness = build_data_freshness(scored_df)
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

market_cols = st.columns(4)
market_cols[0].metric(
    "Market Regime",
    clean_text(market_context.get("Market_Regime")),
    clean_text(market_context.get("Breadth_Label")),
    help="Konteks IHSG dan market breadth. Tidak mengubah Score, tetapi memengaruhi checklist risiko entry.",
)
market_cols[1].metric(
    "IHSG 20D / 60D",
    f"{format_percent(market_context.get('IHSG_Return_20D_%'))}",
    f"60D {format_percent(market_context.get('IHSG_Return_60D_%'))}",
)
market_cols[2].metric(
    "Breadth MA50 / MA200",
    f"{format_percent(market_context.get('Above_MA50_%'))}",
    f"MA200 {format_percent(market_context.get('Above_MA200_%'))}",
)
market_cols[3].metric(
    "Freshness",
    clean_text(data_freshness.get("Freshness_Label")),
    f"Lag {data_freshness.get('Online_Data_Lag_Days') if pd.notna(data_freshness.get('Online_Data_Lag_Days')) else '-'} hari",
)
if market_context.get("Market_Error"):
    st.warning(f"Market regime memakai fallback/terbatas. Detail: {market_context.get('Market_Error')}")
if market_context.get("Breadth_Error"):
    st.caption(f"Market breadth terbatas: {market_context.get('Breadth_Error')}")

tab_summary, tab_reco, tab_history, tab_explore, tab_sector, tab_quality, tab_method = st.tabs(
    ["Ringkasan", "Rekomendasi", "Harga & Teknikal", "Explorer", "Sektor", "Kualitas Data", "Metodologi"]
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
        summary_chart_data = chart_market_frame(summary_data, "Ringkasan grafik utama")
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

        strong_buy_count = int(summary_chart_data["Recommendation"].eq("Strong Buy").sum())
        buy_or_better_count = int(summary_chart_data["Recommendation"].isin(["Strong Buy", "Buy"]).sum())
        low_risk_count = int(summary_chart_data["Risk_Level"].eq("Low").sum())
        top_candidates = summary_chart_data.sort_values(["Score", "Threshold_Pass_Ratio", "Liquidity_Score"], ascending=False).head(3)
        top_codes = ", ".join(top_candidates["Kode"].tolist()) if not top_candidates.empty else "-"
        source_ratio = online_price / max(len(summary_data), 1) * 100

        with st.container(border=True):
            st.markdown("**Insight Utama**")
            insight_cols = st.columns(4)
            insight_cols[0].metric("Layak dicermati", f"{buy_or_better_count:,}", f"{strong_buy_count:,} Strong Buy")
            insight_cols[1].metric("Risiko rendah", f"{low_risk_count:,}")
            insight_cols[2].metric("Cakupan online", f"{source_ratio:.0f}%")
            insight_cols[3].metric("Top 3", top_codes)
            if top_candidates.empty:
                st.caption("Belum ada kandidat kuat pada filter ini. Longgarkan filter atau pilih cakupan Semua universe.")
            else:
                st.caption("Gunakan insight ini sebagai daftar awal; validasi detail tetap ada di tab Rekomendasi dan Kualitas Data.")

        with st.container(border=True):
            st.markdown("**Konteks market & freshness**")
            regime_cols = st.columns(5)
            regime_cols[0].metric("Regime", clean_text(market_context.get("Market_Regime")), clean_text(market_context.get("Breadth_Label")))
            regime_cols[1].metric("IHSG", format_number(market_context.get("IHSG_Close")), f"MA200 {format_number(market_context.get('IHSG_MA200'))}")
            regime_cols[2].metric("Return IHSG 20D", format_percent(market_context.get("IHSG_Return_20D_%")), f"60D {format_percent(market_context.get('IHSG_Return_60D_%'))}")
            regime_cols[3].metric("Breadth MA50", format_percent(market_context.get("Above_MA50_%")), f"MA200 {format_percent(market_context.get('Above_MA200_%'))}")
            regime_cols[4].metric("Freshness", clean_text(data_freshness.get("Freshness_Label")), f"Online harga {format_percent(data_freshness.get('Online_Price_Coverage_%'), 0)}")
            st.caption(clean_text(market_context.get("Regime_Reason"), "Konteks market belum tersedia."))

        chart_cols = st.columns([1, 1, 1])
        with chart_cols[0]:
            reco_counts = summary_chart_data["Recommendation"].value_counts().reindex(["Strong Buy", "Buy", "Watchlist", "Speculative", "Avoid"]).dropna().reset_index()
            reco_counts.columns = ["Recommendation", "Jumlah"]
            fig = px.bar(
                reco_counts,
                x="Recommendation",
                y="Jumlah",
                color="Recommendation",
                title="Distribusi rekomendasi",
                color_discrete_map=RECOMMENDATION_COLORS,
            )
            fig.update_layout(height=330, showlegend=False, margin=dict(l=20, r=20, t=60, b=40))
            show_chart(fig)
        with chart_cols[1]:
            risk_counts = summary_chart_data["Risk_Level"].value_counts().reindex(["Low", "Medium", "High"]).dropna().reset_index()
            risk_counts.columns = ["Risk_Level", "Jumlah"]
            fig = px.pie(
                risk_counts,
                names="Risk_Level",
                values="Jumlah",
                hole=0.45,
                title="Komposisi risiko",
                color="Risk_Level",
                color_discrete_map=RISK_COLORS,
            )
            fig.update_layout(height=330, margin=dict(l=20, r=20, t=60, b=40))
            show_chart(fig)
        with chart_cols[2]:
            source_mix = build_source_mix(summary_chart_data)
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
                    color_discrete_map=SOURCE_COLORS,
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
                "Decision_Summary",
                "Top_Strengths",
                "Top_Risks",
                "Score",
                "Sector_Relative_Score",
                "Sector_Relative_Label",
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
            top_summary = summary_chart_data.sort_values(["Score", "Threshold_Pass_Ratio", "Liquidity_Score"], ascending=False).head(15)
            show_table(
                top_summary[[column for column in top_summary_columns if column in top_summary.columns]],
                hide_index=True,
                column_config={
                    "Score": st.column_config.ProgressColumn("Score", min_value=0, max_value=100, format="%.1f", help=HELP_TEXT["score"]),
                    "Sector_Relative_Score": st.column_config.ProgressColumn("Relatif Sektor", min_value=0, max_value=100, format="%.1f", help=HELP_TEXT["sector_relative"]),
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
            factor_matrix = summary_chart_data.sort_values("Score", ascending=False).head(12).set_index("Kode")[
                [column for column in factor_columns if column in summary_chart_data.columns]
            ]
            if factor_matrix.empty:
                st.info("Matriks faktor belum tersedia.")
            else:
                fig = px.imshow(
                    factor_matrix,
                    aspect="auto",
                    text_auto=".0f",
                    color_continuous_scale=SCORE_SCALE,
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
            reco_max = max(1, min(100, len(filtered)))
            reco_min = min(5, reco_max)
            reco_default = min(25, reco_max)
            reco_step = 5 if reco_max >= 10 else 1
            reco_limit = safe_slider("Jumlah tampil", reco_min, reco_max, reco_default, step=reco_step, help=HELP_TEXT["reco_limit"])
        with reco_controls[1]:
            reco_sort = st.selectbox(
                "Urutkan berdasarkan",
                ["Score", "Sector_Relative_Score", "Threshold_Pass_Ratio", "Valuation_Score", "Quality_Score", "Risk_Score", "Liquidity_Score", "Momentum_Score", "Return_52W", "Volume", "Turnover"],
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

        reco_chart_view = chart_market_frame(reco_view, "Grafik rekomendasi")

        left, right = st.columns([1.25, 1])
        with left:
            chart_data = prepare_chart_frame(reco_chart_view.sort_values(reco_sort), reco_sort)
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
                    color_discrete_map=RECOMMENDATION_COLORS,
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
                "Sector_Relative_Score",
            ]
            radar_base = reco_chart_view.head(5)
            fig = go.Figure()
            for index, (_, row) in enumerate(radar_base.iterrows()):
                values = [row[col] for col in component_cols]
                color = STOCK_LINE_COLORS[index % len(STOCK_LINE_COLORS)]
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
                            "Relatif Sektor",
                            "Valuasi",
                        ],
                        fill="toself",
                        name=row["Kode"],
                        line=dict(color=color),
                        fillcolor=color,
                        opacity=0.32,
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
            "Decision_Summary",
            "Top_Strengths",
            "Top_Risks",
            "Action_Checklist",
            "Market_Regime",
            "Market_Breadth",
            "Safety_Notes",
            "Score",
            "Sector_Relative_Score",
            "Sector_Relative_Label",
            "Sector_Valuation_Score",
            "Sector_Quality_Score",
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
            "Online_Fundamental_Field_Count",
            "Excel_Fundamental_Field_Count",
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
            "Risk_Level",
            "Clean_Data",
            "Decision_Summary",
            "Top_Strengths",
            "Top_Risks",
            "Market_Regime",
            "Score",
            "Sector_Relative_Score",
            "Threshold_Pass_Ratio",
            "Penutupan",
            "PER",
            "PBV",
            "ROE",
            "DER",
            "Return_52W",
            "Volume",
            "Price_Source",
            "Fundamental_Source",
            "Safety_Notes",
            "Sektor",
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
            table,
            hide_index=True,
            column_config={
                "Score": st.column_config.ProgressColumn("Score", min_value=0, max_value=100, format="%.1f", help=HELP_TEXT["score"]),
                "Sector_Relative_Score": st.column_config.ProgressColumn("Relatif Sektor", min_value=0, max_value=100, format="%.1f", help=HELP_TEXT["sector_relative"]),
                "Sector_Valuation_Score": st.column_config.NumberColumn("Valuasi Sektor", format="%.1f", help=HELP_TEXT["sector_relative"]),
                "Sector_Quality_Score": st.column_config.NumberColumn("Kualitas Sektor", format="%.1f", help=HELP_TEXT["sector_relative"]),
                "Decision_Summary": st.column_config.TextColumn("Ringkasan Keputusan", help=HELP_TEXT["explainability"]),
                "Top_Strengths": st.column_config.TextColumn("Faktor Kuat", help=HELP_TEXT["explainability"]),
                "Top_Risks": st.column_config.TextColumn("Faktor Lemah", help=HELP_TEXT["explainability"]),
                "Action_Checklist": st.column_config.TextColumn("Checklist", help=HELP_TEXT["explainability"]),
                "Market_Regime": st.column_config.TextColumn("Market Regime", help="Konteks IHSG saat ini. Tidak mengubah Score, tetapi menambah checklist risiko."),
                "Market_Breadth": st.column_config.TextColumn("Market Breadth", help="Kesehatan pasar dari persentase saham di atas MA50/MA200 pada sample breadth."),
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
                "Online_Fundamental_Field_Count": st.column_config.NumberColumn("Field Online", format="%.0f", help="Jumlah rasio fundamental utama yang tersedia dari TradingView scanner online."),
                "Excel_Fundamental_Field_Count": st.column_config.NumberColumn("Field Excel", format="%.0f", help="Jumlah rasio fundamental yang masih diisi dari Excel fallback karena sumber online kosong."),
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
    explorer_chart_data = chart_market_frame(explorer_data, "Explorer")
    explore_controls = st.columns([1, 1, 1, 1])
    with explore_controls[0]:
        explore_x = st.selectbox("Sumbu X", ANALYSIS_COLUMNS, index=ANALYSIS_COLUMNS.index("PER"), help=HELP_TEXT["explorer_axis"])
    with explore_controls[1]:
        explore_y = st.selectbox("Sumbu Y", ANALYSIS_COLUMNS, index=ANALYSIS_COLUMNS.index("ROE"), help=HELP_TEXT["explorer_axis"])
    with explore_controls[2]:
        explore_color = st.selectbox("Warna", ["Score", "Quality_Score", "Risk_Score", "Threshold_Pass_Ratio", "Recommendation", "Risk_Level", "Sektor"], help=HELP_TEXT["explore_color"])
    with explore_controls[3]:
        explore_size = st.selectbox("Ukuran bubble", ["Volume", "Turnover", "Score", "Liquidity_Score", "Index_Count"], help=HELP_TEXT["explore_size"])

    explore_max = max(1, min(500, len(explorer_chart_data)))
    explore_min = min(50, explore_max)
    explore_default = min(250, explore_max)
    explore_step = 25 if explore_max >= 50 else 1
    explore_limit = safe_slider("Jumlah titik Explorer", explore_min, explore_max, explore_default, step=explore_step, help=HELP_TEXT["explore_limit"])
    explore_plot = explorer_chart_data.sort_values("Score", ascending=False).head(explore_limit)

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
            **chart_color_kwargs(explore_color),
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
            **chart_color_kwargs("Quality_Score"),
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
    for index, (column, container) in enumerate(zip(histogram_columns, hist_cols)):
        with container:
            fig = px.histogram(
                explore_plot,
                x=column,
                nbins=35,
                title=f"Distribusi {column}",
                color_discrete_sequence=[FACTOR_COLORS.get(column, "#2563eb")],
            )
            fig.update_layout(height=340)
            show_chart(fig)

with tab_history:
    st.subheader("Histori harga")
    history_source = filtered if not filtered.empty else scored_df
    history_mode = st.radio(
        "Sumber grafik histori",
        ["Online yfinance KODE.JK", "Excel Metrik 4W-52W"],
        horizontal=True,
        help=HELP_TEXT["history_source"],
    )
    chart_scope = st.radio(
        "Cakupan grafik",
        ["Saham pilihan", "All/top N hasil filter"],
        horizontal=True,
        help=HELP_TEXT["history_scope"],
    )
    history_max = max(1, min(100, len(filtered) if not filtered.empty else len(scored_df)))
    history_min = min(5, history_max)
    history_default = min(25, history_max)
    history_step = 5 if history_max >= 10 else 1
    top_n_history = safe_slider("Jumlah saham untuk grafik all/top N", history_min, history_max, history_default, step=history_step, help=HELP_TEXT["history_top_n"])
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

    render_excel_fallback_history = history_mode == "Excel Metrik 4W-52W"

    if history_mode == "Online yfinance KODE.JK":
        period = st.selectbox(
            "Rentang data online",
            ONLINE_PERIOD_OPTIONS,
            index=ONLINE_PERIOD_OPTIONS.index("1y"),
            format_func=lambda value: ONLINE_PERIOD_LABELS.get(value, value),
            help=HELP_TEXT["history_period"],
        )
        online_history, online_error, online_source = fetch_yahoo_history(selected_codes, period=period)
        if online_error:
            st.warning(online_error)
            st.caption("Dashboard memakai cache lokal bila tersedia. Excel Metrik hanya dipakai jika online/cache kosong.")
        if online_history.empty:
            st.warning("Data online/cache kosong. Menampilkan Excel Metrik sebagai fallback.")
            render_excel_fallback_history = True
        else:
            st.caption(f"Sumber histori aktif: {online_source}")
            last_dates = online_history.groupby("Kode")["Date"].max().reset_index()
            last_dates["Last Update Online"] = last_dates["Date"].dt.strftime("%Y-%m-%d")
            if show_history_table:
                show_table(last_dates[["Kode", "Last Update Online"]], hide_index=True)

            chart_func = px.area if history_chart_type == "Area" else px.line
            fig = chart_func(
                online_history,
                x="Date",
                y="Close",
                color="Kode",
                title="Harga penutupan historis online",
                labels={"Close": "Harga penutupan", "Date": "Tanggal"},
                color_discrete_sequence=STOCK_LINE_COLORS,
            )
            fig.update_layout(height=480)
            show_chart(fig)

            fig = chart_func(
                online_history,
                x="Date",
                y="Normalized",
                color="Kode",
                title="Perbandingan performa, indeks awal = 100",
                labels={"Normalized": "Indeks performa", "Date": "Tanggal"},
                color_discrete_sequence=STOCK_LINE_COLORS,
            )
            fig.add_hline(y=100, line_dash="dash", line_color=CHART_AXIS_COLOR)
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
                    summary.sort_values("Return_Total_%", ascending=False),
                    hide_index=True,
                    column_config={
                        "Start": st.column_config.DateColumn("Awal"),
                        "End": st.column_config.DateColumn("Akhir"),
                        "Start_Close": st.column_config.NumberColumn("Harga Awal", format="%.0f"),
                        "Last_Close": st.column_config.NumberColumn("Harga Terakhir", format="%.0f"),
                        "Return_Total_%": st.column_config.NumberColumn("Return Total", format="%.1f%%", help=HELP_TEXT["return"]),
                    },
                )

        st.caption("Sumber online memakai ticker IDX format KODE.JK, misalnya BBCA.JK. Jika data live gagal, dashboard mencoba fallback dan cache lokal.")
    if render_excel_fallback_history:
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
                    color_discrete_sequence=STOCK_LINE_COLORS,
                )
                fig.add_hline(y=0, line_dash="dash", line_color=CHART_AXIS_COLOR)
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
                        compare,
                        hide_index=True,
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
                    color_continuous_scale=SCORE_SCALE,
                )
                fig.add_vline(x=0, line_dash="dash", line_color=CHART_AXIS_COLOR)
                fig.update_layout(height=440, xaxis_title="Return 52 minggu (%)", yaxis_title="Score")
                show_chart(fig)

with tab_history:
    st.subheader("Analisa teknikal")
    st.caption("Teknikal memakai OHLCV yfinance/cache sebagai konfirmasi timing. Score fundamental utama tidak berubah.")
    technical_source = filtered if not filtered.empty else scored_df
    technical_codes = technical_source["Kode"].dropna().astype(str).str.upper().unique().tolist()

    if not technical_codes:
        st.warning("Tidak ada kode saham pada filter saat ini.")
    else:
        tech_controls = st.columns([1, 1, 1])
        with tech_controls[0]:
            technical_code = st.selectbox("Kode saham", technical_codes, index=0, help=HELP_TEXT["technical_code"])
        with tech_controls[1]:
            technical_period = st.selectbox(
                "Periode teknikal",
                ONLINE_PERIOD_OPTIONS,
                index=ONLINE_PERIOD_OPTIONS.index("2y"),
                format_func=lambda value: ONLINE_PERIOD_LABELS.get(value, value),
                help=HELP_TEXT["technical_period"],
            )
        with tech_controls[2]:
            chart_style = st.segmented_control("Chart harga", ["Candlestick", "Line"], default="Candlestick")

        load_technical = st.toggle("Tampilkan hasil analisa teknikal online/cache", value=True, help="Aktif untuk menampilkan OHLCV, indikator, Entry Action, dan Position Action. Matikan bila ingin menghindari refresh online sementara.")
        show_auto_scan = st.toggle("Tampilkan ringkasan teknikal top kandidat", value=True, help="Menghitung teknikal otomatis untuk maksimal 5 saham teratas dari hasil filter agar keputusan entry/posisi langsung terlihat.")
        if not load_technical:
            st.info("Aktifkan toggle di atas untuk menampilkan Entry Action, Position Action, candlestick, MA, RSI, MACD, ATR, dan Technical Score.")
        else:
            if show_auto_scan:
                auto_scan_codes = technical_source.head(min(5, len(technical_codes)))["Kode"].tolist()
                if auto_scan_codes:
                    with st.spinner("Menghitung ringkasan teknikal top kandidat..."):
                        auto_history, auto_error, auto_source_label = fetch_yahoo_history(auto_scan_codes, period=technical_period)
                        auto_indicators = build_technical_indicators(auto_history)
                        auto_summary = summarize_technical_indicators(auto_indicators)
                    if auto_error:
                        st.warning(auto_error)
                    if not auto_summary.empty:
                        auto_summary = add_entry_decision(auto_summary, scored_df)
                        auto_summary["Entry_Action"] = pd.Categorical(auto_summary["Entry_Action"], categories=ENTRY_ACTION_ORDER, ordered=True)
                        auto_summary["Position_Action"] = pd.Categorical(auto_summary["Position_Action"], categories=POSITION_ACTION_ORDER, ordered=True)
                        auto_summary = auto_summary.sort_values(
                            ["Entry_Action", "Position_Action", "Technical_Score", "Score"],
                            ascending=[True, True, False, False],
                        )
                        st.markdown("**Ringkasan teknikal top kandidat**")
                        st.caption(f"Sumber ringkasan: {auto_source_label}. Maksimal 5 saham teratas dari hasil filter/ranking saat ini.")
                        summary_cols = [
                            "Kode",
                            "Nama Perusahaan",
                            "Score",
                            "Recommendation",
                            "Entry_Action",
                            "Position_Action",
                            "Exit_Risk",
                            "Technical_Score",
                            "Technical_Signal",
                            "RSI14",
                            "ATR_Stop_2x",
                            "ATR_Stop_Distance_%",
                            "Position_Risk_Bucket",
                            "Timing_Reason",
                            "Position_Reason",
                        ]
                        show_table(
                            auto_summary[[column for column in summary_cols if column in auto_summary.columns]],
                            hide_index=True,
                            column_config={
                                "Score": st.column_config.NumberColumn("Fundamental Score", format="%.1f"),
                                "Technical_Score": st.column_config.NumberColumn("Technical Score", format="%.1f"),
                                "RSI14": st.column_config.NumberColumn("RSI", format="%.1f"),
                                "ATR_Stop_2x": st.column_config.NumberColumn("ATR Stop 2x", format="%.0f", help=HELP_TEXT["atr_stop"]),
                                "ATR_Stop_Distance_%": st.column_config.NumberColumn("Jarak Stop", format="%.1f%%", help=HELP_TEXT["atr_stop"]),
                            },
                        )

            tech_history, tech_error, tech_source_label = fetch_yahoo_history([technical_code], period=technical_period)
            if tech_error:
                st.warning(tech_error)
            if tech_history.empty:
                st.info("Data OHLCV online/cache belum tersedia untuk teknikal. Coba periode lain atau refresh cache histori.")
            else:
                tech_history = build_technical_indicators(tech_history)
                tech_summary = summarize_technical_indicators(tech_history)
                tech_decision = add_entry_decision(tech_summary, scored_df)
                latest_tech = tech_decision.iloc[0] if not tech_decision.empty else pd.Series(dtype=object)

                metric_cols = st.columns(6)
                metric_cols[0].metric("Entry Action", clean_text(latest_tech.get("Entry_Action")), help=HELP_TEXT["entry_action"])
                metric_cols[1].metric("Position Action", clean_text(latest_tech.get("Position_Action")), help=HELP_TEXT["position_action"])
                metric_cols[2].metric("Technical Score", format_number(latest_tech.get("Technical_Score")), help=HELP_TEXT["technical_score"])
                metric_cols[3].metric("Sinyal", clean_text(latest_tech.get("Technical_Signal")))
                metric_cols[4].metric("Exit Risk", clean_text(latest_tech.get("Exit_Risk")))
                metric_cols[5].metric("RSI 14", format_number(latest_tech.get("RSI14")))
                tech_start = tech_history["Date"].min()
                tech_end = tech_history["Date"].max()
                tech_range_label = f"{tech_start:%Y-%m-%d} s.d. {tech_end:%Y-%m-%d}" if pd.notna(tech_start) and pd.notna(tech_end) else "-"
                st.caption(
                    f"Sumber teknikal aktif: {tech_source_label}. Data tampil: {len(tech_history):,} baris, {tech_range_label}. "
                    f"Entry: {clean_text(latest_tech.get('Timing_Reason'))} Posisi: {clean_text(latest_tech.get('Position_Reason'))}"
                )

                decision_columns = [
                    "Kode",
                    "Nama Perusahaan",
                    "Score",
                    "Recommendation",
                    "Technical_Score",
                    "Technical_Signal",
                    "Entry_Action",
                    "Position_Action",
                    "Exit_Risk",
                    "ATR_Stop_2x",
                    "ATR_Stop_Distance_%",
                    "Position_Risk_Bucket",
                    "Timing_Reason",
                    "Position_Reason",
                ]
                st.markdown("**Ringkasan keputusan teknikal**")
                show_table(
                    tech_decision[[column for column in decision_columns if column in tech_decision.columns]],
                    hide_index=True,
                    column_config={
                        "Score": st.column_config.NumberColumn("Fundamental Score", format="%.1f"),
                        "Technical_Score": st.column_config.NumberColumn("Technical Score", format="%.1f"),
                        "ATR_Stop_2x": st.column_config.NumberColumn("ATR Stop 2x", format="%.0f", help=HELP_TEXT["atr_stop"]),
                        "ATR_Stop_Distance_%": st.column_config.NumberColumn("Jarak Stop", format="%.1f%%", help=HELP_TEXT["atr_stop"]),
                    },
                )

                price_panel = tech_history.copy()
                fig = go.Figure()
                if chart_style == "Candlestick" and {"Open", "High", "Low", "Close"}.issubset(price_panel.columns):
                    fig.add_trace(
                        go.Candlestick(
                            x=price_panel["Date"],
                            open=price_panel["Open"],
                            high=price_panel["High"],
                            low=price_panel["Low"],
                            close=price_panel["Close"],
                            name="OHLC",
                            increasing_line_color="#15803d",
                            decreasing_line_color="#dc2626",
                        )
                    )
                else:
                    fig.add_trace(go.Scatter(x=price_panel["Date"], y=price_panel["Close"], mode="lines", name="Close", line=dict(color="#2563eb")))
                for ma_column, color in [("MA20", "#0891b2"), ("MA50", "#7c3aed"), ("MA200", "#475569")]:
                    if ma_column in price_panel.columns:
                        fig.add_trace(go.Scatter(x=price_panel["Date"], y=price_panel[ma_column], mode="lines", name=ma_column, line=dict(color=color, width=1.6)))
                fig.update_layout(
                    title=f"{technical_code}: harga, MA20/50/200 ({technical_period})",
                    height=520,
                    xaxis_title="Tanggal",
                    yaxis_title="Harga",
                    xaxis_rangeslider_visible=False,
                    margin=dict(l=20, r=20, t=60, b=40),
                )
                show_chart(fig)

                lower_cols = st.columns([1, 1])
                with lower_cols[0]:
                    momentum_panel = price_panel[["Date", "RSI14", "MACD", "MACD_Signal"]].copy()
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(x=momentum_panel["Date"], y=momentum_panel["RSI14"], mode="lines", name="RSI14", line=dict(color="#2563eb")))
                    fig.add_hline(y=70, line_dash="dash", line_color="#ea580c")
                    fig.add_hline(y=30, line_dash="dash", line_color="#15803d")
                    fig.update_layout(title="RSI 14", height=320, yaxis_title="RSI", margin=dict(l=20, r=20, t=60, b=40))
                    show_chart(fig)
                with lower_cols[1]:
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(x=momentum_panel["Date"], y=momentum_panel["MACD"], mode="lines", name="MACD", line=dict(color="#7c3aed")))
                    fig.add_trace(go.Scatter(x=momentum_panel["Date"], y=momentum_panel["MACD_Signal"], mode="lines", name="Signal", line=dict(color="#ca8a04")))
                    fig.add_hline(y=0, line_dash="dash", line_color=CHART_AXIS_COLOR)
                    fig.update_layout(title="MACD", height=320, yaxis_title="MACD", margin=dict(l=20, r=20, t=60, b=40))
                    show_chart(fig)

                detail_columns = [
                    "Date",
                    "Close",
                    "MA20",
                    "MA50",
                    "MA200",
                    "RSI14",
                    "MACD",
                    "MACD_Signal",
                    "Volume_Ratio",
                    "ATR_%",
                    "Distance_52W_High_%",
                    "Distance_52W_Low_%",
                ]
                with st.expander("Detail indikator terakhir", expanded=False):
                    show_table(
                        tech_history[[column for column in detail_columns if column in tech_history.columns]].tail(60).sort_values("Date", ascending=False),
                        hide_index=True,
                        column_config={
                            "Date": st.column_config.DateColumn("Tanggal"),
                            "Close": st.column_config.NumberColumn("Close", format="%.0f"),
                            "MA20": st.column_config.NumberColumn("MA20", format="%.0f"),
                            "MA50": st.column_config.NumberColumn("MA50", format="%.0f"),
                            "MA200": st.column_config.NumberColumn("MA200", format="%.0f"),
                            "RSI14": st.column_config.NumberColumn("RSI", format="%.1f"),
                            "MACD": st.column_config.NumberColumn("MACD", format="%.2f"),
                            "MACD_Signal": st.column_config.NumberColumn("Signal", format="%.2f"),
                            "Volume_Ratio": st.column_config.NumberColumn("Volume Ratio", format="%.2f"),
                            "ATR_%": st.column_config.NumberColumn("ATR", format="%.1f%%"),
                            "Distance_52W_High_%": st.column_config.NumberColumn("Jarak 52W High", format="%.1f%%"),
                            "Distance_52W_Low_%": st.column_config.NumberColumn("Jarak 52W Low", format="%.1f%%"),
                        },
                    )

        with st.expander("Scan teknikal top saham", expanded=False):
            scan_cols = st.columns([1, 1, 1])
            with scan_cols[0]:
                scan_n = safe_slider("Jumlah scan", 5, min(30, len(technical_codes)), min(10, len(technical_codes)), step=5, help="Scan manual agar app tidak lambat saat dibuka.")
            with scan_cols[1]:
                signal_filter = st.multiselect(
                    "Filter sinyal",
                    ["Bullish", "Constructive", "Neutral", "Overbought", "Weak"],
                    default=[],
                    help=HELP_TEXT["technical_filter"],
                )
            with scan_cols[2]:
                run_scan = st.button("Hitung teknikal top N")
            action_filter = st.multiselect(
                "Filter aksi entry",
                ENTRY_ACTION_ORDER,
                default=[],
                help=HELP_TEXT["entry_action"],
            )
            position_filter = st.multiselect(
                "Filter aksi posisi",
                POSITION_ACTION_ORDER,
                default=[],
                help=HELP_TEXT["position_action"],
            )

            if run_scan:
                scan_codes = technical_source.head(scan_n)["Kode"].tolist()
                with st.spinner("Mengambil OHLCV dan menghitung indikator teknikal..."):
                    scan_history, scan_error, scan_source_label = fetch_yahoo_history(scan_codes, period=technical_period)
                    scan_indicators = build_technical_indicators(scan_history)
                    scan_summary = summarize_technical_indicators(scan_indicators)
                if scan_error:
                    st.warning(scan_error)
                if scan_summary.empty:
                    st.info("Scan teknikal belum menghasilkan data.")
                else:
                    scan_summary = add_entry_decision(scan_summary, scored_df)
                    if signal_filter:
                        scan_summary = scan_summary[scan_summary["Technical_Signal"].isin(signal_filter)]
                    if action_filter:
                        scan_summary = scan_summary[scan_summary["Entry_Action"].isin(action_filter)]
                    if position_filter:
                        scan_summary = scan_summary[scan_summary["Position_Action"].isin(position_filter)]
                    if scan_summary.empty:
                        st.info("Tidak ada saham yang cocok dengan filter sinyal, aksi entry, dan aksi posisi saat ini.")
                    else:
                        scan_summary["Entry_Action"] = pd.Categorical(scan_summary["Entry_Action"], categories=ENTRY_ACTION_ORDER, ordered=True)
                        scan_summary["Position_Action"] = pd.Categorical(scan_summary["Position_Action"], categories=POSITION_ACTION_ORDER, ordered=True)
                        scan_summary = scan_summary.sort_values(["Position_Action", "Entry_Action", "Technical_Score", "Score"], ascending=[True, True, False, False])
                        st.caption(f"Sumber scan: {scan_source_label}. Entry Action untuk calon pembelian; Position Action untuk saham yang sudah dimiliki.")
                        fig = px.bar(
                            scan_summary.head(20),
                            x="Technical_Score",
                            y="Kode",
                            color="Position_Action",
                            orientation="h",
                            title="Ranking aksi posisi top saham",
                            color_discrete_map=POSITION_ACTION_COLORS,
                            hover_data=["Nama Perusahaan", "Sektor", "Score", "Recommendation", "Entry_Action", "Technical_Signal", "RSI14", "ATR_%", "Position_Reason"],
                        )
                        fig.update_layout(height=430, xaxis_title="Technical Score", yaxis_title="")
                        show_chart(fig)
                        show_table(
                            scan_summary[
                                [
                                    "Kode",
                                    "Nama Perusahaan",
                                    "Sektor",
                                    "Score",
                                    "Recommendation",
                                    "Entry_Action",
                                    "Position_Action",
                                    "Exit_Risk",
                                    "Combined_View",
                                    "Technical_Score",
                                    "Technical_Signal",
                                    "RSI14",
                                    "Volume_Ratio",
                                    "ATR_%",
                                    "ATR_Stop_2x",
                                    "ATR_Stop_Distance_%",
                                    "Position_Risk_Bucket",
                                    "Distance_52W_High_%",
                                    "Timing_Reason",
                                    "Position_Reason",
                                ]
                            ].head(50),
                            hide_index=True,
                            column_config={
                                "Score": st.column_config.NumberColumn("Fundamental Score", format="%.1f"),
                                "Technical_Score": st.column_config.NumberColumn("Technical Score", format="%.1f"),
                                "RSI14": st.column_config.NumberColumn("RSI", format="%.1f"),
                                "Volume_Ratio": st.column_config.NumberColumn("Volume Ratio", format="%.2f"),
                                "ATR_%": st.column_config.NumberColumn("ATR", format="%.1f%%"),
                                "ATR_Stop_2x": st.column_config.NumberColumn("ATR Stop 2x", format="%.0f", help=HELP_TEXT["atr_stop"]),
                                "ATR_Stop_Distance_%": st.column_config.NumberColumn("Jarak Stop", format="%.1f%%", help=HELP_TEXT["atr_stop"]),
                                "Distance_52W_High_%": st.column_config.NumberColumn("Jarak 52W High", format="%.1f%%"),
                            },
                        )

with tab_sector:
    sector_chart_base = chart_market_frame(scored_df, "Grafik sektor")
    sector_controls = st.columns([1, 1, 1, 1])
    with sector_controls[0]:
        sector_group_options = [column for column in ["Sektor", "Subsektor", "Industri", "Subindustri", "Industry"] if column in sector_chart_base.columns]
        sector_group = st.selectbox("Kelompok", sector_group_options, help=HELP_TEXT["sector_group"])
    with sector_controls[1]:
        sector_count_max = max(1, min(25, int(sector_chart_base.groupby(sector_group, dropna=False)["Kode"].count().max())))
        sector_min_default = min(3, sector_count_max)
        sector_min_count = safe_slider("Minimum saham per kelompok", 1, sector_count_max, sector_min_default, help=HELP_TEXT["sector_min"])
    with sector_controls[2]:
        sector_sort = st.selectbox("Urutkan sektor", ["Median_Score", "Strong_Buy", "Total_Market_Cap", "Total_Revenue", "Total_Turnover", "Avg_ROE", "Saham"], help=HELP_TEXT["sector_sort"])
    with sector_controls[3]:
        sector_chart = st.selectbox("Visual utama", ["Bar", "Treemap", "Scatter"], help=HELP_TEXT["sector_chart"])

    sector_summary = (
        sector_chart_base.groupby(sector_group, dropna=False)
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
                color_continuous_scale=COUNT_SCALE,
            )
        elif sector_chart == "Treemap":
            fig = px.treemap(
                sector_summary,
                path=[sector_group],
                values="Total_Market_Cap",
                color="Median_Score",
                title=f"Peta market cap dan score {sector_group.lower()}",
                color_continuous_scale=SCORE_SCALE,
                range_color=[0, 100],
            )
        else:
            fig = px.bar(
                sector_summary,
                x=sector_sort,
                y=sector_group,
                orientation="h",
                color="Strong_Buy",
                title=f"Ranking {sector_group.lower()} berdasarkan {sector_sort}",
                color_continuous_scale=COUNT_SCALE,
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
            color_continuous_scale=SCORE_SCALE,
            range_color=[0, 100],
        )
        fig.update_layout(xaxis_title="Total market cap", yaxis_title="")
        fig.update_layout(height=520)
        show_chart(fig)

    show_table(
        sector_summary,
        hide_index=True,
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
    st.subheader("Kualitas data & workflow update")
    quality_report = build_data_quality_report(scored_df, raw_df)
    review_count = int((quality_report["Rows"].gt(0) & ~quality_report["Severity"].eq("Info")).sum())
    high_count = int(((quality_report["Rows"] > 0) & quality_report["Severity"].eq("High")).sum())
    quality_cols = st.columns(4)
    quality_cols[0].metric("Status check", f"{len(quality_report) - review_count}/{len(quality_report)} OK")
    quality_cols[1].metric("Perlu review", f"{review_count}")
    quality_cols[2].metric("High severity", f"{high_count}")
    quality_cols[3].metric("Lolos data bersih", f"{scored_df['Clean_Data'].sum():,}")

    with st.expander("Market regime & freshness audit", expanded=True):
        freshness_rows = pd.DataFrame(
            [
                {"Area": "Market Regime", "Metric": "Regime", "Value": clean_text(market_context.get("Market_Regime")), "Detail": clean_text(market_context.get("Regime_Reason"))},
                {"Area": "Market Regime", "Metric": "IHSG Last Date", "Value": clean_text(market_context.get("IHSG_Last_Date")), "Detail": f"Source: {clean_text(market_context.get('Market_Source'))}"},
                {"Area": "Market Breadth", "Metric": "Breadth Label", "Value": clean_text(market_context.get("Breadth_Label")), "Detail": f"{market_context.get('Breadth_Count', 0)} kode; MA50 {format_percent(market_context.get('Above_MA50_%'))}; MA200 {format_percent(market_context.get('Above_MA200_%'))}"},
                {"Area": "Freshness", "Metric": "Freshness Label", "Value": clean_text(data_freshness.get("Freshness_Label")), "Detail": f"Lag online {data_freshness.get('Online_Data_Lag_Days') if pd.notna(data_freshness.get('Online_Data_Lag_Days')) else '-'} hari"},
                {"Area": "Freshness", "Metric": "Online Price Coverage", "Value": format_percent(data_freshness.get("Online_Price_Coverage_%"), 0), "Detail": f"Stale rows: {data_freshness.get('Stale_Price_Rows', 0):,}"},
                {"Area": "Freshness", "Metric": "Online Fundamental Coverage", "Value": format_percent(data_freshness.get("Online_Fundamental_Coverage_%"), 0), "Detail": f"Excel fallback fields {format_percent(data_freshness.get('Excel_Fundamental_Coverage_%'), 0)}"},
            ]
        )
        show_table(freshness_rows, hide_index=True)
        stale_price = scored_df.copy()
        if "Online_Last_Date" in stale_price.columns:
            stale_price["Online_Last_Date"] = pd.to_datetime(stale_price["Online_Last_Date"], errors="coerce")
            today = pd.Timestamp.today().normalize()
            stale_price["Online_Lag_Days"] = (today - stale_price["Online_Last_Date"].dt.normalize()).dt.days
            stale_price = stale_price[stale_price["Online_Lag_Days"].gt(5)].sort_values("Online_Lag_Days", ascending=False)
            if not stale_price.empty:
                with st.expander("Kode dengan harga online/cache mulai stale", expanded=False):
                    show_table(
                        stale_price[["Kode", "Nama Perusahaan", "Online_Last_Date", "Online_Lag_Days", "Price_Source", "Volume_Source", "Score", "Recommendation"]].head(100),
                        hide_index=True,
                        column_config={
                            "Online_Last_Date": st.column_config.DateColumn("Tanggal Online"),
                            "Online_Lag_Days": st.column_config.NumberColumn("Lag Hari", format="%d"),
                            "Score": st.column_config.NumberColumn("Score", format="%.1f"),
                        },
                    )

    priority_issues = quality_report[
        quality_report["Rows"].gt(0) & quality_report["Severity"].isin(["High", "Medium"])
    ].copy()
    if priority_issues.empty:
        st.success("Tidak ada issue high/medium pada check utama.")
    else:
        severity_rank = {"High": 0, "Medium": 1}
        priority_issues["Severity_Rank"] = priority_issues["Severity"].map(severity_rank).fillna(9)
        priority_issues = priority_issues.sort_values(["Severity_Rank", "Rows"], ascending=[True, False]).head(5)
        with st.container(border=True):
            st.markdown("**Prioritas Perbaikan**")
            show_table(
                priority_issues[["Severity", "Area", "Check", "Rows", "Action"]],
                hide_index=True,
                column_config={
                    "Rows": st.column_config.NumberColumn("Rows", format="%d"),
                },
            )

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
            universe_summary,
            hide_index=True,
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
                color_continuous_scale=SCORE_SCALE,
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
                    color_discrete_map=SOURCE_COLORS,
                )
                fig.update_layout(height=360, yaxis_title="", margin=dict(l=20, r=20, t=60, b=40))
                show_chart(fig)

        show_table(
            completeness_report.sort_values(["Coverage", "Grup", "Kolom"]),
            hide_index=True,
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
                audit_view,
                hide_index=True,
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
                selected_detail,
                hide_index=True,
                column_config={
                    "Status": st.column_config.TextColumn("Status"),
                    "Actual": st.column_config.TextColumn("Nilai Saat Ini"),
                    "Required": st.column_config.TextColumn("Kriteria"),
                },
            )
        else:
            st.info("Pilih minimal satu kode saham untuk audit.")

    show_table(
        quality_report,
        hide_index=True,
        column_config={
            "Rows": st.column_config.NumberColumn("Rows", format="%d"),
        },
    )

    issue_options = quality_report.loc[quality_report["Rows"].gt(0), "Check"].tolist()
    if issue_options:
        selected_issue = st.selectbox("Lihat detail masalah", issue_options, help=HELP_TEXT["quality_issue"])
        detail = get_quality_detail(scored_df, selected_issue)
        show_table(
            detail.head(200),
            hide_index=True,
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
        2. Lengkapi harga/histori dari yfinance dan fundamental massal dari TradingView scanner; nilai online valid diprioritaskan di atas Excel.
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
        - Relatif sektor: PER/PBV dan ROE/ROA/NPM dibanding saham lain dalam sektor yang sama agar valuasi tidak dibaca terlalu absolut lintas sektor.
        - Threshold sheet: rasio dibandingkan dengan batas dari sheet NonBank atau Banking sebagai cadangan metodologi fundamental.

        Penalti diterapkan untuk PER/PBV negatif, profitabilitas negatif, NPM negatif, volume rendah, harga nol, pergerakan harian ekstrem, dan kelulusan threshold yang terlalu rendah.

        Layer teknikal terpisah dari Score fundamental:
        - Technical Score memakai trend MA20/50/200, RSI, MACD, volume ratio, dan ATR.
        - Entry Action dipakai untuk calon pembelian: fundamental menjadi gerbang awal, teknikal menentukan timing.
        - Position Action dipakai untuk saham yang sudah dimiliki: Hold, Add on Pullback, Review Position, Tight Stop, Take Profit, Reduce, atau Exit / Sell.
        - Tidak ada harga beli pribadi yang dipakai; sinyal posisi adalah arahan umum berbasis data pasar terbaru.
        - ATR Stop 2x adalah zona risiko teknikal berbasis volatilitas, bukan instruksi order otomatis.

        Explainability:
        - Decision Summary merangkum rekomendasi, posisi relatif sektor, dan risiko.
        - Top Strengths menunjukkan faktor skor tertinggi.
        - Top Risks menunjukkan faktor skor terlemah.
        - Action Checklist menunjukkan hal yang perlu dikonfirmasi sebelum entry/posisi.
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
            ["Score", "Sector_Relative_Score", "Valuation_Score", "Quality_Score", "Risk_Score", "Liquidity_Score", "Momentum_Score", "Index_Score", "Penalty"],
            help=HELP_TEXT["factor_inspect"],
        )
    with method_view_cols[1]:
        factor_max = max(1, min(50, len(scored_df)))
        factor_min = min(5, factor_max)
        factor_default = min(15, factor_max)
        factor_step = 5 if factor_max >= 10 else 1
        factor_top_n = safe_slider("Jumlah contoh faktor", factor_min, factor_max, factor_default, step=factor_step, help=HELP_TEXT["factor_top_n"])

    factor_examples = scored_df.sort_values(factor_to_inspect, ascending=False).head(factor_top_n)
    fig = px.histogram(
        scored_df,
        x=factor_to_inspect,
        nbins=40,
        title=f"Distribusi {factor_to_inspect}",
        color_discrete_sequence=[FACTOR_COLORS.get(factor_to_inspect, "#2563eb")],
    )
    fig.update_layout(height=320)
    show_chart(fig)
    factor_example_columns = list(
        dict.fromkeys(
            ["Kode", "Nama Perusahaan", "Recommendation", "Risk_Level", factor_to_inspect, "Score", "Threshold_Pass_Ratio", "Sektor"]
        )
    )
    show_table(
        factor_examples[factor_example_columns],
        hide_index=True,
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
