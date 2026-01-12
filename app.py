import streamlit as st
import pandas as pd
import os
import io
from XlsxWriter import Workbook
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# =========================
# Konfigurasi
# =========================
RIWAYAT_CSV = "riwayat_prediksi_telur.csv"
DATA_HARGA = [
    {"provinsi": "Semua Provinsi", "tanggal": "2026-01-06", "harga_per_butir": 1956},
    {"provinsi": "Aceh", "tanggal": "2026-01-06", "harga_per_butir": 1956},
    {"provinsi": "Sumatera Utara", "tanggal": "2026-01-06", "harga_per_butir": 2154},
    {"provinsi": "Sumatera Barat", "tanggal": "2026-01-06", "harga_per_butir": 1860},   
    {"provinsi": "Riau", "tanggal": "2026-01-06", "harga_per_butir": 2007},
    {"provinsi": "Kepulauan Riau", "tanggal": "2026-01-06", "harga_per_butir": 1851},
    {"provinsi": "Jambi", "tanggal": "2026-01-06", "harga_per_butir": 1716},
    {"provinsi": "Bengkulu", "tanggal": "2026-01-06", "harga_per_butir": 1878},
    {"provinsi": "Sumatera Selatan", "tanggal": "2026-01-06", "harga_per_butir": 1824},
    {"provinsi": "Kepulauan Bangka Belitung", "tanggal": "2026-01-06", "harga_per_butir": 1968},
    {"provinsi": "Lampung", "tanggal": "2026-01-06", "harga_per_butir": 1749},
    {"provinsi": "Banten", "tanggal": "2026-01-06", "harga_per_butir": 1791},
    {"provinsi": "Jawa Barat", "tanggal": "2026-01-06", "harga_per_butir": 1815},
    {"provinsi": "DKI Jakarta", "tanggal": "2026-01-06", "harga_per_butir": 1830},
    {"provinsi": "Jawa Tengah", "tanggal": "2026-01-06", "harga_per_butir": 1755},
    {"provinsi": "DI Yogyakarta", "tanggal": "2026-01-06", "harga_per_butir": 1800},
    {"provinsi": "Jawa Timur", "tanggal": "2026-01-06", "harga_per_butir": 1731},
    {"provinsi": "Bali", "tanggal": "2026-01-06", "harga_per_butir": 1836},
    {"provinsi": "Nusa Tenggara Barat", "tanggal": "2026-01-06", "harga_per_butir": 1818},
    {"provinsi": "Nusa Tenggara Timur", "tanggal": "2026-01-06", "harga_per_butir": 2037},
    {"provinsi": "Kalimantan Barat", "tanggal": "2026-01-06", "harga_per_butir": 2088},
    {"provinsi": "Kalimantan Selatan", "tanggal": "2026-01-06", "harga_per_butir": 1959},
    {"provinsi": "Kalimantan Tengah", "tanggal": "2026-01-06", "harga_per_butir": 2028},
    {"provinsi": "Kalimantan Timur", "tanggal": "2026-01-06", "harga_per_butir": 1929},
    {"provinsi": "Kalimantan Utara", "tanggal": "2026-01-06", "harga_per_butir": 2007},
    {"provinsi": "Gorontalo", "tanggal": "2026-01-06", "harga_per_butir": 2262},
    {"provinsi": "Sulawesi Selatan", "tanggal": "2026-01-06", "harga_per_butir": 1722},
    {"provinsi": "Sulawesi Tenggara", "tanggal": "2026-01-06", "harga_per_butir": 1962},
    {"provinsi": "Sulawesi Tengah", "tanggal": "2026-01-06", "harga_per_butir": 1956},
    {"provinsi": "Sulawesi Utara", "tanggal": "2026-01-06", "harga_per_butir": 2091},
    {"provinsi": "Sulawesi Barat", "tanggal": "2026-01-06", "harga_per_butir": 1641},
    {"provinsi": "Maluku", "tanggal": "2026-01-06", "harga_per_butir": 2304},
    {"provinsi": "Maluku Utara", "tanggal": "2026-01-06", "harga_per_butir": 2325},
    {"provinsi": "Papua", "tanggal": "2026-01-06", "harga_per_butir": 2478},
    {"provinsi": "Papua Barat", "tanggal": "2026-01-06", "harga_per_butir": 2400}
]

df_harga = pd.DataFrame(DATA_HARGA)
df_harga.set_index("provinsi", inplace=True)

JENIS_TELUR = ["Ayam Ras", "Ayam Kampung", "Bebek", "Puyuh"]
UKURAN_TO_BERAT = {"Kecil": 45, "Sedang": 60, "Besar": 75}

JENIS_MULTIPLIER = {
    "Ayam Ras": 1.0,
    "Ayam Kampung": 1.85,
    "Bebek": 1.35,
    "Puyuh": 0.48
}

JENIS_EMOJI = {
    "Ayam Ras": "🐔",
    "Ayam Kampung": "🐓",
    "Bebek": "🦆",
    "Puyuh": "🐦"
}

# =========================
# Fungsi Riwayat (CSV)
# =========================
def muat_riwayat():
    if os.path.exists(RIWAYAT_CSV):
        return pd.read_csv(RIWAYAT_CSV)
    return pd.DataFrame(columns=["provinsi", "jenis", "berat", "ukuran", "harga_prediksi", "timestamp"])

def simpan_riwayat(entry):
    df = muat_riwayat()
    entry["timestamp"] = datetime.now().isoformat()
    df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
    df.to_csv(RIWAYAT_CSV, index=False)

def hapus_riwayat(index):
    df = muat_riwayat()
    if 0 <= index < len(df):
        removed = df.iloc[index].to_dict()
        df = df.drop(index).reset_index(drop=True)
        df.to_csv(RIWAYAT_CSV, index=False)
        return True, removed
    return False, None

def edit_riwayat(index, entry_baru):
    df = muat_riwayat()
    if 0 <= index < len(df):
        entry_baru["timestamp"] = df.iloc[index]["timestamp"]
        df.iloc[index] = entry_baru
        df.to_csv(RIWAYAT_CSV, index=False)
        return True
    return False

# =========================
# Prediksi & Fuzzy Logic
# =========================
def prediksi_harga(provinsi):
    if provinsi in df_harga.index:
        return df_harga.loc[provinsi, "harga_per_butir"]
    return df_harga.loc["Semua Provinsi", "harga_per_butir"]

def mf_ringan(x):
    return max(0, min(1, (60 - x) / 20)) if x <= 60 else 0

def mf_sedang(x):
    if x <= 50 or x >= 80:
        return 0
    elif 50 < x <= 65:
        return (x - 50) / 15
    elif 65 < x < 80:
        return (80 - x) / 15

def mf_berat(x):
    return max(0, min(1, (x - 70) / 20)) if x >= 60 else 0

def mf_uk_kecil(val):
    return 1 if val == "Kecil" else 0

def mf_uk_sedang(val):
    return 1 if val == "Sedang" else 0

def mf_uk_besar(val):
    return 1 if val == "Besar" else 0

def fuzzy_prediksi_harga(berat, ukuran, harga_dasar, jenis):
    ringan = mf_ringan(berat)
    sedang = mf_sedang(berat)
    berat_mf = mf_berat(berat)
    uk_kecil = mf_uk_kecil(ukuran)
    uk_sedang = mf_uk_sedang(ukuran)
    uk_besar = mf_uk_besar(ukuran)
    
    r1 = min(ringan, uk_kecil)
    r2 = min(sedang, uk_sedang)
    r3 = min(berat_mf, uk_besar)
    
    harga_murah = harga_dasar * 0.9
    harga_normal = harga_dasar * 1.0
    harga_mahal = harga_dasar * 1.15
    
    pembilang = (r1 * harga_murah) + (r2 * harga_normal) + (r3 * harga_mahal)
    penyebut = r1 + r2 + r3
    
    hasil = harga_dasar if penyebut == 0 else pembilang / penyebut
    hasil *= JENIS_MULTIPLIER[jenis]
    return round(hasil)

# =========================
# Page Config & Custom CSS
# =========================
st.set_page_config(
    page_title="EggPrice Pro - Prediksi Harga Telur",
    page_icon="🥚",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Initialize dark mode state
if 'dark_mode' not in st.session_state:
    st.session_state.dark_mode = False

# Function to toggle dark mode
def toggle_dark_mode():
    st.session_state.dark_mode = not st.session_state.dark_mode

# Get current theme
is_dark = st.session_state.dark_mode

# Custom CSS untuk UI Modern dengan Dark Mode & Responsive
st.markdown(f"""
<style>
/* Import Google Fonts */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* ========================================
   THEME VARIABLES - Controlled by Toggle
   ======================================== */
:root {{
    --primary-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    --secondary-gradient: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
    --success-gradient: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
    --warning-gradient: linear-gradient(135deg, #F2994A 0%, #F2C94C 100%);
    --card-bg: {'rgba(30, 30, 45, 0.95)' if is_dark else 'rgba(255, 255, 255, 0.95)'};
    --card-bg-solid: {'#1e1e2d' if is_dark else '#ffffff'};
    --card-shadow: {'0 20px 60px rgba(0, 0, 0, 0.4)' if is_dark else '0 20px 60px rgba(0, 0, 0, 0.1)'};
    --border-radius: 20px;
    --transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    --text-primary: {'#ffffff' if is_dark else '#1a1a2e'};
    --text-secondary: {'#b0b0b0' if is_dark else '#666666'};
    --text-muted: {'#808080' if is_dark else '#999999'};
    --bg-main: {'linear-gradient(135deg, #0f0f1a 0%, #1a1a2e 100%)' if is_dark else 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)'};
    --border-color: {'rgba(255, 255, 255, 0.1)' if is_dark else 'rgba(0, 0, 0, 0.05)'};
    --input-bg: {'#2a2a3d' if is_dark else '#ffffff'};
    --input-border: {'#3a3a4d' if is_dark else '#e8e8e8'};
    --glass-border: {'rgba(255, 255, 255, 0.1)' if is_dark else 'rgba(255, 255, 255, 0.8)'};
}}

/* Global Styles */
.stApp {{
    background: var(--bg-main) !important;
    font-family: 'Inter', sans-serif;
}}

.stApp > header {{
    background: transparent !important;
}}

/* Hide Streamlit Branding */
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}
header {{visibility: hidden;}}

/* Tabs Styling */
.stTabs [data-baseweb="tab-list"] {{
    background: var(--card-bg-solid) !important;
    border: 1px solid var(--border-color) !important;
    gap: 4px;
    padding: 6px;
    border-radius: 16px;
    box-shadow: var(--card-shadow);
    flex-wrap: wrap;
    justify-content: center;
}}

.stTabs [data-baseweb="tab"] {{
    color: var(--text-secondary) !important;
    border-radius: 12px;
    padding: 10px 16px;
    font-weight: 600;
    transition: var(--transition);
    font-size: clamp(0.75rem, 2vw, 0.9rem);
    white-space: nowrap;
}}

.stTabs [aria-selected="true"] {{
    background: var(--primary-gradient) !important;
    color: white !important;
}}

/* Input Fields */
.stSelectbox > div > div,
.stTextInput > div > div > input,
.stNumberInput > div > div > input {{
    background: var(--input-bg) !important;
    border-color: var(--input-border) !important;
    color: var(--text-primary) !important;
    border-radius: 12px;
    border: 2px solid var(--input-border);
    transition: var(--transition);
}}

.stSelectbox [data-baseweb="select"] > div {{
    background: var(--input-bg) !important;
    border-color: var(--input-border) !important;
}}

.stSelectbox svg {{
    fill: var(--text-secondary) !important;
}}

.stSelectbox label,
.stTextInput label,
.stNumberInput label {{
    color: var(--text-primary) !important;
}}

/* Radio Buttons */
.stRadio > div {{
    background: var(--card-bg-solid) !important;
    border: 1px solid var(--border-color) !important;
    padding: 1rem;
    border-radius: 12px;
    box-shadow: var(--card-shadow);
}}

.stRadio > div > div {{
    flex-wrap: wrap;
    gap: 0.5rem;
}}

.stRadio label {{
    color: var(--text-primary) !important;
    font-size: clamp(0.8rem, 2vw, 0.9rem) !important;
}}

/* Slider */
.stSlider > div > div > div {{
    background: var(--primary-gradient);
}}

.stSlider label {{
    color: var(--text-primary) !important;
}}

.stSlider [data-testid="stTickBarMin"],
.stSlider [data-testid="stTickBarMax"] {{
    color: var(--text-secondary) !important;
}}

/* DataFrame */
.stDataFrame {{
    border-radius: 16px;
    overflow: hidden;
    box-shadow: var(--card-shadow);
    background: var(--card-bg-solid) !important;
}}

[data-testid="stDataFrame"] > div {{
    background: var(--card-bg-solid) !important;
}}

/* Expander */
.streamlit-expanderHeader {{
    background: var(--card-bg-solid) !important;
    color: var(--text-primary) !important;
    border: 1px solid var(--border-color) !important;
    border-radius: 12px;
    font-weight: 600;
}}

.streamlit-expanderContent {{
    background: var(--card-bg-solid) !important;
    border: 1px solid var(--border-color) !important;
    border-top: none !important;
    border-radius: 0 0 12px 12px !important;
}}

/* Metrics */
[data-testid="stMetricLabel"] {{
    color: var(--text-secondary) !important;
    font-size: clamp(0.75rem, 2vw, 0.9rem) !important;
}}

[data-testid="stMetricValue"] {{
    font-size: clamp(1.5rem, 4vw, 2.5rem) !important;
    font-weight: 800 !important;
    background: var(--primary-gradient);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}}

/* Alert Messages */
.stSuccess, .stInfo, .stWarning, .stError {{
    border-radius: 12px;
    border: none;
    background: var(--card-bg-solid) !important;
}}

/* Markdown Text */
.stMarkdown, .stMarkdown p, .stMarkdown li {{
    color: var(--text-primary) !important;
}}

/* Number Input */
.stNumberInput [data-testid="stDecrement"],
.stNumberInput [data-testid="stIncrement"] {{
    background: var(--card-bg-solid) !important;
    border-color: var(--border-color) !important;
    color: var(--text-primary) !important;
}}

/* Button Secondary */
.stButton > button[kind="secondary"] {{
    background: var(--card-bg-solid) !important;
    color: var(--text-primary) !important;
    border: 1px solid var(--border-color) !important;
}}

/* Main Container - Responsive */
.main .block-container {{
    padding: 2rem 3rem;
    max-width: 1400px;
}}

@media (max-width: 1200px) {{
    .main .block-container {{
        padding: 1.5rem 2rem;
    }}
}}

@media (max-width: 768px) {{
    .main .block-container {{
        padding: 1rem 1rem;
    }}
}}

@media (max-width: 480px) {{
    .main .block-container {{
        padding: 0.75rem 0.5rem;
    }}
}}

/* Hero Header - Responsive */
.hero-header {{
    background: var(--primary-gradient);
    border-radius: var(--border-radius);
    padding: 2.5rem 3rem;
    margin-bottom: 2rem;
    color: white;
    box-shadow: var(--card-shadow);
    position: relative;
    overflow: hidden;
}}

.hero-header::before {{
    content: '';
    position: absolute;
    top: -50%;
    right: -20%;
    width: 400px;
    height: 400px;
    background: rgba(255, 255, 255, 0.1);
    border-radius: 50%;
}}

.hero-header::after {{
    content: '🥚';
    position: absolute;
    right: 40px;
    top: 50%;
    transform: translateY(-50%);
    font-size: 80px;
    opacity: 0.3;
}}

.hero-title {{
    font-size: clamp(1.5rem, 5vw, 2.5rem);
    font-weight: 800;
    margin: 0;
    letter-spacing: -1px;
}}

.hero-subtitle {{
    font-size: clamp(0.85rem, 2.5vw, 1.1rem);
    opacity: 0.9;
    margin-top: 0.5rem;
    font-weight: 400;
}}

/* Hero Header Responsive */
@media (max-width: 768px) {{
    .hero-header {{
        padding: 1.5rem 1.5rem;
        border-radius: 16px;
        margin-bottom: 1.5rem;
    }}
    
    .hero-header::after {{
        font-size: 50px;
        right: 15px;
        opacity: 0.2;
    }}
    
    .hero-header::before {{
        width: 200px;
        height: 200px;
    }}
}}

@media (max-width: 480px) {{
    .hero-header {{
        padding: 1.25rem 1rem;
        border-radius: 12px;
    }}
    
    .hero-header::after {{
        display: none;
    }}
}}

/* Dark Mode Toggle Button */
.dark-mode-toggle {{
    position: fixed;
    top: 70px;
    right: 20px;
    z-index: 9999;
    background: var(--card-bg-solid);
    border: 2px solid var(--border-color);
    border-radius: 50px;
    padding: 8px 16px;
    cursor: pointer;
    display: flex;
    align-items: center;
    gap: 8px;
    box-shadow: var(--card-shadow);
    transition: var(--transition);
    font-size: 0.9rem;
    color: var(--text-primary);
}}

.dark-mode-toggle:hover {{
    transform: scale(1.05);
}}

@media (max-width: 768px) {{
    .dark-mode-toggle {{
        top: 10px;
        right: 10px;
        padding: 6px 12px;
        font-size: 0.8rem;
    }}
}}

/* Glass Card - Responsive */
.glass-card {{
    background: var(--card-bg);
    backdrop-filter: blur(20px);
    border-radius: var(--border-radius);
    padding: 2rem;
    box-shadow: var(--card-shadow);
    border: 1px solid var(--glass-border);
    transition: var(--transition);
    margin-bottom: 1.5rem;
    color: var(--text-primary);
}}

.glass-card:hover {{
    transform: translateY(-5px);
    box-shadow: 0 30px 80px rgba(0, 0, 0, 0.15);
}}

.glass-card h4 {{
    color: var(--text-primary);
}}

.glass-card p {{
    color: var(--text-secondary);
}}

@media (max-width: 768px) {{
    .glass-card {{
        padding: 1.5rem;
        border-radius: 16px;
    }}
}}

@media (max-width: 480px) {{
    .glass-card {{
        padding: 1rem;
        border-radius: 12px;
    }}
    
    .glass-card:hover {{
        transform: none;
    }}
}}

/* Price Display Card - Responsive */
.price-card {{
    background: var(--primary-gradient);
    border-radius: var(--border-radius);
    padding: 2.5rem;
    color: white;
    text-align: center;
    box-shadow: var(--card-shadow);
    position: relative;
    overflow: hidden;
}}

.price-card::before {{
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23ffffff' fill-opacity='0.05'%3E%3Cpath d='M36 34v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zm0-30V0h-2v4h-4v2h4v4h2V6h4V4h-4zM6 34v-4H4v4H0v2h4v4h2v-4h4v-2H6zM6 4V0H4v4H0v2h4v4h2V6h4V4H6z'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E");
}}

.price-label {{
    font-size: clamp(0.75rem, 2vw, 0.9rem);
    text-transform: uppercase;
    letter-spacing: 2px;
    opacity: 0.9;
    margin-bottom: 0.5rem;
    position: relative;
    z-index: 1;
}}

.price-value {{
    font-size: clamp(2rem, 6vw, 3rem);
    font-weight: 800;
    margin: 0.5rem 0;
    text-shadow: 0 4px 20px rgba(0, 0, 0, 0.2);
    position: relative;
    z-index: 1;
}}

.price-info {{
    font-size: clamp(0.75rem, 2vw, 0.85rem);
    opacity: 0.85;
    margin-top: 1rem;
    position: relative;
    z-index: 1;
}}

@media (max-width: 768px) {{
    .price-card {{
        padding: 1.5rem;
        border-radius: 16px;
    }}
}}

@media (max-width: 480px) {{
    .price-card {{
        padding: 1.25rem;
        border-radius: 12px;
    }}
}}

/* Stat Cards - Responsive */
.stat-card {{
    background: var(--card-bg-solid);
    border-radius: 16px;
    padding: 1.5rem;
    text-align: center;
    box-shadow: var(--card-shadow);
    transition: var(--transition);
    border: 1px solid var(--border-color);
}}

.stat-card:hover {{
    transform: translateY(-3px);
    box-shadow: 0 15px 50px rgba(0, 0, 0, 0.12);
}}

.stat-icon {{
    font-size: clamp(1.8rem, 4vw, 2.5rem);
    margin-bottom: 0.5rem;
}}

.stat-value {{
    font-size: clamp(1.2rem, 3.5vw, 1.8rem);
    font-weight: 700;
    color: var(--text-primary);
    margin: 0.3rem 0;
    word-break: break-word;
}}

.stat-label {{
    font-size: clamp(0.65rem, 2vw, 0.85rem);
    color: var(--text-secondary);
    text-transform: uppercase;
    letter-spacing: 1px;
}}

@media (max-width: 768px) {{
    .stat-card {{
        padding: 1rem;
        border-radius: 12px;
    }}
    
    .stat-card:hover {{
        transform: none;
    }}
}}

@media (max-width: 480px) {{
    .stat-card {{
        padding: 0.75rem;
        border-radius: 10px;
    }}
}}

/* Section Title */
.section-title {{
    font-size: clamp(1.1rem, 3vw, 1.5rem);
    font-weight: 700;
    color: var(--text-primary);
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    gap: 0.75rem;
    flex-wrap: wrap;
}}

.section-title-icon {{
    background: var(--primary-gradient);
    width: clamp(32px, 8vw, 40px);
    height: clamp(32px, 8vw, 40px);
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: clamp(1rem, 2.5vw, 1.2rem);
    flex-shrink: 0;
}}

@media (max-width: 480px) {{
    .section-title {{
        margin-bottom: 1rem;
    }}
}}

/* Custom Buttons - Responsive */
.stButton > button {{
    background: var(--primary-gradient);
    color: white;
    border: none;
    border-radius: 12px;
    padding: 0.75rem 1.5rem;
    font-weight: 600;
    font-size: clamp(0.85rem, 2vw, 1rem);
    transition: var(--transition);
    box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
    width: 100%;
}}

.stButton > button:hover {{
    transform: translateY(-2px);
    box-shadow: 0 8px 25px rgba(102, 126, 234, 0.5);
}}

.stButton > button:active {{
    transform: translateY(0);
}}

@media (max-width: 768px) {{
    .stButton > button {{
        padding: 0.6rem 1rem;
    }}
    
    .stButton > button:hover {{
        transform: none;
    }}
    
    .stTabs [data-baseweb="tab-list"] {{
        padding: 4px;
        border-radius: 12px;
        gap: 2px;
    }}
    
    .stTabs [data-baseweb="tab"] {{
        padding: 8px 12px;
        border-radius: 10px;
        flex: 1 1 auto;
        text-align: center;
        min-width: 0;
    }}
    
    .stRadio > div {{
        padding: 0.75rem;
    }}
}}

@media (max-width: 480px) {{
    .stTabs [data-baseweb="tab"] {{
        padding: 8px 8px;
        font-size: 0.7rem;
    }}
}}

/* Select Box Focus */
.stSelectbox > div > div:focus-within {{
    border-color: #667eea;
    box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.2);
}}

/* DataFrame Responsive */
@media (max-width: 768px) {{
    .stDataFrame {{
        border-radius: 12px;
        font-size: 0.8rem;
    }}
}}

/* Download Buttons - Responsive */
.stDownloadButton > button {{
    background: var(--success-gradient);
    border: none;
    border-radius: 12px;
    padding: 0.75rem 1.5rem;
    font-weight: 600;
    box-shadow: 0 4px 15px rgba(17, 153, 142, 0.4);
    width: 100%;
    font-size: clamp(0.8rem, 2vw, 0.95rem);
}}

@media (max-width: 768px) {{
    .stDownloadButton > button {{
        padding: 0.6rem 1rem;
    }}
}}

/* Input Fields Responsive */
@media (max-width: 768px) {{
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input {{
        padding: 0.6rem 0.8rem;
        font-size: 0.9rem;
    }}
    
    .streamlit-expanderHeader {{
        font-size: 0.9rem;
        padding: 0.75rem !important;
    }}
}}

/* Footer */
.custom-footer {{
    text-align: center;
    padding: 2rem 1rem;
    color: var(--text-secondary);
    font-size: clamp(0.75rem, 2vw, 0.9rem);
    margin-top: 3rem;
    border-top: 1px solid var(--border-color);
}}

.custom-footer p {{
    color: var(--text-secondary) !important;
    margin: 0.5rem 0;
}}

@media (max-width: 768px) {{
    .custom-footer {{
        padding: 1.5rem 0.5rem;
        margin-top: 2rem;
    }}
}}

/* Animations */
@keyframes fadeInUp {{
    from {{
        opacity: 0;
        transform: translateY(20px);
    }}
    to {{
        opacity: 1;
        transform: translateY(0);
    }}
}}

.animate-fade-in {{
    animation: fadeInUp 0.6s ease-out;
}}

/* Disable animations on mobile */
@media (max-width: 768px) {{
    .animate-fade-in {{
        animation: none;
    }}
    
    .glass-card:hover,
    .stat-card:hover,
    .stButton > button:hover {{
        transform: none !important;
    }}
}}

/* Mobile Optimizations */
@media (max-width: 480px) {{
    [data-testid="column"] {{
        width: 100% !important;
        flex: 1 1 100% !important;
    }}
    
    .stMarkdown {{
        margin-bottom: 0.5rem;
    }}
    
    .glass-card,
    .stat-card,
    .price-card {{
        border-radius: 12px !important;
    }}
}}

/* Touch-friendly targets */
@media (hover: none) and (pointer: coarse) {{
    .stButton > button,
    .stDownloadButton > button {{
        min-height: 44px;
    }}
    
    .stRadio label,
    .stSelectbox > div > div {{
        min-height: 44px;
    }}
}}

/* Column Gap Responsive */
[data-testid="stHorizontalBlock"] {{
    gap: 1rem;
    flex-wrap: wrap;
}}

@media (max-width: 768px) {{
    [data-testid="stHorizontalBlock"] {{
        gap: 0.75rem;
    }}
}}

@media (max-width: 480px) {{
    [data-testid="stHorizontalBlock"] {{
        gap: 0.5rem;
    }}
}}

/* Empty State */
.empty-state {{
    text-align: center;
    padding: 3rem 1.5rem;
    color: var(--text-secondary);
}}

.empty-state h3 {{
    color: var(--text-primary);
}}
</style>
""", unsafe_allow_html=True)

# =========================
# Dark Mode Toggle
# =========================
col_header, col_toggle = st.columns([6, 1])
with col_toggle:
    dark_icon = "🌙" if not is_dark else "☀️"
    dark_label = "Dark" if not is_dark else "Light"
    if st.button(f"{dark_icon} {dark_label}", key="dark_mode_toggle", use_container_width=True):
        toggle_dark_mode()
        st.rerun()

# =========================
# Hero Header
# =========================
st.markdown("""
<div class="hero-header animate-fade-in">
    <p class="hero-title">🥚 EggPrice Pro</p>
    <p class="hero-subtitle">Sistem Prediksi Harga Telur Cerdas dengan Fuzzy Logic • Data 06 Januari 2026</p>
</div>
""", unsafe_allow_html=True)

# =========================
# Quick Stats - Responsive Grid
# =========================
avg_price = int(df_harga["harga_per_butir"].mean())
max_price = df_harga["harga_per_butir"].max()
min_price = df_harga["harga_per_butir"].min()
total_provinsi = len(df_harga) - 1

col_stat1, col_stat2 = st.columns(2)
col_stat3, col_stat4 = st.columns(2)

with col_stat1:
    st.markdown(f"""
    <div class="stat-card animate-fade-in">
        <div class="stat-icon">📊</div>
        <div class="stat-value">Rp {avg_price:,}</div>
        <div class="stat-label">Rata-rata Nasional</div>
    </div>
    """, unsafe_allow_html=True)

with col_stat2:
    st.markdown(f"""
    <div class="stat-card animate-fade-in">
        <div class="stat-icon">📈</div>
        <div class="stat-value">Rp {max_price:,}</div>
        <div class="stat-label">Harga Tertinggi</div>
    </div>
    """, unsafe_allow_html=True)

with col_stat3:
    st.markdown(f"""
    <div class="stat-card animate-fade-in">
        <div class="stat-icon">📉</div>
        <div class="stat-value">Rp {min_price:,}</div>
        <div class="stat-label">Harga Terendah</div>
    </div>
    """, unsafe_allow_html=True)

with col_stat4:
    st.markdown(f"""
    <div class="stat-card animate-fade-in">
        <div class="stat-icon">🗺️</div>
        <div class="stat-value">{total_provinsi}</div>
        <div class="stat-label">Provinsi</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# =========================
# Navigation Tabs
# =========================
tab_prediksi, tab_analisis, tab_riwayat, tab_ekspor = st.tabs([
    "🔮 Prediksi Harga", 
    "📊 Analisis Data", 
    "📜 Riwayat", 
    "📤 Ekspor"
])

# =========================
# Tab 1: Prediksi
# =========================
with tab_prediksi:
    # Input controls in a single row
    col_prov, col_jenis, col_ukuran = st.columns([2, 2, 1])
    
    with col_prov:
        provinsi = st.selectbox(
            "📍 Provinsi",
            sorted(df_harga.index.tolist()),
            help="Pilih provinsi untuk melihat harga telur"
        )
    
    with col_jenis:
        if 'selected_jenis' not in st.session_state:
            st.session_state.selected_jenis = "Ayam Ras"
        jenis = st.selectbox(
            "🥚 Jenis Telur",
            JENIS_TELUR,
            index=JENIS_TELUR.index(st.session_state.selected_jenis)
        )
        st.session_state.selected_jenis = jenis
    
    with col_ukuran:
        ukuran = st.selectbox(
            "📏 Ukuran",
            ["Kecil", "Sedang", "Besar"],
            index=1
        )
    
    berat = UKURAN_TO_BERAT[ukuran]
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Results
    col_price, col_details = st.columns([1, 1], gap="large")
    
    with col_price:
        # Calculate price
        harga_dasar = prediksi_harga(provinsi)
        harga = fuzzy_prediksi_harga(berat, ukuran, harga_dasar, jenis)
        harga_per_kg = round(harga * (1000 / berat))
        harga_per_tray = harga * 30
        
        # Price Card
        st.markdown(f"""
        <div class="price-card animate-fade-in">
            <div class="price-label">Harga Per Butir</div>
            <div class="price-value">Rp {harga:,}</div>
            <div class="price-info">
                {JENIS_EMOJI[jenis]} {jenis} • {ukuran} • {berat}g
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_details:
        col_kg, col_tray = st.columns(2)
        
        with col_kg:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-icon">📦</div>
                <div class="stat-value" style="font-size: 1.4rem;">Rp {harga_per_kg:,}</div>
                <div class="stat-label">Per Kilogram</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col_tray:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-icon">🥚</div>
                <div class="stat-value" style="font-size: 1.4rem;">Rp {harga_per_tray:,}</div>
                <div class="stat-label">Per Tray (30 butir)</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Save Button
        if st.button("💾 Simpan ke Riwayat", use_container_width=True, type="primary"):
            entry = {
                "provinsi": provinsi,
                "jenis": jenis,
                "berat": berat,
                "ukuran": ukuran,
                "harga_prediksi": harga
            }
            simpan_riwayat(entry)
            st.success("✅ Berhasil disimpan ke riwayat!")
            st.balloons()


# =========================
# Tab 2: Analisis Data
# =========================
with tab_analisis:
    st.markdown("""
    <div class="section-title">
        <span class="section-title-icon">📊</span>
        Analisis Harga Telur Nasional
    </div>
    """, unsafe_allow_html=True)
    
    col_chart1, col_chart2 = st.columns(2)
    
    with col_chart1:
        # Bar Chart - Harga per Provinsi
        df_chart = df_harga.reset_index()
        df_chart = df_chart[df_chart['provinsi'] != 'Semua Provinsi'].sort_values('harga_per_butir', ascending=True)
        
        fig_bar = px.bar(
            df_chart,
            x='harga_per_butir',
            y='provinsi',
            orientation='h',
            title='Perbandingan Harga Telur per Provinsi',
            labels={'harga_per_butir': 'Harga (Rp)', 'provinsi': 'Provinsi'},
            color='harga_per_butir',
            color_continuous_scale='Viridis'
        )
        fig_bar.update_layout(
            height=600,
            showlegend=False,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Inter"),
            title_font_size=16,
            coloraxis_showscale=False
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    
    with col_chart2:
        # Top 10 Termahal
        df_top = df_chart.nlargest(10, 'harga_per_butir')
        
        fig_top = go.Figure(go.Bar(
            x=df_top['harga_per_butir'],
            y=df_top['provinsi'],
            orientation='h',
            marker=dict(
                color=df_top['harga_per_butir'],
                colorscale='Reds'
            ),
            text=[f"Rp {x:,}" for x in df_top['harga_per_butir']],
            textposition='outside'
        ))
        fig_top.update_layout(
            title='🔥 Top 10 Provinsi Termahal',
            height=350,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Inter"),
            title_font_size=16,
            xaxis_title='Harga (Rp)',
            yaxis_title=''
        )
        st.plotly_chart(fig_top, use_container_width=True)
        
        # Top 10 Termurah
        df_bottom = df_chart.nsmallest(10, 'harga_per_butir')
        
        fig_bottom = go.Figure(go.Bar(
            x=df_bottom['harga_per_butir'],
            y=df_bottom['provinsi'],
            orientation='h',
            marker=dict(
                color=df_bottom['harga_per_butir'],
                colorscale='Greens'
            ),
            text=[f"Rp {x:,}" for x in df_bottom['harga_per_butir']],
            textposition='outside'
        ))
        fig_bottom.update_layout(
            title='💚 Top 10 Provinsi Termurah',
            height=350,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Inter"),
            title_font_size=16,
            xaxis_title='Harga (Rp)',
            yaxis_title=''
        )
        st.plotly_chart(fig_bottom, use_container_width=True)
    
    # Distribusi Harga
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_dist1, col_dist2 = st.columns(2)
    
    with col_dist1:
        fig_hist = px.histogram(
            df_chart,
            x='harga_per_butir',
            nbins=15,
            title='Distribusi Harga Telur',
            labels={'harga_per_butir': 'Harga (Rp)', 'count': 'Jumlah Provinsi'},
            color_discrete_sequence=['#667eea']
        )
        fig_hist.update_layout(
            height=350,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Inter"),
            title_font_size=16
        )
        st.plotly_chart(fig_hist, use_container_width=True)
    
    with col_dist2:
        # Pie Chart by Region
        def get_region(prov):
            sumatera = ['Aceh', 'Sumatera Utara', 'Sumatera Barat', 'Riau', 'Kepulauan Riau', 
                       'Jambi', 'Bengkulu', 'Sumatera Selatan', 'Kepulauan Bangka Belitung', 'Lampung']
            jawa = ['Banten', 'Jawa Barat', 'DKI Jakarta', 'Jawa Tengah', 'DI Yogyakarta', 'Jawa Timur']
            kalimantan = ['Kalimantan Barat', 'Kalimantan Selatan', 'Kalimantan Tengah', 
                         'Kalimantan Timur', 'Kalimantan Utara']
            sulawesi = ['Gorontalo', 'Sulawesi Selatan', 'Sulawesi Tenggara', 'Sulawesi Tengah', 
                       'Sulawesi Utara', 'Sulawesi Barat']
            
            if prov in sumatera: return 'Sumatera'
            elif prov in jawa: return 'Jawa & Bali'
            elif prov in kalimantan: return 'Kalimantan'
            elif prov in sulawesi: return 'Sulawesi'
            else: return 'Indonesia Timur'
        
        df_chart['region'] = df_chart['provinsi'].apply(get_region)
        df_region = df_chart.groupby('region')['harga_per_butir'].mean().reset_index()
        
        fig_pie = px.pie(
            df_region,
            values='harga_per_butir',
            names='region',
            title='Rata-rata Harga per Wilayah',
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        fig_pie.update_layout(
            height=350,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Inter"),
            title_font_size=16
        )
        st.plotly_chart(fig_pie, use_container_width=True)


# =========================
# Tab 3: Riwayat
# =========================
with tab_riwayat:
    st.markdown("""
    <div class="section-title">
        <span class="section-title-icon">📜</span>
        Riwayat Prediksi
    </div>
    """, unsafe_allow_html=True)
    
    df_riwayat = muat_riwayat()
    
    if df_riwayat.empty:
        st.markdown("""
        <div class="glass-card empty-state">
            <div style="font-size: clamp(2.5rem, 8vw, 4rem); margin-bottom: 1rem;">📭</div>
            <h3 style="margin-bottom: 0.5rem;">Belum Ada Riwayat</h3>
            <p>Mulai prediksi harga telur dan simpan hasilnya di sini</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        # Search & Filter
        col_search, col_filter = st.columns([2, 1])
        
        with col_search:
            cari = st.text_input(
                "🔍 Cari riwayat...",
                placeholder="Ketik nama provinsi atau jenis telur",
                label_visibility="collapsed"
            )
        
        with col_filter:
            filter_jenis = st.selectbox(
                "Filter Jenis",
                ["Semua"] + JENIS_TELUR,
                label_visibility="collapsed"
            )
        
        # Apply filters
        filtered = df_riwayat.copy()
        if cari:
            filtered = filtered[
                filtered["provinsi"].str.contains(cari, case=False) |
                filtered["jenis"].str.contains(cari, case=False)
            ]
        if filter_jenis != "Semua":
            filtered = filtered[filtered["jenis"] == filter_jenis]
        
        # Stats
        if not filtered.empty:
            col_s1, col_s2, col_s3 = st.columns(3)
            with col_s1:
                st.metric("Total Riwayat", len(filtered))
            with col_s2:
                st.metric("Rata-rata Harga", f"Rp {int(filtered['harga_prediksi'].mean()):,}")
            with col_s3:
                st.metric("Total Nilai", f"Rp {int(filtered['harga_prediksi'].sum()):,}")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Display as cards
        if not filtered.empty:
            for idx, row in filtered.iterrows():
                emoji = JENIS_EMOJI.get(row['jenis'], '🥚')
                timestamp = row['timestamp'][:10] if pd.notna(row['timestamp']) else '-'
                
                with st.expander(f"{emoji} {row['provinsi']} - {row['jenis']} | Rp {int(row['harga_prediksi']):,}", expanded=False):
                    col_detail, col_action = st.columns([3, 1])
                    
                    with col_detail:
                        st.markdown(f"""
                        - **Provinsi:** {row['provinsi']}
                        - **Jenis:** {row['jenis']}
                        - **Ukuran:** {row['ukuran']}
                        - **Berat:** {row['berat']} gram
                        - **Harga:** Rp {int(row['harga_prediksi']):,}
                        - **Tanggal:** {timestamp}
                        """)
                    
                    with col_action:
                        if st.button("🗑️ Hapus", key=f"del_{idx}", use_container_width=True):
                            success, _ = hapus_riwayat(idx)
                            if success:
                                st.success("Berhasil dihapus!")
                                st.rerun()
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Bulk Actions
        with st.expander("⚙️ Kelola Riwayat"):
            st.warning("⚠️ Hati-hati! Tindakan ini tidak dapat dibatalkan.")
            
            col_edit, col_del = st.columns(2)
            
            with col_edit:
                st.markdown("**✏️ Edit Entri**")
                if len(df_riwayat) > 0:
                    edit_idx = st.number_input(
                        "Pilih nomor entri",
                        min_value=0,
                        max_value=len(df_riwayat)-1,
                        value=0,
                        key="edit_idx"
                    )
                    
                    if edit_idx < len(df_riwayat):
                        row = df_riwayat.iloc[edit_idx]
                        prov_edit = st.selectbox(
                            "Provinsi",
                            sorted(df_harga.index.tolist()),
                            index=sorted(df_harga.index.tolist()).index(row["provinsi"]) if row["provinsi"] in df_harga.index else 0,
                            key="edit_prov"
                        )
                        jenis_edit = st.selectbox(
                            "Jenis",
                            JENIS_TELUR,
                            index=JENIS_TELUR.index(row["jenis"]) if row["jenis"] in JENIS_TELUR else 0,
                            key="edit_jenis"
                        )
                        ukuran_edit = st.selectbox(
                            "Ukuran",
                            ["Kecil", "Sedang", "Besar"],
                            index=["Kecil", "Sedang", "Besar"].index(row["ukuran"]) if row["ukuran"] in ["Kecil", "Sedang", "Besar"] else 1,
                            key="edit_ukuran"
                        )
                        berat_edit = st.number_input(
                            "Berat (gram)",
                            min_value=30,
                            max_value=100,
                            value=int(row["berat"]),
                            key="edit_berat"
                        )
                        
                        if st.button("💾 Simpan Perubahan", type="primary"):
                            harga_baru = fuzzy_prediksi_harga(
                                berat_edit, 
                                ukuran_edit, 
                                prediksi_harga(prov_edit), 
                                jenis_edit
                            )
                            edit_riwayat(edit_idx, {
                                "provinsi": prov_edit,
                                "jenis": jenis_edit,
                                "berat": berat_edit,
                                "ukuran": ukuran_edit,
                                "harga_prediksi": harga_baru,
                                "timestamp": row["timestamp"]
                            })
                            st.success("✅ Perubahan disimpan!")
                            st.rerun()
            
            with col_del:
                st.markdown("**🗑️ Hapus Semua**")
                st.markdown("Hapus seluruh riwayat prediksi")
                if st.button("🗑️ Hapus Semua Riwayat", type="secondary"):
                    if os.path.exists(RIWAYAT_CSV):
                        os.remove(RIWAYAT_CSV)
                        st.success("Semua riwayat telah dihapus!")
                        st.rerun()

# =========================
# Tab 4: Ekspor
# =========================
with tab_ekspor:
    st.markdown("""
    <div class="section-title">
        <span class="section-title-icon">📤</span>
        Ekspor Data
    </div>
    """, unsafe_allow_html=True)
    
    df_riwayat = muat_riwayat()
    
    col_exp1, col_exp2 = st.columns(2)
    
    with col_exp1:
        st.markdown("""
        <div class="glass-card">
            <h4>📋 Data Riwayat Prediksi</h4>
            <p>Ekspor semua riwayat prediksi yang telah disimpan</p>
        </div>
        """, unsafe_allow_html=True)
        
        if df_riwayat.empty:
            st.info("📭 Belum ada data riwayat untuk diekspor")
        else:
            st.dataframe(
                df_riwayat,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "provinsi": "Provinsi",
                    "jenis": "Jenis Telur",
                    "berat": st.column_config.NumberColumn("Berat (g)", format="%d g"),
                    "ukuran": "Ukuran",
                    "harga_prediksi": st.column_config.NumberColumn("Harga", format="Rp %d"),
                    "timestamp": "Waktu"
                }
            )
            
            col_csv, col_xlsx = st.columns(2)
            
            with col_csv:
                csv = df_riwayat.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "⬇️ Unduh CSV",
                    data=csv,
                    file_name="riwayat_prediksi_telur.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            with col_xlsx:
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                    df_riwayat.to_excel(writer, index=False, sheet_name="Riwayat")
                excel_data = excel_buffer.getvalue()
                
                st.download_button(
                    "⬇️ Unduh Excel",
                    data=excel_data,
                    file_name="riwayat_prediksi_telur.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    
    with col_exp2:
        st.markdown("""
        <div class="glass-card">
            <h4>🗺️ Data Harga Nasional</h4>
            <p>Ekspor data harga telur seluruh provinsi</p>
        </div>
        """, unsafe_allow_html=True)
        
        df_nasional = df_harga.reset_index()
        st.dataframe(
            df_nasional,
            use_container_width=True,
            hide_index=True,
            column_config={
                "provinsi": "Provinsi",
                "tanggal": "Tanggal Data",
                "harga_per_butir": st.column_config.NumberColumn("Harga/Butir", format="Rp %d")
            }
        )
        
        col_csv2, col_xlsx2 = st.columns(2)
        
        with col_csv2:
            csv_nasional = df_nasional.to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇️ Unduh CSV",
                data=csv_nasional,
                file_name="harga_telur_nasional.csv",
                mime="text/csv",
                use_container_width=True,
                key="csv_nasional"
            )
        
        with col_xlsx2:
            excel_buffer2 = io.BytesIO()
            with pd.ExcelWriter(excel_buffer2, engine="xlsxwriter") as writer:
                df_nasional.to_excel(writer, index=False, sheet_name="Harga Nasional")
            excel_data2 = excel_buffer2.getvalue()
            
            st.download_button(
                "⬇️ Unduh Excel",
                data=excel_data2,
                file_name="harga_telur_nasional.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="xlsx_nasional"
            )

# =========================
# Footer
# =========================
st.markdown("""
<div class="custom-footer">
    <p>🥚 <strong>EggPrice Pro</strong> - Sistem Prediksi Harga Telur Cerdas</p>
    <p style="font-size: 0.8rem; color: #999;">
        Dibuat dengan ❤️ menggunakan Streamlit & Fuzzy Logic | Data: 06 Januari 2026
    </p>
</div>
""", unsafe_allow_html=True)
