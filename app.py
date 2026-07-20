import calendar
import copy
import hashlib
import html
import json
import math
import posixpath
import re
import urllib.parse
import uuid
import zipfile
from datetime import date, datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

try:
    from st_keyup import st_keyup
except ImportError:
    st_keyup = None


# =========================
# 基本設定
# =========================
APP_TITLE = "顧客カルテ"

# Streamlitでは、st.set_page_config は他の st.* 呼び出しより先に実行する
st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="collapsed",
)

EXCEL_FILE = "配車予定 次郎_修正版.xlsm"
SHEET_NAME = "Sheet1"

# secrets.toml に入れる設定
DROPBOX_APP_KEY = st.secrets.get("DROPBOX_APP_KEY", "")
DROPBOX_APP_SECRET = st.secrets.get("DROPBOX_APP_SECRET", "")
DROPBOX_REFRESH_TOKEN = st.secrets.get("DROPBOX_REFRESH_TOKEN", "")
# 移行期間用。Streamlit CloudではRefresh Token方式の3項目を使う。
DROPBOX_ACCESS_TOKEN = st.secrets.get("DROPBOX_ACCESS_TOKEN", "")
DROPBOX_DEFAULT_FILE_PATH = "/1共有　青山商店　本社/配車表-北海道-/配車予定 次郎.xlsm"
DROPBOX_FILE_PATH = st.secrets.get("DROPBOX_FILE_PATH", DROPBOX_DEFAULT_FILE_PATH)
DROPBOX_BACKUP_FOLDER = "/1共有　青山商店　本社/配車表-北海道-/Backups"
DROPBOX_FAST_CACHE_FILE = "/1共有　青山商店　本社/配車表-北海道-/顧客検索キャッシュ.json"
# Excelの列構成や読み込み処理を変更した時は、この番号を上げて古いJSONを無効化する。
DROPBOX_FAST_CACHE_VERSION = 2
DISPATCH_DROPBOX_DEFAULT_FILE_PATH = "/1共有　青山商店　本社/配車表-次郎-/配車表1.xlsm"
DISPATCH_DROPBOX_FILE_PATH = st.secrets.get(
    "DISPATCH_DROPBOX_FILE_PATH",
    DISPATCH_DROPBOX_DEFAULT_FILE_PATH,
)
DISPATCH_LOCAL_FILE = st.secrets.get(
    "DISPATCH_LOCAL_FILE",
    r"C:\Users\jiroa\Aoyama Dropbox\bulu jack\1共有　青山商店　本社\配車表-次郎-\配車表1.xlsm",
)
DISPATCH_MONTH_SHEETS = [f"{month}月" for month in range(1, 13)]
SOLUBLE_SHEET_NAME = "ソリュブル"
SOLUBLE_FILE_NAME = "aoベンチャーグレイン配車表.xlsx"
SOLUBLE_DROPBOX_DEFAULT_FILE_PATH = (
    str(DROPBOX_DEFAULT_FILE_PATH).rsplit("/", 1)[0] + "/" + SOLUBLE_FILE_NAME
)
SOLUBLE_DROPBOX_FILE_PATH = st.secrets.get(
    "SOLUBLE_DROPBOX_FILE_PATH",
    str(DROPBOX_FILE_PATH).rsplit("/", 1)[0] + "/" + SOLUBLE_FILE_NAME,
)
SOLUBLE_LOCAL_FILE = st.secrets.get(
    "SOLUBLE_LOCAL_FILE",
    r"C:\Users\jiroa\Aoyama Dropbox\bulu jack\1共有　青山商店　本社\配車表-北海道-\aoベンチャーグレイン配車表.xlsx",
)
SOLUBLE_BACKUP_FOLDER = str(SOLUBLE_DROPBOX_FILE_PATH).rsplit("/", 1)[0] + "/Backups"
SOLUBLE_LOCATIONS = {
    "ノベルズ": {"usage": 3, "delivery": 4, "inventory": 5},
    "コスモアグリ": {"usage": 6, "delivery": 7, "inventory": 8},
}
SOLUBLE_CUSTOMER_NAMES = ("三谷牧場", "熊林牧場")
SOLUBLE_CUSTOMER_COLUMNS = {
    "customer_name": 2,       # B列
    "delivery_date": 5,       # E列：配達日
    "delivery_quantity": 6,   # F列：配達数量
    "next_delivery": 7,       # G列：次回配達予定（数式・表示のみ）
    "usage": 8,               # H列：使用数量/日
}
DISPATCH_REQUIRED_COLUMNS = [
    "発注番号",
    "引取日",
    "引取先",
    "商品名",
    "数量",
    "運送会社",
    "納品先",
    "着日",
]
DELIVERY_SHEET_NAME = "次回配達日"
SHEET1_CUSTOMER_COLUMN = 2   # B列：顧客名
SHEET1_HIRAGANA_COLUMN = 9   # I列：ひらがな
SHEET1_ADDRESS_COLUMN = 10   # J列：住所
SHEET1_MAP_COLUMN = 11       # K列：マップ位置
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "")
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_SECRET_KEY = st.secrets.get("SUPABASE_SECRET_KEY", "")
SUPABASE_SERVICE_ROLE_KEY = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", "")
SUPABASE_ANON_KEY = st.secrets.get("SUPABASE_ANON_KEY", "")
SUPABASE_NOTES_TABLE = st.secrets.get("SUPABASE_NOTES_TABLE", "notes")
SUPABASE_CUSTOMER_INFO_TABLE = st.secrets.get(
    "SUPABASE_CUSTOMER_INFO_TABLE",
    "customer_information",
)
LINE_STATUS_NOTE_PREFIX = "line_status_"
LINE_STATUS_BODY = "__LINE_CONNECTED__"
VOICE_INPUT_HELP = "スマホではキーボードのマイクを押して音声入力できます。"
PAST_PRODUCT_NOTE_PREFIX = "__past_product_note__:"

REQUIRED_COLUMNS = [
    "ID",
    "顧客名",
    "地域",
    "商品名",
    "使用数量/日",
    "次回配達予定",
    "残数",
    "ひらがな",
]

REQUIRED_COLUMN_CANDIDATES = {
    "ID": ["ID", "id", "顧客ID", "顧客コード", "コード", "No", "NO", "番号"],
    "顧客名": ["顧客名", "牧場名", "取引先名", "得意先名", "お客様名", "名前", "名称"],
    "地域": ["地域", "地区", "エリア", "住所", "市町村"],
    "商品名": ["商品名", "商品", "品名", "製品名"],
    "使用数量/日": ["使用数量/日", "使用数量", "使用量/日", "一日使用量", "数量/日", "日量"],
    "次回配達予定": ["次回配達予定", "配達予定日", "配送予定日", "配達日", "配送日", "納品日", "予定日", "日付"],
    "残数": ["残数", "残量", "残", "在庫", "残り"],
    "ひらがな": ["ひらがな", "ふりがな", "フリガナ", "かな", "カナ", "よみがな", "読み仮名"],
}

ADDRESS_COLUMN_CANDIDATES = [
    "住所",
    "所在地",
    "配達先住所",
    "配送先住所",
    "納品先住所",
    "顧客住所",
    "牧場住所",
]

MAP_LOCATION_COLUMN_CANDIDATES = [
    "マップ位置",
    "地図位置",
    "Googleマップ",
    "GoogleマップURL",
    "Google Maps",
    "Google Map",
    "マップURL",
    "地図URL",
    "位置情報",
    "緯度経度",
    "緯度・経度",
    "座標",
]


# =========================
# ログイン認証
# =========================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

# HTMLリンクで画面遷移するとStreamlitのセッションが切れる場合があるため、
# ログアウトするまではURLパラメータでもログイン状態を保持する。
try:
    if st.query_params.get("logged_in", "") == "1":
        st.session_state.authenticated = True
except Exception:
    pass

if not st.session_state.authenticated:
    st.title("🔒 顧客カルテ")

    # ログイン画面は下の共通CSSより前に停止するため、パスワード欄の枠線だけここで指定する。
    st.markdown(
        """
        <style>
        div[data-testid="stTextInputRootElement"],
        div[data-baseweb="input"] {
            border: 1px solid rgba(15, 23, 42, 0.28) !important;
            border-radius: 14px !important;
            background: #ffffff !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    if not APP_PASSWORD:
        st.error("APP_PASSWORD が設定されていません。Streamlit Cloud の Secrets に APP_PASSWORD を追加してください。")
        st.stop()

    password = st.text_input("パスワード", type="password")

    if st.button("ログイン"):
        if password == APP_PASSWORD:
            st.session_state.authenticated = True
            st.session_state.page = "home"
            st.session_state.selected_customer = None
            st.session_state.pop("customer_search_input", None)
            st.session_state.pop("customer_search_live", None)
            try:
                st.query_params.clear()
                st.query_params["logged_in"] = "1"
                st.query_params["page"] = "home"
            except Exception:
                pass
            st.rerun()
        else:
            st.error("パスワードが違います。")

    st.stop()



# =========================
# 共通CSS
# =========================
st.markdown(
    """
    <style>
    :root {
        --aoyama-bg: #f6f7fb;
        --aoyama-card: rgba(255, 255, 255, 0.92);
        --aoyama-line: rgba(15, 23, 42, 0.10);
        --aoyama-text: #172033;
        --aoyama-muted: #667085;
        --aoyama-blue: #2563eb;
        --aoyama-green: #0f766e;
        --aoyama-shadow: 0 12px 32px rgba(15, 23, 42, 0.08);
    }

    .stApp {
        background:
            radial-gradient(circle at top left, rgba(37, 99, 235, 0.13), transparent 28rem),
            radial-gradient(circle at top right, rgba(15, 118, 110, 0.11), transparent 24rem),
            linear-gradient(180deg, #f8fafc 0%, #eef2f7 100%);
        color: var(--aoyama-text);
    }

    [data-testid="stHeader"] {
        background: rgba(255, 255, 255, 0.72);
        backdrop-filter: blur(12px);
        border-bottom: 1px solid rgba(15, 23, 42, 0.06);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
    }
    [data-testid="stSidebar"] * {
        color: #f8fafc !important;
    }
    [data-testid="stSidebar"] hr {
        border-color: rgba(255, 255, 255, 0.15);
    }

    .block-container {
        padding-top: 2.2rem;
        padding-bottom: 3rem;
        max-width: 1120px;
    }

    h1, h2, h3 {
        letter-spacing: -0.03em;
    }

    .app-nav-link {
        display: flex;
        align-items: center;
        justify-content: center;
        width: 100%;
        min-height: 3.4rem;
        box-sizing: border-box;
        text-align: center;
        text-decoration: none !important;
        padding: 0.75rem 0.9rem;
        margin: 0.32rem 0;
        border: 1px solid rgba(15, 23, 42, 0.08);
        border-radius: 16px;
        color: #172033 !important;
        background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
        font-weight: 800;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.07);
        transition: transform 0.14s ease, box-shadow 0.14s ease, border-color 0.14s ease;
    }
    .app-nav-link:hover {
        transform: translateY(-1px);
        border-color: rgba(37, 99, 235, 0.32);
        box-shadow: 0 14px 32px rgba(37, 99, 235, 0.13);
        background: linear-gradient(180deg, #ffffff 0%, #eff6ff 100%);
    }

    [data-testid="stSidebar"] .app-nav-link {
        justify-content: flex-start;
        min-height: 2.9rem;
        color: #f8fafc !important;
        background: rgba(255, 255, 255, 0.08);
        border-color: rgba(255, 255, 255, 0.12);
        box-shadow: none;
    }
    [data-testid="stSidebar"] .app-nav-link:hover {
        background: rgba(255, 255, 255, 0.16);
        border-color: rgba(255, 255, 255, 0.28);
        transform: none;
    }

    .stButton > button {
        border-radius: 14px !important;
        border: 1px solid rgba(15, 23, 42, 0.10) !important;
        background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%) !important;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.07);
        font-weight: 700 !important;
    }
    .stButton > button,
    .stButton > button *,
    .stFormSubmitButton > button,
    .stFormSubmitButton > button * {
        color: #172033 !important;
        -webkit-text-fill-color: #172033 !important;
        opacity: 1 !important;
    }
    .stButton > button:hover {
        border-color: rgba(37, 99, 235, 0.35) !important;
        box-shadow: 0 10px 24px rgba(37, 99, 235, 0.12);
    }

    .customer-information-row {
        display: grid;
        grid-template-columns: minmax(6.5rem, 1.4fr) minmax(0, 3fr);
        align-items: start;
        column-gap: 0.9rem;
        padding: 0.35rem 0;
        color: #172033;
        font-size: 1rem;
        line-height: 1.55;
    }
    .customer-information-label,
    .customer-information-content {
        margin: 0;
        padding: 0;
        overflow-wrap: anywhere;
    }
    .customer-information-label {
        font-weight: 800;
    }

    .customer-directory {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 0.65rem;
        margin-top: 0.75rem;
    }
    .customer-directory-item {
        display: flex;
        flex-direction: column;
        justify-content: center;
        min-height: 4.3rem;
        box-sizing: border-box;
        padding: 0.72rem 0.9rem;
        border: 1px solid rgba(15, 23, 42, 0.10);
        border-radius: 14px;
        background: rgba(255, 255, 255, 0.92);
        color: #172033 !important;
        text-decoration: none !important;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.06);
    }
    .customer-directory-item:hover {
        border-color: rgba(37, 99, 235, 0.35);
        background: #eff6ff;
    }
    .customer-directory-name {
        color: #172033;
        font-size: 1rem;
        font-weight: 800;
        line-height: 1.4;
        overflow-wrap: anywhere;
    }
    .customer-directory-meta {
        margin-top: 0.22rem;
        color: #667085;
        font-size: 0.82rem;
        line-height: 1.35;
        overflow-wrap: anywhere;
    }

    .stTextInput input {
        border-radius: 14px !important;
        border: 1px solid rgba(15, 23, 42, 0.13) !important;
        background: rgba(255,255,255,0.92) !important;
        padding: 0.72rem 0.9rem !important;
    }

    [data-testid="stMetricValue"], .stCaptionContainer {
        color: var(--aoyama-muted);
    }

    hr {
        margin: 1.4rem 0;
        border-color: rgba(15, 23, 42, 0.08);
    }

    .note-card {
        border: 1px solid rgba(15, 23, 42, 0.10);
        border-radius: 16px;
        background: rgba(255, 255, 255, 0.92);
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06);
        padding: 0.9rem 1rem;
        margin: 0.65rem 0;
    }
    .note-meta {
        color: #667085;
        font-size: 0.86rem;
        margin-bottom: 0.35rem;
    }
    .note-body {
        color: #172033;
        font-size: 1rem;
        line-height: 1.65;
        white-space: pre-wrap;
        word-break: break-word;
    }

    .customer-name-row {
        display: flex;
        align-items: center;
        gap: 0.45rem;
        margin: 0.1rem 0 0.45rem;
        color: #172033;
        font-size: 1.25rem;
        font-weight: 700;
    }
    .line-status {
        font-size: 0.85rem;
        font-weight: 600;
        line-height: 1;
        white-space: nowrap;
    }
    .line-status-connected {
        color: #4f8f68;
    }
    .line-status-disconnected {
        color: #98a2b3;
    }
    .customer-detail-name-row {
        font-size: 1.65rem;
        margin-top: 0.15rem;
    }
    .line-detail-static {
        color: #4f8f68;
        font-size: 0.85rem;
        font-weight: 600;
        margin-top: 0.8rem;
        white-space: nowrap;
    }
    [data-testid="stPopover"] button {
        min-height: 0 !important;
        margin-top: 0.35rem !important;
        padding: 0.3rem 0.55rem !important;
        color: #667085 !important;
        font-size: 0.85rem !important;
        box-shadow: none !important;
        white-space: nowrap !important;
    }
    [data-testid="stPopover"] button p {
        white-space: nowrap !important;
    }


    @media (max-width: 640px) {
        .block-container {
            padding-left: 0.8rem;
            padding-right: 0.8rem;
            padding-top: 1.4rem;
        }

        /* Streamlit 1.38で列が縦に崩れないようにする */
        [data-testid="stHorizontalBlock"] {
            flex-wrap: nowrap !important;
            align-items: center !important;
            gap: 0.6rem !important;
        }
        [data-testid="stHorizontalBlock"] > div {
            min-width: 0 !important;
        }

        .app-nav-link {
            min-height: 3.1rem;
            border-radius: 14px;
            font-size: 0.95rem;
        }
        h1 {
            font-size: 1.55rem !important;
        }
        h2 {
            font-size: 1.25rem !important;
        }
        h3 {
            font-size: 1.08rem !important;
        }

        .customer-information-row {
            grid-template-columns: minmax(5.5rem, 1.35fr) minmax(0, 2.65fr);
            column-gap: 0.7rem;
            align-items: baseline;
        }

        .customer-directory {
            grid-template-columns: 1fr;
            gap: 0.55rem;
        }
        .customer-directory-item {
            min-height: 3.9rem;
        }
    }
    
    /* サイドバーの更新ボタン文字を白背景でも見える色に固定 */
    [data-testid="stSidebar"] .stButton > button,
    [data-testid="stSidebar"] .stButton > button *,
    [data-testid="stSidebar"] button[kind] *,
    [data-testid="stSidebar"] button[kind] {
        color: #1f2937 !important;
        -webkit-text-fill-color: #1f2937 !important;
        opacity: 1 !important;
    }

</style>
    """,
    unsafe_allow_html=True,
)

# =========================
# 表示用の整形
# =========================
def clean_value(value, blank_text="未設定"):
    if value is None:
        return blank_text

    if isinstance(value, float) and math.isnan(value):
        return blank_text

    text = str(value).strip()

    if text == "" or text.lower() == "nan":
        return blank_text

    if text.startswith("#"):
        return blank_text

    return text


def render_customer_name_with_line(customer_name, connected, detail=False):
    """顧客名の横に控えめなLINE ○／×を表示する。"""
    line_mark = "○" if connected else "×"
    line_class = "line-status-connected" if connected else "line-status-disconnected"
    detail_class = " customer-detail-name-row" if detail else ""
    st.markdown(
        f'<div class="customer-name-row{detail_class}">'
        f'<span>👤 {html.escape(clean_value(customer_name))}</span>'
        f'<span class="line-status {line_class}">LINE {line_mark}</span>'
        "</div>",
        unsafe_allow_html=True,
    )


def format_date(value):
    if value is None:
        return "未設定"

    if isinstance(value, float) and math.isnan(value):
        return "未設定"

    text = str(value).strip()

    if text == "" or text.lower() == "nan" or text.startswith("#"):
        return "未設定"

    try:
        dt = pd.to_datetime(value)
        return f"{dt.year}/{dt.month}/{dt.day}"
    except Exception:
        return text


def to_date(value):
    if value is None:
        return None

    if isinstance(value, float) and math.isnan(value):
        return None

    text = str(value).strip()

    if text == "" or text.lower() == "nan" or text.startswith("#"):
        return None

    try:
        return pd.to_datetime(value).date()
    except Exception:
        return None


def format_number(value, decimal=1, blank_text="未設定"):
    if value is None:
        return blank_text

    if isinstance(value, float) and math.isnan(value):
        return blank_text

    text = str(value).strip()

    if text == "" or text.lower() == "nan":
        return blank_text

    if text.startswith("#"):
        return "計算不可"

    try:
        num = float(value)
        if num.is_integer():
            return str(int(num))
        return f"{num:.{decimal}f}"
    except Exception:
        return text


def is_blank_or_zero(value):
    """空白・NaN・0ならTrue。使用数量/日を非表示にする判定用。

    Excelから「0」「0.0」「０」「０．０」「0 kg」のような文字列で来ても
    0として扱えるように少し広めに判定する。
    """
    if value is None:
        return True

    if isinstance(value, float) and math.isnan(value):
        return True

    text = str(value).strip()

    if text == "" or text.lower() == "nan" or text.startswith("#"):
        return True

    # 全角数字・全角小数点・カンマを整理
    normalized = text.translate(str.maketrans({
        "０": "0", "１": "1", "２": "2", "３": "3", "４": "4",
        "５": "5", "６": "6", "７": "7", "８": "8", "９": "9",
        "．": ".", "，": ",",
    })).replace(",", "")

    # 単位などが付いていても、先頭の数値だけ取り出して判定
    import re
    match = re.match(r"^[-+]?\d+(?:\.\d+)?", normalized)
    if not match:
        return False

    try:
        return float(match.group(0)) == 0
    except Exception:
        return False


def find_existing_column(df, candidates):
    """候補名の中から、Excelに存在する列名を1つ探す"""
    for candidate in candidates:
        if candidate in df.columns:
            return candidate
    return None


def find_required_column_mapping(column_names):
    """Excelの列名候補を、アプリ内で使う標準列名へ対応させる"""
    normalized_columns = [str(col).strip() for col in column_names]
    mapping = {}

    for required_column, candidates in REQUIRED_COLUMN_CANDIDATES.items():
        for candidate in candidates:
            if candidate in normalized_columns:
                mapping[required_column] = candidate
                break

    return mapping


def get_first_nonblank_column_value(df, column_name):
    """指定列から最初の空でない値を取り出す"""
    if not column_name or column_name not in df.columns:
        return ""

    for value in df[column_name].tolist():
        text = clean_value(value, blank_text="")
        if text:
            return text

    return ""


def parse_lat_lng(value):
    """「緯度,経度」形式なら緯度経度を返す"""
    import re

    try:
        text = clean_value(value, blank_text="")
        if not text:
            return None

        normalized = text.translate(str.maketrans({
            "０": "0", "１": "1", "２": "2", "３": "3", "４": "4",
            "５": "5", "６": "6", "７": "7", "８": "8", "９": "9",
            "．": ".", "，": ",",
        }))
        normalized = normalized.replace("、", ",")

        match = re.match(
            r"^\s*(?:緯度\s*[:：]?\s*)?([-+]?\d+(?:\.\d+)?)\s*[, ]\s*(?:経度\s*[:：]?\s*)?([-+]?\d+(?:\.\d+)?)\s*$",
            normalized,
        )
        if not match:
            return None

        lat = float(match.group(1))
        lng = float(match.group(2))
    except Exception:
        return None

    if -90 <= lat <= 90 and -180 <= lng <= 180:
        return lat, lng

    return None


def build_google_maps_url(value):
    """住所・緯度経度・URLからGoogleマップで開くURLを作る"""
    try:
        text = clean_value(value, blank_text="")
        if not text:
            return ""

        parsed = urllib.parse.urlparse(text)
        if parsed.scheme in ("http", "https") and parsed.netloc:
            return text

        lat_lng = parse_lat_lng(text)
        if lat_lng:
            lat, lng = lat_lng
            return f"https://www.google.com/maps/search/?api=1&query={lat},{lng}"

        query = urllib.parse.quote(text)
        return f"https://www.google.com/maps/search/?api=1&query={query}"
    except Exception:
        return ""


def get_customer_map_info(detail):
    """顧客詳細で使う住所・地図情報を取り出す"""
    try:
        map_column = find_existing_column(detail, MAP_LOCATION_COLUMN_CANDIDATES)
        address_column = find_existing_column(detail, ADDRESS_COLUMN_CANDIDATES)
        map_value = get_first_nonblank_column_value(detail, map_column)
        address_value = get_first_nonblank_column_value(detail, address_column)
        target_value = map_value or address_value

        if not target_value:
            return None

        display_value = address_value or map_value
        display_label = "住所" if address_value else "マップ位置"
        target_column = map_column if map_value else address_column

        return {
            "display_label": display_label,
            "display_value": display_value,
            "target_column": target_column,
            "map_url": build_google_maps_url(target_value),
        }
    except Exception:
        return None


def show_google_maps_button(url):
    """Googleマップを開くボタンを表示する"""
    safe_url = clean_value(url, blank_text="")
    if not safe_url:
        return

    try:
        st.link_button("📍 Googleマップ", safe_url)
    except Exception:
        st.markdown(f"[📍 Googleマップ]({safe_url})")


def find_date_column(df):
    """配車カレンダーで使う日付列を探す"""
    candidates = [
        "配車日",
        "配車予定日",
        "配達日",
        "配達予定日",
        "配送日",
        "配送予定日",
        "納品日",
        "予定日",
        "日付",
        "次回配達予定",
    ]

    found = find_existing_column(df, candidates)
    if found:
        return found

    keywords = ["配車", "配達", "配送", "納品", "予定", "日付"]

    for col in df.columns:
        col_text = str(col)
        if "数量" in col_text:
            continue

        if any(keyword in col_text for keyword in keywords):
            parsed = pd.to_datetime(df[col], errors="coerce")
            if parsed.notna().any():
                return col

    return None


# =========================
# Dropbox API
# =========================
def make_dropbox_api_arg(path_or_id):
    """
    Dropbox-API-Argヘッダー用の文字列を作る。
    日本語パスでも送れるようにASCII化してからlatin1で渡す。
    """
    return json.dumps({"path": path_or_id}, ensure_ascii=True).encode("utf-8").decode("latin1")


@st.cache_data(ttl=3300, show_spinner=False)
def request_dropbox_access_token(app_key, app_secret, refresh_token):
    """Refresh Tokenから短期アクセストークンを取得する"""
    url = "https://api.dropboxapi.com/oauth2/token"
    data = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
    }

    try:
        response = requests.post(
            url,
            data=data,
            auth=(app_key, app_secret),
            timeout=30,
        )
    except Exception as e:
        return None, None, str(e)

    if response.status_code != 200:
        return None, response.status_code, response.text

    return response.json().get("access_token"), None, None


def has_dropbox_auth_config():
    return bool(
        DROPBOX_APP_KEY
        or DROPBOX_APP_SECRET
        or DROPBOX_REFRESH_TOKEN
        or DROPBOX_ACCESS_TOKEN
    )


def get_dropbox_access_token():
    """
    Streamlit CloudではRefresh Token方式で短期アクセストークンを取得する。
    DROPBOX_ACCESS_TOKENは既存環境をすぐ壊さないための移行用。
    """
    if DROPBOX_APP_KEY or DROPBOX_APP_SECRET or DROPBOX_REFRESH_TOKEN:
        missing = []
        if not DROPBOX_APP_KEY:
            missing.append("DROPBOX_APP_KEY")
        if not DROPBOX_APP_SECRET:
            missing.append("DROPBOX_APP_SECRET")
        if not DROPBOX_REFRESH_TOKEN:
            missing.append("DROPBOX_REFRESH_TOKEN")

        if missing:
            st.error("Dropbox API設定が不足しています。")
            st.write("Streamlit Cloud の Secrets に以下を追加してください。")
            st.code("\n".join(missing))
            st.stop()

        access_token, status_code, error_text = request_dropbox_access_token(
            DROPBOX_APP_KEY,
            DROPBOX_APP_SECRET,
            DROPBOX_REFRESH_TOKEN,
        )

        if not access_token:
            st.error("Dropboxのアクセストークン更新に失敗しました。")
            if status_code:
                st.write(f"Dropboxからの応答コード：{status_code}")
            st.code(error_text or "access_token が返りませんでした。")
            st.stop()

        return access_token

    if DROPBOX_ACCESS_TOKEN:
        return DROPBOX_ACCESS_TOKEN

    st.error("Dropbox API設定が不足しています。")
    st.write("secrets.toml に DROPBOX_APP_KEY / DROPBOX_APP_SECRET / DROPBOX_REFRESH_TOKEN を設定してください。")
    st.stop()


def download_dropbox_file(path_or_id, access_token):
    """Dropbox APIでExcelをダウンロードする"""
    url = "https://content.dropboxapi.com/2/files/download"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Dropbox-API-Arg": make_dropbox_api_arg(path_or_id),
    }

    try:
        response = requests.post(url, headers=headers, timeout=30)
    except Exception as e:
        st.error("Dropbox APIへの接続に失敗しました。")
        st.exception(e)
        st.stop()

    if response.status_code != 200:
        return None, response

    return response.content, response


@st.cache_data(ttl=3600, show_spinner=False)
def get_dropbox_root_info(access_token):
    """チームDropboxを含むルート名前空間と、メンバーフォルダのパスを取得する。"""
    response = requests.post(
        "https://api.dropboxapi.com/2/users/get_current_account",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
        data="null",
        timeout=30,
    )
    if response.status_code != 200:
        return None, response

    root_info = response.json().get("root_info", {})
    if not root_info.get("root_namespace_id"):
        return None, response
    return root_info, response


def download_dropbox_team_file(path_or_id, access_token):
    """チームルートを明示し、メンバーフォルダ内のファイルを取得する。"""
    root_info, response = get_dropbox_root_info(access_token)
    if root_info is None:
        return None, response

    rooted_path = str(path_or_id or "").strip()
    home_path = str(root_info.get("home_path") or "").rstrip("/")
    if home_path and rooted_path.startswith("/") and not rooted_path.startswith(home_path + "/"):
        rooted_path = home_path + rooted_path

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Dropbox-API-Path-Root": json.dumps(
            {".tag": "root", "root": root_info["root_namespace_id"]},
            ensure_ascii=False,
        ),
        "Dropbox-API-Arg": make_dropbox_api_arg(rooted_path),
    }
    try:
        response = requests.post(
            "https://content.dropboxapi.com/2/files/download",
            headers=headers,
            timeout=60,
        )
    except Exception as error:
        raise RuntimeError("Dropboxのチームルートへ接続できませんでした。") from error

    if response.status_code != 200:
        return None, response
    return response.content, response


def dropbox_error_text(response):
    """Dropboxのエラーを画面表示できる文字列にする。"""
    if response is None:
        return "Dropboxから応答がありませんでした。"
    try:
        body = json.dumps(response.json(), ensure_ascii=False, indent=2)
    except Exception:
        body = response.text
    return f"HTTP {response.status_code}\n{body}"


def get_download_revision(response):
    """download応答に含まれるファイルrevを取り出す。"""
    try:
        metadata = json.loads(response.headers.get("Dropbox-API-Result", "{}"))
        return metadata.get("rev", "")
    except Exception:
        return ""


def upload_dropbox_file(path, content, access_token, mode="add", rev=""):
    """競合ファイルを作らずDropboxへファイルをアップロードする。"""
    mode_arg = {".tag": "update", "update": rev} if mode == "update" else mode
    api_arg = {
        "path": path,
        "mode": mode_arg,
        "autorename": False,
        "mute": False,
        "strict_conflict": True,
    }
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/octet-stream",
        "Dropbox-API-Arg": json.dumps(api_arg, ensure_ascii=True).encode("utf-8").decode("latin1"),
    }
    try:
        return requests.post(
            "https://content.dropboxapi.com/2/files/upload",
            headers=headers,
            data=content,
            timeout=120,
        )
    except Exception as exc:
        raise RuntimeError(f"Dropboxへのアップロードに失敗しました: {exc}") from exc


def call_dropbox_rpc(endpoint, payload, access_token):
    """DropboxのメタデータAPIを呼び出す。"""
    try:
        return requests.post(
            f"https://api.dropboxapi.com/2/{endpoint}",
            headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
            json=payload,
            timeout=30,
        )
    except Exception as exc:
        raise RuntimeError(f"Dropbox APIへの接続に失敗しました: {exc}") from exc


def get_dropbox_revision(path, access_token):
    """Dropboxファイルの現在のrevだけを軽量に取得する。"""
    response = call_dropbox_rpc("files/get_metadata", {"path": path}, access_token)
    if response.status_code != 200:
        raise RuntimeError("Dropboxのファイル情報を取得できませんでした。\n" + dropbox_error_text(response))
    return str(response.json().get("rev", ""))


def ensure_dropbox_backup_folder(access_token):
    """Backupsフォルダがなければ作成する。既に存在する場合は成功扱いにする。"""
    response = call_dropbox_rpc(
        "files/create_folder_v2",
        {"path": DROPBOX_BACKUP_FOLDER, "autorename": False},
        access_token,
    )
    if response.status_code == 200:
        return
    if response.status_code == 409:
        try:
            error_data = response.json()
            summary = str(error_data.get("error_summary", ""))
            # path/conflict/folder/ は「同名フォルダが既にある」という正常状態。
            if "conflict" in summary and "folder" in summary:
                return
        except Exception:
            pass
    raise RuntimeError(
        "Dropboxにバックアップフォルダを作成できませんでした。\n"
        + dropbox_error_text(response)
    )


def trim_old_dropbox_backups(access_token, keep=30):
    """対象ブックのバックアップを新しい順にkeep件だけ残す。"""
    entries = []
    response = call_dropbox_rpc(
        "files/list_folder",
        {"path": DROPBOX_BACKUP_FOLDER, "recursive": False, "include_deleted": False},
        access_token,
    )
    if response.status_code != 200:
        return f"バックアップ一覧を取得できませんでした。\n{dropbox_error_text(response)}"

    data = response.json()
    entries.extend(data.get("entries", []))
    while data.get("has_more"):
        response = call_dropbox_rpc("files/list_folder/continue", {"cursor": data["cursor"]}, access_token)
        if response.status_code != 200:
            return f"バックアップ一覧の続きが取得できませんでした。\n{dropbox_error_text(response)}"
        data = response.json()
        entries.extend(data.get("entries", []))

    pattern = re.compile(r"^配車予定 次郎_\d{8}_\d{6}(?:_\d+)?\.xlsm$")
    backups = [item for item in entries if item.get(".tag") == "file" and pattern.match(item.get("name", ""))]
    backups.sort(key=lambda item: (item.get("server_modified", ""), item.get("name", "")), reverse=True)
    warnings = []
    for item in backups[keep:]:
        delete_response = call_dropbox_rpc("files/delete_v2", {"path": item["path_lower"]}, access_token)
        if delete_response.status_code != 200:
            warnings.append(item.get("name", "不明なファイル"))
    return "削除できなかった古いバックアップ: " + ", ".join(warnings) if warnings else ""


def normalize_match_value(value):
    return clean_value(value, blank_text="").strip()


def find_sheet1_customer_rows(workbook, customer_name):
    """
    Sheet1のB列は顧客名の値ではなく、例: =次回配達日!B7 の数式。
    数式の参照先（次回配達日シートB列）をたどって対象顧客の行を返す。
    値が直接入っている場合にも対応する。
    """
    if SHEET_NAME not in workbook.sheetnames or DELIVERY_SHEET_NAME not in workbook.sheetnames:
        return []

    sheet1 = workbook[SHEET_NAME]
    delivery_ws = workbook[DELIVERY_SHEET_NAME]
    target = normalize_match_value(customer_name)
    rows = []
    formula_pattern = re.compile(
        r"^=\s*(?:'次回配達日'|次回配達日)!\$?B\$?(\d+)\s*$",
        re.IGNORECASE,
    )

    for row in range(2, sheet1.max_row + 1):
        value = sheet1.cell(row, SHEET1_CUSTOMER_COLUMN).value

        # 顧客名が直接入っている場合
        if normalize_match_value(value) == target:
            rows.append(row)
            continue

        # =次回配達日!B7 のような数式の場合
        if isinstance(value, str):
            match = formula_pattern.match(value.strip())
            if match:
                source_row = int(match.group(1))
                source_customer = normalize_match_value(
                    delivery_ws.cell(source_row, 2).value
                )
                if source_customer == target:
                    rows.append(row)

    return rows


def find_header_column_in_worksheet(ws, candidates, max_rows=50):
    """見出し行を走査し、候補名に完全一致する列番号を返す。"""
    candidate_set = {str(item).strip() for item in candidates}
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows)):
        for cell in row:
            if normalize_match_value(cell.value) in candidate_set:
                return cell.column
    return None


@st.cache_data(ttl=60, show_spinner=False)
def read_edit_values_from_bytes(content, customer_name, product_name):
    """最新ブックから編集欄の現在値を取得する。"""
    workbook = load_workbook(BytesIO(content), keep_vba=True, data_only=False, read_only=False)
    try:
        if DELIVERY_SHEET_NAME not in workbook.sheetnames or SHEET_NAME not in workbook.sheetnames:
            raise ValueError("必要なシート（次回配達日 または Sheet1）が見つかりません。")
        delivery_ws = workbook[DELIVERY_SHEET_NAME]
        matches = [
            row for row in range(1, delivery_ws.max_row + 1)
            if normalize_match_value(delivery_ws.cell(row, 2).value) == customer_name
            and normalize_match_value(delivery_ws.cell(row, 5).value) == product_name
        ]
        product_values = {}
        if len(matches) == 1:
            row = matches[0]
            product_values = {
                "メーカー": delivery_ws.cell(row, 6).value,
                "本数": delivery_ws.cell(row, 8).value,
                "kg/本": delivery_ws.cell(row, 9).value,
                "配達日": delivery_ws.cell(row, 10).value,
            }

        customer_ws = workbook[SHEET_NAME]
        customer_rows = find_sheet1_customer_rows(workbook, customer_name)
        first_row = customer_rows[0] if customer_rows else None
        return {
            **product_values,
            "住所": customer_ws.cell(first_row, SHEET1_ADDRESS_COLUMN).value if first_row else None,
            "マップ位置": customer_ws.cell(first_row, SHEET1_MAP_COLUMN).value if first_row else None,
            "商品一致件数": len(matches),
            "顧客一致件数": len(customer_rows),
        }
    finally:
        workbook.close()


@st.cache_data(ttl=60, show_spinner=False)
def read_customer_edit_bundle_from_bytes(content, customer_name):
    """顧客詳細に必要な地図と全商品の編集値を、Excelを1回開いてまとめて読む。"""
    workbook = load_workbook(BytesIO(content), keep_vba=True, data_only=False, read_only=False)
    try:
        if DELIVERY_SHEET_NAME not in workbook.sheetnames or SHEET_NAME not in workbook.sheetnames:
            raise ValueError("必要なシート（次回配達日 または Sheet1）が見つかりません。")

        delivery_ws = workbook[DELIVERY_SHEET_NAME]
        target = normalize_match_value(customer_name)
        product_rows = {}
        for row in range(1, delivery_ws.max_row + 1):
            if normalize_match_value(delivery_ws.cell(row, 2).value) != target:
                continue
            product = normalize_match_value(delivery_ws.cell(row, 5).value)
            if product:
                product_rows.setdefault(product, []).append(row)

        products = {}
        for product, rows in product_rows.items():
            first_row = rows[0]
            products[product] = {
                "メーカー": delivery_ws.cell(first_row, 6).value,
                "本数": delivery_ws.cell(first_row, 8).value,
                "kg/本": delivery_ws.cell(first_row, 9).value,
                "配達日": delivery_ws.cell(first_row, 10).value,
                "商品一致件数": len(rows),
            }

        customer_rows = find_sheet1_customer_rows(workbook, customer_name)
        first_customer_row = customer_rows[0] if customer_rows else None
        ws = workbook[SHEET_NAME]
        map_values = {
            "住所": ws.cell(first_customer_row, SHEET1_ADDRESS_COLUMN).value if first_customer_row else None,
            "マップ位置": ws.cell(first_customer_row, SHEET1_MAP_COLUMN).value if first_customer_row else None,
            "顧客一致件数": len(customer_rows),
        }
        for values in products.values():
            values.update(map_values)
        return {"map": map_values, "products": products}
    finally:
        workbook.close()


def parse_optional_nonnegative_number(text, integer=False):
    value = str(text).strip().translate(str.maketrans("０１２３４５６７８９．，", "0123456789.,"))
    # 音声入力で付きやすい単位を許可する。
    value = re.sub(r"\s*(?:本|kg|KG|ｋｇ|キロ|キログラム)\s*$", "", value, flags=re.IGNORECASE)
    value = value.replace(",", "")
    if value == "":
        return None
    try:
        number = float(value)
    except Exception as exc:
        raise ValueError("数値で入力してください。") from exc
    if not math.isfinite(number) or number < 0:
        raise ValueError("0以上の数値で入力してください。")
    if integer and not number.is_integer():
        raise ValueError("整数で入力してください。")
    return int(number) if integer else number


def validate_map_location(value):
    text = str(value).strip()
    if not text:
        return ""
    parsed = urllib.parse.urlparse(text)
    if parsed.scheme or "://" in text:
        if parsed.scheme not in ("http", "https") or not parsed.netloc:
            raise ValueError("マップ位置のURLが正しくありません。")
    # 数字と区切りだけで座標らしく見える入力に限って、緯度経度として検証する。
    # 通常の住所・施設名にカンマが含まれていても許可する。
    coordinate_like = bool(re.match(r"^[\s+\-\d０-９.,，．、:：緯度経]+$", text))
    if coordinate_like and not parse_lat_lng(text):
        raise ValueError("緯度,経度は例のように入力してください（43.123456, 143.123456）。")
    return text


def same_excel_value(old, new):
    if old is None and new in (None, ""):
        return True
    if isinstance(old, (datetime, date)) and isinstance(new, (datetime, date)):
        return old.date() == new.date() if isinstance(old, datetime) else old == (new.date() if isinstance(new, datetime) else new)
    return old == new


def enable_excel_recalculation(workbook):
    """Excelで開いたときに残数・次回配達予定などの数式を必ず再計算させる。"""
    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except Exception:
        # openpyxlの版によって属性名が異なる場合でも、セル保存は続行する。
        pass


def verify_remote_workbook_changes(content, changed_cells):
    """Dropboxから再取得したブックに、変更値が実在することを検証する。"""
    workbook = load_workbook(BytesIO(content), keep_vba=True, data_only=False, read_only=False)
    try:
        for sheet, row, column, expected in changed_cells:
            if sheet not in workbook.sheetnames:
                raise RuntimeError(f"Dropbox保存後の確認でシート「{sheet}」が見つかりません。")
            cell = workbook[sheet].cell(row, column)
            if not same_excel_value(cell.value, expected):
                raise RuntimeError(
                    f"Dropbox保存後の確認で {sheet}!{cell.coordinate} が更新されていません。"
                )
    finally:
        workbook.close()


def confirm_dropbox_upload(target_path, access_token, changed_cells):
    """アップロード後にDropbox本体を読み直し、保存完了を保証する。"""
    uploaded_content, response = download_dropbox_file(target_path, access_token)
    if uploaded_content is None:
        raise RuntimeError(
            "Dropboxへ送信後、更新済みExcelを再取得できませんでした。\n"
            + dropbox_error_text(response)
        )
    verify_remote_workbook_changes(uploaded_content, changed_cells)


def update_workbook_bytes(original_content, customer_name, product_name, proposed):
    """指定された6列の値だけ変更し、再オープン検証したbytesを返す。"""
    workbook = load_workbook(BytesIO(original_content), keep_vba=True, data_only=False, read_only=False)
    original_sheets = list(workbook.sheetnames)
    changed_cells = []
    try:
        if DELIVERY_SHEET_NAME not in workbook.sheetnames or SHEET_NAME not in workbook.sheetnames:
            raise ValueError("必要なシート（次回配達日 または Sheet1）が見つかりません。")
        delivery_ws = workbook[DELIVERY_SHEET_NAME]
        product_rows = [
            row for row in range(1, delivery_ws.max_row + 1)
            if normalize_match_value(delivery_ws.cell(row, 2).value) == customer_name
            and normalize_match_value(delivery_ws.cell(row, 5).value) == product_name
        ]
        if not product_rows:
            raise ValueError("顧客名・商品名が一致する行が見つかりません。")
        if len(product_rows) > 1:
            raise ValueError("同じ顧客名・商品名の行が複数見つかりました。")

        product_row = product_rows[0]
        for label, column in {"メーカー": 6, "本数": 8, "kg/本": 9, "配達日": 10}.items():
            cell = delivery_ws.cell(product_row, column)
            new_value = proposed[label]
            if not same_excel_value(cell.value, new_value):
                cell.value = new_value
                changed_cells.append((DELIVERY_SHEET_NAME, product_row, column, new_value))

        customer_ws = workbook[SHEET_NAME]
        customer_rows = find_sheet1_customer_rows(workbook, customer_name)
        if not customer_rows:
            raise ValueError("Sheet1のB列数式が参照する顧客名に一致する行が見つかりません。")
        for row in customer_rows:
            for label, column in {"住所": SHEET1_ADDRESS_COLUMN, "マップ位置": SHEET1_MAP_COLUMN}.items():
                cell = customer_ws.cell(row, column)
                new_value = proposed[label]
                if not same_excel_value(cell.value, new_value):
                    cell.value = new_value
                    changed_cells.append((SHEET_NAME, row, column, new_value))

        if not changed_cells:
            raise ValueError("変更された項目がありません。")
        enable_excel_recalculation(workbook)
        output = BytesIO()
        workbook.save(output)
    finally:
        workbook.close()

    saved_content = output.getvalue()
    verified = load_workbook(BytesIO(saved_content), keep_vba=True, data_only=False, read_only=False)
    try:
        if list(verified.sheetnames) != original_sheets:
            raise ValueError("保存後にシート構成が変わったため、更新を中止しました。")
        if DELIVERY_SHEET_NAME not in verified.sheetnames or SHEET_NAME not in verified.sheetnames:
            raise ValueError("保存後の検証で必要なシートが見つかりません。")
        if verified.vba_archive is None:
            raise ValueError("保存後の検証でVBAプロジェクトを確認できません。")
        for sheet, row, column, expected in changed_cells:
            actual = verified[sheet].cell(row, column).value
            if not same_excel_value(actual, expected):
                raise ValueError(f"保存後の検証で{sheet}!{verified[sheet].cell(row, column).coordinate}の値が一致しません。")
    finally:
        verified.close()
    return saved_content, changed_cells


def save_customer_excel_changes(customer_name, product_name, proposed):
    """バックアップ、検証、rev競合防止を含む一連の保存処理。"""
    access_token = get_dropbox_access_token()
    target_path = get_dropbox_file_path()
    original_content, download_response = download_dropbox_file(target_path, access_token)
    if original_content is None:
        raise RuntimeError("最新のExcelを取得できませんでした。\n" + dropbox_error_text(download_response))
    revision = get_download_revision(download_response)
    if not revision:
        raise RuntimeError("Dropboxのrevを取得できないため、安全のため更新を中止しました。")

    ensure_dropbox_backup_folder(access_token)
    timestamp = get_jst_now().strftime("%Y%m%d_%H%M%S_%f")
    backup_path = f"{DROPBOX_BACKUP_FOLDER}/配車予定 次郎_{timestamp}.xlsm"
    backup_response = upload_dropbox_file(backup_path, original_content, access_token, mode="add")
    if backup_response.status_code != 200:
        raise RuntimeError("バックアップを作成できないため、本番ファイルは更新しません。\n" + dropbox_error_text(backup_response))

    saved_content, changed_cells = update_workbook_bytes(original_content, customer_name, product_name, proposed)
    upload_response = upload_dropbox_file(target_path, saved_content, access_token, mode="update", rev=revision)
    if upload_response.status_code == 409:
        raise RuntimeError("PCまたは別端末でExcelが更新されています。再読み込みしてからやり直してください")
    if upload_response.status_code != 200:
        raise RuntimeError("本番Excelを更新できませんでした。必要なDropbox権限は files.content.write です。\n" + dropbox_error_text(upload_response))

    # 「APIが200を返した」だけで成功扱いにせず、Dropbox上の実ファイルを再確認する。
    confirm_dropbox_upload(target_path, access_token, changed_cells)

    cleanup_warning = trim_old_dropbox_backups(access_token, keep=30)
    st.cache_data.clear()
    return {
        "backup_path": backup_path,
        "updated_at": get_jst_now(),
        "changed_cells": changed_cells,
        "cleanup_warning": cleanup_warning,
    }



@st.cache_data(ttl=60, show_spinner=False)
def read_customer_map_values_from_bytes(content, customer_name):
    """Sheet1の表示値で顧客を探し、J列住所・K列マップ位置を返す。"""
    workbook = load_workbook(BytesIO(content), keep_vba=True, data_only=False, read_only=False)
    try:
        if SHEET_NAME not in workbook.sheetnames:
            raise ValueError("Sheet1が見つかりません。")
        ws = workbook[SHEET_NAME]
        rows = find_sheet1_customer_rows(workbook, customer_name)
        if not rows:
            raise ValueError("Sheet1のB列に表示されている顧客名と一致する行が見つかりません。")
        first_row = rows[0]
        return {
            "住所": ws.cell(first_row, SHEET1_ADDRESS_COLUMN).value,
            "マップ位置": ws.cell(first_row, SHEET1_MAP_COLUMN).value,
            "顧客一致件数": len(rows),
        }
    finally:
        workbook.close()


def update_customer_map_workbook_bytes(original_content, customer_name, address, map_location):
    """Sheet1のJ列住所・K列マップ位置だけを更新する。"""
    workbook = load_workbook(BytesIO(original_content), keep_vba=True, data_only=False, read_only=False)
    original_sheets = list(workbook.sheetnames)
    changed_cells = []
    try:
        if SHEET_NAME not in workbook.sheetnames:
            raise ValueError("Sheet1が見つかりません。")
        ws = workbook[SHEET_NAME]
        rows = find_sheet1_customer_rows(workbook, customer_name)
        if not rows:
            raise ValueError("Sheet1のB列に表示されている顧客名と一致する行が見つかりません。")

        if not normalize_match_value(ws.cell(1, SHEET1_ADDRESS_COLUMN).value):
            ws.cell(1, SHEET1_ADDRESS_COLUMN).value = "住所"
            changed_cells.append((SHEET_NAME, 1, SHEET1_ADDRESS_COLUMN, "住所"))
        if not normalize_match_value(ws.cell(1, SHEET1_MAP_COLUMN).value):
            ws.cell(1, SHEET1_MAP_COLUMN).value = "マップ位置"
            changed_cells.append((SHEET_NAME, 1, SHEET1_MAP_COLUMN, "マップ位置"))

        for row in rows:
            for label, column, new_value in (
                ("住所", SHEET1_ADDRESS_COLUMN, address),
                ("マップ位置", SHEET1_MAP_COLUMN, map_location),
            ):
                cell = ws.cell(row, column)
                if not same_excel_value(cell.value, new_value):
                    cell.value = new_value
                    changed_cells.append((SHEET_NAME, row, column, new_value))

        if not changed_cells:
            raise ValueError("変更された項目がありません。")

        enable_excel_recalculation(workbook)
        output = BytesIO()
        workbook.save(output)
    finally:
        workbook.close()

    saved_content = output.getvalue()
    verified = load_workbook(BytesIO(saved_content), keep_vba=True, data_only=False, read_only=False)
    try:
        if list(verified.sheetnames) != original_sheets:
            raise ValueError("保存後にシート構成が変わったため、更新を中止しました。")
        if verified.vba_archive is None:
            raise ValueError("保存後の検証でVBAプロジェクトを確認できません。")
        for sheet, row, column, expected in changed_cells:
            actual = verified[sheet].cell(row, column).value
            if not same_excel_value(actual, expected):
                coordinate = verified[sheet].cell(row, column).coordinate
                raise ValueError(f"保存後の検証で{sheet}!{coordinate}の値が一致しません。")
    finally:
        verified.close()

    return saved_content, changed_cells


def save_customer_map_changes(customer_name, address, map_location):
    """更新前バックアップ、rev競合防止付きで住所・マップ位置を保存する。"""
    access_token = get_dropbox_access_token()
    target_path = get_dropbox_file_path()
    original_content, download_response = download_dropbox_file(target_path, access_token)
    if original_content is None:
        raise RuntimeError("最新のExcelを取得できませんでした。\n" + dropbox_error_text(download_response))

    revision = get_download_revision(download_response)
    if not revision:
        raise RuntimeError("Dropboxのrevを取得できないため、安全のため更新を中止しました。")

    ensure_dropbox_backup_folder(access_token)
    timestamp = get_jst_now().strftime("%Y%m%d_%H%M%S_%f")
    backup_path = f"{DROPBOX_BACKUP_FOLDER}/配車予定 次郎_{timestamp}.xlsm"
    backup_response = upload_dropbox_file(
        backup_path,
        original_content,
        access_token,
        mode="add",
    )
    if backup_response.status_code != 200:
        raise RuntimeError(
            "バックアップを作成できないため、本番ファイルは更新しません。\n"
            + dropbox_error_text(backup_response)
        )

    saved_content, changed_cells = update_customer_map_workbook_bytes(
        original_content,
        customer_name,
        address,
        map_location,
    )
    upload_response = upload_dropbox_file(
        target_path,
        saved_content,
        access_token,
        mode="update",
        rev=revision,
    )
    if upload_response.status_code == 409:
        raise RuntimeError(
            "PCまたは別端末でExcelが更新されています。再読み込みしてからやり直してください"
        )
    if upload_response.status_code != 200:
        raise RuntimeError(
            "本番Excelを更新できませんでした。\n"
            + dropbox_error_text(upload_response)
        )

    # Dropbox上のファイルを読み直し、住所・地図が正しいセルへ入ったことを確認する。
    confirm_dropbox_upload(target_path, access_token, changed_cells)

    cleanup_warning = trim_old_dropbox_backups(access_token, keep=30)
    st.cache_data.clear()
    return {
        "backup_path": backup_path,
        "updated_at": get_jst_now(),
        "changed_cells": changed_cells,
        "cleanup_warning": cleanup_warning,
    }


def get_dropbox_file_path():
    """Dropboxから直接取得するファイルパスを返す"""
    path = str(DROPBOX_FILE_PATH or "").strip()

    if not path:
        return DROPBOX_DEFAULT_FILE_PATH

    return path


def show_dropbox_download_error(path, response):
    st.error("DropboxからExcelファイルを直接取得できませんでした。")
    st.write("Dropboxの指定パスを確認してください。")

    st.write("現在アプリが取得しようとしたパス：")
    st.code(path)
    st.write("Secretsに設定すべき正しい値：")
    st.code(f'DROPBOX_FILE_PATH = "{DROPBOX_DEFAULT_FILE_PATH}"')

    if response is not None:
        st.write(f"Dropboxからの応答コード：{response.status_code}")
        try:
            error_body = json.dumps(response.json(), ensure_ascii=False, indent=2)
        except Exception:
            error_body = response.text
        st.code(error_body)

    st.write("よくある原因：")
    st.write("- パスの先頭に不要なフォルダ名が入っている")
    st.write("- 全角スペースが半角スペースに変わっている")
    st.write("- Dropbox上でファイル名またはフォルダ名が変更された")
    st.stop()


@st.cache_data(ttl=60, show_spinner=False)
def get_cached_dropbox_excel_content():
    """同じExcelを画面操作ごとに再取得せず、1分間だけ共有する。"""
    access_token = get_dropbox_access_token()
    dropbox_file_path = get_dropbox_file_path()
    content, response = download_dropbox_file(dropbox_file_path, access_token)
    if content is None:
        raise RuntimeError(dropbox_error_text(response))
    return content


def read_excel_from_dropbox_api():
    """Dropbox APIでExcelをダウンロードして読み込む"""
    if not has_dropbox_auth_config():
        st.error("Dropbox API設定が不足しています。")
        st.write("secrets.toml に DROPBOX_APP_KEY / DROPBOX_APP_SECRET / DROPBOX_REFRESH_TOKEN を設定してください。")
        st.stop()

    dropbox_file_path = get_dropbox_file_path()
    try:
        content = get_cached_dropbox_excel_content()
    except Exception:
        # 従来の詳しいエラー表示を維持するため、失敗時だけ直接取得する。
        access_token = get_dropbox_access_token()
        content, response = download_dropbox_file(dropbox_file_path, access_token)
        if content is None:
            show_dropbox_download_error(dropbox_file_path, response)

    return BytesIO(content)


def read_excel_local():
    """同じフォルダにあるローカルExcelを読み込む"""
    excel_path = Path(EXCEL_FILE)

    if not excel_path.exists():
        st.error(f"Excelファイルが見つかりません：{EXCEL_FILE}")
        st.stop()

    return excel_path

# =========================
# メモ帳（Supabase保存）
# =========================
def get_jst_now():
    """日本時間の現在日時を返す"""
    return datetime.now(timezone(timedelta(hours=9)))


def format_note_datetime(value):
    """ISO形式の日時を見やすく表示する"""
    try:
        dt = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        if dt.tzinfo is not None:
            dt = dt.astimezone(timezone(timedelta(hours=9)))
        return dt.strftime("%Y/%m/%d %H:%M")
    except Exception:
        return clean_value(value, blank_text="")


def get_supabase_key():
    """Supabaseへ接続するキーを返す。StreamlitではSecrets内だけに置く。"""
    return str(SUPABASE_SECRET_KEY or SUPABASE_SERVICE_ROLE_KEY or SUPABASE_ANON_KEY or "").strip()


def has_supabase_config():
    return bool(str(SUPABASE_URL or "").strip() and get_supabase_key())


def get_supabase_notes_table():
    table_name = str(SUPABASE_NOTES_TABLE or "notes").strip()

    if not table_name.replace("_", "").isalnum():
        st.error("Supabaseのメモ帳テーブル名が正しくありません。")
        st.write("SUPABASE_NOTES_TABLE は英数字とアンダースコアだけで指定してください。")
        st.stop()

    return table_name


def get_supabase_notes_url():
    base_url = str(SUPABASE_URL or "").strip().rstrip("/")
    table_name = urllib.parse.quote(get_supabase_notes_table(), safe="")
    return f"{base_url}/rest/v1/{table_name}"


def get_supabase_headers(prefer=None):
    key = get_supabase_key()
    headers = {
        "apikey": key,
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    # 新しいsb_secret_/sb_publishable_キーはJWTではないため、
    # Authorization: Bearerには入れず、apikeyヘッダーだけで送る。
    # 旧service_role/anonのJWTキーは従来どおりBearerにも設定する。
    if not key.startswith(("sb_secret_", "sb_publishable_")):
        headers["Authorization"] = f"Bearer {key}"
    if prefer:
        headers["Prefer"] = prefer
    return headers


def make_line_status_id(customer_name):
    """顧客名から、LINE状態保存用の重複しないIDを作る。"""
    customer = clean_value(customer_name, blank_text="")
    digest = hashlib.sha256(customer.encode("utf-8")).hexdigest()
    return f"{LINE_STATUS_NOTE_PREFIX}{digest}"


@st.cache_data(ttl=30, show_spinner=False)
def load_line_statuses_from_supabase():
    """LINE接続中の顧客名だけを小さなデータとして読み込む。"""
    if not has_supabase_config():
        return {}

    try:
        response = requests.get(
            get_supabase_notes_url(),
            headers=get_supabase_headers(),
            params={
                "select": "customer_name",
                "id": f"like.{LINE_STATUS_NOTE_PREFIX}*",
                "limit": "5000",
            },
            timeout=15,
        )
    except Exception:
        return {}

    if response.status_code != 200:
        return {}

    try:
        rows = response.json()
    except Exception:
        return {}

    if not isinstance(rows, list):
        return {}

    return {
        clean_value(row.get("customer_name"), blank_text=""): True
        for row in rows
        if clean_value(row.get("customer_name"), blank_text="")
    }


def get_line_connected(customer_name):
    customer = clean_value(customer_name, blank_text="")
    return bool(load_line_statuses_from_supabase().get(customer, False))


def save_line_connected(customer_name, connected):
    """Excelを変更せず、LINE状態だけをSupabaseへ保存する。"""
    if not has_supabase_config():
        st.error("LINE状態を保存するための接続設定がありません。")
        return False

    customer = clean_value(customer_name, blank_text="")
    status_id = make_line_status_id(customer)

    try:
        if connected:
            response = requests.post(
                get_supabase_notes_url(),
                headers=get_supabase_headers(
                    prefer="resolution=merge-duplicates,return=minimal"
                ),
                json={
                    "id": status_id,
                    "customer_name": customer,
                    "body": LINE_STATUS_BODY,
                    "created_at": get_jst_now().isoformat(),
                },
                timeout=15,
            )
            success = response.status_code in (200, 201)
        else:
            response = requests.delete(
                get_supabase_notes_url(),
                headers=get_supabase_headers(prefer="return=minimal"),
                params={"id": f"eq.{status_id}"},
                timeout=15,
            )
            success = response.status_code in (200, 204)
    except Exception as exc:
        st.error(f"LINE状態を保存できませんでした：{exc}")
        return False

    if not success:
        st.error("LINE状態を保存できませんでした。")
        return False

    load_line_statuses_from_supabase.clear()
    return True


def show_supabase_config_error():
    st.error("メモ帳を使うにはSupabase設定が必要です。")
    st.write("Streamlit Cloud の Secrets に以下を追加してください。")
    st.code(
        '\n'.join(
            [
                'SUPABASE_URL = "https://xxxx.supabase.co"',
                'SUPABASE_SECRET_KEY = "SupabaseのSecret key"',
                'SUPABASE_NOTES_TABLE = "notes"',
            ]
        )
    )
    st.stop()


def show_supabase_response_error(action, response):
    st.error(f"メモ帳をSupabaseに{action}できませんでした。")
    if response is not None:
        st.write(f"Supabaseからの応答コード：{response.status_code}")
        try:
            st.code(json.dumps(response.json(), ensure_ascii=False, indent=2))
        except Exception:
            st.code(response.text)
    st.stop()


@st.cache_data(ttl=30, show_spinner=False)
def load_notes_from_supabase(customer_name=None, limit=500):
    """Supabaseのnotesテーブルからメモを新しい順で読み込む"""
    if not has_supabase_config():
        show_supabase_config_error()

    params = {
        "select": "id,customer_name,body,created_at",
        "order": "created_at.desc",
        "limit": str(limit),
        "id": f"not.like.{LINE_STATUS_NOTE_PREFIX}*",
    }

    if customer_name is not None:
        target = clean_value(customer_name, blank_text="")
        params["customer_name"] = f"eq.{target}"

    try:
        response = requests.get(
            get_supabase_notes_url(),
            headers=get_supabase_headers(),
            params=params,
            timeout=30,
        )
    except Exception as e:
        st.error("メモ帳の読み込み中にSupabaseへの接続に失敗しました。")
        st.exception(e)
        st.stop()

    if response.status_code != 200:
        show_supabase_response_error("読み込み", response)

    try:
        notes = response.json()
    except Exception:
        st.error("Supabaseから返ったメモ帳データの形式が正しくありません。")
        st.stop()

    if not isinstance(notes, list):
        return []

    return notes


def insert_note_to_supabase(note):
    """Supabaseのnotesテーブルへメモを1件追加する"""
    if not has_supabase_config():
        show_supabase_config_error()

    try:
        response = requests.post(
            get_supabase_notes_url(),
            headers=get_supabase_headers(prefer="return=minimal"),
            json=note,
            timeout=30,
        )
    except Exception as e:
        st.error("メモ帳の保存中にSupabaseへの接続に失敗しました。")
        st.exception(e)
        st.stop()

    if response.status_code not in (200, 201):
        show_supabase_response_error("保存", response)
    load_notes_from_supabase.clear()


def delete_note_from_supabase(note_id):
    """Supabaseのnotesテーブルからメモを1件削除する"""
    target_id = clean_value(note_id, blank_text="")
    if not target_id:
        st.warning("削除するメモが見つかりません。")
        return False

    if not has_supabase_config():
        show_supabase_config_error()

    try:
        response = requests.delete(
            get_supabase_notes_url(),
            headers=get_supabase_headers(prefer="return=minimal"),
            params={"id": f"eq.{target_id}"},
            timeout=30,
        )
    except Exception as e:
        st.error("メモ帳の削除中にSupabaseへの接続に失敗しました。")
        st.exception(e)
        st.stop()

    if response.status_code not in (200, 204):
        show_supabase_response_error("削除", response)

    load_notes_from_supabase.clear()
    return True


def make_note_id():
    return get_jst_now().strftime("%Y%m%d%H%M%S%f")


def add_note(customer_name, body):
    """顧客名に紐づくメモを1件追加する"""
    note_text = str(body or "").strip()
    if not note_text:
        st.warning("メモ本文を入力してください。")
        return False

    now = get_jst_now().isoformat()
    note = {
        "id": make_note_id(),
        "customer_name": clean_value(customer_name, blank_text=""),
        "body": note_text,
        "created_at": now,
    }

    insert_note_to_supabase(note)
    return True


def get_notes_for_customer(customer_name):
    return load_notes_from_supabase(customer_name=customer_name)


def render_note_card(note, show_customer=True):
    customer_name = clean_value(note.get("customer_name"), blank_text="未設定")
    created_at = format_note_datetime(note.get("created_at", ""))
    body = html.escape(clean_value(note.get("body"), blank_text="")).replace("\n", "<br>")

    if show_customer:
        customer_link = build_customer_detail_link(customer_name, class_name="dispatch-month-link")
        meta = f"{html.escape(created_at)}　{customer_link}"
    else:
        meta = html.escape(created_at)

    st.markdown(
        f"""
        <div class="note-card">
            <div class="note-meta">{meta}</div>
            <div class="note-body">{body}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_note_delete_controls(note):
    note_id = clean_value(note.get("id"), blank_text="")
    if not note_id:
        return

    confirm_key = f"confirm_delete_note_{note_id}"

    if st.session_state.get(confirm_key):
        st.warning("このメモを削除しますか？")
        col1, col2 = st.columns(2)

        with col1:
            if st.button("本当に削除", key=f"delete_note_confirm_{note_id}"):
                if delete_note_from_supabase(note_id):
                    st.session_state.pop(confirm_key, None)
                    st.success("メモを削除しました。")
                    st.rerun()

        with col2:
            if st.button("キャンセル", key=f"delete_note_cancel_{note_id}"):
                st.session_state.pop(confirm_key, None)
                st.rerun()
    else:
        if st.button("削除", key=f"delete_note_start_{note_id}"):
            st.session_state[confirm_key] = True
            st.rerun()


def show_customer_notes(customer_name):
    st.markdown("---")
    st.subheader("📝 この顧客のメモ")
    st.caption(f"🎤 {VOICE_INPUT_HELP}")

    note_key = f"customer_note_input_{customer_name}"
    clear_note_key = f"clear_{note_key}"

    # Streamlitでは生成済みウィジェットの値を同じ実行中に変更できないため、
    # 保存後の次回実行で入力欄を空にする。
    if st.session_state.pop(clear_note_key, False):
        st.session_state[note_key] = ""

    note_text = st.text_area(
        "メモ本文",
        key=note_key,
        height=120,
        placeholder="例：次回は午前中希望。サンプル持参。など",
        help=VOICE_INPUT_HELP,
    )

    if st.button("メモを保存", key=f"save_customer_note_{customer_name}"):
        if add_note(customer_name, note_text):
            st.session_state[clear_note_key] = True
            st.session_state["note_save_success"] = customer_name
            st.rerun()

    if st.session_state.pop("note_save_success", None) == customer_name:
        st.success("メモを保存しました。")

    customer_notes = get_notes_for_customer(customer_name)

    if not customer_notes:
        st.info("この顧客のメモはまだありません。")
        return

    st.markdown("#### メモ履歴")
    for note in customer_notes:
        render_note_card(note, show_customer=False)
        render_note_delete_controls(note)


def show_notes_page(df):
    show_back_home_button("notes_back_home")

    st.markdown("---")
    st.header("📝 メモ帳")
    st.caption("全顧客のメモを新しい順で表示します。")

    notes = load_notes_from_supabase()

    if not notes:
        st.info("メモはまだありません。顧客詳細画面から保存できます。")
        return

    for note in notes:
        render_note_card(note, show_customer=True)
        render_note_delete_controls(note)


# =========================
# 顧客情報（Supabase保存）
# =========================
def get_supabase_customer_information_table():
    table_name = str(SUPABASE_CUSTOMER_INFO_TABLE or "customer_information").strip()
    if not table_name.replace("_", "").isalnum():
        raise RuntimeError("Supabaseの顧客情報テーブル名が正しくありません。")
    return table_name


def get_supabase_customer_information_url():
    base_url = str(SUPABASE_URL or "").strip().rstrip("/")
    table_name = urllib.parse.quote(get_supabase_customer_information_table(), safe="")
    return f"{base_url}/rest/v1/{table_name}"


def get_stable_customer_key(detail):
    """同一顧客の全行でIDが1種類だけの場合に限り、安定キーとして使う。"""
    if "ID" not in detail.columns:
        return None
    customer_ids = {
        clean_value(value, blank_text="")
        for value in detail["ID"].tolist()
        if clean_value(value, blank_text="")
    }
    return next(iter(customer_ids)) if len(customer_ids) == 1 else None


def customer_information_query(customer_name, customer_key=None):
    params = {
        "select": "id,customer_key,customer_name,field_name,content,sort_order,created_at,updated_at",
        "order": "sort_order.asc,created_at.asc,id.asc",
    }
    if customer_key:
        params["customer_key"] = f"eq.{customer_key}"
    else:
        params["customer_key"] = "is.null"
        params["customer_name"] = f"eq.{customer_name}"
    return params


def check_customer_information_response(action, response, success_codes):
    if response.status_code in success_codes:
        return
    detail = str(response.text or "").strip()[:500]
    message = f"顧客情報を{action}できませんでした（{response.status_code}）。"
    if detail:
        message += f" {detail}"
    raise RuntimeError(message)


@st.cache_data(ttl=30, show_spinner=False)
def load_customer_information(customer_name, customer_key=None):
    """件数上限を設けず、Supabaseからページ単位で顧客情報を読む。"""
    if not has_supabase_config():
        raise RuntimeError("Supabase設定がありません。")

    rows = []
    page_size = 1000
    offset = 0
    while True:
        params = customer_information_query(customer_name, customer_key)
        params.update({"limit": str(page_size), "offset": str(offset)})
        try:
            response = requests.get(
                get_supabase_customer_information_url(),
                headers=get_supabase_headers(),
                params=params,
                timeout=30,
            )
        except Exception as exc:
            raise RuntimeError("顧客情報の読み込み中にSupabaseへ接続できませんでした。") from exc
        check_customer_information_response("読み込み", response, (200,))
        page = response.json()
        if not isinstance(page, list):
            raise RuntimeError("Supabaseから返った顧客情報の形式が正しくありません。")
        rows.extend(page)
        if len(page) < page_size:
            break
        offset += page_size
    return rows


def clear_customer_information_cache():
    load_customer_information.clear()


def insert_customer_information(customer_name, customer_key, field_name, content, sort_order):
    now = get_jst_now().isoformat()
    payload = {
        "id": str(uuid.uuid4()),
        "customer_key": customer_key or None,
        "customer_name": clean_value(customer_name, blank_text=""),
        "field_name": str(field_name).strip(),
        "content": str(content or ""),
        "sort_order": int(sort_order),
        "created_at": now,
        "updated_at": now,
    }
    try:
        response = requests.post(
            get_supabase_customer_information_url(),
            headers=get_supabase_headers(prefer="return=minimal"),
            json=payload,
            timeout=30,
        )
    except Exception as exc:
        raise RuntimeError("顧客情報の保存中にSupabaseへ接続できませんでした。") from exc
    check_customer_information_response("保存", response, (200, 201))
    clear_customer_information_cache()


def update_customer_information(item_id, field_name, content):
    try:
        response = requests.patch(
            get_supabase_customer_information_url(),
            headers=get_supabase_headers(prefer="return=minimal"),
            params={"id": f"eq.{item_id}"},
            json={
                "field_name": str(field_name).strip(),
                "content": str(content or ""),
            },
            timeout=30,
        )
    except Exception as exc:
        raise RuntimeError("顧客情報の更新中にSupabaseへ接続できませんでした。") from exc
    check_customer_information_response("更新", response, (200, 204))
    clear_customer_information_cache()


def delete_customer_information(item_id):
    try:
        response = requests.delete(
            get_supabase_customer_information_url(),
            headers=get_supabase_headers(prefer="return=minimal"),
            params={"id": f"eq.{item_id}"},
            timeout=30,
        )
    except Exception as exc:
        raise RuntimeError("顧客情報の削除中にSupabaseへ接続できませんでした。") from exc
    check_customer_information_response("削除", response, (200, 204))
    clear_customer_information_cache()


def reorder_customer_information(first_item, second_item):
    """隣接2行を1回のupsertで入れ替え、並び順をまとめて保存する。"""
    payload = []
    for item, new_order in (
        (first_item, second_item.get("sort_order", 0)),
        (second_item, first_item.get("sort_order", 0)),
    ):
        payload.append(
            {
                "id": item["id"],
                "customer_key": item.get("customer_key"),
                "customer_name": item["customer_name"],
                "field_name": item["field_name"],
                "content": item.get("content", ""),
                "sort_order": int(new_order),
                "created_at": item["created_at"],
            }
        )
    try:
        response = requests.post(
            get_supabase_customer_information_url(),
            headers=get_supabase_headers(
                prefer="resolution=merge-duplicates,return=minimal"
            ),
            params={"on_conflict": "id"},
            json=payload,
            timeout=30,
        )
    except Exception as exc:
        raise RuntimeError("顧客情報の並び替え中にSupabaseへ接続できませんでした。") from exc
    check_customer_information_response("並び替え", response, (200, 201))
    clear_customer_information_cache()



def make_past_product_note_field(product_name):
    """顧客情報テーブル内で、過去商品メモを通常項目と分けるための内部項目名を作る。"""
    product = clean_value(product_name, blank_text="").strip()
    return f"{PAST_PRODUCT_NOTE_PREFIX}{product}"


def extract_past_product_name(field_name):
    """内部項目名から商品名だけを取り出す。"""
    field = clean_value(field_name, blank_text="")
    if not field.startswith(PAST_PRODUCT_NOTE_PREFIX):
        return ""
    return field[len(PAST_PRODUCT_NOTE_PREFIX):].strip()


def is_past_product_note_item(item):
    """顧客情報テーブル上の商品メモ専用レコードか判定する。"""
    return clean_value(item.get("field_name"), blank_text="").startswith(PAST_PRODUCT_NOTE_PREFIX)


def get_past_product_names(detail, visible_detail):
    """使用数量/日が0または空白の商品を、過去に使用した商品として抽出する。"""
    active_products = {
        clean_value(value, blank_text="").strip()
        for value in visible_detail.get("商品名", []).tolist()
        if clean_value(value, blank_text="").strip()
    }

    past_products = []
    for _, row in detail.iterrows():
        product_name = clean_value(row.get("商品名"), blank_text="").strip()
        if not product_name:
            continue
        if product_name in active_products:
            continue
        if not is_blank_or_zero(row.get("使用数量/日")):
            continue
        if product_name not in past_products:
            past_products.append(product_name)

    return past_products


def get_past_product_note_items(customer_name, customer_key):
    """過去商品メモを商品名ごとの辞書で返す。"""
    items = load_customer_information(customer_name, customer_key)
    result = {}
    for item in items:
        if not is_past_product_note_item(item):
            continue
        product_name = extract_past_product_name(item.get("field_name"))
        if product_name and product_name not in result:
            result[product_name] = item
    return result


def save_past_product_note(customer_name, customer_key, product_name, content):
    """過去商品の商品別メモを、既存の顧客情報テーブルへ内部項目として保存する。"""
    field_name = make_past_product_note_field(product_name)
    items = load_customer_information(customer_name, customer_key)
    existing = next(
        (
            item for item in items
            if clean_value(item.get("field_name"), blank_text="") == field_name
        ),
        None,
    )

    if existing:
        update_customer_information(existing["id"], field_name, content)
        return

    next_order = max(
        (int(item.get("sort_order", 0)) for item in items),
        default=0,
    ) + 10
    insert_customer_information(
        customer_name,
        customer_key,
        field_name,
        content,
        next_order,
    )


def delete_past_product_note(note_item):
    """過去商品の商品別メモを削除する。"""
    item_id = clean_value(note_item.get("id"), blank_text="")
    if not item_id:
        raise RuntimeError("削除する商品メモが見つかりません。")
    delete_customer_information(item_id)


def render_past_products_section(customer_name, customer_key, detail, visible_detail):
    """顧客詳細の最下部に、過去に使用した商品と商品別メモを表示する。"""
    past_products = get_past_product_names(detail, visible_detail)
    if not past_products:
        return

    st.markdown("---")
    st.subheader("過去に使用した商品")

    if not has_supabase_config():
        st.warning("商品メモを使うにはSupabase設定が必要です。")
        return

    try:
        note_items = get_past_product_note_items(customer_name, customer_key)
    except Exception as exc:
        st.warning(f"商品メモを読み込めませんでした：{exc}")
        note_items = {}

    identity = customer_key or customer_name
    for product_name in past_products:
        note_item = note_items.get(product_name)
        current_content = clean_value(
            note_item.get("content") if note_item else "",
            blank_text="",
        )
        state_suffix = hashlib.sha256(
            f"past-product|{identity}|{product_name}".encode("utf-8")
        ).hexdigest()[:16]
        delete_confirm_key = f"past_product_delete_confirm_{state_suffix}"

        with st.expander(product_name):
            memo = st.text_area(
                "メモ",
                value=current_content,
                key=f"past_product_note_{state_suffix}",
                height=110,
                placeholder="例：値上げのため中止、効果が薄かった、別商品へ変更 など",
            )

            save_col, delete_col = st.columns(2)
            with save_col:
                if st.button(
                    "保存",
                    key=f"past_product_note_save_{state_suffix}",
                    type="primary",
                    use_container_width=True,
                ):
                    try:
                        save_past_product_note(
                            customer_name,
                            customer_key,
                            product_name,
                            memo,
                        )
                        st.success("商品メモを保存しました。")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"商品メモを保存できませんでした：{exc}")

            with delete_col:
                if note_item:
                    if st.session_state.get(delete_confirm_key):
                        if st.button(
                            "削除する",
                            key=f"past_product_note_delete_confirm_{state_suffix}",
                            use_container_width=True,
                        ):
                            try:
                                delete_past_product_note(note_item)
                                st.session_state.pop(delete_confirm_key, None)
                                st.success("商品メモを削除しました。")
                                st.rerun()
                            except Exception as exc:
                                st.error(f"商品メモを削除できませんでした：{exc}")
                        if st.button(
                            "キャンセル",
                            key=f"past_product_note_delete_cancel_{state_suffix}",
                            use_container_width=True,
                        ):
                            st.session_state.pop(delete_confirm_key, None)
                            st.rerun()
                    else:
                        if st.button(
                            "削除",
                            key=f"past_product_note_delete_{state_suffix}",
                            use_container_width=True,
                        ):
                            st.session_state[delete_confirm_key] = True
                            st.rerun()


def render_customer_information_form(customer_name, customer_key, items, state_suffix):
    add_key = f"customer_information_add_{state_suffix}"
    if not st.session_state.get(add_key):
        if st.button("＋ 項目を追加", key=f"customer_information_add_button_{state_suffix}"):
            st.session_state[add_key] = True
            st.rerun()
        return

    st.markdown("**新しい項目**")
    with st.form(f"customer_information_add_form_{state_suffix}"):
        field_name = st.text_input("項目名", placeholder="例：担当者")
        content = st.text_area(
            "内容",
            placeholder="内容を入力（複数行可）",
            height=120,
        )
        save_col, cancel_col = st.columns(2)
        with save_col:
            save = st.form_submit_button(
                "保存", type="primary", use_container_width=True
            )
        with cancel_col:
            cancel = st.form_submit_button("キャンセル", use_container_width=True)

    if cancel:
        st.session_state.pop(add_key, None)
        st.rerun()
    if save:
        if not str(field_name).strip():
            st.warning("項目名を入力してください。")
            return
        next_order = max(
            (int(item.get("sort_order", 0)) for item in items),
            default=0,
        ) + 10
        try:
            insert_customer_information(
                customer_name,
                customer_key,
                field_name,
                content,
                next_order,
            )
            st.session_state.pop(add_key, None)
            st.session_state[f"customer_information_success_{state_suffix}"] = "項目を追加しました。"
            st.rerun()
        except RuntimeError as exc:
            st.error(str(exc))


def render_customer_information_card(customer_name, customer_key=None):
    identity = customer_key or customer_name
    state_suffix = hashlib.sha256(str(identity).encode("utf-8")).hexdigest()[:16]
    edit_mode_key = f"customer_information_edit_mode_{state_suffix}"
    editing_item_key = f"customer_information_editing_item_{state_suffix}"
    deleting_item_key = f"customer_information_deleting_item_{state_suffix}"

    with st.container(border=True):
        title_col, action_col = st.columns([4, 1])
        with title_col:
            st.subheader("顧客情報")
        with action_col:
            edit_mode = bool(st.session_state.get(edit_mode_key))
            if st.button(
                "完了" if edit_mode else "編集",
                key=f"customer_information_mode_button_{state_suffix}",
                use_container_width=True,
            ):
                st.session_state[edit_mode_key] = not edit_mode
                st.session_state.pop(editing_item_key, None)
                st.session_state.pop(deleting_item_key, None)
                st.rerun()

        if not has_supabase_config():
            st.warning("顧客情報を使うにはSupabase設定が必要です。")
            return

        try:
            items = load_customer_information(customer_name, customer_key)
            items = [item for item in items if not is_past_product_note_item(item)]
        except Exception as exc:
            st.warning(str(exc))
            return

        success_key = f"customer_information_success_{state_suffix}"
        success_message = st.session_state.pop(success_key, None)
        if success_message:
            st.success(success_message)

        if not items:
            st.info("登録されている情報はありません。")

        edit_mode = bool(st.session_state.get(edit_mode_key))
        active_edit_id = st.session_state.get(editing_item_key)
        active_delete_id = st.session_state.get(deleting_item_key)

        for index, item in enumerate(items):
            item_id = str(item.get("id", ""))
            field_name = clean_value(item.get("field_name"), blank_text="")
            content = clean_value(item.get("content"), blank_text="")

            if edit_mode and active_edit_id == item_id:
                with st.form(f"customer_information_edit_form_{item_id}"):
                    edited_name = st.text_input("項目名", value=field_name)
                    edited_content = st.text_area(
                        "内容", value=content, height=120
                    )
                    save_col, cancel_col = st.columns(2)
                    with save_col:
                        save = st.form_submit_button(
                            "保存", type="primary", use_container_width=True
                        )
                    with cancel_col:
                        cancel = st.form_submit_button(
                            "キャンセル", use_container_width=True
                        )
                if cancel:
                    st.session_state.pop(editing_item_key, None)
                    st.rerun()
                if save:
                    if not str(edited_name).strip():
                        st.warning("項目名を入力してください。")
                    else:
                        try:
                            update_customer_information(
                                item_id, edited_name, edited_content
                            )
                            st.session_state.pop(editing_item_key, None)
                            st.session_state[success_key] = "項目を更新しました。"
                            st.rerun()
                        except RuntimeError as exc:
                            st.error(str(exc))
                continue

            if edit_mode and active_delete_id == item_id:
                st.warning(f"「{field_name}」を削除しますか？")
                delete_col, cancel_col = st.columns(2)
                with delete_col:
                    if st.button(
                        "削除する",
                        key=f"customer_information_delete_confirm_{item_id}",
                        use_container_width=True,
                    ):
                        try:
                            delete_customer_information(item_id)
                            st.session_state.pop(deleting_item_key, None)
                            st.session_state[success_key] = "項目を削除しました。"
                            st.rerun()
                        except RuntimeError as exc:
                            st.error(str(exc))
                with cancel_col:
                    if st.button(
                        "キャンセル",
                        key=f"customer_information_delete_cancel_{item_id}",
                        use_container_width=True,
                    ):
                        st.session_state.pop(deleting_item_key, None)
                        st.rerun()
                continue

            if edit_mode:
                up_col, down_col, name_col, content_col, edit_col, delete_col = st.columns(
                    [0.55, 0.55, 1.5, 3, 0.8, 0.8]
                )
                with up_col:
                    if st.button(
                        "↑",
                        key=f"customer_information_up_{item_id}",
                        disabled=index == 0,
                    ):
                        try:
                            reorder_customer_information(item, items[index - 1])
                            st.rerun()
                        except RuntimeError as exc:
                            st.error(str(exc))
                with down_col:
                    if st.button(
                        "↓",
                        key=f"customer_information_down_{item_id}",
                        disabled=index == len(items) - 1,
                    ):
                        try:
                            reorder_customer_information(item, items[index + 1])
                            st.rerun()
                        except RuntimeError as exc:
                            st.error(str(exc))
                with name_col:
                    st.markdown(f"**{html.escape(field_name)}**", unsafe_allow_html=True)
                with content_col:
                    safe_content = html.escape(content).replace("\n", "<br>")
                    st.markdown(
                        f'<div style="overflow-wrap:anywhere">{safe_content}</div>',
                        unsafe_allow_html=True,
                    )
                with edit_col:
                    if st.button("編集", key=f"customer_information_edit_{item_id}"):
                        st.session_state[editing_item_key] = item_id
                        st.session_state.pop(deleting_item_key, None)
                        st.rerun()
                with delete_col:
                    if st.button("削除", key=f"customer_information_delete_{item_id}"):
                        st.session_state[deleting_item_key] = item_id
                        st.session_state.pop(editing_item_key, None)
                        st.rerun()
            else:
                safe_name = html.escape(field_name)
                safe_content = html.escape(content).replace("\n", "<br>")
                st.markdown(
                    (
                        '<div class="customer-information-row">'
                        f'<div class="customer-information-label">{safe_name}</div>'
                        f'<div class="customer-information-content">{safe_content}</div>'
                        '</div>'
                    ),
                    unsafe_allow_html=True,
                )

        if edit_mode:
            st.markdown("---")
            render_customer_information_form(
                customer_name,
                customer_key,
                items,
                state_suffix,
            )



# =========================
# Excel読み込み・整形
# =========================
def calculate_delivery_values(delivery_values_ws, row):
    """Excelの数式キャッシュが消えている場合に、L列とO列をPythonで再計算する。"""
    usage = delivery_values_ws.cell(row, 7).value
    kg_per_bottle = delivery_values_ws.cell(row, 9).value
    delivery_date = delivery_values_ws.cell(row, 10).value
    delivery_quantity = delivery_values_ws.cell(row, 11).value
    next_delivery = delivery_values_ws.cell(row, 12).value
    remaining = delivery_values_ws.cell(row, 15).value

    if isinstance(next_delivery, str) and next_delivery.startswith("="):
        next_delivery = None
    if isinstance(remaining, str) and remaining.startswith("="):
        remaining = None

    try:
        if next_delivery is None and delivery_date is not None:
            next_delivery = delivery_date + timedelta(
                days=math.floor(float(delivery_quantity) / float(usage))
            )
    except Exception:
        next_delivery = None

    try:
        if remaining is None and next_delivery is not None:
            target_date = next_delivery.date() if isinstance(next_delivery, datetime) else next_delivery
            remaining = (target_date - date.today()).days * float(usage) / float(kg_per_bottle)
    except Exception:
        remaining = None

    return next_delivery, remaining


def rebuild_sheet1_from_formula_references(excel_source):
    """openpyxl保存後に数式キャッシュが消えても、参照元からSheet1相当を復元する。"""
    if isinstance(excel_source, BytesIO):
        content = excel_source.getvalue()
    else:
        content = Path(excel_source).read_bytes()

    formula_wb = load_workbook(BytesIO(content), keep_vba=True, data_only=False, read_only=True)
    try:
        if SHEET_NAME not in formula_wb.sheetnames or DELIVERY_SHEET_NAME not in formula_wb.sheetnames:
            return pd.DataFrame()

        sheet1 = formula_wb[SHEET_NAME]
        # 顧客名などの元データは数式ではなく直接値なので、数式を保持する側から読む。
        # data_only側はopenpyxl保存後に数式キャッシュが消え、空欄になるため使わない。
        delivery = formula_wb[DELIVERY_SHEET_NAME]
        rows = []
        for sheet1_row in range(2, sheet1.max_row + 1):
            source_row = None
            for column in (1, 2):
                formula = sheet1.cell(sheet1_row, column).value
                if isinstance(formula, str) and formula.startswith("="):
                    # シート名の引用符・全角文字に依存せず、参照式末尾の行番号を使う。
                    match = re.search(r"(\d+)\s*$", formula.strip())
                    if match:
                        source_row = int(match.group(1))
                        break
            if source_row is None:
                continue

            customer_name = delivery.cell(source_row, 2).value
            product_name = delivery.cell(source_row, 5).value
            if not normalize_match_value(customer_name) or not normalize_match_value(product_name):
                continue

            next_delivery, remaining = calculate_delivery_values(delivery, source_row)
            rows.append({
                "ID": delivery.cell(source_row, 1).value,
                "顧客名": customer_name,
                "地域": delivery.cell(source_row, 3).value,
                "商品名": product_name,
                "使用数量/日": delivery.cell(source_row, 7).value,
                "次回配達予定": next_delivery,
                "残数": remaining,
                "ひらがな": sheet1.cell(sheet1_row, SHEET1_HIRAGANA_COLUMN).value,
                "住所": sheet1.cell(sheet1_row, SHEET1_ADDRESS_COLUMN).value,
                "マップ位置": sheet1.cell(sheet1_row, SHEET1_MAP_COLUMN).value,
                "メーカー": delivery.cell(source_row, 6).value,
                "本数": delivery.cell(source_row, 8).value,
                "kg/本": delivery.cell(source_row, 9).value,
                "配達日": delivery.cell(source_row, 10).value,
            })
        return pd.DataFrame(rows)
    finally:
        formula_wb.close()


def normalize_excel_table(excel_source):
    """
    ExcelのSheet1から、顧客一覧表を取り出す。

    対応できる形：
    1) 1行目が見出し
       ID / 顧客名 / 地域 / 商品名 / 使用数量/日 / 次回配達予定 / 残数 / ひらがな

    2) 上部に大きな表示があり、途中の行に見出しがある形
       9行目などに ID / 顧客名 / 地域 / 商品名 ... がある
    """
    # 現在のブック構造ではこちらが最短経路。数式キャッシュの有無にも影響されない。
    try:
        rebuilt = rebuild_sheet1_from_formula_references(excel_source)
        if not rebuilt.empty:
            return rebuilt
    except Exception:
        pass

    try:
        raw = pd.read_excel(
            excel_source,
            sheet_name=SHEET_NAME,
            header=None,
            engine="openpyxl",
        )
    except Exception as e:
        st.error("Excelファイルを読み込めませんでした。")
        st.exception(e)
        st.stop()

    header_row_index = None

    for idx, row in raw.iterrows():
        values = [str(v).strip() for v in row.tolist() if not pd.isna(v)]
        column_mapping = find_required_column_mapping(values)
        score = len(column_mapping)

        # IDだけは上部表示にも出るので、顧客名・ひらがなに相当する列がある行を重視
        if "顧客名" in column_mapping and "ひらがな" in column_mapping and score >= 5:
            header_row_index = idx
            break

    if header_row_index is None:
        # openpyxlでxlsmを保存すると数式セルの前回計算結果が消える。
        # その場合はSheet1の数式参照先である「次回配達日」から一覧を復元する。
        rebuilt = rebuild_sheet1_from_formula_references(excel_source)
        if rebuilt.empty:
            st.error("必要な見出し行が見つかりません。")
            st.write("必要な列：", REQUIRED_COLUMNS)
            st.stop()
        return rebuilt

    header = raw.iloc[header_row_index].tolist()
    df = raw.iloc[header_row_index + 1:].copy()
    df.columns = header

    # 列名の空白除去
    df.columns = [str(c).strip() for c in df.columns]

    column_mapping = find_required_column_mapping(df.columns)
    missing = [col for col in REQUIRED_COLUMNS if col not in column_mapping]
    if missing:
        st.error("必要な列が見つかりません。")
        st.write("見つからない列：", missing)
        st.write("使用できる列名候補：")
        for col in missing:
            st.write(f"- {col}: {', '.join(REQUIRED_COLUMN_CANDIDATES[col])}")
        st.write("Excelから読み取れた列：", list(df.columns))
        st.stop()

    rename_mapping = {
        actual_column: required_column
        for required_column, actual_column in column_mapping.items()
        if actual_column != required_column
    }
    df = df.rename(columns=rename_mapping)

    # 既存機能に必要な列を先頭に置きつつ、備考などの追加列も残す
    ordered_columns = REQUIRED_COLUMNS.copy()
    for col in df.columns:
        if col not in ordered_columns:
            ordered_columns.append(col)

    df = df[ordered_columns].copy()

    # 検索に必要な行だけ残す
    df = df.dropna(subset=["顧客名", "ひらがな"])

    df["顧客名"] = df["顧客名"].astype(str).str.strip()
    df["ひらがな"] = df["ひらがな"].astype(str).str.strip()

    df = df[(df["顧客名"] != "") & (df["ひらがな"] != "")]
    return df


@st.cache_data(ttl=60, show_spinner=False)
def load_fast_dropbox_data():
    """通常表示は小さなJSONを使い、Excelが変わった時だけ再生成する。"""
    access_token = get_dropbox_access_token()
    excel_path = get_dropbox_file_path()
    excel_revision = get_dropbox_revision(excel_path, access_token)

    cache_content, cache_response = download_dropbox_file(
        DROPBOX_FAST_CACHE_FILE,
        access_token,
    )
    if cache_content is not None:
        try:
            payload = json.loads(cache_content.decode("utf-8"))
            if (
                payload.get("cache_version") == DROPBOX_FAST_CACHE_VERSION
                and payload.get("excel_revision") == excel_revision
            ):
                records = payload.get("records", [])
                if isinstance(records, list) and records:
                    return pd.DataFrame(records)
        except Exception:
            pass

    # Excelが更新された時だけ、1回だけ重い解析を行ってJSONを作り直す。
    excel_content, response = download_dropbox_file(excel_path, access_token)
    if excel_content is None:
        raise RuntimeError("Excelを取得できませんでした。\n" + dropbox_error_text(response))
    df = normalize_excel_table(BytesIO(excel_content))
    records = json.loads(df.to_json(orient="records", date_format="iso", force_ascii=False))
    payload = json.dumps(
        {
            "cache_version": DROPBOX_FAST_CACHE_VERSION,
            "excel_revision": excel_revision,
            "records": records,
        },
        ensure_ascii=False,
        separators=(",", ":"),
    ).encode("utf-8")
    upload_response = upload_dropbox_file(
        DROPBOX_FAST_CACHE_FILE,
        payload,
        access_token,
        mode="overwrite",
    )
    if upload_response.status_code != 200:
        # キャッシュ作成失敗でも、取得済みデータで画面表示は続ける。
        return df
    return df


@st.cache_data(ttl=60)
def load_data():
    """
    Dropbox API設定があればDropbox上のExcelを読む。
    設定がなければ同じフォルダのローカルExcelを読む。
    """
    if has_dropbox_auth_config():
        return load_fast_dropbox_data()

    return normalize_excel_table(read_excel_local())


# =========================
# 画面遷移
# =========================
def get_query_value(key, default=""):
    """URLパラメータを安全に1つ取り出す"""
    try:
        value = st.query_params.get(key, default)
    except Exception:
        return default

    if isinstance(value, list):
        return value[0] if value else default

    return value if value is not None else default


def update_query_params(**params):
    """ブラウザの戻るボタンで戻れるようにURLへ現在画面を残す"""
    try:
        # ログイン状態はURLにも残す。ログアウト時だけ消す。
        st.query_params["logged_in"] = "1"

        for key, value in params.items():
            if value is None or value == "":
                if key in st.query_params:
                    del st.query_params[key]
            else:
                st.query_params[key] = str(value)
    except Exception:
        pass



def make_app_url(page="home", customer=None, customer_search=None, region_search=None):
    """ブラウザの戻るボタンで戻れるように、通常リンク用URLを作る。"""
    params = {"logged_in": "1", "page": page}
    if customer:
        params["customer"] = str(customer)
    if customer_search:
        params["customer_search"] = str(customer_search)
    if region_search:
        params["region_search"] = str(region_search)
    return "?" + urllib.parse.urlencode(params)


def render_page_link(label, page="home", customer=None, customer_search=None, region_search=None, class_name="app-nav-link"):
    """st.buttonではなくHTMLリンクで画面遷移する。これによりブラウザ戻るが効く。"""
    url = make_app_url(
        page=page,
        customer=customer,
        customer_search=customer_search,
        region_search=region_search,
    )
    return f'<a class="{class_name}" href="{url}" target="_self">{html.escape(str(label))}</a>'

def sync_page_from_query_params():
    """URLのpage/customerを読んで、ブラウザ戻る・進むに追従する"""
    page = str(get_query_value("page", "home")).strip() or "home"
    customer = str(get_query_value("customer", "")).strip()

    valid_pages = {
        "home",
        "customer_list",
        "customer",
        "region",
        "calendar",
        "dispatch_table",
        "soluble_inventory",
        "notes",
        "detail",
    }

    raw_page = str(get_query_value("page", "")).strip()
    if page not in valid_pages:
        page = "home"
    if customer and not raw_page:
        page = "detail"

    st.session_state["page"] = page

    if page == "detail" and customer:
        st.session_state["selected_customer"] = customer
    elif page != "detail":
        st.session_state["selected_customer"] = None


def set_page(page_name, rerun=False):
    st.session_state["page"] = page_name

    if page_name != "detail":
        st.session_state["selected_customer"] = None

    update_query_params(page=page_name, customer=None)

    if rerun:
        st.rerun()


def select_customer(customer_name, page_name="detail"):
    st.session_state["selected_customer"] = customer_name
    st.session_state["page"] = page_name
    update_query_params(page=page_name, customer=customer_name)


def show_back_home_button(key):
    """各画面からホームへ戻るための共通リンク。ブラウザ履歴にも残る。"""
    st.markdown(render_page_link("← ホームへ戻る", page="home"), unsafe_allow_html=True)


def show_detail_search_shortcuts():
    """顧客詳細から、次の検索をすぐ始めるためのショートカット。"""
    col_customer, col_region = st.columns(2)
    with col_customer:
        st.markdown(
            render_page_link("🔍 顧客名で検索", page="customer"),
            unsafe_allow_html=True,
        )
    with col_region:
        st.markdown(
            render_page_link("📍 地域名で検索", page="region"),
            unsafe_allow_html=True,
        )


# =========================
# 顧客詳細
# =========================
def value_for_input(value):
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value)


def date_for_input(value):
    parsed = to_date(value)
    return parsed.strftime("%Y/%m/%d") if parsed else ""


def parse_optional_date(text):
    value = str(text).strip()
    if not value:
        return None
    try:
        value = value.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
        # 「2026年7月15日」のような音声入力結果も受け付ける。
        value = value.replace("年", "/").replace("月", "/").replace("日", "")
        # 「7/14」「7月14日」のように年がなければ、現在の年を自動補完する。
        if re.fullmatch(r"\d{1,2}\s*[/\-]\s*\d{1,2}", value):
            value = f"{date.today().year}/{value}"
        parsed = pd.to_datetime(value, errors="raise")
        return parsed.to_pydatetime().replace(hour=0, minute=0, second=0, microsecond=0)
    except Exception as exc:
        raise ValueError("配達日は 2026/07/15 のように入力してください。") from exc


def display_change_value(value):
    if value is None or value == "":
        return "（空欄）"
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y/%m/%d")
    return str(value)


def render_customer_excel_editor(customer_name, product_name, current):
    """商品カード内に、確認画面付きのExcel編集欄を追加する。"""
    identity = f"{customer_name}|{product_name}"
    key_suffix = str(abs(hash(identity)))
    edit_key = f"excel_edit_{key_suffix}"
    confirm_key = f"excel_confirm_{key_suffix}"

    if current.get("商品一致件数") == 0:
        st.error("顧客名・商品名が一致する行が見つからないため編集できません。")
        return
    if current.get("商品一致件数", 0) > 1:
        st.error("同じ顧客名・商品名の行が複数見つかりました")
        return

    st.caption("メーカー")
    st.markdown(f"**{clean_value(current.get('メーカー'))}**")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.caption("本数")
        st.markdown(f"**{format_number(current.get('本数'))}**")
    with col_b:
        st.caption("kg/本")
        st.markdown(f"**{format_number(current.get('kg/本'))}**")
    with col_c:
        st.caption("配達日")
        st.markdown(f"**{format_date(current.get('配達日'))}**")

    if not st.session_state.get(edit_key) and not st.session_state.get(confirm_key):
        if st.button("編集", key=f"edit_button_{key_suffix}"):
            st.session_state[edit_key] = True
            st.rerun()
        return

    if st.session_state.get(confirm_key):
        pending = st.session_state[confirm_key]
        st.markdown("**保存前の確認**")
        for label, values in pending["changes"].items():
            st.write(f"{label}：{display_change_value(values[0])} → {display_change_value(values[1])}")
        save_col, cancel_col = st.columns(2)
        with save_col:
            if st.button("保存", key=f"save_confirm_{key_suffix}", type="primary", use_container_width=True):
                try:
                    with st.spinner("バックアップを作成して保存しています…"):
                        result = save_customer_excel_changes(customer_name, product_name, pending["proposed"])
                    st.session_state.pop(confirm_key, None)
                    st.session_state.pop(edit_key, None)
                    st.session_state["excel_save_success"] = {
                        **result,
                        "customer_name": customer_name,
                        "product_name": product_name,
                    }
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
        with cancel_col:
            if st.button("キャンセル", key=f"cancel_confirm_{key_suffix}", use_container_width=True):
                st.session_state.pop(confirm_key, None)
                st.session_state.pop(edit_key, None)
                st.rerun()
        return

    with st.form(f"excel_edit_form_{key_suffix}"):
        st.caption(f"🎤 {VOICE_INPUT_HELP} 入力欄は毎回空白から始まります。")
        maker = st.text_input(
            "メーカー",
            value="",
            placeholder="メーカー名を入力",
            help=VOICE_INPUT_HELP,
        )
        bottles = st.text_input(
            "本数",
            value="",
            placeholder="例：44本",
            help=VOICE_INPUT_HELP,
        )
        kg_per_bottle = st.text_input(
            "kg/本",
            value="",
            placeholder="例：450キロ",
            help=VOICE_INPUT_HELP,
        )
        delivery_date = st.text_input(
            "配達日",
            value="",
            placeholder="例：2026年7月15日",
            help=VOICE_INPUT_HELP,
        )
        save_col, cancel_col = st.columns(2)
        with save_col:
            proceed = st.form_submit_button("保存", type="primary", use_container_width=True)
        with cancel_col:
            cancel = st.form_submit_button("キャンセル", use_container_width=True)

    if cancel:
        st.session_state.pop(edit_key, None)
        st.rerun()
    if proceed:
        try:
            proposed = {
                # 空欄は既存値を維持し、入力された項目だけ更新する。
                "メーカー": str(maker).strip() or current.get("メーカー"),
                "本数": (
                    parse_optional_nonnegative_number(bottles, integer=True)
                    if str(bottles).strip() else current.get("本数")
                ),
                "kg/本": (
                    parse_optional_nonnegative_number(kg_per_bottle, integer=False)
                    if str(kg_per_bottle).strip() else current.get("kg/本")
                ),
                "配達日": (
                    parse_optional_date(delivery_date)
                    if str(delivery_date).strip() else current.get("配達日")
                ),
                "住所": current.get("住所"),
                "マップ位置": current.get("マップ位置"),
            }
            changes = {
                label: (current.get(label), proposed[label])
                for label in proposed
                if not same_excel_value(current.get(label), proposed[label])
            }
            if not changes:
                st.warning("変更された項目がありません。")
            else:
                with st.spinner("DropboxのExcelへ保存しています…"):
                    result = save_customer_excel_changes(
                        customer_name,
                        product_name,
                        proposed,
                    )
                st.session_state.pop(confirm_key, None)
                st.session_state.pop(edit_key, None)
                st.session_state["excel_save_success"] = {
                    **result,
                    "customer_name": customer_name,
                    "product_name": product_name,
                }
                st.rerun()
        except ValueError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"保存できませんでした：{exc}")



def render_customer_map_editor(customer_name, current):
    """顧客単位の住所・マップ位置専用編集欄。"""
    key_suffix = str(abs(hash(f"map|{customer_name}")))
    edit_key = f"map_edit_{key_suffix}"
    confirm_key = f"map_confirm_{key_suffix}"

    st.markdown("### 📍 住所・マップ位置")
    st.write(f"**住所：** {clean_value(current.get('住所'))}")
    st.write(f"**マップ位置：** {clean_value(current.get('マップ位置'))}")

    if not st.session_state.get(edit_key) and not st.session_state.get(confirm_key):
        if st.button("住所・マップ位置を編集", key=f"map_edit_button_{key_suffix}"):
            st.session_state[edit_key] = True
            st.rerun()
        return

    if st.session_state.get(confirm_key):
        pending = st.session_state[confirm_key]
        st.markdown("**保存前の確認**")
        for label, values in pending["changes"].items():
            st.write(
                f"{label}：{display_change_value(values[0])} → "
                f"{display_change_value(values[1])}"
            )

        save_col, cancel_col = st.columns(2)
        with save_col:
            if st.button(
                "保存",
                key=f"map_save_confirm_{key_suffix}",
                type="primary",
                use_container_width=True,
            ):
                try:
                    with st.spinner("バックアップを作成して保存しています…"):
                        result = save_customer_map_changes(
                            customer_name,
                            pending["住所"],
                            pending["マップ位置"],
                        )
                    st.session_state.pop(confirm_key, None)
                    st.session_state.pop(edit_key, None)
                    st.session_state["excel_save_success"] = {
                        **result,
                        "customer_name": customer_name,
                        "product_name": "住所・マップ位置",
                    }
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

        with cancel_col:
            if st.button(
                "キャンセル",
                key=f"map_cancel_confirm_{key_suffix}",
                use_container_width=True,
            ):
                st.session_state.pop(confirm_key, None)
                st.session_state.pop(edit_key, None)
                st.rerun()
        return

    with st.form(f"map_edit_form_{key_suffix}"):
        st.caption(f"🎤 {VOICE_INPUT_HELP}")
        address = st.text_input(
            "住所",
            value=value_for_input(current.get("住所")),
            help=VOICE_INPUT_HELP,
        )
        map_location = st.text_input(
            "マップ位置",
            value=value_for_input(current.get("マップ位置")),
            help=f"緯度,経度／Googleマップ共有URL／文字列を入力できます。{VOICE_INPUT_HELP}",
        )
        save_col, cancel_col = st.columns(2)
        with save_col:
            proceed = st.form_submit_button(
                "保存",
                type="primary",
                use_container_width=True,
            )
        with cancel_col:
            cancel = st.form_submit_button(
                "キャンセル",
                use_container_width=True,
            )

    if cancel:
        st.session_state.pop(edit_key, None)
        st.rerun()

    if proceed:
        try:
            proposed_address = str(address)
            proposed_map = validate_map_location(map_location)
            changes = {}
            if not same_excel_value(current.get("住所"), proposed_address):
                changes["住所"] = (current.get("住所"), proposed_address)
            if not same_excel_value(current.get("マップ位置"), proposed_map):
                changes["マップ位置"] = (
                    current.get("マップ位置"),
                    proposed_map,
                )

            if not changes:
                st.warning("変更された項目がありません。")
            else:
                with st.spinner("DropboxのExcelへ保存しています…"):
                    result = save_customer_map_changes(
                        customer_name,
                        proposed_address,
                        proposed_map,
                    )
                st.session_state.pop(confirm_key, None)
                st.session_state.pop(edit_key, None)
                st.session_state["excel_save_success"] = {
                    **result,
                    "customer_name": customer_name,
                    "product_name": "住所・マップ位置",
                }
                st.rerun()
        except ValueError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"保存できませんでした：{exc}")


def show_customer_detail(df, customer_name):
    detail = df[df["顧客名"] == customer_name].copy()

    if detail.empty:
        st.warning("選択した顧客の情報が見つかりません。")
        return

    show_back_home_button("detail_back_home")
    show_detail_search_shortcuts()

    # 使用数量/日が0・空白・NaNの商品行は、商品名ごと表示しない。
    visible_detail = detail[~detail["使用数量/日"].apply(is_blank_or_zero)].copy()

    region = clean_value(detail.iloc[0]["地域"])

    st.markdown("---")
    line_connected = get_line_connected(customer_name)

    name_col, line_col, _ = st.columns([6, 3, 3])
    with name_col:
        st.markdown(
            '<div class="customer-name-row customer-detail-name-row">'
            f'<span>👤 {html.escape(clean_value(customer_name))}</span>'
            "</div>",
            unsafe_allow_html=True,
        )
    with line_col:
        if line_connected:
            st.markdown(
                '<div class="line-detail-static">LINE ○</div>',
                unsafe_allow_html=True,
            )
        else:
            with st.popover("LINE ×"):
                st.caption("LINEを○にしますか？")
                if st.button(
                    "○にする",
                    key=f"line_status_{make_line_status_id(customer_name)}",
                ):
                    with st.spinner("保存しています…"):
                        if save_line_connected(customer_name, True):
                            st.toast("LINEを○にしました。")
                            st.rerun()

    st.write(f"**地域：** {region}")
    st.write(f"**商品数：** {len(visible_detail)}件")

    success = st.session_state.pop("excel_save_success", None)
    if success:
        st.success("保存しました")
        st.success("バックアップを作成しました")
        st.write(f"**更新日時：** {success['updated_at'].strftime('%Y/%m/%d %H:%M:%S')}")
        st.write(f"**更新した顧客名：** {success['customer_name']}")
        st.write(f"**更新した商品名：** {success['product_name']}")
        if success.get("cleanup_warning"):
            st.warning(success["cleanup_warning"])

    try:
        map_info = get_customer_map_info(detail)
        if map_info and map_info["map_url"]:
            show_google_maps_button(map_info["map_url"])
    except Exception:
        pass

    # 詳細表示では重いExcelを開かず、高速JSON内の現在値を使う。
    first_detail = detail.iloc[0]
    current_map_values = {
        "住所": first_detail.get("住所"),
        "マップ位置": first_detail.get("マップ位置"),
        "顧客一致件数": len(detail),
    }
    render_customer_map_editor(customer_name, current_map_values)

    customer_key = get_stable_customer_key(detail)
    render_customer_information_card(customer_name, customer_key)

    if visible_detail.empty:
        st.info("表示対象の商品はありません。使用数量/日が0または空白の商品は非表示にしています。")

    for _, row in visible_detail.iterrows():
        product_name = clean_value(row["商品名"])
        customer_id = clean_value(row["ID"])
        usage = format_number(row["使用数量/日"])
        next_date = format_date(row["次回配達予定"])
        remaining = format_number(row["残数"])

        with st.container(border=True):
            st.subheader(f"📦 {product_name}")

            col1, col2 = st.columns(2)

            with col1:
                st.caption("ID")
                st.markdown(f"**{customer_id}**")

                st.caption("使用数量/日")
                st.markdown(f"**{usage}**")

            with col2:
                st.caption("次回配達予定")
                st.markdown(f"**{next_date}**")

                st.caption("残数")
                st.markdown(f"**{remaining}**")

            product_match_count = int((detail["商品名"].astype(str).str.strip() == product_name).sum())
            current_edit_values = {
                "メーカー": row.get("メーカー"),
                "本数": row.get("本数"),
                "kg/本": row.get("kg/本"),
                "配達日": row.get("配達日"),
                "住所": current_map_values["住所"],
                "マップ位置": current_map_values["マップ位置"],
                "商品一致件数": product_match_count,
                "顧客一致件数": len(detail),
            }
            render_customer_excel_editor(customer_name, product_name, current_edit_values)


    if normalize_soluble_customer_name(customer_name) in {
        normalize_soluble_customer_name(name) for name in SOLUBLE_CUSTOMER_NAMES
    }:
        try:
            soluble_content, _ = load_soluble_workbook_content()
            soluble_customer_summary = get_soluble_customer_summary(
                soluble_content,
                customer_name,
            )
            if soluble_customer_summary is not None:
                render_soluble_customer_product_card(
                    customer_name,
                    soluble_customer_summary,
                    key_scope="customer_detail",
                )
            else:
                st.warning(f"ソリュブルシートに「{customer_name}」が見つかりません。")
        except Exception as exc:
            st.warning(f"ソリュブル情報を読み込めませんでした：{exc}")

    show_customer_notes(customer_name)
    render_past_products_section(customer_name, customer_key, detail, visible_detail)

# =========================
# 顧客名一覧
# =========================
CUSTOMER_DIRECTORY_GROUPS = {
    "あ行": set("あいうえおぁぃぅぇぉ"),
    "か行": set("かきくけこがぎぐげご"),
    "さ行": set("さしすせそざじずぜぞ"),
    "た行": set("たちつてとだぢづでどっ"),
    "な行": set("なにぬねの"),
    "は行": set("はひふへほばびぶべぼぱぴぷぺぽ"),
    "ま行": set("まみむめも"),
    "や行": set("やゆよゃゅょ"),
    "ら行": set("らりるれろ"),
    "わ行": set("わをんゎ"),
}


def normalize_directory_kana(value):
    """顧客名一覧の並び替え用に、カタカナをひらがなへ寄せる。"""
    text = clean_value(value, blank_text="").strip()
    return "".join(
        chr(ord(char) - 0x60) if "ァ" <= char <= "ヶ" else char
        for char in text
    )


def get_customer_directory_group(value):
    kana = normalize_directory_kana(value)
    if not kana:
        return "その他"
    first = kana[0]
    for group_name, characters in CUSTOMER_DIRECTORY_GROUPS.items():
        if first in characters:
            return group_name
    return "その他"


def show_customer_directory(df=None):
    st.subheader("👥 顧客名一覧")
    show_back_home_button("customer_directory_back_home")
    st.caption("Sheet1の顧客を五十音順で表示します。顧客名を押すと詳細を開きます。")

    if df is None:
        with st.spinner("顧客データを読み込んでいます…"):
            df = load_data()

    directory = df[["顧客名", "地域", "商品名", "ひらがな"]].copy()
    for column in directory.columns:
        directory[column] = directory[column].fillna("").astype(str).str.strip()
    directory = directory[directory["顧客名"] != ""]

    if directory.empty:
        st.info("表示できる顧客がありません。")
        return

    customers = (
        directory.groupby("顧客名", as_index=False)
        .agg(
            地域=("地域", "first"),
            ひらがな=("ひらがな", "first"),
            商品数=("商品名", lambda values: values[values != ""].nunique()),
        )
    )
    customers["並び順"] = customers.apply(
        lambda row: normalize_directory_kana(row["ひらがな"] or row["顧客名"]),
        axis=1,
    )
    customers["五十音"] = customers["並び順"].map(get_customer_directory_group)

    kana_filter = st.selectbox(
        "五十音で絞り込み",
        ["すべて", *CUSTOMER_DIRECTORY_GROUPS.keys(), "その他"],
        key="customer_directory_kana_filter",
    )
    keyword = st.text_input(
        "一覧を絞り込み",
        placeholder="顧客名・ひらがな・地域",
        key="customer_directory_keyword",
    ).strip()

    filtered = customers
    if kana_filter != "すべて":
        filtered = filtered[filtered["五十音"] == kana_filter]
    if keyword:
        keyword_kana = normalize_directory_kana(keyword)
        name_text = filtered["顧客名"].astype(str)
        region_text = filtered["地域"].astype(str)
        kana_text = filtered["並び順"].astype(str)
        filtered = filtered[
            name_text.str.contains(keyword, case=False, na=False, regex=False)
            | region_text.str.contains(keyword, case=False, na=False, regex=False)
            | kana_text.str.contains(keyword_kana, case=False, na=False, regex=False)
        ]

    filtered = filtered.sort_values(["並び順", "顧客名"]).reset_index(drop=True)
    st.write(f"顧客：{len(filtered)}件")

    if filtered.empty:
        st.info("条件に一致する顧客がありません。")
        return

    parts = ['<div class="customer-directory">']
    for _, row in filtered.iterrows():
        name = clean_value(row["顧客名"])
        region = clean_value(row["地域"], blank_text="未設定")
        product_count = int(row["商品数"])
        url = html.escape(make_app_url(page="detail", customer=name), quote=True)
        parts.append(
            (
                f'<a class="customer-directory-item" href="{url}" target="_self">'
                f'<span class="customer-directory-name">{html.escape(name)}</span>'
                f'<span class="customer-directory-meta">地域：{html.escape(region)}　商品：{product_count}件</span>'
                '</a>'
            )
        )
    parts.append("</div>")
    st.markdown("".join(parts), unsafe_allow_html=True)


# =========================
# 顧客検索
# =========================
def show_customer_search(df=None, show_home_link=False):
    st.subheader("🔍 顧客検索")
    if show_home_link:
        show_back_home_button("customer_back_home")
    st.caption(f"🎤 {VOICE_INPUT_HELP} 漢字の顧客名でも検索できます。")

    page_name = "customer" if show_home_link else "home"

    default_keyword = str(get_query_value("customer_search", "")).strip()
    if st_keyup is not None:
        keyword = str(
            st_keyup(
                "ひらがな・漢字で検索",
                value=default_keyword,
                placeholder="入力すると候補がすぐ表示されます",
                debounce=250,
                key="customer_search_live",
            )
            or ""
        ).strip()
    else:
        # 追加部品がまだインストールされていない環境でもアプリを止めない。
        keyword = st.text_input(
            "ひらがな・漢字で検索",
            value=default_keyword,
            placeholder="例：こ、こも、むら",
            key="customer_search_input",
            help=VOICE_INPUT_HELP,
        ).strip()

    if keyword:
        update_query_params(page=page_name, customer_search=keyword)
    else:
        update_query_params(page=page_name, customer_search=None)

    if not keyword:
        st.info("顧客名のひらがなを入力してください。")
        return

    # ログイン直後はExcelを読まず、実際に検索が始まった時だけ取得する。
    if df is None:
        with st.spinner("顧客データを読み込んでいます…"):
            df = load_data()

    hit = df[
        df["ひらがな"].str.startswith(keyword, na=False)
        | df["顧客名"].str.contains(keyword, na=False, regex=False)
    ]

    if hit.empty:
        st.warning("該当する顧客がありません。")
        return

    customers = hit[["顧客名", "地域"]].drop_duplicates().reset_index(drop=True)
    line_by_customer = load_line_statuses_from_supabase()

    st.write(f"候補：{len(customers)}件")

    for i, row in customers.iterrows():
        name = clean_value(row["顧客名"])
        region = clean_value(row["地域"])
        line_connected = line_by_customer.get(row["顧客名"], False)

        with st.container(border=True):
            render_customer_name_with_line(name, line_connected)
            st.write(f"地域：{region}")

            st.markdown(
                render_page_link("この顧客を見る", page="detail", customer=name, customer_search=keyword),
                unsafe_allow_html=True,
            )


# =========================
# 地域検索
# =========================
def show_region_search(df):
    st.subheader("📍 地域検索")
    show_back_home_button("region_back_home")
    st.caption(f"🎤 {VOICE_INPUT_HELP}")

    default_keyword = str(get_query_value("region_search", "")).strip()
    keyword = st.text_input(
        "地域名で検索",
        value=default_keyword,
        placeholder="例：帯広、芽室、釧路",
        key="region_search_input",
        help=VOICE_INPUT_HELP,
    ).strip()

    if keyword:
        update_query_params(page="region", region_search=keyword)
    else:
        update_query_params(page="region", region_search=None)

    if not keyword:
        st.info("地域名を入力してください。")
        return

    region_text = df["地域"].fillna("").astype(str).str.strip()
    hit = df[region_text.str.contains(keyword, na=False)]

    if hit.empty:
        st.warning("該当する地域の顧客が見つかりません。")
        return

    customers = (
        hit[["顧客名", "地域"]]
        .drop_duplicates()
        .sort_values(["地域", "顧客名"])
        .reset_index(drop=True)
    )

    st.write(f"候補：{len(customers)}件")

    for i, row in customers.iterrows():
        name = clean_value(row["顧客名"])
        region = clean_value(row["地域"])

        with st.container(border=True):
            st.markdown(f"### 👤 {name}")
            st.write(f"地域：{region}")

            st.markdown(
                render_page_link("この顧客を見る", page="detail", customer=name, region_search=keyword),
                unsafe_allow_html=True,
            )


# =========================
# 配車カレンダー
# =========================
DISPATCH_COLUMN_CANDIDATES = {
    "date": ["次回配達予定", "配達予定日", "配送予定日", "配達日", "配送日", "納品日", "予定日", "日付"],
    "customer": ["顧客名", "牧場名", "取引先名", "得意先名", "お客様名", "名前", "名称"],
    "region": ["地域", "地区", "エリア", "住所", "市町村"],
    "product": ["商品名", "商品", "品名", "製品名"],
    "maker": ["メーカー", "製造元", "製造メーカー"],
}

DISPATCH_REQUIRED_LABELS = {
    "date": "日付（例：次回配達予定）",
    "customer": "顧客名・牧場名",
    "region": "地域",
    "product": "商品名",
}

WEEKDAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]


def find_dispatch_columns(df):
    """配車カレンダーに必要な列を、候補名から探す"""
    return {
        key: find_existing_column(df, candidates)
        for key, candidates in DISPATCH_COLUMN_CANDIDATES.items()
    }


def show_missing_dispatch_columns_error(df, dispatch_columns):
    missing = [
        label
        for key, label in DISPATCH_REQUIRED_LABELS.items()
        if not dispatch_columns.get(key)
    ]

    if not missing:
        return False

    st.error("配車カレンダーに必要な列が見つかりません。")
    st.write("見つからない項目：", missing)
    st.write("次のような列名が使えます。")
    st.code(
        "日付: 次回配達予定 / 配達予定日 / 日付\n"
        "顧客: 顧客名 / 牧場名 / 取引先名\n"
        "地域: 地域 / 地区 / エリア\n"
        "商品: 商品名 / 商品 / 品名"
    )
    st.write("Excelから読み取れた列：")
    st.write(list(df.columns))
    return True


def get_default_dispatch_date(df, date_column):
    """予定がある日付のうち、今日以降で一番近い日を初期表示にする"""
    parsed = pd.to_datetime(df[date_column], errors="coerce").dropna()
    if parsed.empty:
        return date.today()

    available_dates = sorted(set(parsed.dt.date))
    today = date.today()

    for target_date in available_dates:
        if target_date >= today:
            return target_date

    return available_dates[-1]


def get_calendar_month_start(df, date_column):
    """表示中の月をsession_stateで保持する"""
    if "dispatch_calendar_year" not in st.session_state or "dispatch_calendar_month" not in st.session_state:
        default_date = get_default_dispatch_date(df, date_column)
        st.session_state["dispatch_calendar_year"] = default_date.year
        st.session_state["dispatch_calendar_month"] = default_date.month

    return date(
        st.session_state["dispatch_calendar_year"],
        st.session_state["dispatch_calendar_month"],
        1,
    )


def change_dispatch_month(delta):
    current = date(
        st.session_state["dispatch_calendar_year"],
        st.session_state["dispatch_calendar_month"],
        1,
    )
    month = current.month + delta
    year = current.year

    if month < 1:
        month = 12
        year -= 1
    elif month > 12:
        month = 1
        year += 1

    st.session_state["dispatch_calendar_year"] = year
    st.session_state["dispatch_calendar_month"] = month


def clean_dispatch_maker(value):
    """カレンダーに表示するメーカー名を整える。空白と数値の0は表示しない。"""
    maker = clean_value(value, blank_text="").strip()
    if not maker:
        return ""

    try:
        if float(maker.replace(",", "")) == 0:
            return ""
    except ValueError:
        pass

    return maker


def format_month_day(target_day):
    weekday = WEEKDAYS_JA[target_day.weekday()]
    return f"{target_day.month}/{target_day.day}（{weekday}）"


def make_dispatch_items_by_day(df, month_start, dispatch_columns):
    last_day = calendar.monthrange(month_start.year, month_start.month)[1]
    start_day = date(month_start.year, month_start.month, 1)
    end_day = date(month_start.year, month_start.month, last_day)

    rows_by_day = {
        date(month_start.year, month_start.month, day_num): []
        for day_num in range(1, last_day + 1)
    }

    date_column = dispatch_columns["date"]
    customer_column = dispatch_columns["customer"]
    region_column = dispatch_columns["region"]
    product_column = dispatch_columns["product"]
    maker_column = dispatch_columns.get("maker")
    parsed_dates = pd.to_datetime(df[date_column], errors="coerce").dt.date

    for idx, row in df.iterrows():
        delivery_date = parsed_dates.loc[idx]

        if pd.isna(delivery_date) or not (start_day <= delivery_date <= end_day):
            continue

        item = {
            "顧客名": clean_value(row[customer_column]),
            "地域": clean_value(row[region_column]),
            "商品名": clean_value(row[product_column]),
            "メーカー": clean_dispatch_maker(row[maker_column]) if maker_column else "",
        }
        rows_by_day[delivery_date].append(item)

    for delivery_date, items in rows_by_day.items():
        rows_by_day[delivery_date] = sorted(
            items,
            key=lambda item: (item["顧客名"], item["地域"], item["商品名"]),
        )

    return rows_by_day


def inject_dispatch_calendar_css():
    st.markdown(
        """
        <style>
        .dispatch-month-title {
            text-align: center;
            font-size: 1.2rem;
            font-weight: 700;
            padding-top: 0.35rem;
        }
        .dispatch-two-day-row {
            display: grid;
            grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
            gap: 0.75rem;
            margin: 0.75rem 0 1.25rem;
        }
        .dispatch-day-panel {
            border: 1px solid rgba(49, 51, 63, 0.18);
            border-radius: 16px;
            color: #111827 !important;
            overflow: visible;
            padding: 0.8rem;
            background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.07);
            min-width: 0;
        }
        .dispatch-day-title {
            display: block;
            color: #111827 !important;
            background: linear-gradient(135deg, #eff6ff 0%, #ecfdf5 100%);
            border-radius: 12px;
            font-size: 0.95rem;
            font-weight: 700;
            line-height: 1.35;
            margin-bottom: 0.55rem;
            overflow-wrap: anywhere;
            padding: 0.25rem 0.4rem;
            text-align: center;
            white-space: normal;
        }
        .dispatch-item {
            border-top: 1px solid rgba(49, 51, 63, 0.12);
            display: block;
            overflow: visible;
            padding: 0.55rem 0;
        }
        .dispatch-item:first-of-type {
            border-top: 0;
            padding-top: 0;
        }
        .dispatch-name {
            color: #111827 !important;
            display: block;
            font-size: 0.95rem;
            font-weight: 700;
            line-height: 1.35;
            overflow-wrap: anywhere;
            white-space: normal;
            word-break: normal;
        }
        .dispatch-name a,
        .dispatch-month-link {
            color: #2563eb !important;
            font-weight: 700;
            text-decoration: none;
        }
        .dispatch-name a:hover,
        .dispatch-month-link:hover {
            text-decoration: underline;
        }
        .dispatch-month-product {
            color: #374151 !important;
            font-size: 0.82rem;
            white-space: normal;
        }
        .dispatch-line,
        .dispatch-empty {
            color: #374151 !important;
            display: block;
            font-size: 0.9rem;
            line-height: 1.45;
            overflow-wrap: anywhere;
            white-space: normal;
            word-break: normal;
        }
        .dispatch-month-scroll {
            width: 100%;
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
            border: 1px solid rgba(15, 23, 42, 0.10);
            border-radius: 16px;
            background: #ffffff;
            box-shadow: 0 10px 26px rgba(15, 23, 42, 0.08);
        }
        .dispatch-month-table {
            border-collapse: collapse;
            min-width: 760px;
            width: max-content;
            color: #111827 !important;
            table-layout: auto;
        }
        .dispatch-month-table th,
        .dispatch-month-table td {
            border-bottom: 1px solid rgba(49, 51, 63, 0.12);
            border-right: 1px solid rgba(49, 51, 63, 0.08);
            padding: 0.45rem 0.6rem;
            text-align: left;
            vertical-align: top;
            white-space: nowrap;
            min-width: 130px;
            font-size: 0.9rem;
            position: static !important;
        }
        .dispatch-month-table th:first-child,
        .dispatch-month-table td:first-child {
            min-width: 86px;
            position: sticky !important;
            left: 0;
            z-index: 3;
            background: #ffffff;
        }
        .dispatch-month-table th:first-child {
            z-index: 4;
            background: #f3f4f6;
        }
        .dispatch-month-table th {
            background: #eff6ff;
            font-weight: 800;
        }
        @media (max-width: 420px) {
            .dispatch-two-day-row {
                gap: 0.45rem;
            }
            .dispatch-day-panel {
                padding: 0.6rem;
            }
            .dispatch-name,
            .dispatch-line,
            .dispatch-empty {
                font-size: 0.85rem;
            }
            .dispatch-day-title {
                font-size: 0.82rem;
                padding-left: 0.25rem;
                padding-right: 0.25rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def escape_html(value):
    return html.escape(clean_value(value), quote=True)


def build_customer_detail_link(customer_name, label=None, class_name="dispatch-month-link"):
    """配車カレンダーから顧客詳細へ移動するリンクを作る"""
    customer = clean_value(customer_name, blank_text="").strip()

    if not customer:
        return escape_html(label or customer_name)

    link_label = label or customer
    url = make_app_url(page="detail", customer=customer)
    return f'<a class="{class_name}" href="{url}" target="_self">{escape_html(link_label)}</a>'


def handle_customer_query_param():
    """旧リンク互換用。URLの顧客名を消さず、ブラウザ戻るに使えるよう保持する。"""
    sync_page_from_query_params()


def build_two_day_panel_html(target_day, items):
    parts = [
        '<div class="dispatch-day-panel">',
        f'<div class="dispatch-day-title">{html.escape(format_month_day(target_day))}</div>',
    ]

    if not items:
        parts.append('<div class="dispatch-empty">予定なし</div>')
    else:
        for item in items:
            customer_link = build_customer_detail_link(item.get("顧客名"), class_name="dispatch-month-link")
            parts.append('<div class="dispatch-item">')
            parts.append(f'<div class="dispatch-name">👤 {customer_link}</div>')
            parts.append(f'<div class="dispatch-line">地域：{escape_html(item.get("地域"))}</div>')
            parts.append(f'<div class="dispatch-line">商品：{escape_html(item.get("商品名"))}</div>')
            maker = clean_dispatch_maker(item.get("メーカー"))
            if maker:
                parts.append(f'<div class="dispatch-line">メーカー：{escape_html(maker)}</div>')
            parts.append('</div>')

    parts.append('</div>')
    return "".join(parts)


def show_dispatch_month_switcher(month_start):
    col_prev, col_month, col_next = st.columns([1, 2, 1])

    with col_prev:
        st.button(
            "◀",
            key="dispatch_prev_month",
            use_container_width=True,
            on_click=change_dispatch_month,
            args=(-1,),
        )

    with col_month:
        st.markdown(
            f'<div class="dispatch-month-title">{month_start.year}年{month_start.month}月</div>',
            unsafe_allow_html=True,
        )

    with col_next:
        st.button(
            "▶",
            key="dispatch_next_month",
            use_container_width=True,
            on_click=change_dispatch_month,
            args=(1,),
        )


def show_two_day_dispatch_calendar(rows_by_day, month_start):
    st.subheader("📱 2日表示")
    st.caption("スマホでも2日分を横並びで表示します。")

    last_day = calendar.monthrange(month_start.year, month_start.month)[1]

    for day_num in range(1, last_day + 1, 2):
        day1 = date(month_start.year, month_start.month, day_num)
        day2 = date(month_start.year, month_start.month, day_num + 1) if day_num + 1 <= last_day else None

        left_panel = build_two_day_panel_html(day1, rows_by_day.get(day1, []))
        right_panel = build_two_day_panel_html(day2, rows_by_day.get(day2, [])) if day2 else '<div></div>'

        st.markdown(
            f'<div class="dispatch-two-day-row">{left_panel}{right_panel}</div>',
            unsafe_allow_html=True,
        )

def format_month_cell_item(item):
    customer_name = clean_value(item.get("顧客名"))
    product_name = clean_value(item.get("商品名"), blank_text="").strip()
    maker = clean_dispatch_maker(item.get("メーカー"))
    customer_link = build_customer_detail_link(customer_name, class_name="dispatch-month-link")

    if not product_name and not maker:
        return customer_link

    product_label = f"{product_name}/{maker}" if product_name and maker else product_name or maker
    return f'{customer_link}<br><span class="dispatch-month-product">{escape_html(product_label)}</span>'


def make_month_dispatch_table(rows_by_day, month_start):
    last_day = calendar.monthrange(month_start.year, month_start.month)[1]
    max_count = max((len(items) for items in rows_by_day.values()), default=0)
    farm_column_count = max(5, max_count)

    table_rows = []

    for day_num in range(1, last_day + 1):
        target_day = date(month_start.year, month_start.month, day_num)
        items = rows_by_day.get(target_day, [])

        row_data = {"月/日": format_month_day(target_day)}

        for item_index in range(farm_column_count):
            column_name = f"牧場名{item_index + 1}"
            row_data[column_name] = format_month_cell_item(items[item_index]) if item_index < len(items) else ""

        table_rows.append(row_data)

    return pd.DataFrame(table_rows)


def show_month_dispatch_calendar(rows_by_day, month_start):
    st.subheader("🗓 月表示")
    st.caption("横スクロールで1か月分を確認できます。日付列は固定表示します。")

    month_df = make_month_dispatch_table(rows_by_day, month_start)

    header_cells = "".join(
        f'<th>{html.escape(str(column))}</th>'
        for column in month_df.columns
    )

    body_rows = []
    for _, row in month_df.iterrows():
        row_cells = []
        for column in month_df.columns:
            value = row[column]
            if str(value) == "nan":
                cell_value = ""
            elif column == "月/日":
                cell_value = html.escape(str(value))
            else:
                cell_value = str(value)
            row_cells.append(f"<td>{cell_value}</td>")
        cells = "".join(row_cells)
        body_rows.append(f"<tr>{cells}</tr>")

    table_html = f"""
    <div class="dispatch-month-scroll">
      <table class="dispatch-month-table">
        <thead><tr>{header_cells}</tr></thead>
        <tbody>{''.join(body_rows)}</tbody>
      </table>
    </div>
    """

    st.markdown(table_html, unsafe_allow_html=True)

def show_dispatch_calendar(df):
    st.markdown("---")
    st.header("🗓 配車カレンダー")
    show_back_home_button("calendar_back_home")

    inject_dispatch_calendar_css()

    if df.empty:
        st.warning("Excelから読み込めるデータがありません。")
        return

    dispatch_columns = find_dispatch_columns(df)

    if show_missing_dispatch_columns_error(df, dispatch_columns):
        return

    month_start = get_calendar_month_start(df, dispatch_columns["date"])
    show_dispatch_month_switcher(month_start)
    rows_by_day = make_dispatch_items_by_day(df, month_start, dispatch_columns)

    st.caption(
        f"使用している日付列：{dispatch_columns['date']} / "
        f"顧客名：{dispatch_columns['customer']} / "
        f"地域：{dispatch_columns['region']} / "
        f"商品名：{dispatch_columns['product']}"
    )

    total_count = sum(len(items) for items in rows_by_day.values())
    st.write(f"{month_start.month}月の予定：{total_count}件")

    view = st.radio(
        "表示切替",
        ["📱 2日表示", "🗓 月表示"],
        horizontal=True,
        key="dispatch_calendar_view",
    )

    if view == "📱 2日表示":
        show_two_day_dispatch_calendar(rows_by_day, month_start)
    else:
        show_month_dispatch_calendar(rows_by_day, month_start)


# =========================
# 配車表（配車表1.xlsm・1月～12月）
# =========================
def normalize_dispatch_text(value):
    """配車表の表示・絞り込み用に前後空白と連続空白をそろえる。"""
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u3000", " ")).strip()


def read_dispatch_month_sheets(excel_source):
    """配車表1.xlsmの1月～12月シートからA～H列だけを結合する。"""
    if isinstance(excel_source, BytesIO):
        source = BytesIO(excel_source.getvalue())
    else:
        source = excel_source

    workbook = load_workbook(source, read_only=True, data_only=True)
    rows = []
    try:
        missing_sheets = [name for name in DISPATCH_MONTH_SHEETS if name not in workbook.sheetnames]
        if missing_sheets:
            raise ValueError("月別シートが見つかりません：" + "、".join(missing_sheets))

        for sheet_name in DISPATCH_MONTH_SHEETS:
            ws = workbook[sheet_name]
            headers = [normalize_dispatch_text(ws.cell(1, column).value) for column in range(1, 9)]
            if headers != DISPATCH_REQUIRED_COLUMNS:
                raise ValueError(
                    f"{sheet_name}のA～H列の見出しが想定と異なります。\n"
                    f"読み取った見出し：{' / '.join(headers)}"
                )

            for values in ws.iter_rows(min_row=2, max_col=8, values_only=True):
                if not any(value is not None and normalize_dispatch_text(value) for value in values):
                    continue

                record = dict(zip(DISPATCH_REQUIRED_COLUMNS, values))
                record["参照シート"] = sheet_name
                rows.append(record)
    finally:
        workbook.close()

    df = pd.DataFrame(rows, columns=DISPATCH_REQUIRED_COLUMNS + ["参照シート"])
    if df.empty:
        return df

    for column in ["引取先", "商品名", "数量", "運送会社", "納品先"]:
        df[column] = df[column].map(normalize_dispatch_text)

    pickup_dates = pd.to_datetime(df["引取日"], errors="coerce")
    arrival_dates = pd.to_datetime(df["着日"], errors="coerce")
    df["_引取日"] = pickup_dates.map(lambda value: value.date() if pd.notna(value) else None)
    df["_着日"] = arrival_dates.map(lambda value: value.date() if pd.notna(value) else None)
    return df


@st.cache_data(ttl=60, show_spinner=False)
def get_cached_dispatch_dropbox_content():
    access_token = get_dropbox_access_token()
    content, response = download_dropbox_team_file(
        str(DISPATCH_DROPBOX_FILE_PATH or DISPATCH_DROPBOX_DEFAULT_FILE_PATH).strip(),
        access_token,
    )
    if content is None:
        raise RuntimeError("配車表1.xlsmをDropboxから取得できませんでした。\n" + dropbox_error_text(response))
    return content


@st.cache_data(ttl=300, show_spinner=False)
def load_dispatch_board_data():
    """本番はDropbox、設定がない場合は指定されたローカルファイルから読む。"""
    dropbox_error = None
    if has_dropbox_auth_config():
        try:
            return read_dispatch_month_sheets(BytesIO(get_cached_dispatch_dropbox_content()))
        except Exception as error:
            # Dropbox側に配車表フォルダがまだ共有されていないPCでは、同期済みローカル版を使う。
            dropbox_error = error

    local_path = Path(str(DISPATCH_LOCAL_FILE or "").strip())
    if not local_path.exists():
        message = f"配車表1.xlsmが見つかりません：{local_path}"
        if dropbox_error is not None:
            message += f"\nDropbox取得エラー：{dropbox_error}"
        raise FileNotFoundError(message)
    return read_dispatch_month_sheets(local_path)


def dispatch_date_label(value):
    target = to_date(value)
    if target is None:
        return "未入力"
    weekdays = "月火水木金土日"
    return f"{target.month}/{target.day}（{weekdays[target.weekday()]}）"


def dispatch_filter_options(series):
    values = sorted({normalize_dispatch_text(value) for value in series if normalize_dispatch_text(value)})
    if any(not normalize_dispatch_text(value) for value in series):
        values.append("（空白）")
    return values


def apply_dispatch_choice_filter(df, column, selected):
    if not selected:
        return df
    selected_values = {value for value in selected if value != "（空白）"}
    include_blank = "（空白）" in selected
    normalized = df[column].map(normalize_dispatch_text)
    mask = normalized.isin(selected_values)
    if include_blank:
        mask = mask | normalized.eq("")
    return df[mask]


def apply_dispatch_date_filter(df, column, mode, range_value):
    if mode == "すべて":
        return df

    today = date.today()
    values = df[column]
    if mode == "今日":
        return df[values == today]
    if mode == "明日":
        return df[values == today + timedelta(days=1)]
    if mode == "今週":
        start = today - timedelta(days=today.weekday())
        end = start + timedelta(days=6)
        return df[values.map(lambda value: pd.notna(value) and start <= value <= end)]
    if mode == "未入力":
        return df[values.isna()]
    if mode == "期間指定" and isinstance(range_value, (tuple, list)) and len(range_value) == 2:
        start, end = range_value
        return df[values.map(lambda value: pd.notna(value) and start <= value <= end)]
    return df


def show_dispatch_filters(df):
    """Excelフィルターに近いAND条件の絞り込みを表示する。"""
    with st.expander("🔎 絞り込み", expanded=False):
        pickup_mode = st.selectbox(
            "引取日",
            ["すべて", "今日", "明日", "今週", "期間指定", "未入力"],
            key="dispatch_filter_pickup_mode",
        )
        pickup_range = None
        if pickup_mode == "期間指定":
            pickup_range = st.date_input(
                "引取日の期間",
                value=(date.today(), date.today()),
                key="dispatch_filter_pickup_range",
            )

        arrival_mode = st.selectbox(
            "着日",
            ["すべて", "今日", "明日", "今週", "期間指定", "未入力"],
            key="dispatch_filter_arrival_mode",
        )
        arrival_range = None
        if arrival_mode == "期間指定":
            arrival_range = st.date_input(
                "着日の期間",
                value=(date.today(), date.today()),
                key="dispatch_filter_arrival_range",
            )

        selected_pickups = st.multiselect(
            "引取先",
            dispatch_filter_options(df["引取先"]),
            key="dispatch_filter_pickup_places",
            placeholder="入力して候補を検索",
        )
        selected_products = st.multiselect(
            "商品名",
            dispatch_filter_options(df["商品名"]),
            key="dispatch_filter_products",
            placeholder="入力して候補を検索",
        )
        quantity_keyword = st.text_input(
            "数量",
            key="dispatch_filter_quantity",
            placeholder="例：450㎏、44本",
        ).strip()
        selected_carriers = st.multiselect(
            "運送会社",
            dispatch_filter_options(df["運送会社"]),
            key="dispatch_filter_carriers",
            placeholder="入力して候補を検索",
        )
        selected_destinations = st.multiselect(
            "納品先",
            dispatch_filter_options(df["納品先"]),
            key="dispatch_filter_destinations",
            placeholder="入力して候補を検索",
        )

        if st.button("条件をすべて解除", use_container_width=True, key="dispatch_filter_clear"):
            for key in list(st.session_state.keys()):
                if key.startswith("dispatch_filter_"):
                    del st.session_state[key]
            st.rerun()

    filtered = apply_dispatch_date_filter(df, "_引取日", pickup_mode, pickup_range)
    filtered = apply_dispatch_date_filter(filtered, "_着日", arrival_mode, arrival_range)
    filtered = apply_dispatch_choice_filter(filtered, "引取先", selected_pickups)
    filtered = apply_dispatch_choice_filter(filtered, "商品名", selected_products)
    filtered = apply_dispatch_choice_filter(filtered, "運送会社", selected_carriers)
    filtered = apply_dispatch_choice_filter(filtered, "納品先", selected_destinations)
    if quantity_keyword:
        quantity_text = filtered["数量"].map(normalize_dispatch_text)
        filtered = filtered[quantity_text.str.contains(quantity_keyword, regex=False, na=False)]
    return filtered


def render_dispatch_board_card(row):
    pickup_date = dispatch_date_label(row.get("_引取日"))
    arrival_date = dispatch_date_label(row.get("_着日"))
    pickup_place = normalize_dispatch_text(row.get("引取先")) or "未入力"
    destination = normalize_dispatch_text(row.get("納品先")) or "未入力"
    product = normalize_dispatch_text(row.get("商品名")) or "未入力"
    quantity = normalize_dispatch_text(row.get("数量")) or "未入力"
    carrier = normalize_dispatch_text(row.get("運送会社")) or "未入力"
    order_number = normalize_dispatch_text(row.get("発注番号"))

    with st.container(border=True):
        st.markdown(f"### {pickup_date} 引取 → {arrival_date} 着")
        st.write(f"**引取先：** {pickup_place}")
        st.write(f"**納品先：** {destination}")
        st.write(f"**商品名：** {product}")
        st.write(f"**数量：** {quantity}")
        st.write(f"**運送会社：** {carrier}")
        if order_number:
            st.caption(f"発注番号：{order_number}")


def show_dispatch_day_cards(df, basis_column, selected_day):
    day_rows = df[df[basis_column] == selected_day].copy()
    if day_rows.empty:
        st.info("この日の配車はありません。")
        return

    day_rows = day_rows.sort_values(
        ["_引取日", "_着日", "引取先", "納品先"],
        na_position="last",
    )
    st.subheader(f"{dispatch_date_label(selected_day)}：{len(day_rows)}件")
    for _, row in day_rows.iterrows():
        render_dispatch_board_card(row)


def show_dispatch_month_calendar(df):
    basis_label = st.radio(
        "カレンダー基準",
        ["引取日", "着日"],
        horizontal=True,
        key="dispatch_board_basis",
    )
    basis_column = "_引取日" if basis_label == "引取日" else "_着日"
    available_dates = sorted({value for value in df[basis_column] if pd.notna(value)})
    if not available_dates:
        st.info(f"{basis_label}が入力された配車はありません。")
        return

    periods = sorted({(value.year, value.month) for value in available_dates})
    today_period = (date.today().year, date.today().month)
    default_index = periods.index(today_period) if today_period in periods else len(periods) - 1
    selected_period = st.selectbox(
        "表示月",
        periods,
        index=default_index,
        format_func=lambda value: f"{value[0]}年{value[1]}月",
        key=f"dispatch_board_month_{basis_column}",
    )
    year, month = selected_period
    month_rows = df[
        df[basis_column].map(
            lambda value: pd.notna(value) and value.year == year and value.month == month
        )
    ]
    counts = month_rows.groupby(basis_column).size().to_dict()

    weekday_names = ["月", "火", "水", "木", "金", "土", "日"]
    header_columns = st.columns(7)
    for column, label in zip(header_columns, weekday_names):
        column.markdown(f"**{label}**")

    selected_key = f"dispatch_selected_day_{basis_column}"
    for week in calendar.Calendar(firstweekday=0).monthdayscalendar(year, month):
        columns = st.columns(7)
        for column, day_number in zip(columns, week):
            if day_number == 0:
                column.write("")
                continue
            target_day = date(year, month, day_number)
            count = int(counts.get(target_day, 0))
            label = f"{day_number}\n{count}件" if count else str(day_number)
            if column.button(
                label,
                key=f"dispatch_day_{basis_column}_{target_day.isoformat()}",
                use_container_width=True,
                disabled=count == 0,
            ):
                st.session_state[selected_key] = target_day

    selected_day = st.session_state.get(selected_key)
    if not isinstance(selected_day, date) or (selected_day.year, selected_day.month) != selected_period:
        selected_day = date.today() if date.today() in counts else min(counts)
        st.session_state[selected_key] = selected_day
    show_dispatch_day_cards(df, basis_column, selected_day)


def show_dispatch_filtered_list(df):
    st.subheader(f"絞り込み結果：{len(df)}件")
    if df.empty:
        st.info("条件に一致する配車はありません。")
        return

    sorted_df = df.sort_values(["_引取日", "_着日", "引取先"], na_position="last")
    for _, row in sorted_df.iterrows():
        render_dispatch_board_card(row)


def render_dispatch_responsive_list(display_df):
    """PCはExcel風一覧、スマホは横スクロール不要の縦型カードで表示する。"""
    st.markdown(
        """
        <style>
        .dispatch-desktop-view {
            display: block;
            max-height: 760px;
            overflow-y: auto;
            overflow-x: hidden;
            border: 1px solid #cbd5e1;
            border-radius: 10px;
            background: #ffffff;
        }
        .dispatch-excel-table {
            width: 100%;
            min-width: 0;
            table-layout: fixed;
            border-collapse: separate;
            border-spacing: 0;
            font-size: 13px;
            color: #172033;
        }
        .dispatch-excel-table th:nth-child(1),
        .dispatch-excel-table td:nth-child(1) { width: 9%; }
        .dispatch-excel-table th:nth-child(2),
        .dispatch-excel-table td:nth-child(2) { width: 20%; }
        .dispatch-excel-table th:nth-child(3),
        .dispatch-excel-table td:nth-child(3) { width: 18%; }
        .dispatch-excel-table th:nth-child(4),
        .dispatch-excel-table td:nth-child(4) { width: 10%; }
        .dispatch-excel-table th:nth-child(5),
        .dispatch-excel-table td:nth-child(5) { width: 15%; }
        .dispatch-excel-table th:nth-child(6),
        .dispatch-excel-table td:nth-child(6) { width: 19%; }
        .dispatch-excel-table th:nth-child(7),
        .dispatch-excel-table td:nth-child(7) { width: 9%; }
        .dispatch-excel-table th {
            position: sticky;
            top: 0;
            z-index: 2;
            padding: 9px 6px;
            background: #dbeaf7;
            border-right: 1px solid #94a3b8;
            border-bottom: 2px solid #64748b;
            text-align: center;
            white-space: nowrap;
        }
        .dispatch-excel-table td {
            padding: 7px 6px;
            border-right: 1px solid #cbd5e1;
            border-bottom: 1px solid #cbd5e1;
            vertical-align: middle;
            background: #ffffff;
            overflow-wrap: anywhere;
            word-break: break-word;
        }
        .dispatch-excel-table tr:nth-child(even) td { background: #f8fafc; }
        .dispatch-excel-table .date-cell,
        .dispatch-excel-table .quantity-cell { text-align: center; white-space: nowrap; }
        .dispatch-mobile-view { display: none; }

        @media (max-width: 768px) {
            .dispatch-desktop-view { display: none; }
            .dispatch-mobile-view { display: block; }
            .dispatch-day-group { margin: 0 0 18px 0; }
            .dispatch-day-heading {
                position: sticky;
                top: 0;
                z-index: 2;
                margin: 0 0 7px 0;
                padding: 8px 10px;
                border-left: 5px solid #2563eb;
                border-radius: 7px;
                background: #eaf2ff;
                color: #172033;
                font-size: 16px;
                font-weight: 800;
            }
            .dispatch-mobile-card {
                margin: 0 0 8px 0;
                padding: 10px 11px;
                border: 1px solid #cbd5e1;
                border-radius: 10px;
                background: #ffffff;
                box-shadow: 0 1px 3px rgba(15, 23, 42, 0.08);
                color: #172033;
            }
            .dispatch-date-line {
                display: grid;
                grid-template-columns: minmax(0, 1fr) 20px minmax(0, 1fr);
                align-items: center;
                gap: 4px;
                margin-bottom: 8px;
            }
            .dispatch-date-box {
                padding: 6px 7px;
                border-radius: 7px;
                text-align: center;
                line-height: 1.25;
            }
            .dispatch-pickup-date { background: #e8f1ff; color: #174ea6; }
            .dispatch-arrival-date { background: #e8f8ef; color: #176b3a; }
            .dispatch-date-label { display: block; font-size: 10px; font-weight: 700; }
            .dispatch-date-value { display: block; margin-top: 2px; font-size: 14px; font-weight: 800; }
            .dispatch-date-arrow { text-align: center; color: #64748b; font-weight: 800; }
            .dispatch-route {
                display: grid;
                grid-template-columns: minmax(0, 1fr) 20px minmax(0, 1fr);
                align-items: stretch;
                gap: 4px;
                margin-bottom: 8px;
            }
            .dispatch-route-box {
                min-width: 0;
                padding: 6px 7px;
                border: 1px solid #e2e8f0;
                border-radius: 7px;
                background: #fafcff;
            }
            .dispatch-route-label,
            .dispatch-detail-label { display: block; color: #64748b; font-size: 10px; font-weight: 700; }
            .dispatch-route-value { display: block; margin-top: 2px; font-size: 13px; font-weight: 800; overflow-wrap: anywhere; }
            .dispatch-route-arrow { align-self: center; text-align: center; color: #64748b; font-weight: 800; }
            .dispatch-details {
                display: grid;
                grid-template-columns: minmax(0, 1.4fr) minmax(0, 0.8fr);
                gap: 6px;
                margin-bottom: 6px;
            }
            .dispatch-detail-box {
                min-width: 0;
                padding: 5px 7px;
                border-radius: 6px;
                background: #f8fafc;
            }
            .dispatch-detail-value { display: block; margin-top: 1px; font-size: 13px; font-weight: 700; overflow-wrap: anywhere; }
            .dispatch-carrier {
                padding: 6px 7px;
                border-radius: 6px;
                background: #fff7e6;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    def safe_value(value):
        text = normalize_dispatch_text(value) or "未入力"
        return html.escape(text)

    columns = ["引取日", "引取先", "商品名", "数量", "運送会社", "納品先", "着日"]
    desktop_parts = [
        '<div class="dispatch-desktop-view">',
        '<table class="dispatch-excel-table"><thead><tr>',
    ]
    desktop_parts.extend(f"<th>{html.escape(column)}</th>" for column in columns)
    desktop_parts.append("</tr></thead><tbody>")
    for _, row in display_df.iterrows():
        desktop_parts.append("<tr>")
        for column in columns:
            css_class = "date-cell" if column in ["引取日", "着日"] else "quantity-cell" if column == "数量" else ""
            desktop_parts.append(f'<td class="{css_class}">{safe_value(row.get(column))}</td>')
        desktop_parts.append("</tr>")
    desktop_parts.append("</tbody></table></div>")

    mobile_parts = ['<div class="dispatch-mobile-view">']
    for pickup_date, day_rows in display_df.groupby("引取日", sort=False, dropna=False):
        pickup_label = safe_value(pickup_date)
        mobile_parts.append('<section class="dispatch-day-group">')
        mobile_parts.append(
            f'<div class="dispatch-day-heading">{pickup_label}　引取 {len(day_rows)}件</div>'
        )
        for _, row in day_rows.iterrows():
            mobile_parts.extend(
                [
                    '<article class="dispatch-mobile-card">',
                    '<div class="dispatch-date-line">',
                    f'<div class="dispatch-date-box dispatch-pickup-date"><span class="dispatch-date-label">引取日</span><span class="dispatch-date-value">{safe_value(row.get("引取日"))}</span></div>',
                    '<div class="dispatch-date-arrow">→</div>',
                    f'<div class="dispatch-date-box dispatch-arrival-date"><span class="dispatch-date-label">着日</span><span class="dispatch-date-value">{safe_value(row.get("着日"))}</span></div>',
                    '</div>',
                    '<div class="dispatch-route">',
                    f'<div class="dispatch-route-box"><span class="dispatch-route-label">引取先</span><span class="dispatch-route-value">{safe_value(row.get("引取先"))}</span></div>',
                    '<div class="dispatch-route-arrow">→</div>',
                    f'<div class="dispatch-route-box"><span class="dispatch-route-label">納品先</span><span class="dispatch-route-value">{safe_value(row.get("納品先"))}</span></div>',
                    '</div>',
                    '<div class="dispatch-details">',
                    f'<div class="dispatch-detail-box"><span class="dispatch-detail-label">商品名</span><span class="dispatch-detail-value">{safe_value(row.get("商品名"))}</span></div>',
                    f'<div class="dispatch-detail-box"><span class="dispatch-detail-label">数量</span><span class="dispatch-detail-value">{safe_value(row.get("数量"))}</span></div>',
                    '</div>',
                    f'<div class="dispatch-carrier"><span class="dispatch-detail-label">運送会社</span><span class="dispatch-detail-value">{safe_value(row.get("運送会社"))}</span></div>',
                    '</article>',
                ]
            )
        mobile_parts.append("</section>")
    mobile_parts.append("</div>")

    st.markdown("".join(desktop_parts + mobile_parts), unsafe_allow_html=True)


def show_dispatch_board():
    st.markdown("---")
    st.header("🚚 配車表")
    show_back_home_button("dispatch_board_back_home")
    st.caption("配車表1.xlsmの月別シートを、元のExcelに近い一覧で表示します。")

    with st.spinner("配車表を読み込んでいます…"):
        df = load_dispatch_board_data()
    if df.empty:
        st.warning("1月～12月シートに表示できるデータがありません。")
        return

    current_month_name = f"{date.today().month}月"
    default_month_index = (
        DISPATCH_MONTH_SHEETS.index(current_month_name)
        if current_month_name in DISPATCH_MONTH_SHEETS
        else 0
    )
    selected_month = st.selectbox(
        "表示する月",
        DISPATCH_MONTH_SHEETS,
        index=default_month_index,
        key="dispatch_table_month",
    )

    previous_month = st.session_state.get("_dispatch_table_previous_month")
    if previous_month is not None and previous_month != selected_month:
        for key in list(st.session_state.keys()):
            if key.startswith("dispatch_filter_"):
                del st.session_state[key]
        st.session_state["_dispatch_table_previous_month"] = selected_month
        st.rerun()
    st.session_state["_dispatch_table_previous_month"] = selected_month

    month_df = df[df["参照シート"] == selected_month].copy()
    filtered = show_dispatch_filters(month_df)

    st.markdown(
        f"**参照：{selected_month}シート　｜　全 {len(month_df)}件　｜　条件一致 {len(filtered)}件**"
    )
    st.caption("※ 1件は、元のExcelの月別シートにある明細1行です。")

    if filtered.empty:
        st.info("条件に一致する配車はありません。")
        return

    display_df = filtered.sort_values(
        ["_引取日", "_着日", "発注番号"],
        na_position="last",
    )[["引取日", "引取先", "商品名", "数量", "運送会社", "納品先", "着日"]].copy()

    display_df["引取日"] = filtered.loc[display_df.index, "_引取日"].map(dispatch_date_label)
    display_df["着日"] = filtered.loc[display_df.index, "_着日"].map(dispatch_date_label)
    for column in ["引取先", "商品名", "数量", "運送会社", "納品先"]:
        display_df[column] = display_df[column].map(normalize_dispatch_text)

    render_dispatch_responsive_list(display_df)


# =========================
# ソリュブル在庫（aoベンチャーグレイン配車表.xlsx）
# =========================
def soluble_cell_is_manual(cell):
    """Excelで黄色に塗られたセルを手入力値として扱う。"""
    if cell.fill.fill_type != "solid":
        return False
    color = cell.fill.fgColor
    if color.type == "rgb":
        return str(color.rgb or "").upper().endswith("FFFF00")
    return False


def soluble_formula_value(formula_ws, value_ws, row, column, memo=None, visiting=None):
    """保存後にExcelの計算キャッシュが空でも、対象表の単純な加減式を表示できるようにする。"""
    memo = memo if memo is not None else {}
    visiting = visiting if visiting is not None else set()
    key = (row, column)
    if key in memo:
        return memo[key]
    if key in visiting:
        return None

    raw = formula_ws.cell(row, column).value
    cached = value_ws.cell(row, column).value
    if not (isinstance(raw, str) and raw.startswith("=")):
        memo[key] = raw
        return raw
    if cached is not None:
        memo[key] = cached
        return cached

    expression = raw[1:].replace(" ", "").replace("$", "").upper()
    tokens = re.findall(r"[A-Z]+\d+|\d+(?:\.\d+)?|[+-]", expression)
    if not tokens or "".join(tokens) != expression:
        return None

    visiting.add(key)
    try:
        def token_value(token):
            match = re.fullmatch(r"([A-Z]+)(\d+)", token)
            if match:
                letters, target_row = match.groups()
                target_column = 0
                for letter in letters:
                    target_column = target_column * 26 + (ord(letter) - 64)
                return soluble_formula_value(
                    formula_ws,
                    value_ws,
                    int(target_row),
                    target_column,
                    memo,
                    visiting,
                )
            return float(token)

        result = token_value(tokens[0])
        index = 1
        while index < len(tokens):
            operator = tokens[index]
            right = token_value(tokens[index + 1])
            if result is None:
                result = 0
            if right is None:
                right = 0
            if isinstance(result, (date, datetime)) and isinstance(right, (int, float)):
                result = result + timedelta(days=right if operator == "+" else -right)
            else:
                result = result + right if operator == "+" else result - right
            index += 2
        memo[key] = result
        return result
    except Exception:
        return None
    finally:
        visiting.discard(key)


def soluble_date_value(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        return date(1899, 12, 30) + timedelta(days=int(value))
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.date()


def soluble_number_label(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    if isinstance(value, (int, float)):
        number = float(value)
        if number.is_integer():
            return f"{int(number):,}"
        return f"{number:,.2f}".rstrip("0").rstrip(".")
    return str(value)


def soluble_input_value(value):
    if value is None:
        return ""
    if isinstance(value, (int, float)) and float(value).is_integer():
        return str(int(value))
    return str(value)


def parse_soluble_number(text, label):
    cleaned = str(text or "").strip().replace(",", "").replace("，", "")
    if not cleaned:
        return None
    try:
        number = float(cleaned)
    except ValueError as exc:
        raise ValueError(f"{label}は数字で入力してください。") from exc
    return int(number) if number.is_integer() else number


def same_soluble_value(left, right):
    if left is None and right is None:
        return True
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return math.isclose(float(left), float(right), rel_tol=0, abs_tol=1e-9)
    return left == right


def normalize_soluble_customer_name(value):
    """ソリュブル上段の顧客名照合用。半角・全角空白の違いだけを吸収する。"""
    return re.sub(r"[\s　]+", "", clean_value(value, blank_text=""))


def find_soluble_customer_row(ws, customer_name):
    """ソリュブルシート上段から顧客名で対象行を探す。"""
    target = normalize_soluble_customer_name(customer_name)
    if not target:
        return None

    customer_column = SOLUBLE_CUSTOMER_COLUMNS["customer_name"]
    # 上段の顧客一覧は2行目の見出しから、日別表が始まる10行目より前にある。
    for row_number in range(3, min(ws.max_row, 10) + 1):
        if normalize_soluble_customer_name(ws.cell(row_number, customer_column).value) == target:
            return row_number
    return None


def calculate_soluble_customer_next_delivery(delivery_date_value, delivery_quantity, usage):
    """G列の「配達数量÷使用数量/日＋配達日」と同じ表示日を計算する。"""
    delivery_day = soluble_date_value(delivery_date_value)
    if delivery_day is None:
        return None
    try:
        quantity = float(delivery_quantity)
        daily_usage = float(usage)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(quantity) or not math.isfinite(daily_usage) or daily_usage <= 0:
        return None
    calculated = datetime.combine(delivery_day, datetime.min.time()) + timedelta(
        days=quantity / daily_usage
    )
    return calculated.date()


@st.cache_data(ttl=60, show_spinner=False)
def read_soluble_customer_summaries(content):
    """三谷牧場・熊林牧場の上段4項目をまとめて読む。"""
    formula_wb = load_workbook(BytesIO(content), data_only=False, read_only=False)
    value_wb = load_workbook(BytesIO(content), data_only=True, read_only=False)
    try:
        if SOLUBLE_SHEET_NAME not in formula_wb.sheetnames:
            raise ValueError("ソリュブルシートが見つかりません。")
        formula_ws = formula_wb[SOLUBLE_SHEET_NAME]
        value_ws = value_wb[SOLUBLE_SHEET_NAME]
        result = {}

        for customer_name in SOLUBLE_CUSTOMER_NAMES:
            row_number = find_soluble_customer_row(formula_ws, customer_name)
            if row_number is None:
                continue

            delivery_date_value = formula_ws.cell(
                row_number, SOLUBLE_CUSTOMER_COLUMNS["delivery_date"]
            ).value
            delivery_quantity = formula_ws.cell(
                row_number, SOLUBLE_CUSTOMER_COLUMNS["delivery_quantity"]
            ).value
            usage = formula_ws.cell(
                row_number, SOLUBLE_CUSTOMER_COLUMNS["usage"]
            ).value
            next_delivery_value = value_ws.cell(
                row_number, SOLUBLE_CUSTOMER_COLUMNS["next_delivery"]
            ).value
            next_delivery = soluble_date_value(next_delivery_value)
            if next_delivery is None:
                next_delivery = calculate_soluble_customer_next_delivery(
                    delivery_date_value,
                    delivery_quantity,
                    usage,
                )

            result[customer_name] = {
                "row": row_number,
                "顧客名": customer_name,
                "配達日": soluble_date_value(delivery_date_value),
                "配達数量": delivery_quantity,
                "次回配達予定": next_delivery,
                "使用数量/日": usage,
            }
        return result
    finally:
        formula_wb.close()
        value_wb.close()


def get_soluble_customer_summary(content, customer_name):
    """選択中の顧客に対応するソリュブル上段情報を返す。"""
    target = normalize_soluble_customer_name(customer_name)
    for name, summary in read_soluble_customer_summaries(content).items():
        if normalize_soluble_customer_name(name) == target:
            return summary
    return None


@st.cache_data(ttl=60, show_spinner=False)
def load_soluble_workbook_content():
    """Dropboxを優先し、開発用PCでは同期済みローカルファイルも利用する。"""
    target_path = str(SOLUBLE_DROPBOX_FILE_PATH or SOLUBLE_DROPBOX_DEFAULT_FILE_PATH).strip()
    if has_dropbox_auth_config():
        access_token = get_dropbox_access_token()
        content, response = download_dropbox_file(target_path, access_token)
        if content is not None:
            return content, "Dropbox"
        local_path = Path(str(SOLUBLE_LOCAL_FILE))
        if not local_path.exists():
            raise RuntimeError(
                "aoベンチャーグレイン配車表.xlsxをDropboxから取得できませんでした。\n"
                + dropbox_error_text(response)
            )
        return local_path.read_bytes(), "同期済みローカルファイル"

    local_path = Path(str(SOLUBLE_LOCAL_FILE))
    if not local_path.exists():
        raise FileNotFoundError(f"対象ファイルが見つかりません：{local_path}")
    return local_path.read_bytes(), "同期済みローカルファイル"


def read_soluble_rows(content):
    formula_wb = load_workbook(BytesIO(content), data_only=False, read_only=False)
    value_wb = load_workbook(BytesIO(content), data_only=True, read_only=False)
    try:
        if SOLUBLE_SHEET_NAME not in formula_wb.sheetnames:
            raise ValueError("ソリュブルシートが見つかりません。")
        formula_ws = formula_wb[SOLUBLE_SHEET_NAME]
        value_ws = value_wb[SOLUBLE_SHEET_NAME]
        memo = {}
        rows = []
        for row_number in range(11, formula_ws.max_row + 1):
            day_value = soluble_formula_value(formula_ws, value_ws, row_number, 2, memo)
            day = soluble_date_value(day_value)
            if day is None:
                continue
            record = {"row": row_number, "date": day}
            for location, columns in SOLUBLE_LOCATIONS.items():
                for field, column in columns.items():
                    record[f"{location}_{field}"] = soluble_formula_value(
                        formula_ws, value_ws, row_number, column, memo
                    )
                    record[f"{location}_{field}_manual"] = soluble_cell_is_manual(
                        formula_ws.cell(row_number, column)
                    )
                    record[f"{location}_{field}_formula"] = (
                        formula_ws.cell(row_number, column).value
                        if isinstance(formula_ws.cell(row_number, column).value, str)
                        and formula_ws.cell(row_number, column).value.startswith("=")
                        else ""
                    )
            rows.append(record)
        return rows
    finally:
        formula_wb.close()
        value_wb.close()


def _disabled_unsafe_xml_builder(content, row_number, location, updates):
    """XLSX全体を再生成せず、対象セルのXMLだけを変更して既存の計算結果を保つ。"""
    if location not in SOLUBLE_LOCATIONS:
        raise ValueError("対象の会社が正しくありません。")
    if row_number < 11:
        raise ValueError("更新する行が正しくありません。")
    if not updates:
        raise ValueError("変更された項目がありません。")

    columns = SOLUBLE_LOCATIONS[location]
    for field in updates:
        if field not in columns:
            raise ValueError("更新項目が正しくありません。")

    original_rows = read_soluble_rows(content)
    current_row = next((row for row in original_rows if row["row"] == row_number), None)
    previous_row = next((row for row in original_rows if row["row"] == row_number - 1), None)
    if current_row is None:
        raise ValueError("更新する日付行が見つかりません。")

    # openpyxlは共有数式を各セルの通常の式へ展開して読めるため、表示値の再計算に利用する。
    formula_book = load_workbook(BytesIO(content), data_only=False, read_only=False)
    try:
        if SOLUBLE_SHEET_NAME not in formula_book.sheetnames:
            raise ValueError("ソリュブルシートが見つかりません。")
        formula_sheet = formula_book[SOLUBLE_SHEET_NAME]
        expanded_formulas = {
            cell.coordinate: cell.value[1:]
            for row in formula_sheet.iter_rows(min_row=11, min_col=2, max_col=8)
            for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=")
        }
    finally:
        formula_book.close()

    resolved_updates = {}
    cached_values = {}
    for field, requested_value in updates.items():
        if requested_value == "__AUTO_INVENTORY__":
            if field != "inventory" or previous_row is None:
                raise ValueError("この日は在庫を自動計算にできません。")
            inventory_column = columns["inventory"]
            usage_column = columns["usage"]
            delivery_column = columns["delivery"]
            inventory_letter = chr(64 + inventory_column)
            usage_letter = chr(64 + usage_column)
            delivery_letter = chr(64 + delivery_column)
            formula = f"={inventory_letter}{row_number - 1}-{usage_letter}{row_number}+{delivery_letter}{row_number}"
            previous_inventory = previous_row.get(f"{location}_inventory") or 0
            current_usage = updates.get("usage", current_row.get(f"{location}_usage")) or 0
            current_delivery = updates.get("delivery", current_row.get(f"{location}_delivery")) or 0
            if not all(isinstance(value, (int, float)) for value in (previous_inventory, current_usage, current_delivery)):
                raise ValueError("在庫の自動計算に使う値が数字ではありません。")
            resolved_updates[field] = formula
            cached_values[field] = previous_inventory - current_usage + current_delivery
        else:
            resolved_updates[field] = requested_value
            cached_values[field] = requested_value

    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    office_rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", main_ns)

    with zipfile.ZipFile(BytesIO(content), "r") as source_zip:
        workbook_root = ET.fromstring(source_zip.read("xl/workbook.xml"))
        relationship_id = ""
        for sheet_node in workbook_root.findall(f".//{{{main_ns}}}sheet"):
            if sheet_node.get("name") == SOLUBLE_SHEET_NAME:
                relationship_id = sheet_node.get(f"{{{office_rel_ns}}}id", "")
                break
        if not relationship_id:
            raise ValueError("ソリュブルシートが見つかりません。")

        relationships_root = ET.fromstring(source_zip.read("xl/_rels/workbook.xml.rels"))
        sheet_target = ""
        for relationship in relationships_root.findall(f"{{{package_rel_ns}}}Relationship"):
            if relationship.get("Id") == relationship_id:
                sheet_target = relationship.get("Target", "")
                break
        if not sheet_target:
            raise ValueError("ソリュブルシートの保存先を確認できません。")
        sheet_part = (
            sheet_target.lstrip("/")
            if sheet_target.startswith("/")
            else posixpath.normpath(posixpath.join("xl", sheet_target))
        )

        sheet_root = ET.fromstring(source_zip.read(sheet_part))
        styles_root = ET.fromstring(source_zip.read("xl/styles.xml"))
        fills = styles_root.find(f"{{{main_ns}}}fills")
        cell_xfs = styles_root.find(f"{{{main_ns}}}cellXfs")
        if fills is None or cell_xfs is None:
            raise ValueError("Excelの表示形式を確認できません。")

        yellow_fill_id = None
        for fill_index, fill in enumerate(list(fills)):
            foreground = fill.find(f".//{{{main_ns}}}fgColor")
            if foreground is not None and str(foreground.get("rgb", "")).upper().endswith("FFFF00"):
                yellow_fill_id = fill_index
                break
        if yellow_fill_id is None:
            fill = ET.Element(f"{{{main_ns}}}fill")
            pattern = ET.SubElement(fill, f"{{{main_ns}}}patternFill", {"patternType": "solid"})
            ET.SubElement(pattern, f"{{{main_ns}}}fgColor", {"rgb": "FFFFFF00"})
            ET.SubElement(pattern, f"{{{main_ns}}}bgColor", {"indexed": "64"})
            fills.append(fill)
            yellow_fill_id = len(fills) - 1
            fills.set("count", str(len(fills)))

        style_cache = {}

        def style_with_manual_fill(style_id, manual):
            style_id = int(style_id or 0)
            if style_id >= len(cell_xfs):
                style_id = 0
            original_xf = cell_xfs[style_id]
            current_fill_id = int(original_xf.get("fillId", "0"))
            is_yellow = current_fill_id == yellow_fill_id
            if is_yellow == manual:
                return style_id
            cache_key = (style_id, manual)
            if cache_key in style_cache:
                return style_cache[cache_key]
            new_xf = copy.deepcopy(original_xf)
            new_xf.set("fillId", str(yellow_fill_id if manual else 0))
            new_xf.set("applyFill", "1" if manual else "0")
            cell_xfs.append(new_xf)
            new_style_id = len(cell_xfs) - 1
            cell_xfs.set("count", str(len(cell_xfs)))
            style_cache[cache_key] = new_style_id
            return new_style_id

        sheet_data = sheet_root.find(f"{{{main_ns}}}sheetData")
        if sheet_data is None:
            raise ValueError("ソリュブルシートのセルを確認できません。")
        row_node = sheet_data.find(f"{{{main_ns}}}row[@r='{row_number}']")
        if row_node is None:
            raise ValueError("更新する日付行が見つかりません。")

        changed = []
        for field, new_value in resolved_updates.items():
            coordinate = f"{chr(64 + columns[field])}{row_number}"
            cell_node = row_node.find(f"{{{main_ns}}}c[@r='{coordinate}']")
            if cell_node is None:
                cell_node = ET.SubElement(row_node, f"{{{main_ns}}}c", {"r": coordinate})
            is_formula = isinstance(new_value, str) and new_value.startswith("=")
            should_be_manual = not is_formula
            cell_node.set("s", str(style_with_manual_fill(cell_node.get("s", "0"), should_be_manual)))
            cell_node.attrib.pop("t", None)
            for child in list(cell_node):
                if child.tag in {f"{{{main_ns}}}f", f"{{{main_ns}}}v", f"{{{main_ns}}}is"}:
                    cell_node.remove(child)
            if is_formula:
                formula_node = ET.SubElement(cell_node, f"{{{main_ns}}}f")
                formula_node.text = new_value[1:]
            cached_value = cached_values[field]
            if cached_value is not None:
                value_node = ET.SubElement(cell_node, f"{{{main_ns}}}v")
                value_node.text = str(int(cached_value)) if isinstance(cached_value, float) and cached_value.is_integer() else str(cached_value)
            changed.append((coordinate, new_value, should_be_manual))

        # 変更後の数値を使って、ソリュブル表内の単純な加減式の表示値も更新する。
        # これによりExcelを開く前でも、アプリとExcelプレビューで最新在庫を確認できる。
        cell_nodes = {
            cell.get("r", ""): cell
            for cell in sheet_root.findall(f".//{{{main_ns}}}c")
            if cell.get("r")
        }
        calculated = {}
        calculating = set()

        def calculate_xml_cell(coordinate):
            if coordinate in calculated:
                return calculated[coordinate]
            if coordinate in calculating:
                return 0
            node = cell_nodes.get(coordinate)
            if node is None:
                return 0
            formula_node = node.find(f"{{{main_ns}}}f")
            value_node = node.find(f"{{{main_ns}}}v")
            formula_text = (
                formula_node.text
                if formula_node is not None and formula_node.text
                else expanded_formulas.get(coordinate, "") if formula_node is not None else ""
            )
            if not formula_text:
                try:
                    value = float(value_node.text) if value_node is not None and value_node.text else 0
                except ValueError:
                    value = 0
                calculated[coordinate] = value
                return value

            expression = formula_text.replace(" ", "").replace("$", "").upper()
            tokens = re.findall(r"[A-Z]+\d+|\d+(?:\.\d+)?|[+-]", expression)
            if not tokens or "".join(tokens) != expression:
                return 0
            calculating.add(coordinate)
            try:
                def token_number(token):
                    return calculate_xml_cell(token) if re.fullmatch(r"[A-Z]+\d+", token) else float(token)

                result = token_number(tokens[0])
                index = 1
                while index < len(tokens):
                    right = token_number(tokens[index + 1])
                    result = result + right if tokens[index] == "+" else result - right
                    index += 2
                calculated[coordinate] = result
                return result
            finally:
                calculating.discard(coordinate)

        for coordinate, node in cell_nodes.items():
            match = re.fullmatch(r"([B-H])(\d+)", coordinate)
            formula_node = node.find(f"{{{main_ns}}}f")
            if not match or int(match.group(2)) < 11 or formula_node is None:
                continue
            result = calculate_xml_cell(coordinate)
            value_node = node.find(f"{{{main_ns}}}v")
            if value_node is None:
                value_node = ET.SubElement(node, f"{{{main_ns}}}v")
            value_node.text = str(int(result)) if float(result).is_integer() else str(result)

        calculation_properties = workbook_root.find(f"{{{main_ns}}}calcPr")
        if calculation_properties is not None:
            calculation_properties.set("calcMode", "auto")
            calculation_properties.set("fullCalcOnLoad", "1")
            calculation_properties.set("forceFullCalc", "1")

        replacement_parts = {
            sheet_part: ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True),
            "xl/styles.xml": ET.tostring(styles_root, encoding="utf-8", xml_declaration=True),
            "xl/workbook.xml": ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True),
        }
        output = BytesIO()
        with zipfile.ZipFile(output, "w") as target_zip:
            for item in source_zip.infolist():
                target_zip.writestr(item, replacement_parts.get(item.filename, source_zip.read(item.filename)))

    saved_content = output.getvalue()
    formula_wb = load_workbook(BytesIO(saved_content), data_only=False, read_only=False)
    value_wb = load_workbook(BytesIO(saved_content), data_only=True, read_only=False)
    try:
        ws = formula_wb[SOLUBLE_SHEET_NAME]
        value_ws = value_wb[SOLUBLE_SHEET_NAME]
        for coordinate, expected, expected_manual in changed:
            cell = ws[coordinate]
            if cell.value != expected or soluble_cell_is_manual(cell) != expected_manual:
                raise ValueError(f"保存確認で{SOLUBLE_SHEET_NAME}!{coordinate}が一致しません。")
            if expected is not None and value_ws[coordinate].value is None:
                raise ValueError(f"保存確認で{SOLUBLE_SHEET_NAME}!{coordinate}の表示値がありません。")
    finally:
        formula_wb.close()
        value_wb.close()
    return saved_content, changed


def build_soluble_updated_workbook(content, row_number, location, updates):
    """openpyxlの標準保存を使い、Excel本体で開ける形式のまま対象セルを更新する。"""
    if location not in SOLUBLE_LOCATIONS:
        raise ValueError("対象の会社が正しくありません。")
    if row_number < 11:
        raise ValueError("更新する行が正しくありません。")
    if not updates:
        raise ValueError("変更された項目がありません。")

    workbook = load_workbook(BytesIO(content), data_only=False, read_only=False)
    original_sheets = list(workbook.sheetnames)
    changed = []
    yellow_fill = PatternFill(fill_type="solid", fgColor="FFFFFF00")
    clear_fill = PatternFill(fill_type=None)
    try:
        if SOLUBLE_SHEET_NAME not in workbook.sheetnames:
            raise ValueError("ソリュブルシートが見つかりません。")
        ws = workbook[SOLUBLE_SHEET_NAME]
        if row_number > ws.max_row:
            raise ValueError("更新する日付行が見つかりません。")
        columns = SOLUBLE_LOCATIONS[location]

        for field, requested_value in updates.items():
            if field not in columns:
                raise ValueError("更新項目が正しくありません。")
            if requested_value == "__AUTO_INVENTORY__":
                if field != "inventory" or row_number <= 11:
                    raise ValueError("この日は在庫を自動計算にできません。")
                inventory_letter = ws.cell(row_number, columns["inventory"]).column_letter
                usage_letter = ws.cell(row_number, columns["usage"]).column_letter
                delivery_letter = ws.cell(row_number, columns["delivery"]).column_letter
                new_value = (
                    f"={inventory_letter}{row_number - 1}-{usage_letter}{row_number}+{delivery_letter}{row_number}"
                )
                manual = False
            else:
                new_value = requested_value
                manual = True

            cell = ws.cell(row_number, columns[field])
            cell.value = new_value
            cell.fill = yellow_fill if manual else clear_fill
            changed.append((cell.coordinate, new_value, manual))

        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
        output = BytesIO()
        workbook.save(output)
    finally:
        workbook.close()

    saved_content = output.getvalue()
    verified = load_workbook(BytesIO(saved_content), data_only=False, read_only=False)
    try:
        if list(verified.sheetnames) != original_sheets:
            raise ValueError("保存後にシート構成が変わったため、更新を中止しました。")
        ws = verified[SOLUBLE_SHEET_NAME]
        for coordinate, expected, expected_manual in changed:
            cell = ws[coordinate]
            if cell.value != expected or soluble_cell_is_manual(cell) != expected_manual:
                raise ValueError(f"保存確認で{SOLUBLE_SHEET_NAME}!{coordinate}が一致しません。")
    finally:
        verified.close()
    return saved_content, changed


def ensure_soluble_backup_folder(access_token):
    response = call_dropbox_rpc(
        "files/create_folder_v2",
        {"path": SOLUBLE_BACKUP_FOLDER, "autorename": False},
        access_token,
    )
    if response.status_code == 200:
        return
    if response.status_code == 409 and "conflict" in str(response.text).lower():
        return
    raise RuntimeError("Dropboxにバックアップフォルダを作成できませんでした。\n" + dropbox_error_text(response))


def save_soluble_changes(row_number, location, updates):
    target_path = str(SOLUBLE_DROPBOX_FILE_PATH or SOLUBLE_DROPBOX_DEFAULT_FILE_PATH).strip()
    timestamp = get_jst_now().strftime("%Y%m%d_%H%M%S_%f")

    if has_dropbox_auth_config():
        access_token = get_dropbox_access_token()
        original_content, response = download_dropbox_file(target_path, access_token)
        if original_content is None:
            raise RuntimeError("最新の対象Excelを取得できませんでした。\n" + dropbox_error_text(response))
        revision = get_download_revision(response)
        if not revision:
            raise RuntimeError("Dropboxの更新番号を取得できないため、保存を中止しました。")
        saved_content, changed = build_soluble_updated_workbook(
            original_content, row_number, location, updates
        )
        ensure_soluble_backup_folder(access_token)
        backup_path = f"{SOLUBLE_BACKUP_FOLDER}/aoベンチャーグレイン配車表_{timestamp}.xlsx"
        backup_response = upload_dropbox_file(backup_path, original_content, access_token, mode="add")
        if backup_response.status_code != 200:
            raise RuntimeError("バックアップを作成できないため、本番ファイルは更新しません。\n" + dropbox_error_text(backup_response))
        upload_response = upload_dropbox_file(
            target_path, saved_content, access_token, mode="update", rev=revision
        )
        if upload_response.status_code == 409:
            raise RuntimeError("保存中にPCなどでExcelが更新されました。再読み込みしてからやり直してください。")
        if upload_response.status_code != 200:
            raise RuntimeError("対象Excelを更新できませんでした。\n" + dropbox_error_text(upload_response))
    else:
        local_path = Path(str(SOLUBLE_LOCAL_FILE))
        if not local_path.exists():
            raise FileNotFoundError(f"対象ファイルが見つかりません：{local_path}")
        original_content = local_path.read_bytes()
        saved_content, changed = build_soluble_updated_workbook(
            original_content, row_number, location, updates
        )
        backup_dir = local_path.parent / "Backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup_path = backup_dir / f"aoベンチャーグレイン配車表_{timestamp}.xlsx"
        backup_path.write_bytes(original_content)
        local_path.write_bytes(saved_content)

    st.cache_data.clear()
    return changed


def build_soluble_customer_updated_workbook(content, customer_name, updates):
    """上段顧客のE/F/H列だけを更新し、G列の数式は保持する。"""
    if normalize_soluble_customer_name(customer_name) not in {
        normalize_soluble_customer_name(name) for name in SOLUBLE_CUSTOMER_NAMES
    }:
        raise ValueError("編集対象の顧客が正しくありません。")
    if not updates:
        raise ValueError("変更された項目がありません。")

    allowed_columns = {
        "delivery_date": SOLUBLE_CUSTOMER_COLUMNS["delivery_date"],
        "delivery_quantity": SOLUBLE_CUSTOMER_COLUMNS["delivery_quantity"],
        "usage": SOLUBLE_CUSTOMER_COLUMNS["usage"],
    }
    if any(field not in allowed_columns for field in updates):
        raise ValueError("更新項目が正しくありません。")

    workbook = load_workbook(BytesIO(content), data_only=False, read_only=False)
    original_sheets = list(workbook.sheetnames)
    changed = []
    try:
        if SOLUBLE_SHEET_NAME not in workbook.sheetnames:
            raise ValueError("ソリュブルシートが見つかりません。")
        ws = workbook[SOLUBLE_SHEET_NAME]
        row_number = find_soluble_customer_row(ws, customer_name)
        if row_number is None:
            raise ValueError(f"ソリュブルシートに「{customer_name}」が見つかりません。")

        next_delivery_cell = ws.cell(
            row_number,
            SOLUBLE_CUSTOMER_COLUMNS["next_delivery"],
        )
        original_next_delivery_formula = next_delivery_cell.value
        if not (
            isinstance(original_next_delivery_formula, str)
            and original_next_delivery_formula.startswith("=")
        ):
            raise ValueError("次回配達予定の数式が見つからないため、更新を中止しました。")

        for field, new_value in updates.items():
            column_number = allowed_columns[field]
            cell = ws.cell(row_number, column_number)
            if not same_excel_value(cell.value, new_value):
                cell.value = new_value
                changed.append((cell.coordinate, new_value))

        if not changed:
            raise ValueError("変更された項目がありません。")

        # 次回配達予定は編集せず、既存のG列数式に任せる。
        if next_delivery_cell.value != original_next_delivery_formula:
            raise ValueError("次回配達予定の数式が変更されたため、保存を中止しました。")
        enable_excel_recalculation(workbook)
        output = BytesIO()
        workbook.save(output)
    finally:
        workbook.close()

    saved_content = output.getvalue()
    verified = load_workbook(BytesIO(saved_content), data_only=False, read_only=False)
    try:
        if list(verified.sheetnames) != original_sheets:
            raise ValueError("保存後にシート構成が変わったため、更新を中止しました。")
        ws = verified[SOLUBLE_SHEET_NAME]
        verified_row = find_soluble_customer_row(ws, customer_name)
        if verified_row is None:
            raise ValueError("保存後に対象顧客の行を確認できません。")
        if ws.cell(
            verified_row,
            SOLUBLE_CUSTOMER_COLUMNS["next_delivery"],
        ).value != original_next_delivery_formula:
            raise ValueError("保存後に次回配達予定の数式が変わっています。")
        for coordinate, expected in changed:
            if not same_excel_value(ws[coordinate].value, expected):
                raise ValueError(f"保存確認で{SOLUBLE_SHEET_NAME}!{coordinate}が一致しません。")
    finally:
        verified.close()
    return saved_content, changed


def verify_soluble_customer_saved_content(content, customer_name, changed):
    """Dropbox保存後に、変更セルとG列数式が残っていることを確認する。"""
    workbook = load_workbook(BytesIO(content), data_only=False, read_only=False)
    try:
        if SOLUBLE_SHEET_NAME not in workbook.sheetnames:
            raise RuntimeError("保存後の確認でソリュブルシートが見つかりません。")
        ws = workbook[SOLUBLE_SHEET_NAME]
        row_number = find_soluble_customer_row(ws, customer_name)
        if row_number is None:
            raise RuntimeError("保存後の確認で対象顧客が見つかりません。")
        formula = ws.cell(row_number, SOLUBLE_CUSTOMER_COLUMNS["next_delivery"]).value
        if not (isinstance(formula, str) and formula.startswith("=")):
            raise RuntimeError("保存後の確認で次回配達予定の数式が見つかりません。")
        for coordinate, expected in changed:
            if not same_excel_value(ws[coordinate].value, expected):
                raise RuntimeError(
                    f"Dropbox保存後の確認で{SOLUBLE_SHEET_NAME}!{coordinate}が更新されていません。"
                )
    finally:
        workbook.close()


def save_soluble_customer_changes(customer_name, updates):
    """上段顧客の変更を、既存ソリュブル保存と同じバックアップ方式で保存する。"""
    target_path = str(SOLUBLE_DROPBOX_FILE_PATH or SOLUBLE_DROPBOX_DEFAULT_FILE_PATH).strip()
    timestamp = get_jst_now().strftime("%Y%m%d_%H%M%S_%f")

    if has_dropbox_auth_config():
        access_token = get_dropbox_access_token()
        original_content, response = download_dropbox_file(target_path, access_token)
        if original_content is None:
            raise RuntimeError(
                "最新の対象Excelを取得できませんでした。\n" + dropbox_error_text(response)
            )
        revision = get_download_revision(response)
        if not revision:
            raise RuntimeError("Dropboxの更新番号を取得できないため、保存を中止しました。")

        saved_content, changed = build_soluble_customer_updated_workbook(
            original_content,
            customer_name,
            updates,
        )
        ensure_soluble_backup_folder(access_token)
        backup_path = (
            f"{SOLUBLE_BACKUP_FOLDER}/"
            f"aoベンチャーグレイン配車表_{timestamp}.xlsx"
        )
        backup_response = upload_dropbox_file(
            backup_path,
            original_content,
            access_token,
            mode="add",
        )
        if backup_response.status_code != 200:
            raise RuntimeError(
                "バックアップを作成できないため、本番ファイルは更新しません。\n"
                + dropbox_error_text(backup_response)
            )
        upload_response = upload_dropbox_file(
            target_path,
            saved_content,
            access_token,
            mode="update",
            rev=revision,
        )
        if upload_response.status_code == 409:
            raise RuntimeError(
                "保存中にPCなどでExcelが更新されました。再読み込みしてからやり直してください。"
            )
        if upload_response.status_code != 200:
            raise RuntimeError(
                "対象Excelを更新できませんでした。\n"
                + dropbox_error_text(upload_response)
            )

        confirmed_content, confirmed_response = download_dropbox_file(target_path, access_token)
        if confirmed_content is None:
            raise RuntimeError(
                "保存後のExcelを再取得できませんでした。\n"
                + dropbox_error_text(confirmed_response)
            )
        verify_soluble_customer_saved_content(
            confirmed_content,
            customer_name,
            changed,
        )
    else:
        local_path = Path(str(SOLUBLE_LOCAL_FILE))
        if not local_path.exists():
            raise FileNotFoundError(f"対象ファイルが見つかりません：{local_path}")
        original_content = local_path.read_bytes()
        saved_content, changed = build_soluble_customer_updated_workbook(
            original_content,
            customer_name,
            updates,
        )
        backup_dir = local_path.parent / "Backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup_path = backup_dir / f"aoベンチャーグレイン配車表_{timestamp}.xlsx"
        backup_path.write_bytes(original_content)
        local_path.write_bytes(saved_content)
        verify_soluble_customer_saved_content(
            local_path.read_bytes(),
            customer_name,
            changed,
        )

    st.cache_data.clear()
    return changed


def render_soluble_customer_editor(customer_name, current, key_scope):
    """既存の商品カードと同じ操作で、配達日・配達数量・使用数量/日だけ編集する。"""
    identity = f"{key_scope}|{customer_name}|soluble"
    key_suffix = hashlib.sha256(identity.encode("utf-8")).hexdigest()[:16]
    edit_key = f"soluble_customer_edit_{key_suffix}"

    if not st.session_state.get(edit_key):
        if st.button("編集", key=f"soluble_customer_edit_button_{key_suffix}"):
            st.session_state[edit_key] = True
            st.rerun()
        return

    with st.form(f"soluble_customer_edit_form_{key_suffix}"):
        st.caption(f"🎤 {VOICE_INPUT_HELP} 入力欄は毎回空白から始まります。")
        delivery_date_text = st.text_input(
            "配達日",
            value="",
            placeholder="例：2026年7月15日",
            help=VOICE_INPUT_HELP,
        )
        delivery_quantity_text = st.text_input(
            "配達数量",
            value="",
            placeholder="例：15000",
            help=VOICE_INPUT_HELP,
        )
        usage_text = st.text_input(
            "使用数量/日",
            value="",
            placeholder="例：1000",
            help=VOICE_INPUT_HELP,
        )
        st.caption("次回配達予定はExcelの数式で計算されるため、直接編集しません。")
        save_col, cancel_col = st.columns(2)
        with save_col:
            save = st.form_submit_button("保存", type="primary", use_container_width=True)
        with cancel_col:
            cancel = st.form_submit_button("キャンセル", use_container_width=True)

    if cancel:
        st.session_state.pop(edit_key, None)
        st.rerun()

    if save:
        try:
            updates = {}
            if str(delivery_date_text).strip():
                new_delivery_date = parse_optional_date(delivery_date_text)
                if not same_excel_value(new_delivery_date, current.get("配達日")):
                    updates["delivery_date"] = new_delivery_date
            if str(delivery_quantity_text).strip():
                new_delivery_quantity = parse_optional_nonnegative_number(
                    delivery_quantity_text,
                    integer=False,
                )
                if not same_soluble_value(
                    new_delivery_quantity,
                    current.get("配達数量"),
                ):
                    updates["delivery_quantity"] = new_delivery_quantity
            if str(usage_text).strip():
                new_usage = parse_optional_nonnegative_number(
                    usage_text,
                    integer=False,
                )
                if not same_soluble_value(new_usage, current.get("使用数量/日")):
                    updates["usage"] = new_usage

            if not updates:
                st.warning("変更された項目がありません。")
                return

            with st.spinner("元ファイルをバックアップして保存しています…"):
                changed = save_soluble_customer_changes(customer_name, updates)
            st.session_state.pop(edit_key, None)
            st.session_state["soluble_customer_save_success"] = {
                "customer_name": customer_name,
                "changed_count": len(changed),
            }
            st.rerun()
        except ValueError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"保存できませんでした：{exc}")


def render_soluble_customer_product_card(customer_name, current, key_scope):
    """三谷牧場・熊林牧場のソリュブル情報を既存商品カードと同じ形で表示する。"""
    with st.container(border=True):
        st.subheader("📦 ソリュブル")

        success = st.session_state.get("soluble_customer_save_success")
        if success and normalize_soluble_customer_name(
            success.get("customer_name")
        ) == normalize_soluble_customer_name(customer_name):
            st.success(f"保存しました（{success.get('changed_count', 0)}セル更新）。")
            st.session_state.pop("soluble_customer_save_success", None)

        col1, col2 = st.columns(2)
        with col1:
            st.caption("配達日")
            st.markdown(f"**{format_date(current.get('配達日'))}**")
            st.caption("配達数量")
            st.markdown(f"**{soluble_number_label(current.get('配達数量'))}**")
        with col2:
            st.caption("次回配達予定")
            st.markdown(f"**{format_date(current.get('次回配達予定'))}**")
            st.caption("使用数量/日")
            st.markdown(f"**{soluble_number_label(current.get('使用数量/日'))}**")

        render_soluble_customer_editor(customer_name, current, key_scope)


def show_soluble_inventory_page():
    st.markdown("---")
    st.header("🧪 ソリュブル在庫")
    show_back_home_button("soluble_back_home")
    st.caption("aoベンチャーグレイン配車表.xlsx の「ソリュブル」シートを表示します。")

    with st.spinner("ソリュブル在庫を読み込んでいます…"):
        content, source = load_soluble_workbook_content()
        rows = read_soluble_rows(content)
        customer_summaries = read_soluble_customer_summaries(content)
    if not rows and not customer_summaries:
        st.warning("ソリュブルシートに表示できるデータがありません。")
        return

    location = st.radio(
        "表示する会社",
        list(SOLUBLE_LOCATIONS.keys()) + list(SOLUBLE_CUSTOMER_NAMES),
        horizontal=True,
        key="soluble_location",
    )

    if location in SOLUBLE_CUSTOMER_NAMES:
        current = customer_summaries.get(location)
        if current is None:
            st.warning(f"{location}の行がソリュブルシートに見つかりません。")
            return
        st.caption(f"参照：{source}")
        render_soluble_customer_product_card(
            location,
            current,
            key_scope="soluble_inventory_page",
        )
        return

    if not rows:
        st.warning("ソリュブルシートに表示できる日付がありません。")
        return

    # 数値がまだ空の日も、ここから新しく入力できるように日付行はすべて表示対象にする。
    active_rows = list(rows)
    if not active_rows:
        st.info(f"{location}の表示データはありません。")
        return

    month_keys = sorted({(row["date"].year, row["date"].month) for row in active_rows})
    month_labels = [f"{year}年{month}月" for year, month in month_keys]
    # Streamlit CloudはUTCで動くため、日本時間の「今日」を使う。
    today = get_jst_now().date()
    today_key = (today.year, today.month)
    today_month_label = f"{today.year}年{today.month}月"
    default_month = month_keys.index(today_key) if today_key in month_keys else len(month_keys) - 1

    # 日付が変わった最初の表示だけ、表示月と開始日を今日へ戻す。
    # 同じ日のうちは、ユーザーが選んだ別の日付・表示月をそのまま維持する。
    month_widget_key = f"soluble_month_{location}"
    daily_reset_key = f"soluble_daily_default_{location}"
    if st.session_state.get(daily_reset_key) != today.isoformat():
        if today_month_label in month_labels:
            st.session_state[month_widget_key] = today_month_label
        st.session_state[daily_reset_key] = today.isoformat()

    selected_month_label = st.selectbox(
        "表示月",
        month_labels,
        index=default_month,
        key=month_widget_key,
    )
    selected_month_key = month_keys[month_labels.index(selected_month_label)]
    month_rows = [
        row for row in active_rows
        if (row["date"].year, row["date"].month) == selected_month_key
    ]

    day_options = [row["date"] for row in month_rows]
    default_day = day_options.index(today) if today in day_options else 0
    start_widget_key = f"soluble_start_{location}_{selected_month_label}"
    start_default_key = f"{start_widget_key}_default_{today.isoformat()}"
    if today in day_options and not st.session_state.get(start_default_key):
        st.session_state[start_widget_key] = today
        st.session_state[start_default_key] = True

    control_left, control_right = st.columns(2)
    with control_left:
        start_day = st.selectbox(
            "開始日",
            day_options,
            index=default_day,
            format_func=lambda day: f"{day.month}/{day.day}（{'月火水木金土日'[day.weekday()]}）",
            key=start_widget_key,
        )
    with control_right:
        period_widget_key = f"soluble_period_{location}"
        period_options = ["7日間", "14日間", "1か月"]
        period_default_key = f"{period_widget_key}_default_one_month_v2"
        if (
            not st.session_state.get(period_default_key)
            or st.session_state.get(period_widget_key) not in period_options
        ):
            st.session_state[period_widget_key] = "1か月"
            st.session_state[period_default_key] = True
        period = st.selectbox(
            "表示期間",
            period_options,
            index=2,
            key=period_widget_key,
        )
    manual_only = st.checkbox("黄色の手入力だけ表示", key=f"soluble_manual_only_{location}")

    if period == "1か月":
        next_month = 1 if start_day.month == 12 else start_day.month + 1
        next_year = start_day.year + 1 if start_day.month == 12 else start_day.year
        one_month_later = date(
            next_year,
            next_month,
            min(start_day.day, calendar.monthrange(next_year, next_month)[1]),
        )
        visible_rows = [
            row for row in active_rows
            if start_day <= row["date"] < one_month_later
        ]
    else:
        visible_rows = [
            row for row in month_rows
            if start_day <= row["date"] < start_day + timedelta(
                days=7 if period == "7日間" else 14
            )
        ]
    if manual_only:
        visible_rows = [
            row for row in visible_rows
            if any(row.get(f"{location}_{field}_manual") for field in ("usage", "delivery", "inventory"))
        ]

    st.markdown(
        """
        <style>
        .soluble-legend {display:flex; gap:.7rem; align-items:center; margin:.35rem 0 1rem; color:#596273;}
        .soluble-yellow-chip {display:inline-block; width:1.25rem; height:1.25rem; background:#fff59d; border:1px solid #e3cb42; border-radius:.3rem;}
        .soluble-card {background:rgba(255,255,255,.78); border:1px solid #cbd5e1; border-radius:16px; padding:14px 16px; margin:.65rem 0 .25rem; box-shadow:0 6px 16px rgba(30,41,59,.05);}
        .soluble-card-date {font-size:1.12rem; font-weight:800; margin-bottom:10px;}
        .soluble-values {display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:8px;}
        .soluble-value {background:#f8fafc; border-radius:12px; padding:9px 10px; min-width:0;}
        .soluble-value.manual {background:#fff59d; border:1px solid #e3cb42;}
        .soluble-value.negative {background:#fee2e2; border:1px solid #f87171;}
        .soluble-label {display:block; color:#697386; font-size:.78rem; margin-bottom:3px;}
        .soluble-number {display:block; color:#182033; font-size:1.04rem; font-weight:800; overflow-wrap:anywhere;}
        @media (max-width: 640px) {
          .soluble-card {padding:13px 12px; border-radius:14px;}
          .soluble-values {gap:6px;}
          .soluble-value {padding:9px 7px;}
          .soluble-label {font-size:.72rem;}
          .soluble-number {font-size:.96rem;}
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div class="soluble-legend"><span class="soluble-yellow-chip"></span><span>黄色は手入力　｜　参照：{html.escape(source)}</span></div>',
        unsafe_allow_html=True,
    )

    if not visible_rows:
        st.info("この条件で表示するデータはありません。")
        return

    weekday = "月火水木金土日"
    for row in visible_rows:
        usage = row.get(f"{location}_usage")
        delivery = row.get(f"{location}_delivery")
        inventory = row.get(f"{location}_inventory")
        cells = []
        for label, field, value in (
            ("使用量/日", "usage", usage),
            ("納品", "delivery", delivery),
            ("在庫", "inventory", inventory),
        ):
            classes = ["soluble-value"]
            if row.get(f"{location}_{field}_manual"):
                classes.append("manual")
            if field == "inventory" and isinstance(value, (int, float)) and value < 0:
                classes.append("negative")
            cells.append(
                f'<div class="{" ".join(classes)}"><span class="soluble-label">{label}</span>'
                f'<span class="soluble-number">{html.escape(soluble_number_label(value))}</span></div>'
            )
        day = row["date"]
        st.markdown(
            f'<section class="soluble-card"><div class="soluble-card-date">{day.month}/{day.day}（{weekday[day.weekday()]}）</div>'
            f'<div class="soluble-values">{"".join(cells)}</div></section>',
            unsafe_allow_html=True,
        )

        with st.expander(f"✏️ {day.month}/{day.day}を入力・修正"):
            form_key = f"soluble_form_{location}_{row['row']}"
            with st.form(form_key):
                usage_text = st.text_input(
                    "使用量/日",
                    value=soluble_input_value(usage),
                    key=f"{form_key}_usage",
                )
                delivery_text = st.text_input(
                    "納品",
                    value=soluble_input_value(delivery),
                    key=f"{form_key}_delivery",
                )
                current_formula = bool(row.get(f"{location}_inventory_formula")) and not bool(
                    row.get(f"{location}_inventory_manual")
                )
                auto_inventory = st.checkbox(
                    "在庫は「前日在庫 − 使用量 + 納品」で自動計算する",
                    value=current_formula,
                    key=f"{form_key}_auto",
                )
                inventory_text = st.text_input(
                    "在庫（自動計算を外した場合に使用）",
                    value=soluble_input_value(inventory),
                    key=f"{form_key}_inventory",
                )
                submitted = st.form_submit_button("バックアップして保存", use_container_width=True)

            if submitted:
                try:
                    new_usage = parse_soluble_number(usage_text, "使用量/日")
                    new_delivery = parse_soluble_number(delivery_text, "納品")
                    new_inventory = None if auto_inventory else parse_soluble_number(inventory_text, "在庫")
                    updates = {}
                    if not same_soluble_value(new_usage, usage):
                        updates["usage"] = new_usage
                    if not same_soluble_value(new_delivery, delivery):
                        updates["delivery"] = new_delivery
                    if auto_inventory:
                        if not current_formula:
                            updates["inventory"] = "__AUTO_INVENTORY__"
                    elif current_formula or not same_soluble_value(new_inventory, inventory):
                        updates["inventory"] = new_inventory
                    with st.spinner("元ファイルをバックアップして保存しています…"):
                        changed = save_soluble_changes(
                            row["row"],
                            location,
                            updates,
                        )
                    st.success(f"保存しました（{len(changed)}セル更新）。黄色は手入力値です。")
                    st.rerun()
                except Exception as error:
                    st.error(str(error))



# =========================
# ホームメニュー
# =========================
def show_home_menu():
    st.subheader("メニュー")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(render_page_link("👥 顧客名一覧", page="customer_list"), unsafe_allow_html=True)
    with col2:
        st.markdown(render_page_link("🔍 顧客検索", page="customer"), unsafe_allow_html=True)

    col3, col4 = st.columns(2)
    with col3:
        st.markdown(render_page_link("📍 地域検索", page="region"), unsafe_allow_html=True)
    with col4:
        st.markdown(render_page_link("🗓 配車カレンダー", page="calendar"), unsafe_allow_html=True)

    col5, col6 = st.columns(2)
    with col5:
        st.markdown(render_page_link("🚚 配車表", page="dispatch_table"), unsafe_allow_html=True)
    with col6:
        st.markdown(render_page_link("🧪 ソリュブル在庫", page="soluble_inventory"), unsafe_allow_html=True)

    col7, _ = st.columns(2)
    with col7:
        st.markdown(render_page_link("📝 メモ帳", page="notes"), unsafe_allow_html=True)

    st.markdown("---")

# =========================
# メイン
# =========================
if "page" not in st.session_state:
    st.session_state["page"] = "home"

if "selected_customer" not in st.session_state:
    st.session_state["selected_customer"] = None

# URLにpage/customerがある場合は、ブラウザの戻る・進むに合わせて画面を復元する。
handle_customer_query_param()


MENU_OPTIONS = {
    "👥 顧客名一覧": "customer_list",
    "🔍 顧客検索": "customer",
    "📍 地域検索": "region",
    "🗓 配車カレンダー": "calendar",
    "🚚 配車表": "dispatch_table",
    "🧪 ソリュブル在庫": "soluble_inventory",
    "📝 メモ帳": "notes",
}

current_page = st.session_state.get("page", "home")

with st.sidebar:
    st.title(f"🚚 {APP_TITLE}")
    st.markdown("### メニュー")
    st.markdown(render_page_link("👥 顧客名一覧", page="customer_list"), unsafe_allow_html=True)
    st.markdown(render_page_link("🔍 顧客検索", page="customer"), unsafe_allow_html=True)
    st.markdown(render_page_link("📍 地域検索", page="region"), unsafe_allow_html=True)
    st.markdown(render_page_link("🗓 配車カレンダー", page="calendar"), unsafe_allow_html=True)
    st.markdown(render_page_link("🚚 配車表", page="dispatch_table"), unsafe_allow_html=True)
    st.markdown(render_page_link("🧪 ソリュブル在庫", page="soluble_inventory"), unsafe_allow_html=True)
    st.markdown(render_page_link("📝 メモ帳", page="notes"), unsafe_allow_html=True)

    st.markdown("---")
    if st.button("🔄 更新", use_container_width=True):
        st.cache_data.clear()
        st.rerun()


col_title, col_logout = st.columns([3, 1])

with col_title:
    st.title(f"🚚 {APP_TITLE}")
    st.caption("顧客名一覧・顧客検索・地域検索・配車カレンダー・配車表・ソリュブル在庫・メモ帳")

with col_logout:
    st.write("")
    if st.button("ログアウト"):
        st.session_state.authenticated = False
        st.session_state.page = "home"
        st.session_state.selected_customer = None
        try:
            st.query_params.clear()
        except Exception:
            pass
        st.rerun()

try:
    if st.session_state["page"] == "home":
        show_home_menu()
        show_customer_search()

    elif st.session_state["page"] == "customer":
        show_customer_search(show_home_link=True)

    elif st.session_state["page"] == "customer_list":
        show_customer_directory()

    elif st.session_state["page"] == "region":
        df = load_data()
        show_region_search(df)

    elif st.session_state["page"] == "calendar":
        df = load_data()
        show_dispatch_calendar(df)

    elif st.session_state["page"] == "dispatch_table":
        show_dispatch_board()

    elif st.session_state["page"] == "soluble_inventory":
        show_soluble_inventory_page()

    elif st.session_state["page"] == "notes":
        show_notes_page(None)

    elif st.session_state["page"] == "detail":
        selected = st.session_state.get("selected_customer")
        if selected:
            df = load_data()
            show_customer_detail(df, selected)
        else:
            set_page("home")
            st.rerun()
except Exception as e:
    st.error("画面表示中にエラーが発生しました。")
    st.write("原因確認のため、エラー内容を表示しています。")
    st.exception(e)
    st.stop()

st.caption("※ 顧客情報はSheet1、配車表は配車表1.xlsm、ソリュブル在庫はaoベンチャーグレイン配車表.xlsxを読み込んで表示しています。")
