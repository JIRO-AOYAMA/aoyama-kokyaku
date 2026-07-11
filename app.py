import calendar
import html
import json
import math
import urllib.parse
from datetime import date, datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
import streamlit as st


# =========================
# 基本設定
# =========================
APP_TITLE = "青山商店 業務アプリ"

# Streamlitでは、st.set_page_config は他の st.* 呼び出しより先に実行する
st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🚚",
    layout="wide",
)

EXCEL_FILE = "配車予定 次郎.xlsm"
SHEET_NAME = "Sheet1"

# secrets.toml に入れる設定
DROPBOX_APP_KEY = st.secrets.get("DROPBOX_APP_KEY", "")
DROPBOX_APP_SECRET = st.secrets.get("DROPBOX_APP_SECRET", "")
DROPBOX_REFRESH_TOKEN = st.secrets.get("DROPBOX_REFRESH_TOKEN", "")
# 移行期間用。Streamlit CloudではRefresh Token方式の3項目を使う。
DROPBOX_ACCESS_TOKEN = st.secrets.get("DROPBOX_ACCESS_TOKEN", "")
DROPBOX_DEFAULT_FILE_PATH = "/1共有　青山商店　本社/配車表-北海道-/配車予定 次郎.xlsm"
DROPBOX_FILE_PATH = st.secrets.get("DROPBOX_FILE_PATH", DROPBOX_DEFAULT_FILE_PATH)
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "")
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_SECRET_KEY = st.secrets.get("SUPABASE_SECRET_KEY", "")
SUPABASE_SERVICE_ROLE_KEY = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", "")
SUPABASE_ANON_KEY = st.secrets.get("SUPABASE_ANON_KEY", "")
SUPABASE_NOTES_TABLE = st.secrets.get("SUPABASE_NOTES_TABLE", "notes")

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
    st.title("🔒 青山商店")
    st.caption("業務アプリ")

    if not APP_PASSWORD:
        st.error("APP_PASSWORD が設定されていません。Streamlit Cloud の Secrets に APP_PASSWORD を追加してください。")
        st.stop()

    password = st.text_input("パスワード", type="password")

    if st.button("ログイン"):
        if password == APP_PASSWORD:
            st.session_state.authenticated = True
            st.session_state.page = "home"
            st.session_state.selected_customer = None
            try:
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

    div[data-testid="stVerticalBlockBorderWrapper"] {
        border: 1px solid var(--aoyama-line) !important;
        border-radius: 18px !important;
        background: var(--aoyama-card) !important;
        box-shadow: var(--aoyama-shadow);
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
    .stButton > button:hover {
        border-color: rgba(37, 99, 235, 0.35) !important;
        box-shadow: 0 10px 24px rgba(37, 99, 235, 0.12);
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


    @media (max-width: 640px) {
        .block-container {
            padding-left: 0.8rem;
            padding-right: 0.8rem;
            padding-top: 1.4rem;
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

    if -90 <= lat <= 90 and -180 <= lng <= 180:
        return lat, lng

    return None


def build_google_maps_url(value):
    """住所・緯度経度・URLからGoogleマップで開くURLを作る"""
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


def get_customer_map_info(detail):
    """顧客詳細で使う住所・地図情報を取り出す"""
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


def render_google_maps_link(url):
    safe_url = html.escape(str(url), quote=True)
    return f'<a class="app-nav-link" href="{safe_url}" target="_blank" rel="noopener noreferrer">📍 Googleマップ</a>'


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


def read_excel_from_dropbox_api():
    """Dropbox APIでExcelをダウンロードして読み込む"""
    if not has_dropbox_auth_config():
        st.error("Dropbox API設定が不足しています。")
        st.write("secrets.toml に DROPBOX_APP_KEY / DROPBOX_APP_SECRET / DROPBOX_REFRESH_TOKEN を設定してください。")
        st.stop()

    access_token = get_dropbox_access_token()
    dropbox_file_path = get_dropbox_file_path()
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
        "Authorization": f"Bearer {key}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    if prefer:
        headers["Prefer"] = prefer
    return headers


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


def load_notes_from_supabase(customer_name=None, limit=500):
    """Supabaseのnotesテーブルからメモを新しい順で読み込む"""
    if not has_supabase_config():
        show_supabase_config_error()

    params = {
        "select": "id,customer_name,body,created_at",
        "order": "created_at.desc",
        "limit": str(limit),
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
    st.caption("スマホではキーボードのマイクから音声入力できます。")

    note_key = f"customer_note_input_{customer_name}"
    note_text = st.text_area(
        "メモ本文",
        key=note_key,
        height=120,
        placeholder="例：次回は午前中希望。サンプル持参。など",
    )

    if st.button("メモを保存", key=f"save_customer_note_{customer_name}"):
        if add_note(customer_name, note_text):
            st.session_state[note_key] = ""
            st.success("メモを保存しました。")
            st.rerun()

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
# Excel読み込み・整形
# =========================
def normalize_excel_table(excel_source):
    """
    ExcelのSheet1から、顧客一覧表を取り出す。

    対応できる形：
    1) 1行目が見出し
       ID / 顧客名 / 地域 / 商品名 / 使用数量/日 / 次回配達予定 / 残数 / ひらがな

    2) 上部に大きな表示があり、途中の行に見出しがある形
       9行目などに ID / 顧客名 / 地域 / 商品名 ... がある
    """
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
        st.error("必要な見出し行が見つかりません。")
        st.write("必要な列：", REQUIRED_COLUMNS)
        st.write("読み取った先頭20行：")
        st.dataframe(raw.head(20))
        st.stop()

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


@st.cache_data(ttl=60)
def load_data():
    """
    Dropbox API設定があればDropbox上のExcelを読む。
    設定がなければ同じフォルダのローカルExcelを読む。
    """
    if has_dropbox_auth_config():
        excel_source = read_excel_from_dropbox_api()
    else:
        excel_source = read_excel_local()

    return normalize_excel_table(excel_source)


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

    valid_pages = {"home", "region", "calendar", "delivery", "notes", "detail"}

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


# =========================
# 顧客詳細
# =========================
def show_customer_detail(df, customer_name):
    detail = df[df["顧客名"] == customer_name].copy()

    if detail.empty:
        st.warning("選択した顧客の情報が見つかりません。")
        return

    show_back_home_button("detail_back_home")

    # 使用数量/日が0・空白・NaNの商品行は、商品名ごと表示しない。
    visible_detail = detail[~detail["使用数量/日"].apply(is_blank_or_zero)].copy()

    region = clean_value(detail.iloc[0]["地域"])

    st.markdown("---")
    st.header(f"👤 {customer_name}")
    st.write(f"**地域：** {region}")
    st.write(f"**商品数：** {len(visible_detail)}件")

    map_info = get_customer_map_info(detail)
    if map_info:
        st.write(f"**{map_info['display_label']}：** {map_info['display_value']}")
        if map_info["map_url"]:
            st.markdown(render_google_maps_link(map_info["map_url"]), unsafe_allow_html=True)

    if visible_detail.empty:
        st.info("表示対象の商品はありません。使用数量/日が0または空白の商品は非表示にしています。")
        return

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


    show_customer_notes(customer_name)

# =========================
# 顧客検索
# =========================
def show_customer_search(df):
    st.subheader("🔍 顧客検索")

    default_keyword = str(get_query_value("customer_search", "")).strip()
    keyword = st.text_input(
        "ひらがなで検索",
        value=default_keyword,
        placeholder="例：こ、こも、むら",
        key="customer_search_input",
    ).strip()

    if keyword:
        update_query_params(page="home", customer_search=keyword)
    else:
        update_query_params(page="home", customer_search=None)

    if not keyword:
        st.info("顧客名のひらがなを入力してください。")
        return

    hit = df[df["ひらがな"].str.startswith(keyword, na=False)]

    if hit.empty:
        st.warning("該当する顧客がありません。")
        return

    customers = hit[["顧客名", "地域"]].drop_duplicates().reset_index(drop=True)

    st.write(f"候補：{len(customers)}件")

    for i, row in customers.iterrows():
        name = clean_value(row["顧客名"])
        region = clean_value(row["地域"])

        with st.container(border=True):
            st.markdown(f"### 👤 {name}")
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

    default_keyword = str(get_query_value("region_search", "")).strip()
    keyword = st.text_input(
        "地域名で検索",
        value=default_keyword,
        placeholder="例：帯広、芽室、釧路",
        key="region_search_input",
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
# 配達予定一覧
# =========================
def make_delivery_list(df):
    today = date.today()
    start_day = today - timedelta(days=7)
    end_day = today + timedelta(days=31)

    rows = []

    for _, row in df.iterrows():
        next_delivery = to_date(row["次回配達予定"])
        if next_delivery is None:
            continue

        if start_day <= next_delivery <= end_day:
            rows.append(
                {
                    "顧客名": clean_value(row["顧客名"]),
                    "地域": clean_value(row["地域"]),
                    "次回配達予定日": next_delivery,
                    "次回配達予定表示": format_date(row["次回配達予定"]),
                }
            )

    if not rows:
        return pd.DataFrame(columns=["顧客名", "地域", "次回配達予定日", "次回配達予定表示"])

    delivery_df = pd.DataFrame(rows)

    # 同じ顧客が複数商品を持つ場合は、一番古い次回配達予定だけを一覧に表示
    delivery_df = (
        delivery_df.sort_values(["次回配達予定日", "顧客名"])
        .drop_duplicates(subset=["顧客名"], keep="first")
        .reset_index(drop=True)
    )

    return delivery_df


def show_delivery_list(df):
    show_back_home_button("delivery_back_home")

    st.markdown("---")
    st.header("📅 配達予定一覧")
    st.caption("過去7日 ～ 1か月後")

    delivery_df = make_delivery_list(df)

    if delivery_df.empty:
        st.info("対象期間内の配達予定はありません。")
        return

    for i, row in delivery_df.iterrows():
        customer_name = row["顧客名"]
        region = clean_value(row["地域"])
        next_date = row["次回配達予定表示"]

        with st.container(border=True):
            col1, col2 = st.columns([3, 2])

            with col1:
                st.markdown(f"### 👤 {customer_name}")
                st.caption(f"地域：{region}")

            with col2:
                st.markdown(f"**{next_date}**")

            st.markdown(
                render_page_link("詳細を見る", page="detail", customer=customer_name),
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
        DISPATCH_REQUIRED_LABELS[key]
        for key, col in dispatch_columns.items()
        if not col
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
    parsed_dates = pd.to_datetime(df[date_column], errors="coerce").dt.date

    for idx, row in df.iterrows():
        delivery_date = parsed_dates.loc[idx]

        if pd.isna(delivery_date) or not (start_day <= delivery_date <= end_day):
            continue

        item = {
            "顧客名": clean_value(row[customer_column]),
            "地域": clean_value(row[region_column]),
            "商品名": clean_value(row[product_column]),
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
            parts.append('</div>')

    parts.append('</div>')
    return "".join(parts)


def show_dispatch_month_switcher(month_start):
    col_prev, col_month, col_next = st.columns([1, 2, 1])

    with col_prev:
        if st.button("◀", key="dispatch_prev_month", use_container_width=True):
            change_dispatch_month(-1)
            st.rerun()

    with col_month:
        st.markdown(
            f'<div class="dispatch-month-title">{month_start.year}年{month_start.month}月</div>',
            unsafe_allow_html=True,
        )

    with col_next:
        if st.button("▶", key="dispatch_next_month", use_container_width=True):
            change_dispatch_month(1)
            st.rerun()


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
    product_name = clean_value(item.get("商品名"))
    customer_link = build_customer_detail_link(customer_name, class_name="dispatch-month-link")

    if product_name == "未設定":
        return customer_link

    return f'{customer_link}<br><span class="dispatch-month-product">{escape_html(product_name)}</span>'


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
    )

    if view == "📱 2日表示":
        show_two_day_dispatch_calendar(rows_by_day, month_start)
    else:
        show_month_dispatch_calendar(rows_by_day, month_start)



# =========================
# ホームメニュー
# =========================
def show_home_menu():
    st.subheader("メニュー")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(render_page_link("🔍 顧客検索", page="home"), unsafe_allow_html=True)
    with col2:
        st.markdown(render_page_link("📍 地域検索", page="region"), unsafe_allow_html=True)

    col3, col4 = st.columns(2)
    with col3:
        st.markdown(render_page_link("🗓 配車カレンダー", page="calendar"), unsafe_allow_html=True)
    with col4:
        st.markdown(render_page_link("📅 配達予定一覧", page="delivery"), unsafe_allow_html=True)

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
    "🔍 顧客検索": "home",
    "📍 地域検索": "region",
    "🗓 配車カレンダー": "calendar",
    "📅 配達予定一覧": "delivery",
    "📝 メモ帳": "notes",
}

current_page = st.session_state.get("page", "home")

with st.sidebar:
    st.title("🚚 青山商店")
    st.caption("業務アプリ")
    st.markdown("### メニュー")
    st.markdown(render_page_link("🔍 顧客検索", page="home"), unsafe_allow_html=True)
    st.markdown(render_page_link("📍 地域検索", page="region"), unsafe_allow_html=True)
    st.markdown(render_page_link("🗓 配車カレンダー", page="calendar"), unsafe_allow_html=True)
    st.markdown(render_page_link("📅 配達予定一覧", page="delivery"), unsafe_allow_html=True)
    st.markdown(render_page_link("📝 メモ帳", page="notes"), unsafe_allow_html=True)

    st.markdown("---")
    if st.button("🔄 データ更新", use_container_width=True):
        st.cache_data.clear()
        st.rerun()


col_title, col_logout = st.columns([3, 1])

with col_title:
    st.title(f"🚚 {APP_TITLE}")
    st.caption("顧客検索・地域検索・配車カレンダー・配達予定・メモ帳")

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
    df = load_data()
except Exception as e:
    st.error("ログイン後のデータ読み込み中にエラーが発生しました。")
    st.write("白画面になる原因を確認するため、エラー内容を表示しています。")
    st.exception(e)
    st.stop()

if st.session_state["page"] == "home":
    show_home_menu()
    show_customer_search(df)

elif st.session_state["page"] == "region":
    show_region_search(df)

elif st.session_state["page"] == "calendar":
    show_dispatch_calendar(df)

elif st.session_state["page"] == "delivery":
    show_delivery_list(df)

elif st.session_state["page"] == "notes":
    show_notes_page(df)

elif st.session_state["page"] == "detail":
    selected = st.session_state.get("selected_customer")
    if selected:
        show_customer_detail(df, selected)
    else:
        set_page("home")
        st.rerun()

st.caption("※ このアプリはExcelのSheet1を読み込んで表示しています。Dropbox API設定がある場合はDropbox上のExcelを読み込みます。")
