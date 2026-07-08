import calendar
import json
import math
from datetime import date, timedelta
from html import escape
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
import streamlit as st


# =========================
# 基本設定
# =========================
APP_TITLE = "青山商店 業務アプリ"
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


# =========================
# Streamlit設定
# =========================
st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🚚",
    layout="centered",
)

# =========================
# ログイン認証
# =========================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

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
            st.rerun()
        else:
            st.error("パスワードが違います。")

    st.stop()


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


@st.cache_data(ttl=300)
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
def set_page(page_name):
    st.session_state["page"] = page_name


def select_customer(customer_name, page_name="detail"):
    st.session_state["selected_customer"] = customer_name
    st.session_state["page"] = page_name


# =========================
# 顧客詳細
# =========================
def show_customer_detail(df, customer_name):
    detail = df[df["顧客名"] == customer_name].copy()

    if detail.empty:
        st.warning("選択した顧客の情報が見つかりません。")
        return

    if st.button("← ホームへ戻る"):
        set_page("home")
        st.rerun()

    region = clean_value(detail.iloc[0]["地域"])

    st.markdown("---")
    st.header(f"👤 {customer_name}")
    st.write(f"**地域：** {region}")
    st.write(f"**商品数：** {len(detail)}件")

    for _, row in detail.iterrows():
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


# =========================
# 顧客検索
# =========================
def show_customer_search(df):
    st.subheader("🔍 顧客検索")

    keyword = st.text_input(
        "ひらがなで検索",
        placeholder="例：こ、こも、むら",
    ).strip()

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

            if st.button("この顧客を見る", key=f"search_select_{i}_{name}"):
                select_customer(name)
                st.rerun()


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
    if st.button("← ホームへ戻る"):
        set_page("home")
        st.rerun()

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

            if st.button("詳細を見る", key=f"delivery_select_{i}_{customer_name}"):
                select_customer(customer_name)
                st.rerun()


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
            border-radius: 8px;
            padding: 0.75rem;
            background: #ffffff;
            min-width: 0;
        }
        .dispatch-day-title {
            font-weight: 700;
            margin-bottom: 0.5rem;
        }
        .dispatch-item {
            border-top: 1px solid rgba(49, 51, 63, 0.12);
            padding: 0.55rem 0;
        }
        .dispatch-item:first-of-type {
            border-top: 0;
            padding-top: 0;
        }
        .dispatch-name {
            font-weight: 700;
            line-height: 1.35;
            word-break: break-word;
        }
        .dispatch-line,
        .dispatch-empty {
            color: rgba(49, 51, 63, 0.72);
            font-size: 0.9rem;
            line-height: 1.45;
            word-break: break-word;
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
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_two_day_panel(target_day, items):
    html_parts = [
        '<div class="dispatch-day-panel">',
        f'<div class="dispatch-day-title">{escape(format_month_day(target_day))}</div>',
    ]

    if not items:
        html_parts.append('<div class="dispatch-empty">予定なし</div>')
    else:
        for item in items:
            html_parts.extend(
                [
                    '<div class="dispatch-item">',
                    f'<div class="dispatch-name">👤 {escape(item["顧客名"])}</div>',
                    f'<div class="dispatch-line">地域：{escape(item["地域"])}</div>',
                    f'<div class="dispatch-line">商品：{escape(item["商品名"])}</div>',
                    '</div>',
                ]
            )

    html_parts.append("</div>")
    return "\n".join(html_parts)


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

    last_day = calendar.monthrange(month_start.year, month_start.month)[1]

    for day_num in range(1, last_day + 1, 2):
        day1 = date(month_start.year, month_start.month, day_num)
        day2 = date(month_start.year, month_start.month, day_num + 1) if day_num + 1 <= last_day else None

        left_panel = render_two_day_panel(day1, rows_by_day.get(day1, []))
        if day2:
            right_panel = render_two_day_panel(day2, rows_by_day.get(day2, []))
        else:
            right_panel = '<div class="dispatch-day-panel"><div class="dispatch-empty">&nbsp;</div></div>'

        st.markdown(
            f'<div class="dispatch-two-day-row">{left_panel}{right_panel}</div>',
            unsafe_allow_html=True,
        )


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
            row_data[column_name] = items[item_index]["顧客名"] if item_index < len(items) else ""

        table_rows.append(row_data)

    return pd.DataFrame(table_rows)


def show_month_dispatch_calendar(rows_by_day, month_start):
    st.subheader("🗓 月表示")
    st.caption("横スクロールで1か月分を確認できます。")

    month_df = make_month_dispatch_table(rows_by_day, month_start)
    column_config = {
        "月/日": st.column_config.TextColumn("月/日", width="small"),
    }

    for column_name in month_df.columns:
        if column_name != "月/日":
            column_config[column_name] = st.column_config.TextColumn(column_name, width="medium")

    st.dataframe(
        month_df,
        hide_index=True,
        use_container_width=True,
        height=min(760, 38 * (len(month_df) + 1)),
        column_config=column_config,
    )


def show_dispatch_calendar(df):
    st.markdown("---")
    st.header("🗓 配車カレンダー")
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
# メイン
# =========================
if "page" not in st.session_state:
    st.session_state["page"] = "home"

if "selected_customer" not in st.session_state:
    st.session_state["selected_customer"] = None


MENU_OPTIONS = {
    "🔍 顧客検索": "home",
    "📅 配達予定一覧": "delivery",
    "🗓 配車カレンダー": "calendar",
}

current_page = st.session_state.get("page", "home")
menu_pages = list(MENU_OPTIONS.values())
menu_labels = list(MENU_OPTIONS.keys())

if current_page in menu_pages:
    current_menu_index = menu_pages.index(current_page)
else:
    current_menu_index = 0

with st.sidebar:
    st.title("🚚 青山商店")
    st.caption("業務アプリ")

    selected_menu = st.radio(
        "メニュー",
        menu_labels,
        index=current_menu_index,
    )

selected_page = MENU_OPTIONS[selected_menu]

if st.session_state["page"] == "detail":
    if selected_page != "home":
        st.session_state["page"] = selected_page
        st.session_state["selected_customer"] = None
        st.rerun()
else:
    st.session_state["page"] = selected_page


col_title, col_logout = st.columns([3, 1])

with col_title:
    st.title(f"🚚 {APP_TITLE}")
    st.caption("顧客検索・配達予定・配車カレンダー")

with col_logout:
    st.write("")
    if st.button("ログアウト"):
        st.session_state.authenticated = False
        st.session_state.page = "home"
        st.session_state.selected_customer = None
        st.rerun()

df = load_data()

if st.session_state["page"] == "home":
    show_customer_search(df)

    st.markdown("---")
    st.subheader("📅 配達予定一覧")
    st.caption("過去7日 ～ 1か月後")

    if st.button("配達予定一覧を見る"):
        set_page("delivery")
        st.rerun()

    st.markdown("---")
    st.subheader("🗓 配車カレンダー")
    st.caption("2日表示 / 月表示で配車予定を確認できます。")

    if st.button("配車カレンダーを見る"):
        set_page("calendar")
        st.rerun()

elif st.session_state["page"] == "delivery":
    show_delivery_list(df)

elif st.session_state["page"] == "calendar":
    show_dispatch_calendar(df)

elif st.session_state["page"] == "detail":
    selected = st.session_state.get("selected_customer")
    if selected:
        show_customer_detail(df, selected)
    else:
        set_page("home")
        st.rerun()

st.caption("※ このアプリはExcelのSheet1を読み込んで表示しています。Dropbox API設定がある場合はDropbox上のExcelを読み込みます。")
