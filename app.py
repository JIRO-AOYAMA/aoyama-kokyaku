import json
import math
from datetime import date, timedelta
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
import streamlit as st


# =========================
# 基本設定
# =========================
APP_TITLE = "青山商店 顧客検索"
EXCEL_FILE = "配車予定 次郎.xlsm"
SHEET_NAME = "Sheet1"

# secrets.toml に入れる設定
DROPBOX_APP_KEY = st.secrets.get("DROPBOX_APP_KEY", "")
DROPBOX_APP_SECRET = st.secrets.get("DROPBOX_APP_SECRET", "")
DROPBOX_REFRESH_TOKEN = st.secrets.get("DROPBOX_REFRESH_TOKEN", "")
# 移行期間用。Streamlit CloudではRefresh Token方式の3項目を使う。
DROPBOX_ACCESS_TOKEN = st.secrets.get("DROPBOX_ACCESS_TOKEN", "")
DROPBOX_FILE_PATH = st.secrets.get("DROPBOX_FILE_PATH", "")
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
    st.caption("顧客検索アプリ")

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


def search_dropbox_file_by_name(filename, access_token):
    """
    DROPBOX_FILE_PATHが空、またはパス指定で失敗した時に、
    Dropbox内からファイル名で検索する。
    """
    url = "https://api.dropboxapi.com/2/files/search_v2"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    data = {
        "query": filename,
        "options": {
            "filename_only": True,
            "max_results": 20,
        },
    }

    try:
        response = requests.post(url, headers=headers, json=data, timeout=30)
    except Exception as e:
        st.error("Dropbox内のファイル検索に失敗しました。")
        st.exception(e)
        st.stop()

    if response.status_code != 200:
        st.error("Dropbox内のファイル検索に失敗しました。")
        st.write("Dropboxからの応答：")
        st.code(response.text)
        st.stop()

    result = response.json()
    matches = result.get("matches", [])

    if not matches:
        st.error(f"Dropbox内でファイルが見つかりません：{filename}")
        st.write("ファイル名が完全一致しているか確認してください。")
        st.stop()

    # 名前が完全一致するファイルを優先
    for match in matches:
        metadata = match.get("metadata", {}).get("metadata", {})
        file_id = metadata.get("id")
        name = metadata.get("name", "")
        path_display = metadata.get("path_display", "")

        if file_id and name == filename:
            return file_id, path_display

    # 完全一致が取れない場合は先頭候補
    metadata = matches[0].get("metadata", {}).get("metadata", {})
    return metadata.get("id"), metadata.get("path_display", "")


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


def read_excel_from_dropbox_api():
    """Dropbox APIでExcelをダウンロードして読み込む"""
    if not has_dropbox_auth_config():
        st.error("Dropbox API設定が不足しています。")
        st.write("secrets.toml に DROPBOX_APP_KEY / DROPBOX_APP_SECRET / DROPBOX_REFRESH_TOKEN を設定してください。")
        st.stop()

    access_token = get_dropbox_access_token()
    content = None
    response = None

    # 1. まずDROPBOX_FILE_PATHがあれば、それを試す
    if DROPBOX_FILE_PATH:
        content, response = download_dropbox_file(DROPBOX_FILE_PATH, access_token)

    # 2. パス指定で失敗したら、ファイル名で検索する
    if content is None:
        file_id, path_display = search_dropbox_file_by_name(EXCEL_FILE, access_token)
        content, response = download_dropbox_file(file_id, access_token)

        if content is None:
            st.error("Dropbox APIからExcelファイルをダウンロードできませんでした。")
            st.write("ファイル検索では見つかりましたが、ダウンロードに失敗しました。")
            st.write("見つかったパス：")
            st.code(path_display)
            st.write("Dropboxからの応答：")
            st.code(response.text)
            st.stop()

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
        score = sum(1 for col in REQUIRED_COLUMNS if col in values)

        # IDだけは上部表示にも出るので、顧客名・ひらがながある行を重視
        if "顧客名" in values and "ひらがな" in values and score >= 5:
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

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        st.error("必要な列が見つかりません。")
        st.write("見つからない列：", missing)
        st.write("Excelから読み取れた列：", list(df.columns))
        st.stop()

    df = df[REQUIRED_COLUMNS].copy()

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
# メイン
# =========================
if "page" not in st.session_state:
    st.session_state["page"] = "home"

if "selected_customer" not in st.session_state:
    st.session_state["selected_customer"] = None


col_title, col_logout = st.columns([3, 1])

with col_title:
    st.title("🚚 青山商店")
    st.caption("顧客検索アプリ")

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

elif st.session_state["page"] == "delivery":
    show_delivery_list(df)

elif st.session_state["page"] == "detail":
    selected = st.session_state.get("selected_customer")
    if selected:
        show_customer_detail(df, selected)
    else:
        set_page("home")
        st.rerun()

st.caption("※ このアプリはExcelのSheet1を読み込んで表示しています。Dropbox API設定がある場合はDropbox上のExcelを読み込みます。")
