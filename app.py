import json
import math
from datetime import date, timedelta
import calendar
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

    password = st.text_input("パスワード", type="password")

    if st.button("ログイン"):
        if APP_PASSWORD and password == APP_PASSWORD:
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


def search_dropbox_file_by_name(filename):
    """
    DROPBOX_FILE_PATHが空、またはパス指定で失敗した時に、
    Dropbox内からファイル名で検索する。
    """
    url = "https://api.dropboxapi.com/2/files/search_v2"

    headers = {
        "Authorization": f"Bearer {DROPBOX_ACCESS_TOKEN}",
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


def download_dropbox_file(path_or_id):
    """Dropbox APIでExcelをダウンロードする"""
    url = "https://content.dropboxapi.com/2/files/download"

    headers = {
        "Authorization": f"Bearer {DROPBOX_ACCESS_TOKEN}",
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
    if not DROPBOX_ACCESS_TOKEN:
        st.error("Dropbox API設定が不足しています。")
        st.write("secrets.toml に DROPBOX_ACCESS_TOKEN を設定してください。")
        st.stop()

    content = None
    response = None

    # 1. まずDROPBOX_FILE_PATHがあれば、それを試す
    if DROPBOX_FILE_PATH:
        content, response = download_dropbox_file(DROPBOX_FILE_PATH)

    # 2. パス指定で失敗したら、ファイル名で検索する
    if content is None:
        file_id, path_display = search_dropbox_file_by_name(EXCEL_FILE)
        content, response = download_dropbox_file(file_id)

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
    if DROPBOX_ACCESS_TOKEN:
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
def get_month_start():
    today = date.today()
    if "calendar_year" not in st.session_state:
        st.session_state["calendar_year"] = today.year
    if "calendar_month" not in st.session_state:
        st.session_state["calendar_month"] = today.month

    return date(st.session_state["calendar_year"], st.session_state["calendar_month"], 1)


def change_month(delta):
    current = get_month_start()
    month = current.month + delta
    year = current.year

    if month < 1:
        month = 12
        year -= 1
    elif month > 12:
        month = 1
        year += 1

    st.session_state["calendar_year"] = year
    st.session_state["calendar_month"] = month


def make_calendar_rows(df, month_start):
    last_day = calendar.monthrange(month_start.year, month_start.month)[1]
    start_day = date(month_start.year, month_start.month, 1)
    end_day = date(month_start.year, month_start.month, last_day)

    rows_by_day = {}

    for day_num in range(1, last_day + 1):
        target_day = date(month_start.year, month_start.month, day_num)
        rows_by_day[target_day] = []

    for _, row in df.iterrows():
        next_delivery = to_date(row["次回配達予定"])
        if next_delivery is None:
            continue

        if start_day <= next_delivery <= end_day:
            customer_name = clean_value(row["顧客名"])
            region = clean_value(row["地域"])
            product_name = clean_value(row["商品名"])

            rows_by_day.setdefault(next_delivery, []).append(
                {
                    "顧客名": customer_name,
                    "地域": region,
                    "商品名": product_name,
                }
            )

    for target_day in rows_by_day:
        rows_by_day[target_day] = sorted(rows_by_day[target_day], key=lambda x: x["顧客名"])

    return rows_by_day


def show_calendar_header():
    month_start = get_month_start()

    if st.button("← ホームへ戻る"):
        set_page("home")
        st.rerun()

    st.markdown("---")
    st.header("🗓 配車カレンダー")

    col_prev, col_month, col_next = st.columns([1, 2, 1])

    with col_prev:
        if st.button("◀ 前月"):
            change_month(-1)
            st.rerun()

    with col_month:
        st.markdown(f"### {month_start.year}年{month_start.month}月")

    with col_next:
        if st.button("翌月 ▶"):
            change_month(1)
            st.rerun()

    return month_start


def show_two_day_calendar(df, month_start):
    rows_by_day = make_calendar_rows(df, month_start)
    last_day = calendar.monthrange(month_start.year, month_start.month)[1]

    st.subheader("📱 2日表示")

    for day_num in range(1, last_day + 1, 2):
        day1 = date(month_start.year, month_start.month, day_num)
        day2 = date(month_start.year, month_start.month, day_num + 1) if day_num + 1 <= last_day else None

        col1, col2 = st.columns(2)

        for col, target_day in [(col1, day1), (col2, day2)]:
            with col:
                if target_day is None:
                    st.write("")
                    continue

                weekday = ["月", "火", "水", "木", "金", "土", "日"][target_day.weekday()]
                st.markdown(f"#### {target_day.month}/{target_day.day}（{weekday}）")

                items = rows_by_day.get(target_day, [])

                if not items:
                    st.caption("予定なし")
                else:
                    for i, item in enumerate(items):
                        name = item["顧客名"]
                        region = item["地域"]
                        product_name = item["商品名"]

                        with st.container(border=True):
                            st.markdown(f"**👤 {name}**")
                            st.caption(f"地域：{region}")
                            st.caption(f"商品：{product_name}")

                            if st.button("詳細", key=f"cal_2day_{target_day}_{i}_{name}"):
                                select_customer(name)
                                st.rerun()

        st.markdown("---")


def show_month_calendar(df, month_start):
    rows_by_day = make_calendar_rows(df, month_start)
    last_day = calendar.monthrange(month_start.year, month_start.month)[1]

    st.subheader("🗓 月表示")
    st.caption("横スクロールできます。")

    max_count = 0
    for items in rows_by_day.values():
        max_count = max(max_count, len(items))

    if max_count == 0:
        max_count = 1

    table_rows = []

    for day_num in range(1, last_day + 1):
        target_day = date(month_start.year, month_start.month, day_num)
        weekday = ["月", "火", "水", "木", "金", "土", "日"][target_day.weekday()]
        items = rows_by_day.get(target_day, [])

        row_data = {
            "月/日": f"{target_day.month}/{target_day.day}（{weekday}）"
        }

        for i in range(max_count):
            col_name = f"牧場名{i + 1}"
            if i < len(items):
                row_data[col_name] = items[i]["顧客名"]
            else:
                row_data[col_name] = ""

        table_rows.append(row_data)

    month_df = pd.DataFrame(table_rows)
    st.dataframe(month_df, use_container_width=True, hide_index=True)


def show_dispatch_calendar(df):
    month_start = show_calendar_header()

    view = st.radio(
        "表示切替",
        ["📱 2日表示", "🗓 月表示"],
        horizontal=True,
    )

    if view == "📱 2日表示":
        show_two_day_calendar(df, month_start)
    else:
        show_month_calendar(df, month_start)


# =========================
# メイン
# =========================
if "page" not in st.session_state:
    st.session_state["page"] = "home"

if "selected_customer" not in st.session_state:
    st.session_state["selected_customer"] = None


st.title("🚚 青山商店")
st.caption("顧客検索アプリ")

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
    st.caption("2日表示 / 月表示")

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
