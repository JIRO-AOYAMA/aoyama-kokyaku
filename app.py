import base64
import calendar
import copy
import gzip
import hashlib
import html
import json
import math
import mimetypes
import posixpath
import re
import threading
import time
import urllib.parse
import unicodedata
import uuid
import zipfile
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

try:
    from st_keyup import st_keyup
except ImportError:
    st_keyup = None


# =========================
# 基本設定
# =========================
APP_TITLE = "取引先カルテ"

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
FULL_DATA_BACKUP_DROPBOX_FOLDER = st.secrets.get(
    "FULL_DATA_BACKUP_DROPBOX_FOLDER",
    str(DROPBOX_FILE_PATH).rsplit("/", 1)[0] + "/顧客カルテ全データ_Backups",
)
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

# 仕入先・運送会社の基本情報を保存する別ブック。
# 配車予定 次郎.xlsm と同じDropboxフォルダに置く。
TRADE_PARTNER_FILE_NAME = "取引先カルテ.xlsx"
TRADE_PARTNER_DROPBOX_DEFAULT_FILE_PATH = (
    str(DROPBOX_DEFAULT_FILE_PATH).rsplit("/", 1)[0] + "/" + TRADE_PARTNER_FILE_NAME
)
TRADE_PARTNER_DROPBOX_FILE_PATH = st.secrets.get(
    "TRADE_PARTNER_DROPBOX_FILE_PATH",
    str(DROPBOX_FILE_PATH).rsplit("/", 1)[0] + "/" + TRADE_PARTNER_FILE_NAME,
)
TRADE_PARTNER_BACKUP_FOLDER = (
    str(TRADE_PARTNER_DROPBOX_FILE_PATH).rsplit("/", 1)[0] + "/取引先カルテ_Backups"
)
TRADE_PARTNER_MASTER_SHEET = "取引先マスター"
TRADE_PARTNER_CONTACT_SHEET = "担当者"
TRADE_PARTNER_PRODUCT_SHEET = "仕入商品"
TRADE_PARTNER_TRANSPORT_SHEET = "運送条件"
SOLUBLE_LOCATIONS = {
    "ノベルズ": {"usage": 3, "delivery": 4, "inventory": 5},
    "コスモアグリ": {"usage": 6, "delivery": 7, "inventory": 8},
}
# Excel内の識別名・既存ロジックは変えず、画面表示だけを分かりやすい名称にする。
SOLUBLE_LOCATION_DISPLAY_NAMES = {
    "ノベルズ": "ノベルズデイリー",
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
ESTIMATE_PREFIX = "__estimate__:"
ESTIMATE_VERSION = 1
CARRIER_FREIGHT_PREFIX = "__carrier_freight__:"
CARRIER_FREIGHT_VERSION = 1
CHANGE_HISTORY_CUSTOMER = "__CHANGE_HISTORY__"
CHANGE_HISTORY_VERSION = 1
CHANGE_HISTORY_PAGE_SIZE = 30
CHANGE_HISTORY_TARGETS = ("顧客", "仕入先", "運送会社")

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
# WATER it接続（読み取り専用）
# =========================
# スマホでWATER itから手動ダウンロードしたCSVを選び、読み取り専用で表示する。
# 選択したCSVはStreamlitのセッション内だけに保持し、Excel・WATER it・Dropboxへは書き込まない。
# 未選択時は、従来どおり同じフォルダのdata.csvを参考表示できる。
WATER_IT_CSV_PATH = st.secrets.get("WATER_IT_CSV_PATH", "data.csv")
WATER_IT_CSV_URL = st.secrets.get("WATER_IT_CSV_URL", "")
WATER_IT_REQUEST_TIMEOUT = 20
WATER_IT_LOGIN_URL = "https://www.dms2.waterit.optex.net/WIA0101/Index01"
WATER_IT_UPLOAD_BYTES_KEY = "water_it_uploaded_csv_bytes"
WATER_IT_UPLOAD_NAME_KEY = "water_it_uploaded_csv_name"
WATER_IT_UPLOAD_HASH_KEY = "water_it_uploaded_csv_hash"
WATER_IT_UPLOAD_PERSISTED_KEY = "water_it_uploaded_csv_persisted"
# 既存のcustomer_informationテーブル内に、通常の顧客情報と衝突しない内部レコードとして保存する。
# 新しいSupabaseテーブルやSecretsは不要。Excel・WATER it・Dropboxには書き込まない。
WATER_IT_STORAGE_CUSTOMER = "__WATER_IT_STORAGE__"
WATER_IT_STORAGE_FIELD = "__water_it_csv_snapshot__"
WATER_IT_STORAGE_ID = str(
    uuid.uuid5(uuid.NAMESPACE_URL, "aoyama-water-it-csv-snapshot-v1")
)
WATER_IT_STORAGE_VERSION = 1
WATER_IT_REQUIRED_COLUMNS = [
    "測定日時",
    "測定項目",
    "測定値",
    "単位",
    "エリア",
    "ポイント",
]
WATER_IT_ALERT_COLUMNS = [
    "HOLD中",
    "メンテナンス時期",
    "校正時期",
    "消耗品交換時期",
    "オーバーホール時期",
    "ローバッテリ",
    "センサまたは変換器異常",
    "通信不良または断線",
    "状態",
]
# WATER it側の元名称は変更せず、画面表示と顧客照合だけを統一する。
WATER_IT_POINT_DISPLAY_NAMES = {
    "ノベルズデイリーファーム": "ノベルズデイリー",
}

# ソリュブル在庫とWATER itの対応。Excel内の識別名や既存列構成は変更しない。
SOLUBLE_WATER_IT_POINT_NAMES = {
    "ノベルズ": "ノベルズデイリー",
    "コスモアグリ": "コスモアグリ",
}
# 実績平均はアプリ上の参考表示だけに使い、Excelの使用量/日へは反映しない。
SOLUBLE_WATER_IT_USAGE_WINDOWS = (3, 7, 20, 30)


# =========================
# OneDrive接続（写真・資料）
# =========================
ONEDRIVE_AUTHORITY = "https://login.microsoftonline.com/consumers"
ONEDRIVE_AUTHORIZE_URL = ONEDRIVE_AUTHORITY + "/oauth2/v2.0/authorize"
ONEDRIVE_TOKEN_URL = ONEDRIVE_AUTHORITY + "/oauth2/v2.0/token"
ONEDRIVE_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
ONEDRIVE_SCOPES = "openid profile offline_access User.Read Files.ReadWrite"
ONEDRIVE_REQUEST_TIMEOUT = 90
ONEDRIVE_AUTH_FLOW_TTL_SECONDS = 15 * 60
ONEDRIVE_PRODUCTION_REDIRECT_URI = "https://aoyama-kokyaku.streamlit.app"
ONEDRIVE_TEST_REDIRECT_HOST = "aoyama-onedrive-test.streamlit.app"
ONEDRIVE_ROOT_FOLDER = "取引先カルテ"
ONEDRIVE_CUSTOMER_FOLDER = "顧客"
ONEDRIVE_ATTACHMENT_PREFIX = "__onedrive_attachment__:"
ONEDRIVE_ATTACHMENT_VERSION = 1
ONEDRIVE_FIXED_TAGS = ("設備", "名刺", "納品場所", "商品", "トラブル")
ONEDRIVE_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}
ONEDRIVE_PDF_EXTENSIONS = {".pdf"}
ONEDRIVE_PAGE_SIZE = 12


def read_onedrive_settings():
    """Streamlit SecretsからOneDrive接続設定を読む。"""
    try:
        settings = st.secrets["onedrive"]
        client_id = str(settings.get("client_id", "")).strip()
        client_secret = str(settings.get("client_secret", "")).strip()
        redirect_uri = str(settings.get("redirect_uri", "")).strip()
    except Exception:
        client_id = ""
        client_secret = ""
        redirect_uri = ""

    # テストアプリのURLが残っていても、本番カルテへ戻るよう安全側で補正する。
    if not redirect_uri or ONEDRIVE_TEST_REDIRECT_HOST in redirect_uri:
        redirect_uri = ONEDRIVE_PRODUCTION_REDIRECT_URI

    missing = []
    if not client_id:
        missing.append("client_id")
    if not client_secret or client_secret == "PASTE_SECRET_VALUE_HERE":
        missing.append("client_secret")
    if missing:
        raise RuntimeError(
            "StreamlitのSecretsにある[onedrive]へ"
            + "、".join(missing)
            + "を設定してください。"
        )
    return client_id, client_secret, redirect_uri


def read_onedrive_configured_refresh_token():
    """通常利用時の自動接続に使う更新トークンをSecretsから読む。"""
    try:
        value = str(st.secrets["onedrive"].get("refresh_token", "")).strip()
    except Exception:
        value = ""
    if value in {"PASTE_REFRESH_TOKEN_HERE", "Microsoftの更新トークン"}:
        return ""
    return value


@st.cache_resource(show_spinner=False)
def get_onedrive_pending_auth_store():
    """外部ログイン中だけ必要な認証情報を一時保持する。"""
    return {"lock": threading.RLock(), "flows": {}}


@st.cache_resource(show_spinner=False)
def get_onedrive_shared_token_store():
    """全利用者が同じOneDriveを使えるよう、サーバー内で認証結果を共有する。"""
    return {"lock": threading.RLock(), "token": None}


def get_onedrive_shared_token_result():
    store = get_onedrive_shared_token_store()
    with store["lock"]:
        token = store.get("token")
        return dict(token) if isinstance(token, dict) else None


def clear_onedrive_shared_token_result():
    store = get_onedrive_shared_token_store()
    with store["lock"]:
        store["token"] = None


def cleanup_onedrive_pending_auth_flows(store):
    now = time.time()
    expired = [
        state
        for state, entry in store["flows"].items()
        if now - float(entry.get("created_at", 0)) > ONEDRIVE_AUTH_FLOW_TTL_SECONDS
    ]
    for state in expired:
        store["flows"].pop(state, None)


def save_onedrive_pending_auth_flow(state, payload):
    store = get_onedrive_pending_auth_store()
    with store["lock"]:
        cleanup_onedrive_pending_auth_flows(store)
        store["flows"][state] = {
            "created_at": time.time(),
            **payload,
        }


def pop_onedrive_pending_auth_flow(state):
    store = get_onedrive_pending_auth_store()
    with store["lock"]:
        cleanup_onedrive_pending_auth_flows(store)
        return store["flows"].pop(state, None)


def get_raw_query_params():
    try:
        return {key: str(value) for key, value in st.query_params.items()}
    except Exception:
        legacy = st.experimental_get_query_params()
        return {
            key: str(value[0] if isinstance(value, list) and value else value)
            for key, value in legacy.items()
        }


def set_query_params_after_onedrive_auth(page="home", customer=""):
    try:
        st.query_params.clear()
        st.query_params["logged_in"] = "1"
        st.query_params["page"] = str(page or "home")
        if customer:
            st.query_params["customer"] = str(customer)
    except Exception:
        params = {"logged_in": "1", "page": str(page or "home")}
        if customer:
            params["customer"] = str(customer)
        st.experimental_set_query_params(**params)


def make_pkce_verifier():
    return uuid.uuid4().hex + uuid.uuid4().hex


def make_pkce_challenge(verifier):
    digest = hashlib.sha256(verifier.encode("ascii")).digest()
    return base64.urlsafe_b64encode(digest).decode("ascii").rstrip("=")


def build_onedrive_sign_in_url(return_page="home", customer_name=""):
    client_id, _, redirect_uri = read_onedrive_settings()
    state = uuid.uuid4().hex
    verifier = make_pkce_verifier()
    save_onedrive_pending_auth_flow(
        state,
        {
            "code_verifier": verifier,
            "redirect_uri": redirect_uri,
            "return_page": str(return_page or "home"),
            "customer_name": str(customer_name or ""),
        },
    )
    params = {
        "client_id": client_id,
        "response_type": "code",
        "redirect_uri": redirect_uri,
        "response_mode": "query",
        "scope": ONEDRIVE_SCOPES,
        "state": state,
        "code_challenge": make_pkce_challenge(verifier),
        "code_challenge_method": "S256",
    }
    return ONEDRIVE_AUTHORIZE_URL + "?" + urllib.parse.urlencode(params)


def save_onedrive_token_result(result):
    token = dict(result or {})
    expires_in = int(token.get("expires_in") or 3600)
    token["expires_at"] = time.time() + max(expires_in, 60)
    st.session_state["onedrive_token_result"] = token
    store = get_onedrive_shared_token_store()
    with store["lock"]:
        store["token"] = dict(token)


def clear_onedrive_auth_state(clear_shared=False):
    st.session_state.pop("onedrive_token_result", None)
    for key in list(st.session_state.keys()):
        if str(key).startswith("onedrive_thumbnail_"):
            st.session_state.pop(key, None)
    if clear_shared:
        clear_onedrive_shared_token_result()


def refresh_onedrive_access_token(token=None):
    refresh_token = str((token or {}).get("refresh_token") or "").strip()
    if not refresh_token:
        refresh_token = read_onedrive_configured_refresh_token()
    if not refresh_token:
        return None
    client_id, client_secret, redirect_uri = read_onedrive_settings()
    response = requests.post(
        ONEDRIVE_TOKEN_URL,
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "redirect_uri": redirect_uri,
            "scope": ONEDRIVE_SCOPES,
        },
        timeout=ONEDRIVE_REQUEST_TIMEOUT,
    )
    if response.status_code != 200:
        return None
    result = response.json()
    if not result.get("refresh_token"):
        result["refresh_token"] = refresh_token
    save_onedrive_token_result(result)
    return result


def get_onedrive_access_token():
    # まずサーバー共有の認証結果を使い、通常利用者にはMicrosoftログインを要求しない。
    token = get_onedrive_shared_token_result()
    if not isinstance(token, dict):
        session_token = st.session_state.get("onedrive_token_result")
        token = dict(session_token) if isinstance(session_token, dict) else None

    access_token = str((token or {}).get("access_token") or "").strip()
    expires_at = float((token or {}).get("expires_at") or 0)
    if access_token and expires_at > time.time() + 60:
        return access_token

    # 同時アクセス時に更新処理が重ならないよう、共有ロック内で再確認する。
    store = get_onedrive_shared_token_store()
    with store["lock"]:
        current = store.get("token")
        if isinstance(current, dict):
            current_access_token = str(current.get("access_token") or "").strip()
            current_expires_at = float(current.get("expires_at") or 0)
            if current_access_token and current_expires_at > time.time() + 60:
                return current_access_token
            token = dict(current)

        refreshed = refresh_onedrive_access_token(token)
        if refreshed and refreshed.get("access_token"):
            return str(refreshed["access_token"])

    clear_onedrive_auth_state(clear_shared=True)
    return None


def process_onedrive_callback_if_present():
    """Microsoftから戻った認証コードを、通常ログイン判定より先に処理する。"""
    params = get_raw_query_params()
    if not params.get("code") and not params.get("error"):
        return
    state = str(params.get("state") or "")
    pending = pop_onedrive_pending_auth_flow(state) if state else None
    return_page = str((pending or {}).get("return_page") or "home")
    customer_name = str((pending or {}).get("customer_name") or "")

    st.session_state.authenticated = True
    st.session_state["page"] = return_page
    st.session_state["selected_customer"] = customer_name if return_page == "detail" else None

    try:
        if params.get("error"):
            description = params.get("error_description") or params.get("error")
            raise RuntimeError(f"Microsoftへのサインインが完了しませんでした：{description}")
        if not pending:
            raise RuntimeError(
                "認証の一時情報が見つかりません。ログイン開始から15分以内に、もう一度サインインしてください。"
            )

        client_id, client_secret, _ = read_onedrive_settings()
        response = requests.post(
            ONEDRIVE_TOKEN_URL,
            data={
                "client_id": client_id,
                "client_secret": client_secret,
                "grant_type": "authorization_code",
                "code": str(params.get("code") or ""),
                "redirect_uri": str(pending.get("redirect_uri") or ""),
                "code_verifier": str(pending.get("code_verifier") or ""),
                "scope": ONEDRIVE_SCOPES,
            },
            timeout=ONEDRIVE_REQUEST_TIMEOUT,
        )
        if response.status_code != 200:
            try:
                detail = response.json().get("error_description") or response.json().get("error")
            except Exception:
                detail = response.text
            raise RuntimeError(f"OneDriveの認証情報を取得できませんでした：{detail}")
        token_result = response.json()
        save_onedrive_token_result(token_result)
        if not read_onedrive_configured_refresh_token():
            setup_refresh_token = str(token_result.get("refresh_token") or "").strip()
            if setup_refresh_token:
                st.session_state["onedrive_refresh_token_setup_value"] = setup_refresh_token
        st.session_state["onedrive_auth_success"] = True
    except Exception as exc:
        st.session_state["onedrive_auth_error"] = str(exc)

    set_query_params_after_onedrive_auth(return_page, customer_name)
    st.rerun()


def onedrive_graph_request(method, path, access_token, expected=(200,), **kwargs):
    headers = dict(kwargs.pop("headers", {}) or {})
    headers["Authorization"] = f"Bearer {access_token}"
    response = requests.request(
        method,
        ONEDRIVE_GRAPH_BASE + path,
        headers=headers,
        timeout=ONEDRIVE_REQUEST_TIMEOUT,
        **kwargs,
    )
    if response.status_code not in expected:
        try:
            payload = response.json()
            message = str(payload.get("error", {}).get("message", "")).strip()
        except Exception:
            message = str(response.text or "").strip()
        if response.status_code == 401:
            clear_onedrive_auth_state(clear_shared=True)
            message = message or "OneDriveの認証が失効しています。管理者が再接続してください。"
        raise RuntimeError(
            f"Microsoft Graphでエラーが発生しました（{response.status_code}）"
            + (f"：{message}" if message else "")
        )
    return response


def get_onedrive_profile(access_token):
    return onedrive_graph_request(
        "GET",
        "/me?$select=displayName,mail,userPrincipalName",
        access_token,
    ).json()


def get_onedrive_path_item(access_token, path):
    encoded = urllib.parse.quote(str(path).strip("/"), safe="/")
    response = onedrive_graph_request(
        "GET",
        f"/me/drive/root:/{encoded}?$select=id,name,folder,webUrl",
        access_token,
        expected=(200, 404),
    )
    return None if response.status_code == 404 else response.json()


def ensure_onedrive_folder_path(access_token, path):
    segments = [segment for segment in str(path).replace("\\", "/").split("/") if segment]
    if not segments:
        raise RuntimeError("OneDriveの保存フォルダが空です。")
    current_path = ""
    parent_id = None
    item = None
    for segment in segments:
        current_path = f"{current_path}/{segment}".strip("/")
        item = get_onedrive_path_item(access_token, current_path)
        if item:
            parent_id = str(item.get("id") or "")
            continue
        target = "/me/drive/root/children" if not parent_id else f"/me/drive/items/{urllib.parse.quote(parent_id, safe='')}/children"
        response = onedrive_graph_request(
            "POST",
            target,
            access_token,
            expected=(200, 201),
            headers={"Content-Type": "application/json"},
            json={
                "name": segment,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "fail",
            },
        )
        item = response.json()
        parent_id = str(item.get("id") or "")
    return item or {}


def upload_onedrive_file(access_token, folder_path, filename, content, content_type):
    clean_name = re.sub(r"[\\/:*?\"<>|]", "_", str(filename or "")).strip().rstrip(".")
    if not clean_name:
        raise RuntimeError("ファイル名が空です。")
    ensure_onedrive_folder_path(access_token, folder_path)
    full_path = f"{str(folder_path).strip('/')}/{clean_name}"
    encoded = urllib.parse.quote(full_path, safe="/")
    response = onedrive_graph_request(
        "PUT",
        f"/me/drive/root:/{encoded}:/content",
        access_token,
        expected=(200, 201),
        headers={"Content-Type": content_type or "application/octet-stream"},
        data=content,
    )
    return response.json()


def delete_onedrive_file(access_token, item_id):
    onedrive_graph_request(
        "DELETE",
        f"/me/drive/items/{urllib.parse.quote(str(item_id), safe='')}",
        access_token,
        expected=(204,),
    )


def download_onedrive_file(access_token, item_id):
    response = onedrive_graph_request(
        "GET",
        f"/me/drive/items/{urllib.parse.quote(str(item_id), safe='')}/content",
        access_token,
        expected=(200,),
    )
    return response.content


def render_onedrive_pdf_inline(pdf_content, filename):
    """OneDriveへ移動せず、現在の顧客カルテ内にPDFを表示する。"""
    encoded = base64.b64encode(bytes(pdf_content or b"")).decode("ascii")
    safe_filename = html.escape(str(filename or "PDF"))
    components.html(
        f"""
        <div style="width:100%; margin:0; padding:0;">
          <object
            data="data:application/pdf;base64,{encoded}"
            type="application/pdf"
            width="100%"
            height="720"
            aria-label="{safe_filename}"
          >
            <div style="padding:16px; font-family:sans-serif; line-height:1.6;">
              この端末ではPDFを画面内に表示できません。下の保存ボタンから確認してください。
            </div>
          </object>
        </div>
        """,
        height=740,
        scrolling=True,
    )


def download_onedrive_thumbnail(access_token, item_id):
    response = onedrive_graph_request(
        "GET",
        f"/me/drive/items/{urllib.parse.quote(str(item_id), safe='')}/thumbnails?$select=medium",
        access_token,
        expected=(200,),
    )
    values = list(response.json().get("value", []))
    if not values:
        return None
    url = str((values[0].get("medium") or {}).get("url") or "").strip()
    if not url:
        return None
    image_response = requests.get(url, timeout=ONEDRIVE_REQUEST_TIMEOUT)
    if image_response.status_code != 200:
        return None
    return image_response.content


# Microsoftから戻った時は、アプリの共通パスワード画面より先に認証を完了する。
process_onedrive_callback_if_present()


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
    st.title("🔒 取引先カルテ")

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


def get_dropbox_response_metadata(response):
    """Return FileMetadata from an upload/copy response."""
    try:
        payload = response.json()
    except Exception:
        return {}
    if not isinstance(payload, dict):
        return {}
    metadata = payload.get("metadata")
    return metadata if isinstance(metadata, dict) else payload


def calculate_dropbox_content_hash(content):
    """Calculate Dropbox content_hash without downloading the remote file again."""
    block_size = 4 * 1024 * 1024
    combined = hashlib.sha256()
    for offset in range(0, len(content), block_size):
        combined.update(hashlib.sha256(content[offset:offset + block_size]).digest())
    return combined.hexdigest()


def verify_dropbox_file_metadata(metadata, expected_content, previous_revision=""):
    """Verify size, content hash, and revision from Dropbox FileMetadata."""
    if not isinstance(metadata, dict):
        raise RuntimeError("Dropbox\u306e\u4fdd\u5b58\u7d50\u679c\u3092\u78ba\u8a8d\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002")

    expected_size = len(expected_content)
    expected_hash = calculate_dropbox_content_hash(expected_content)
    remote_size = metadata.get("size")
    remote_hash = str(metadata.get("content_hash") or "")
    remote_revision = str(metadata.get("rev") or "")

    try:
        size_matches = int(remote_size) == expected_size
    except Exception:
        size_matches = False
    if not size_matches:
        raise RuntimeError("Dropbox\u4fdd\u5b58\u5f8c\u306e\u30d5\u30a1\u30a4\u30eb\u30b5\u30a4\u30ba\u304c\u4e00\u81f4\u3057\u307e\u305b\u3093\u3002")
    if not remote_hash or remote_hash != expected_hash:
        raise RuntimeError("Dropbox\u4fdd\u5b58\u5f8c\u306e\u30d5\u30a1\u30a4\u30eb\u5185\u5bb9\u304c\u4e00\u81f4\u3057\u307e\u305b\u3093\u3002")
    if not remote_revision:
        raise RuntimeError("Dropbox\u4fdd\u5b58\u5f8c\u306erev\u3092\u78ba\u8a8d\u3067\u304d\u307e\u305b\u3093\u3002")
    if previous_revision and remote_revision == previous_revision:
        raise RuntimeError("Dropbox\u306e\u66f4\u65b0\u756a\u53f7\u304c\u5909\u308f\u3063\u3066\u3044\u306a\u3044\u305f\u3081\u3001\u4fdd\u5b58\u3092\u5b8c\u4e86\u3067\u304d\u307e\u305b\u3093\u3002")
    return remote_revision


def get_dropbox_file_metadata(path, access_token):
    """Fetch only metadata; this does not download the Excel bytes."""
    response = call_dropbox_rpc("files/get_metadata", {"path": path}, access_token)
    if response.status_code != 200:
        raise RuntimeError(
            "Dropbox\u306e\u30d5\u30a1\u30a4\u30eb\u60c5\u5831\u3092\u53d6\u5f97\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002\n"
            + dropbox_error_text(response)
        )
    metadata = get_dropbox_response_metadata(response)
    if not metadata:
        raise RuntimeError("Dropbox\u306e\u30d5\u30a1\u30a4\u30eb\u60c5\u5831\u3092\u8aad\u307f\u53d6\u308c\u307e\u305b\u3093\u3067\u3057\u305f\u3002")
    return metadata


def copy_dropbox_file(from_path, to_path, access_token):
    """Create a server-side Dropbox copy without re-uploading the Excel bytes."""
    return call_dropbox_rpc(
        "files/copy_v2",
        {
            "from_path": from_path,
            "to_path": to_path,
            "allow_shared_folder": False,
            "autorename": False,
            "allow_ownership_transfer": False,
        },
        access_token,
    )


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


def create_dropbox_backup(target_path, backup_path, original_content, access_token):
    """Create and verify the pre-save backup, preferring a fast server-side copy."""
    ensure_dropbox_backup_folder(access_token)
    copy_response = copy_dropbox_file(target_path, backup_path, access_token)

    if copy_response.status_code == 200:
        metadata = get_dropbox_response_metadata(copy_response)
        if not metadata.get("content_hash") or metadata.get("size") is None:
            metadata = get_dropbox_file_metadata(backup_path, access_token)
        try:
            verify_dropbox_file_metadata(metadata, original_content)
            return
        except Exception:
            call_dropbox_rpc("files/delete_v2", {"path": backup_path}, access_token)
            raise RuntimeError(
                "PC\u307e\u305f\u306f\u5225\u7aef\u672b\u3067Excel\u304c\u66f4\u65b0\u3055\u308c\u305f\u53ef\u80fd\u6027\u304c\u3042\u308a\u307e\u3059\u3002"
                "\u518d\u8aad\u307f\u8fbc\u307f\u3057\u3066\u304b\u3089\u3084\u308a\u76f4\u3057\u3066\u304f\u3060\u3055\u3044\u3002"
            )

    # Fallback keeps the previous behavior if server-side copy is unavailable.
    backup_response = upload_dropbox_file(
        backup_path,
        original_content,
        access_token,
        mode="add",
    )
    if backup_response.status_code != 200:
        raise RuntimeError(
            "\u30d0\u30c3\u30af\u30a2\u30c3\u30d7\u3092\u4f5c\u6210\u3067\u304d\u306a\u3044\u305f\u3081\u3001\u672c\u756a\u30d5\u30a1\u30a4\u30eb\u306f\u66f4\u65b0\u3057\u307e\u305b\u3093\u3002\n"
            + dropbox_error_text(backup_response)
        )
    metadata = get_dropbox_response_metadata(backup_response)
    if not metadata.get("content_hash") or metadata.get("size") is None:
        metadata = get_dropbox_file_metadata(backup_path, access_token)
    verify_dropbox_file_metadata(metadata, original_content)


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


def find_product_rows_by_usage(ws, customer_name, product_name):
    """同一顧客・同一商品の行を、使用中行と過去行に分けて返す。"""
    matched_rows = [
        row for row in range(1, ws.max_row + 1)
        if normalize_match_value(ws.cell(row, 2).value) == customer_name
        and normalize_match_value(ws.cell(row, 5).value) == product_name
    ]
    active_rows = [
        row for row in matched_rows
        if not is_blank_or_zero(ws.cell(row, 7).value)
    ]
    inactive_rows = [row for row in matched_rows if row not in active_rows]
    return matched_rows, active_rows, inactive_rows


@st.cache_data(ttl=60, show_spinner=False)
def read_edit_values_from_bytes(content, customer_name, product_name):
    """最新ブックから編集欄の現在値を取得する。"""
    workbook = load_workbook(BytesIO(content), keep_vba=True, data_only=False, read_only=False)
    try:
        if DELIVERY_SHEET_NAME not in workbook.sheetnames or SHEET_NAME not in workbook.sheetnames:
            raise ValueError("必要なシート（次回配達日 または Sheet1）が見つかりません。")
        delivery_ws = workbook[DELIVERY_SHEET_NAME]
        matches, active_rows, _ = find_product_rows_by_usage(
            delivery_ws, customer_name, product_name
        )
        product_values = {}
        if len(active_rows) == 1:
            row = active_rows[0]
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
            "商品一致件数": len(active_rows),
            "商品全行件数": len(matches),
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
            active_rows = [
                row for row in rows
                if not is_blank_or_zero(delivery_ws.cell(row, 7).value)
            ]
            selected_row = active_rows[0] if len(active_rows) == 1 else None
            products[product] = {
                "メーカー": delivery_ws.cell(selected_row, 6).value if selected_row else None,
                "本数": delivery_ws.cell(selected_row, 8).value if selected_row else None,
                "kg/本": delivery_ws.cell(selected_row, 9).value if selected_row else None,
                "配達日": delivery_ws.cell(selected_row, 10).value if selected_row else None,
                "商品一致件数": len(active_rows),
                "商品全行件数": len(rows),
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
    return uploaded_content, get_download_revision(response)


def refresh_fast_dropbox_cache_after_save(content, excel_revision, access_token):
    """保存直後のExcelから表示用JSONを更新し、次の再表示で新しい値を出す。"""
    try:
        refreshed_df = rebuild_sheet1_from_formula_references(BytesIO(content))
        if refreshed_df.empty:
            return "保存は完了しましたが、表示用キャッシュを更新できませんでした。更新ボタンを押してください。"

        # Dropbox側の更新番号やJSONの反映待ちに左右されず、保存直後の1回目の
        # 再表示では、今保存したExcelから作った最新データをそのまま使用する。
        # 次の画面実行で1度だけ取り出し、その後は従来どおりDropboxキャッシュを使う。
        st.session_state["customer_excel_immediate_df"] = refreshed_df.copy()

        records = json.loads(
            refreshed_df.to_json(
                orient="records",
                date_format="iso",
                force_ascii=False,
            )
        )
        cache_payload = json.dumps(
            {
                "cache_version": DROPBOX_FAST_CACHE_VERSION,
                "excel_revision": excel_revision,
                "records": records,
            },
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode("utf-8")
        cache_response = upload_dropbox_file(
            DROPBOX_FAST_CACHE_FILE,
            cache_payload,
            access_token,
            mode="overwrite",
        )
        if cache_response.status_code != 200:
            return "保存は完了しましたが、表示用キャッシュを更新できませんでした。更新ボタンを押してください。"
        return ""
    except Exception:
        # 本番Excelの保存と検証は完了しているため、キャッシュ更新失敗だけで保存失敗にはしない。
        return "保存は完了しましたが、表示用キャッシュを更新できませんでした。更新ボタンを押してください。"


def update_workbook_bytes(original_content, customer_name, product_name, proposed):
    """指定項目とK列の配達数量だけを変更し、再オープン検証したbytesを返す。"""
    workbook = load_workbook(BytesIO(original_content), keep_vba=True, data_only=False, read_only=False)
    original_sheets = list(workbook.sheetnames)
    changed_cells = []
    try:
        if DELIVERY_SHEET_NAME not in workbook.sheetnames or SHEET_NAME not in workbook.sheetnames:
            raise ValueError("必要なシート（次回配達日 または Sheet1）が見つかりません。")
        delivery_ws = workbook[DELIVERY_SHEET_NAME]
        product_rows, active_rows, _ = find_product_rows_by_usage(
            delivery_ws, customer_name, product_name
        )
        if not product_rows:
            raise ValueError("顧客名・商品名が一致する行が見つかりません。")
        if len(active_rows) > 1:
            raise ValueError("同じ顧客名・商品名の行が複数見つかりました。確認してください。")
        if not active_rows:
            raise ValueError("使用数量/日に値が入っている行が見つからないため編集できません。")

        product_row = active_rows[0]
        for label, column in {"メーカー": 6, "本数": 8, "kg/本": 9, "配達日": 10}.items():
            cell = delivery_ws.cell(product_row, column)
            new_value = proposed[label]
            if not same_excel_value(cell.value, new_value):
                cell.value = new_value
                changed_cells.append((DELIVERY_SHEET_NAME, product_row, column, new_value))

        # PC版ExcelではマクロがK列「配達数量」を本数×kg/本で記載する。
        # アプリ保存時も同じ値をK列へ入れ、Excelとアプリの次回配達予定を一致させる。
        try:
            delivery_quantity = float(proposed.get("本数")) * float(proposed.get("kg/本"))
            if math.isfinite(delivery_quantity):
                if delivery_quantity.is_integer():
                    delivery_quantity = int(delivery_quantity)
                quantity_cell = delivery_ws.cell(product_row, 11)
                if not same_excel_value(quantity_cell.value, delivery_quantity):
                    quantity_cell.value = delivery_quantity
                    changed_cells.append(
                        (DELIVERY_SHEET_NAME, product_row, 11, delivery_quantity)
                    )
        except (TypeError, ValueError, OverflowError):
            # 既存データで本数またはkg/本が数値でない場合は、従来どおり他項目の保存を続ける。
            pass

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
    """Fast save path with backup, local validation, rev conflict protection, and hash verification."""
    access_token = get_dropbox_access_token()
    target_path = get_dropbox_file_path()
    original_content, download_response = download_dropbox_file(target_path, access_token)
    if original_content is None:
        raise RuntimeError("\u6700\u65b0\u306eExcel\u3092\u53d6\u5f97\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002\n" + dropbox_error_text(download_response))
    revision = get_download_revision(download_response)
    if not revision:
        raise RuntimeError("Dropbox\u306erev\u3092\u53d6\u5f97\u3067\u304d\u306a\u3044\u305f\u3081\u3001\u5b89\u5168\u306e\u305f\u3081\u66f4\u65b0\u3092\u4e2d\u6b62\u3057\u307e\u3057\u305f\u3002")

    timestamp = get_jst_now().strftime("%Y%m%d_%H%M%S_%f")
    backup_path = f"{DROPBOX_BACKUP_FOLDER}/\u914d\u8eca\u4e88\u5b9a \u6b21\u90ce_{timestamp}.xlsm"

    # Dropbox-internal copy avoids uploading the same 1.8 MB file a second time.
    # The copied backup is hash-checked before the production file is touched.
    create_dropbox_backup(
        target_path,
        backup_path,
        original_content,
        access_token,
    )

    saved_content, changed_cells = update_workbook_bytes(
        original_content,
        customer_name,
        product_name,
        proposed,
    )
    upload_response = upload_dropbox_file(
        target_path,
        saved_content,
        access_token,
        mode="update",
        rev=revision,
    )
    if upload_response.status_code == 409:
        raise RuntimeError("PC\u307e\u305f\u306f\u5225\u7aef\u672b\u3067Excel\u304c\u66f4\u65b0\u3055\u308c\u3066\u3044\u307e\u3059\u3002\u518d\u8aad\u307f\u8fbc\u307f\u3057\u3066\u304b\u3089\u3084\u308a\u76f4\u3057\u3066\u304f\u3060\u3055\u3044")
    if upload_response.status_code != 200:
        raise RuntimeError("\u672c\u756aExcel\u3092\u66f4\u65b0\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002\u5fc5\u8981\u306aDropbox\u6a29\u9650\u306f files.content.write \u3067\u3059\u3002\n" + dropbox_error_text(upload_response))

    # The upload response contains size, rev, and content_hash. Verifying those
    # guarantees the bytes without downloading the full workbook again.
    upload_metadata = get_dropbox_response_metadata(upload_response)
    if not upload_metadata.get("content_hash") or upload_metadata.get("size") is None:
        upload_metadata = get_dropbox_file_metadata(target_path, access_token)
    confirmed_revision = verify_dropbox_file_metadata(
        upload_metadata,
        saved_content,
        previous_revision=revision,
    )

    # Use the already verified local bytes to rebuild the immediate-display JSON.
    cache_warning = refresh_fast_dropbox_cache_after_save(
        saved_content,
        confirmed_revision,
        access_token,
    )

    # Keep the existing exact rule: retain the newest 30 backups.
    cleanup_warning = trim_old_dropbox_backups(access_token, keep=30)
    warnings = [warning for warning in (cleanup_warning, cache_warning) if warning]
    st.cache_data.clear()
    return {
        "backup_path": backup_path,
        "updated_at": get_jst_now(),
        "changed_cells": changed_cells,
        "cleanup_warning": "\n".join(warnings),
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
    previous_connected = get_line_connected(customer_name)
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
    warning = record_change_history_safely(
        "顧客",
        "",
        customer,
        "変更",
        {"LINE状態": ("○" if previous_connected else "×", "○" if connected else "×")},
        section="LINE状態",
    )
    remember_change_history_warning(warning)
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


def change_history_value(value):
    """変更前後の値を、JSONと画面表示で安定して扱える文字列へ変換する。"""
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def normalize_change_history_changes(changes):
    """辞書または変更明細リストを、共通の変更明細リストへ整形する。"""
    normalized = []
    if isinstance(changes, dict):
        iterable = []
        for field_name, values in changes.items():
            if isinstance(values, dict):
                before = values.get("before")
                after = values.get("after")
            elif isinstance(values, (list, tuple)) and len(values) >= 2:
                before, after = values[0], values[1]
            else:
                before, after = "", values
            iterable.append(
                {
                    "field": field_name,
                    "before": before,
                    "after": after,
                }
            )
    elif isinstance(changes, list):
        iterable = changes
    else:
        iterable = []

    for item in iterable:
        if not isinstance(item, dict):
            continue
        field_name = clean_value(
            item.get("field") or item.get("項目"),
            blank_text="",
        )
        if not field_name:
            continue
        before = change_history_value(item.get("before"))
        after = change_history_value(item.get("after"))
        if before == after:
            continue
        normalized.append(
            {
                "field": field_name,
                "before": before,
                "after": after,
            }
        )
    return normalized


def clear_change_history_cache():
    try:
        load_change_history_page.clear()
    except Exception:
        pass


def record_change_history(
    target_type,
    target_id,
    target_name,
    action,
    changes,
    section="",
):
    """既存のcustomer_informationテーブルへ、通常顧客情報と分離して履歴を保存する。"""
    if not has_supabase_config():
        raise RuntimeError("Supabase設定がないため変更履歴を保存できません。")

    target_type = clean_value(target_type, blank_text="")
    target_name = clean_value(target_name, blank_text="")
    action = clean_value(action, blank_text="")
    normalized_changes = normalize_change_history_changes(changes)
    if not target_type or not target_name or not action or not normalized_changes:
        return

    now = get_jst_now().isoformat()
    history_id = str(uuid.uuid4())
    content = json.dumps(
        {
            "version": CHANGE_HISTORY_VERSION,
            "target_type": target_type,
            "target_id": clean_value(target_id, blank_text=""),
            "target_name": target_name,
            "action": action,
            "section": clean_value(section, blank_text=""),
            "changes": normalized_changes,
            "source": "app",
        },
        ensure_ascii=False,
        separators=(",", ":"),
    )
    payload = {
        "id": history_id,
        "customer_key": None,
        "customer_name": CHANGE_HISTORY_CUSTOMER,
        "field_name": target_type,
        "content": content,
        "sort_order": 0,
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
        raise RuntimeError("変更履歴の保存中にSupabaseへ接続できませんでした。") from exc
    if response.status_code not in (200, 201):
        detail = str(response.text or "").strip()[:500]
        message = f"変更履歴を保存できませんでした（{response.status_code}）。"
        if detail:
            message += f" {detail}"
        raise RuntimeError(message)
    clear_change_history_cache()


def record_change_history_safely(*args, **kwargs):
    """本体の保存成功を取り消さず、履歴保存だけを警告として返す。"""
    try:
        record_change_history(*args, **kwargs)
        return ""
    except Exception as exc:
        return f"本体の保存は完了しましたが、変更履歴を保存できませんでした：{exc}"


def remember_change_history_warning(warning):
    if warning:
        st.session_state["change_history_warning"] = str(warning)


@st.cache_data(ttl=15, show_spinner=False)
def load_change_history_page(target_type="", start_iso="", offset=0, limit=CHANGE_HISTORY_PAGE_SIZE):
    """変更履歴を新しい順に必要件数だけ取得する。"""
    if not has_supabase_config():
        raise RuntimeError("Supabase設定がありません。")
    params = {
        "select": "id,customer_name,field_name,content,created_at,updated_at",
        "customer_name": f"eq.{CHANGE_HISTORY_CUSTOMER}",
        "order": "created_at.desc,id.desc",
        "limit": str(int(limit) + 1),
        "offset": str(max(0, int(offset))),
    }
    if target_type:
        params["field_name"] = f"eq.{target_type}"
    if start_iso:
        params["created_at"] = f"gte.{start_iso}"
    try:
        response = requests.get(
            get_supabase_customer_information_url(),
            headers=get_supabase_headers(),
            params=params,
            timeout=30,
        )
    except Exception as exc:
        raise RuntimeError("変更履歴の読み込み中にSupabaseへ接続できませんでした。") from exc
    if response.status_code != 200:
        raise RuntimeError(
            f"変更履歴を読み込めませんでした（{response.status_code}）。"
        )
    rows = response.json()
    if not isinstance(rows, list):
        raise RuntimeError("Supabaseから返った変更履歴の形式が正しくありません。")
    return rows


def parse_change_history_row(row):
    payload = {}
    try:
        parsed = json.loads(str(row.get("content") or "{}"))
        if isinstance(parsed, dict):
            payload = parsed
    except Exception:
        payload = {}
    return {
        "id": clean_value(row.get("id"), blank_text=""),
        "created_at": clean_value(row.get("created_at"), blank_text=""),
        "target_type": clean_value(
            payload.get("target_type") or row.get("field_name"),
            blank_text="",
        ),
        "target_id": clean_value(payload.get("target_id"), blank_text=""),
        "target_name": clean_value(payload.get("target_name"), blank_text=""),
        "action": clean_value(payload.get("action"), blank_text=""),
        "section": clean_value(payload.get("section"), blank_text=""),
        "changes": normalize_change_history_changes(payload.get("changes", [])),
        "raw_content": clean_value(row.get("content"), blank_text=""),
    }


def change_history_rows_to_dataframe(rows):
    """変更履歴を、1変更項目につき1行のCSV向け表へ変換する。"""
    records = []
    for row in rows:
        parsed = parse_change_history_row(row)
        changes = parsed["changes"] or [
            {"field": "解析できない履歴", "before": "", "after": parsed["raw_content"]}
        ]
        for change in changes:
            records.append(
                {
                    "変更日時": parsed["created_at"],
                    "対象区分": parsed["target_type"],
                    "対象ID": parsed["target_id"],
                    "対象名": parsed["target_name"],
                    "操作": parsed["action"],
                    "変更箇所": parsed["section"],
                    "項目": change.get("field", ""),
                    "変更前": change.get("before", ""),
                    "変更後": change.get("after", ""),
                    "履歴ID": parsed["id"],
                }
            )
    return backup_dataframe(
        records,
        [
            "変更日時", "対象区分", "対象ID", "対象名", "操作",
            "変更箇所", "項目", "変更前", "変更後", "履歴ID",
        ],
    )


def display_change_history_value(value):
    """変更確認画面では、ISO形式の日時を日付だけで表示する。"""
    if value is None or value == "":
        return "（空欄）"
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y/%m/%d")

    text = str(value).strip()
    iso_date_match = re.fullmatch(
        r"(\d{4})-(\d{2})-(\d{2})(?:[T ]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?)?",
        text,
    )
    if iso_date_match:
        year, month, day = iso_date_match.groups()
        return f"{year}/{month}/{day}"
    return text


def show_change_history_page():
    st.header("🕘 変更確認")
    st.caption("アプリから正常に保存された変更を新しい順に表示します。メモ帳は対象外です。")

    target_label = st.selectbox(
        "対象",
        ["すべて", *CHANGE_HISTORY_TARGETS],
        key="change_history_target_filter",
    )
    period_label = st.selectbox(
        "期間",
        ["今日", "7日間", "30日間", "すべて"],
        index=2,
        key="change_history_period_filter",
    )
    target_type = "" if target_label == "すべて" else target_label
    now = get_jst_now()
    if period_label == "今日":
        start_iso = now.replace(hour=0, minute=0, second=0, microsecond=0).isoformat()
    elif period_label == "7日間":
        start_iso = (now - timedelta(days=7)).replace(
            hour=0, minute=0, second=0, microsecond=0
        ).isoformat()
    elif period_label == "30日間":
        start_iso = (now - timedelta(days=30)).replace(
            hour=0, minute=0, second=0, microsecond=0
        ).isoformat()
    else:
        start_iso = ""

    signature = f"{target_type}|{start_iso}"
    if st.session_state.get("change_history_filter_signature") != signature:
        st.session_state["change_history_filter_signature"] = signature
        st.session_state["change_history_offset"] = 0
    offset = max(0, int(st.session_state.get("change_history_offset", 0)))

    try:
        rows = load_change_history_page(
            target_type=target_type,
            start_iso=start_iso,
            offset=offset,
            limit=CHANGE_HISTORY_PAGE_SIZE,
        )
    except Exception as exc:
        st.error(str(exc))
        return

    has_next = len(rows) > CHANGE_HISTORY_PAGE_SIZE
    visible_rows = rows[:CHANGE_HISTORY_PAGE_SIZE]
    if not visible_rows:
        st.info("該当する変更履歴はありません。")
    else:
        st.caption(
            f"{offset + 1}件目～{offset + len(visible_rows)}件目を表示"
        )
        for row in visible_rows:
            parsed = parse_change_history_row(row)
            title = "　".join(
                part for part in (
                    parsed["target_type"],
                    parsed["target_name"],
                ) if part
            ) or "変更履歴"
            with st.container(border=True):
                st.markdown(f"**{html.escape(title)}**")
                meta = " ｜ ".join(
                    part for part in (
                        format_note_datetime(parsed["created_at"]),
                        parsed["action"],
                    ) if part
                )
                if meta:
                    st.caption(meta)
                if parsed["section"]:
                    st.markdown(
                        f"**変更箇所：{html.escape(parsed['section'])}**"
                    )
                for change in parsed["changes"]:
                    before = display_change_history_value(change.get("before", ""))
                    after = display_change_history_value(change.get("after", ""))
                    field_name = html.escape(clean_value(change.get("field"), blank_text="変更内容"))
                    st.write(f"{field_name}：{before} → {after}")

    previous_col, page_col, next_col = st.columns([1, 1, 1])
    with previous_col:
        if st.button(
            "← 前の30件",
            key="change_history_previous",
            disabled=offset == 0,
            use_container_width=True,
        ):
            st.session_state["change_history_offset"] = max(
                0, offset - CHANGE_HISTORY_PAGE_SIZE
            )
            st.rerun()
    with page_col:
        st.caption(f"ページ {offset // CHANGE_HISTORY_PAGE_SIZE + 1}")
    with next_col:
        if st.button(
            "次の30件 →",
            key="change_history_next",
            disabled=not has_next,
            use_container_width=True,
        ):
            st.session_state["change_history_offset"] = offset + CHANGE_HISTORY_PAGE_SIZE
            st.rerun()


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


def make_onedrive_attachment_field_name():
    return ONEDRIVE_ATTACHMENT_PREFIX + uuid.uuid4().hex


def is_onedrive_attachment_item(item):
    return clean_value(item.get("field_name"), blank_text="").startswith(
        ONEDRIVE_ATTACHMENT_PREFIX
    )


def normalize_attachment_tags(values):
    if isinstance(values, str):
        candidates = re.split(r"[,、\n]+", values)
    elif isinstance(values, (list, tuple, set)):
        candidates = list(values)
    else:
        candidates = []
    result = []
    seen = set()
    for value in candidates:
        tag = clean_value(value, blank_text="").strip().lstrip("#").strip()
        if not tag or tag in seen:
            continue
        seen.add(tag)
        result.append(tag)
    return result


def parse_onedrive_attachment_item(item):
    if not is_onedrive_attachment_item(item):
        return None
    try:
        payload = json.loads(str(item.get("content") or "{}"))
    except Exception:
        return None
    if not isinstance(payload, dict):
        return None
    return {
        "id": str(item.get("id") or ""),
        "field_name": str(item.get("field_name") or ""),
        "customer_key": clean_value(item.get("customer_key"), blank_text=""),
        "customer_name": clean_value(item.get("customer_name"), blank_text=""),
        "file_id": clean_value(payload.get("file_id"), blank_text=""),
        "original_name": clean_value(payload.get("original_name"), blank_text=""),
        "stored_name": clean_value(payload.get("stored_name"), blank_text=""),
        "file_type": clean_value(payload.get("file_type"), blank_text=""),
        "mime_type": clean_value(payload.get("mime_type"), blank_text=""),
        "size": payload.get("size") or 0,
        "onedrive_path": clean_value(payload.get("onedrive_path"), blank_text=""),
        "web_url": clean_value(payload.get("web_url"), blank_text=""),
        "tags": normalize_attachment_tags(payload.get("tags") or []),
        "remarks": clean_value(payload.get("remarks"), blank_text=""),
        "uploaded_by": clean_value(payload.get("uploaded_by"), blank_text=""),
        "created_at": clean_value(payload.get("created_at"), blank_text="")
        or clean_value(item.get("created_at"), blank_text=""),
        "updated_at": clean_value(item.get("updated_at"), blank_text=""),
        "version": payload.get("version") or ONEDRIVE_ATTACHMENT_VERSION,
    }


def serialize_onedrive_attachment(attachment):
    payload = {
        "version": ONEDRIVE_ATTACHMENT_VERSION,
        "file_id": attachment.get("file_id", ""),
        "original_name": attachment.get("original_name", ""),
        "stored_name": attachment.get("stored_name", ""),
        "file_type": attachment.get("file_type", ""),
        "mime_type": attachment.get("mime_type", ""),
        "size": int(attachment.get("size") or 0),
        "onedrive_path": attachment.get("onedrive_path", ""),
        "web_url": attachment.get("web_url", ""),
        "tags": normalize_attachment_tags(attachment.get("tags") or []),
        "remarks": attachment.get("remarks", ""),
        "uploaded_by": attachment.get("uploaded_by", ""),
        "created_at": attachment.get("created_at", ""),
    }
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def get_customer_onedrive_folder_key(customer_key, customer_name):
    if customer_key:
        raw = f"顧客ID_{customer_key}"
    else:
        digest = hashlib.sha256(str(customer_name).encode("utf-8")).hexdigest()[:16]
        raw = f"顧客仮ID_{digest}"
    safe = re.sub(r"[\\/:*?\"<>|]", "_", raw).strip().rstrip(".")
    return safe or "顧客仮ID_未設定"


def get_customer_attachments(customer_name, customer_key):
    items = load_customer_information(customer_name, customer_key)
    attachments = []
    for item in items:
        parsed = parse_onedrive_attachment_item(item)
        if parsed:
            attachments.append(parsed)
    attachments.sort(
        key=lambda row: str(row.get("created_at") or row.get("updated_at") or ""),
        reverse=True,
    )
    return attachments


def attachment_file_kind(filename, mime_type=""):
    suffix = Path(str(filename or "")).suffix.lower()
    mime = str(mime_type or "").lower()
    if mime.startswith("image/") or suffix in ONEDRIVE_IMAGE_EXTENSIONS:
        return "image"
    if mime == "application/pdf" or suffix in ONEDRIVE_PDF_EXTENSIONS:
        return "pdf"
    return ""


def save_customer_onedrive_attachment(
    customer_name,
    customer_key,
    uploaded_name,
    content,
    mime_type,
    tags,
    remarks,
    access_token,
):
    file_kind = attachment_file_kind(uploaded_name, mime_type)
    if not file_kind:
        raise ValueError("画像（JPG・JPEG・PNG・WEBP）またはPDFを選んでください。")
    if not content:
        raise ValueError("選択したファイルが空です。")

    folder_key = get_customer_onedrive_folder_key(customer_key, customer_name)
    category_folder = "写真" if file_kind == "image" else "資料"
    folder_path = "/".join(
        [ONEDRIVE_ROOT_FOLDER, ONEDRIVE_CUSTOMER_FOLDER, folder_key, category_folder]
    )
    original_name = Path(str(uploaded_name or "file")).name
    timestamp = get_jst_now().strftime("%Y%m%d_%H%M%S")
    stored_name = f"{timestamp}_{uuid.uuid4().hex[:8]}_{original_name}"
    uploaded_item = upload_onedrive_file(
        access_token,
        folder_path,
        stored_name,
        content,
        mime_type,
    )
    file_id = clean_value(uploaded_item.get("id"), blank_text="")
    if not file_id:
        raise RuntimeError("OneDriveから保存済みファイルIDを取得できませんでした。")

    profile = {}
    try:
        profile = get_onedrive_profile(access_token)
    except Exception:
        profile = {}
    uploaded_by = (
        clean_value(profile.get("displayName"), blank_text="")
        or clean_value(profile.get("mail"), blank_text="")
        or clean_value(profile.get("userPrincipalName"), blank_text="")
    )
    attachment = {
        "file_id": file_id,
        "original_name": original_name,
        "stored_name": clean_value(uploaded_item.get("name"), blank_text="") or stored_name,
        "file_type": file_kind,
        "mime_type": mime_type or mimetypes.guess_type(original_name)[0] or "application/octet-stream",
        "size": int(uploaded_item.get("size") or len(content)),
        "onedrive_path": folder_path,
        "web_url": clean_value(uploaded_item.get("webUrl"), blank_text=""),
        "tags": normalize_attachment_tags(tags),
        "remarks": str(remarks or "").strip(),
        "uploaded_by": uploaded_by,
        "created_at": get_jst_now().isoformat(),
    }
    try:
        insert_customer_information(
            customer_name,
            customer_key,
            make_onedrive_attachment_field_name(),
            serialize_onedrive_attachment(attachment),
            int(time.time()),
        )
    except Exception:
        try:
            delete_onedrive_file(access_token, file_id)
        except Exception:
            pass
        raise
    return attachment


def update_customer_onedrive_attachment_metadata(attachment, tags, remarks):
    updated = dict(attachment)
    updated["tags"] = normalize_attachment_tags(tags)
    updated["remarks"] = str(remarks or "").strip()
    update_customer_information(
        attachment["id"],
        attachment["field_name"],
        serialize_onedrive_attachment(updated),
    )
    return updated


def format_attachment_size(value):
    try:
        size = float(value or 0)
    except Exception:
        size = 0
    units = ["B", "KB", "MB", "GB"]
    for unit in units:
        if size < 1024 or unit == units[-1]:
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return "0 B"


def format_attachment_datetime(value):
    text = clean_value(value, blank_text="")
    if not text:
        return ""
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        return parsed.astimezone(timezone(timedelta(hours=9))).strftime("%Y/%m/%d %H:%M")
    except Exception:
        return text


def onedrive_attachment_rows_to_dataframe(rows):
    records = []
    for row in rows:
        attachment = parse_onedrive_attachment_item(row)
        if not attachment:
            continue
        records.append(
            {
                "顧客ID": attachment.get("customer_key", ""),
                "顧客名": attachment.get("customer_name", ""),
                "種類": "写真" if attachment.get("file_type") == "image" else "PDF",
                "元ファイル名": attachment.get("original_name", ""),
                "OneDrive保存名": attachment.get("stored_name", ""),
                "OneDrive保存先": attachment.get("onedrive_path", ""),
                "タグ": " ".join(f"#{tag}" for tag in attachment.get("tags", [])),
                "備考": attachment.get("remarks", ""),
                "サイズ": attachment.get("size", ""),
                "登録者": attachment.get("uploaded_by", ""),
                "登録日時": attachment.get("created_at", ""),
                "OneDriveファイルID": attachment.get("file_id", ""),
                "保存ID": attachment.get("id", ""),
            }
        )
    records.sort(key=lambda record: str(record.get("登録日時") or ""), reverse=True)
    return backup_dataframe(
        records,
        [
            "顧客ID", "顧客名", "種類", "元ファイル名", "OneDrive保存名",
            "OneDrive保存先", "タグ", "備考", "サイズ", "登録者", "登録日時",
            "OneDriveファイルID", "保存ID",
        ],
    )


def render_customer_attachments_section(customer_name, customer_key=None):
    identity = customer_key or customer_name
    suffix = hashlib.sha256(str(identity).encode("utf-8")).hexdigest()[:16]
    success_key = f"onedrive_attachment_success_{suffix}"
    edit_key = f"onedrive_attachment_edit_{suffix}"
    delete_key = f"onedrive_attachment_delete_{suffix}"
    limit_key = f"onedrive_attachment_limit_{suffix}"
    preview_key = f"onedrive_attachment_preview_{suffix}"
    preview_data_key = f"onedrive_attachment_preview_data_{suffix}"

    if not has_supabase_config():
        with st.expander("📎 写真・資料"):
            st.warning("写真・資料の管理にはSupabase設定が必要です。")
        return

    try:
        attachments = get_customer_attachments(customer_name, customer_key)
    except Exception as exc:
        with st.expander("📎 写真・資料"):
            st.warning(f"写真・資料の一覧を読み込めませんでした：{exc}")
        return

    with st.expander(f"📎 写真・資料　{len(attachments)}件", expanded=False):
        success_message = st.session_state.pop(success_key, None)
        if success_message:
            st.success(success_message)
        auth_success = st.session_state.pop("onedrive_auth_success", None)
        if auth_success:
            st.success("OneDriveの初回接続が完了しました。")
        auth_error = st.session_state.pop("onedrive_auth_error", None)
        if auth_error:
            st.error(auth_error)

        try:
            read_onedrive_settings()
        except Exception as exc:
            st.warning(str(exc))
            st.code(
                "[onedrive]\n"
                'client_id = "MicrosoftのクライアントID"\n'
                'client_secret = "Microsoftのシークレットの値"\n'
                'redirect_uri = "https://aoyama-kokyaku.streamlit.app"\n'
                'refresh_token = "初回接続後に表示される値"'
            )
            return

        configured_refresh_token = read_onedrive_configured_refresh_token()
        setup_refresh_token = str(
            st.session_state.get("onedrive_refresh_token_setup_value") or ""
        ).strip()
        if setup_refresh_token and not configured_refresh_token:
            st.warning(
                "次回から自動接続するため、下の1行を顧客カルテの"
                "Streamlit Secretsにある[onedrive]の中へ追加してください。"
            )
            st.code(f'refresh_token = "{setup_refresh_token}"', language="toml")
            st.caption("追加して保存するとアプリが再起動し、通常画面から接続ボタンが消えます。")

        access_token = get_onedrive_access_token()
        if not access_token:
            if configured_refresh_token:
                st.warning(
                    "OneDriveへ自動接続できませんでした。Microsoft側で認証が失効した可能性があります。"
                )
                connect_label = "OneDriveを再接続（管理者用）"
            else:
                st.info(
                    "最初の1回だけ管理者がOneDriveへ接続し、表示された更新トークンをSecretsへ追加してください。"
                )
                connect_label = "OneDrive初回設定（管理者用）"
            try:
                st.link_button(
                    connect_label,
                    build_onedrive_sign_in_url("detail", customer_name),
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"OneDriveへの接続を開始できませんでした：{exc}")
        else:
            st.markdown("#### 追加")
            source_mode = st.radio(
                "追加方法",
                ["写真を撮る", "画像・PDFを選ぶ"],
                horizontal=True,
                key=f"onedrive_attachment_source_{suffix}",
            )
            camera_file = None
            selected_file = None
            if source_mode == "写真を撮る":
                camera_file = st.camera_input(
                    "カメラで撮影",
                    key=f"onedrive_attachment_camera_{suffix}",
                )
            else:
                selected_file = st.file_uploader(
                    "画像またはPDFを1つ選択",
                    type=["jpg", "jpeg", "png", "webp", "pdf"],
                    accept_multiple_files=False,
                    key=f"onedrive_attachment_uploader_{suffix}",
                )

            fixed_tags = st.multiselect(
                "固定タグ",
                list(ONEDRIVE_FIXED_TAGS),
                key=f"onedrive_attachment_fixed_tags_{suffix}",
            )
            free_tags = st.text_input(
                "自由タグ",
                placeholder="例：北海道、タンク、要確認",
                key=f"onedrive_attachment_free_tags_{suffix}",
            )
            remarks = st.text_area(
                "備考",
                placeholder="写真や資料について残したい内容",
                height=90,
                key=f"onedrive_attachment_remarks_{suffix}",
            )
            if st.button(
                "OneDriveへ保存",
                type="primary",
                use_container_width=True,
                key=f"onedrive_attachment_upload_{suffix}",
            ):
                uploaded = camera_file if source_mode == "写真を撮る" else selected_file
                if uploaded is None:
                    st.warning("写真を撮るか、画像・PDFを選んでください。")
                else:
                    try:
                        tags = list(fixed_tags) + normalize_attachment_tags(free_tags)
                        with st.spinner("OneDriveへ保存しています…"):
                            saved = save_customer_onedrive_attachment(
                                customer_name,
                                customer_key,
                                uploaded.name,
                                uploaded.getvalue(),
                                uploaded.type or mimetypes.guess_type(uploaded.name)[0] or "application/octet-stream",
                                tags,
                                remarks,
                                access_token,
                            )
                        remember_change_history_warning(
                            record_change_history_safely(
                                "顧客",
                                customer_key or "",
                                customer_name,
                                "追加",
                                {
                                    "ファイル": ("", saved.get("original_name", "")),
                                    "タグ": ("", " ".join(f"#{tag}" for tag in saved.get("tags", []))),
                                },
                                section="写真・資料",
                            )
                        )
                        st.session_state[success_key] = "写真・資料を保存しました。"
                        st.session_state[limit_key] = ONEDRIVE_PAGE_SIZE
                        st.rerun()
                    except Exception as exc:
                        st.error(f"保存できませんでした：{exc}")

        st.markdown("---")
        st.markdown("#### 保存済み")
        if not attachments:
            st.info("保存されている写真・資料はありません。")
            return

        type_filter = st.selectbox(
            "種類",
            ["すべて", "写真", "PDF"],
            key=f"onedrive_attachment_type_filter_{suffix}",
        )
        all_tags = sorted({tag for item in attachments for tag in item.get("tags", [])})
        tag_filter = st.multiselect(
            "タグで絞り込み",
            all_tags,
            key=f"onedrive_attachment_tag_filter_{suffix}",
        ) if all_tags else []

        filtered = []
        for attachment in attachments:
            if type_filter == "写真" and attachment.get("file_type") != "image":
                continue
            if type_filter == "PDF" and attachment.get("file_type") != "pdf":
                continue
            if tag_filter and not set(tag_filter).issubset(set(attachment.get("tags", []))):
                continue
            filtered.append(attachment)

        limit = int(st.session_state.get(limit_key, ONEDRIVE_PAGE_SIZE))
        active_edit_id = st.session_state.get(edit_key)
        active_delete_id = st.session_state.get(delete_key)
        active_preview_id = st.session_state.get(preview_key)

        if not filtered:
            st.info("条件に一致する写真・資料はありません。")

        for attachment in filtered[:limit]:
            item_id = attachment.get("file_id", "")
            metadata_id = attachment.get("id", "")
            filename = attachment.get("original_name", "名称未設定")
            with st.container(border=True):
                if attachment.get("file_type") == "image" and access_token and item_id:
                    thumb_key = f"onedrive_thumbnail_{item_id}"
                    if not isinstance(st.session_state.get(thumb_key), bytes):
                        try:
                            thumbnail = download_onedrive_thumbnail(access_token, item_id)
                            if thumbnail:
                                st.session_state[thumb_key] = thumbnail
                        except Exception:
                            pass
                    if isinstance(st.session_state.get(thumb_key), bytes):
                        st.image(st.session_state[thumb_key], use_column_width=True)

                icon = "🖼" if attachment.get("file_type") == "image" else "📄"
                st.markdown(f"**{icon} {html.escape(filename)}**", unsafe_allow_html=True)
                st.caption(
                    f"{format_attachment_size(attachment.get('size'))}　"
                    f"保存：{format_attachment_datetime(attachment.get('created_at'))}"
                )
                if attachment.get("tags"):
                    st.markdown(" ".join(f"`#{tag}`" for tag in attachment["tags"]))
                if attachment.get("remarks"):
                    st.write(attachment["remarks"])
                preview_label = (
                    "画像を閉じる"
                    if attachment.get("file_type") == "image" and active_preview_id == metadata_id
                    else "画像を大きく表示"
                    if attachment.get("file_type") == "image"
                    else "PDFを閉じる"
                    if active_preview_id == metadata_id
                    else "PDFを表示"
                )
                if st.button(
                    preview_label,
                    key=f"onedrive_attachment_preview_button_{metadata_id}",
                    use_container_width=True,
                ):
                    if active_preview_id == metadata_id:
                        st.session_state.pop(preview_key, None)
                        st.session_state.pop(preview_data_key, None)
                    else:
                        st.session_state[preview_key] = metadata_id
                        st.session_state.pop(preview_data_key, None)
                    st.rerun()

                if active_preview_id == metadata_id:
                    if not access_token or not item_id:
                        st.error("表示するにはOneDriveへ接続してください。")
                    else:
                        preview_data = st.session_state.get(preview_data_key)
                        if not (
                            isinstance(preview_data, dict)
                            and preview_data.get("metadata_id") == metadata_id
                            and isinstance(preview_data.get("content"), bytes)
                        ):
                            try:
                                with st.spinner("ファイルを読み込んでいます…"):
                                    content = download_onedrive_file(access_token, item_id)
                                preview_data = {
                                    "metadata_id": metadata_id,
                                    "content": content,
                                }
                                st.session_state[preview_data_key] = preview_data
                            except Exception as exc:
                                preview_data = None
                                st.error(f"表示できませんでした：{exc}")

                        if isinstance(preview_data, dict):
                            content = preview_data.get("content", b"")
                            if attachment.get("file_type") == "image":
                                st.image(content, caption=filename, use_column_width=True)
                            else:
                                render_onedrive_pdf_inline(content, filename)
                                st.download_button(
                                    "PDFを端末に保存",
                                    data=content,
                                    file_name=filename,
                                    mime=attachment.get("mime_type") or "application/pdf",
                                    use_container_width=True,
                                    key=f"onedrive_attachment_pdf_download_{metadata_id}",
                                )

                if active_edit_id == metadata_id:
                    current_fixed = [tag for tag in attachment.get("tags", []) if tag in ONEDRIVE_FIXED_TAGS]
                    current_free = [tag for tag in attachment.get("tags", []) if tag not in ONEDRIVE_FIXED_TAGS]
                    edited_fixed = st.multiselect(
                        "固定タグを編集",
                        list(ONEDRIVE_FIXED_TAGS),
                        default=current_fixed,
                        key=f"onedrive_attachment_edit_fixed_{metadata_id}",
                    )
                    edited_free = st.text_input(
                        "自由タグを編集",
                        value="、".join(current_free),
                        key=f"onedrive_attachment_edit_free_{metadata_id}",
                    )
                    edited_remarks = st.text_area(
                        "備考を編集",
                        value=attachment.get("remarks", ""),
                        height=90,
                        key=f"onedrive_attachment_edit_remarks_{metadata_id}",
                    )
                    save_col, cancel_col = st.columns(2)
                    with save_col:
                        if st.button(
                            "保存",
                            key=f"onedrive_attachment_edit_save_{metadata_id}",
                            type="primary",
                            use_container_width=True,
                        ):
                            try:
                                old_tags = " ".join(f"#{tag}" for tag in attachment.get("tags", []))
                                new_tags = list(edited_fixed) + normalize_attachment_tags(edited_free)
                                update_customer_onedrive_attachment_metadata(
                                    attachment,
                                    new_tags,
                                    edited_remarks,
                                )
                                changes = {}
                                new_tags_text = " ".join(f"#{tag}" for tag in normalize_attachment_tags(new_tags))
                                if old_tags != new_tags_text:
                                    changes["タグ"] = (old_tags, new_tags_text)
                                if attachment.get("remarks", "") != str(edited_remarks or "").strip():
                                    changes["備考"] = (attachment.get("remarks", ""), str(edited_remarks or "").strip())
                                remember_change_history_warning(
                                    record_change_history_safely(
                                        "顧客",
                                        customer_key or "",
                                        customer_name,
                                        "変更",
                                        changes,
                                        section=f"写真・資料：{filename}",
                                    )
                                )
                                st.session_state.pop(edit_key, None)
                                st.session_state[success_key] = "タグ・備考を更新しました。"
                                st.rerun()
                            except Exception as exc:
                                st.error(f"更新できませんでした：{exc}")
                    with cancel_col:
                        if st.button(
                            "キャンセル",
                            key=f"onedrive_attachment_edit_cancel_{metadata_id}",
                            use_container_width=True,
                        ):
                            st.session_state.pop(edit_key, None)
                            st.rerun()
                    continue

                if active_delete_id == metadata_id:
                    st.warning(f"「{filename}」をOneDriveから削除します。")
                    delete_col, cancel_col = st.columns(2)
                    with delete_col:
                        if st.button(
                            "削除する",
                            key=f"onedrive_attachment_delete_yes_{metadata_id}",
                            type="primary",
                            use_container_width=True,
                        ):
                            if not access_token:
                                st.error("削除するにはOneDriveへ接続してください。")
                            else:
                                try:
                                    with st.spinner("削除しています…"):
                                        delete_onedrive_file(access_token, item_id)
                                        delete_customer_information(metadata_id)
                                    remember_change_history_warning(
                                        record_change_history_safely(
                                            "顧客",
                                            customer_key or "",
                                            customer_name,
                                            "削除",
                                            {"ファイル": (filename, "")},
                                            section="写真・資料",
                                        )
                                    )
                                    st.session_state.pop(delete_key, None)
                                    st.session_state.pop(f"onedrive_thumbnail_{item_id}", None)
                                    if st.session_state.get(preview_key) == metadata_id:
                                        st.session_state.pop(preview_key, None)
                                        st.session_state.pop(preview_data_key, None)
                                    st.session_state[success_key] = "写真・資料を削除しました。"
                                    st.rerun()
                                except Exception as exc:
                                    st.error(f"削除できませんでした：{exc}")
                    with cancel_col:
                        if st.button(
                            "キャンセル",
                            key=f"onedrive_attachment_delete_no_{metadata_id}",
                            use_container_width=True,
                        ):
                            st.session_state.pop(delete_key, None)
                            st.rerun()
                    continue

                action_col, delete_col = st.columns(2)
                with action_col:
                    if st.button(
                        "タグ・備考を編集",
                        key=f"onedrive_attachment_edit_button_{metadata_id}",
                        use_container_width=True,
                    ):
                        st.session_state[edit_key] = metadata_id
                        st.session_state.pop(delete_key, None)
                        st.rerun()
                with delete_col:
                    if st.button(
                        "削除",
                        key=f"onedrive_attachment_delete_button_{metadata_id}",
                        use_container_width=True,
                    ):
                        st.session_state[delete_key] = metadata_id
                        st.session_state.pop(edit_key, None)
                        st.rerun()

        if len(filtered) > limit:
            if st.button(
                "さらに表示",
                key=f"onedrive_attachment_more_{suffix}",
                use_container_width=True,
            ):
                st.session_state[limit_key] = limit + ONEDRIVE_PAGE_SIZE
                st.rerun()

        # 自動接続設定後は利用者ごとのMicrosoft接続操作を表示しない。
        if access_token and not configured_refresh_token:
            st.markdown("---")
            if st.button(
                "初回設定中の一時接続を解除",
                key=f"onedrive_attachment_signout_{suffix}",
                use_container_width=True,
            ):
                clear_onedrive_auth_state(clear_shared=True)
                st.session_state.pop("onedrive_refresh_token_setup_value", None)
                st.rerun()



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


@st.cache_data(ttl=30, show_spinner=False)
def load_all_past_product_notes_from_supabase():
    """取引先メモ画面で使う過去商品メモを、全顧客分読み込む。"""
    if not has_supabase_config():
        return []

    rows = []
    page_size = 1000
    offset = 0
    while True:
        params = {
            "select": "id,customer_key,customer_name,field_name,content,sort_order,created_at,updated_at",
            "field_name": f"like.{PAST_PRODUCT_NOTE_PREFIX}*",
            "order": "updated_at.desc,created_at.desc,id.desc",
            "limit": str(page_size),
            "offset": str(offset),
        }
        try:
            response = requests.get(
                get_supabase_customer_information_url(),
                headers=get_supabase_headers(),
                params=params,
                timeout=30,
            )
        except Exception as exc:
            raise RuntimeError("過去商品メモの読み込み中にSupabaseへ接続できませんでした。") from exc

        check_customer_information_response("読み込み", response, (200,))
        page = response.json()
        if not isinstance(page, list):
            raise RuntimeError("Supabaseから返った過去商品メモの形式が正しくありません。")

        rows.extend(item for item in page if is_past_product_note_item(item))
        if len(page) < page_size:
            break
        offset += page_size

    return rows


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
    else:
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

    load_all_past_product_notes_from_supabase.clear()


def delete_past_product_note(note_item):
    """過去商品の商品別メモを削除する。"""
    item_id = clean_value(note_item.get("id"), blank_text="")
    if not item_id:
        raise RuntimeError("削除する商品メモが見つかりません。")
    delete_customer_information(item_id)
    load_all_past_product_notes_from_supabase.clear()


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
        save_success_key = f"past_product_note_save_success_{state_suffix}"

        save_succeeded = bool(st.session_state.pop(save_success_key, False))
        with st.expander(product_name, expanded=save_succeeded):
            if save_succeeded:
                st.success("メモを保存しました。")

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
                        st.session_state[save_success_key] = True
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



def make_estimate_field_name():
    """顧客情報テーブル内で、見積りを通常項目と分ける内部項目名を作る。"""
    return f"{ESTIMATE_PREFIX}{uuid.uuid4()}"


def is_estimate_item(item):
    """顧客情報テーブル上の提案・見積り専用レコードか判定する。"""
    return clean_value(item.get("field_name"), blank_text="").startswith(ESTIMATE_PREFIX)


def estimate_date_text(value):
    """見積りの日付をYYYY-MM-DDへそろえる。"""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = clean_value(value, blank_text="").strip()
    match = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", text)
    if not match:
        return text
    year, month, day = (int(part) for part in match.groups())
    try:
        return date(year, month, day).isoformat()
    except ValueError:
        return text


def estimate_date_input_value(value):
    """保存済みの日付をst.date_inputで使えるdateへ変換する。"""
    text = estimate_date_text(value)
    try:
        return date.fromisoformat(text)
    except (TypeError, ValueError):
        return get_jst_now().date()


def format_estimate_date(value):
    text = estimate_date_text(value)
    try:
        return date.fromisoformat(text).strftime("%Y/%m/%d")
    except (TypeError, ValueError):
        return text or "未入力"


def serialize_estimate(proposal_date, product_name, manufacturer, unit_price, remarks):
    payload = {
        "version": ESTIMATE_VERSION,
        "proposal_date": estimate_date_text(proposal_date),
        "product_name": clean_value(product_name, blank_text="").strip(),
        "manufacturer": clean_value(manufacturer, blank_text="").strip(),
        "unit_price": clean_value(unit_price, blank_text="").strip(),
        "remarks": clean_value(remarks, blank_text="").strip(),
    }
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def parse_estimate_item(item):
    """Supabase内部レコードを画面表示用の見積り辞書へ変換する。"""
    if not is_estimate_item(item):
        return None
    try:
        payload = json.loads(str(item.get("content") or "{}"))
    except Exception:
        payload = {}
    if not isinstance(payload, dict):
        payload = {}
    return {
        "id": clean_value(item.get("id"), blank_text=""),
        "field_name": clean_value(item.get("field_name"), blank_text=""),
        "customer_key": clean_value(item.get("customer_key"), blank_text=""),
        "customer_name": clean_value(item.get("customer_name"), blank_text=""),
        "proposal_date": estimate_date_text(payload.get("proposal_date")),
        "product_name": clean_value(payload.get("product_name"), blank_text="").strip(),
        "manufacturer": clean_value(payload.get("manufacturer"), blank_text="").strip(),
        "unit_price": clean_value(payload.get("unit_price"), blank_text="").strip(),
        "remarks": clean_value(payload.get("remarks"), blank_text="").strip(),
        "created_at": item.get("created_at", ""),
        "updated_at": item.get("updated_at", ""),
        "sort_order": item.get("sort_order", 0),
    }


def estimate_sort_key(item):
    return (
        estimate_date_text(item.get("proposal_date")),
        str(item.get("updated_at") or item.get("created_at") or ""),
        str(item.get("id") or ""),
    )


def get_customer_estimates(customer_name, customer_key):
    items = load_customer_information(customer_name, customer_key)
    estimates = []
    for item in items:
        parsed = parse_estimate_item(item)
        if parsed:
            estimates.append(parsed)
    estimates.sort(key=estimate_sort_key, reverse=True)
    return estimates


@st.cache_data(ttl=30, show_spinner=False)
def load_all_estimates_from_supabase():
    """ホームの見積り画面で使う全顧客分の見積りを読み込む。"""
    if not has_supabase_config():
        return []

    rows = []
    page_size = 1000
    offset = 0
    while True:
        params = {
            "select": "id,customer_key,customer_name,field_name,content,sort_order,created_at,updated_at",
            "field_name": f"like.{ESTIMATE_PREFIX}*",
            "order": "created_at.desc,id.desc",
            "limit": str(page_size),
            "offset": str(offset),
        }
        try:
            response = requests.get(
                get_supabase_customer_information_url(),
                headers=get_supabase_headers(),
                params=params,
                timeout=30,
            )
        except Exception as exc:
            raise RuntimeError("見積りの読み込み中にSupabaseへ接続できませんでした。") from exc

        check_customer_information_response("読み込み", response, (200,))
        page = response.json()
        if not isinstance(page, list):
            raise RuntimeError("Supabaseから返った見積りの形式が正しくありません。")

        rows.extend(item for item in page if is_estimate_item(item))
        if len(page) < page_size:
            break
        offset += page_size

    return rows


def clear_estimate_cache():
    try:
        load_all_estimates_from_supabase.clear()
    except Exception:
        pass


def estimate_values(item):
    return {
        "提案日": estimate_date_text(item.get("proposal_date")),
        "商品名": clean_value(item.get("product_name"), blank_text=""),
        "メーカー": clean_value(item.get("manufacturer"), blank_text=""),
        "単価": clean_value(item.get("unit_price"), blank_text=""),
        "備考": clean_value(item.get("remarks"), blank_text=""),
    }


def estimate_history_changes(before, after):
    before_values = estimate_values(before or {})
    after_values = estimate_values(after or {})
    return {
        field_name: (before_values.get(field_name, ""), after_values.get(field_name, ""))
        for field_name in after_values
        if before_values.get(field_name, "") != after_values.get(field_name, "")
    }


def save_customer_estimate(
    customer_name,
    customer_key,
    proposal_date,
    product_name,
    manufacturer,
    unit_price,
    remarks,
    existing=None,
):
    content = serialize_estimate(
        proposal_date,
        product_name,
        manufacturer,
        unit_price,
        remarks,
    )
    if existing:
        update_customer_information(existing["id"], existing["field_name"], content)
    else:
        items = load_customer_information(customer_name, customer_key)
        next_order = max(
            (int(item.get("sort_order", 0)) for item in items),
            default=0,
        ) + 10
        insert_customer_information(
            customer_name,
            customer_key,
            make_estimate_field_name(),
            content,
            next_order,
        )
    clear_estimate_cache()


def delete_customer_estimate(item):
    item_id = clean_value(item.get("id"), blank_text="")
    if not item_id:
        raise RuntimeError("削除する見積りが見つかりません。")
    delete_customer_information(item_id)
    clear_estimate_cache()


def estimate_price_label(item):
    price = clean_value(item.get("unit_price"), blank_text="").strip()
    return price or "未入力"


def render_customer_estimates_section(customer_name, customer_key=None):
    """顧客詳細に、折りたたみ式の提案・見積りを表示する。"""
    identity = customer_key or customer_name
    state_suffix = hashlib.sha256(
        f"estimate|{identity}".encode("utf-8")
    ).hexdigest()[:16]
    add_key = f"estimate_add_{state_suffix}"
    edit_key = f"estimate_edit_{state_suffix}"
    delete_key = f"estimate_delete_{state_suffix}"
    success_key = f"estimate_success_{state_suffix}"

    try:
        estimates = get_customer_estimates(customer_name, customer_key)
    except Exception as exc:
        estimates = []
        load_error = str(exc)
    else:
        load_error = ""

    success_message = st.session_state.pop(success_key, None)
    expanded = bool(
        success_message
        or st.session_state.get(add_key)
        or st.session_state.get(edit_key)
        or st.session_state.get(delete_key)
    )

    with st.expander(f"📄 提案・見積り　{len(estimates)}件", expanded=expanded):
        if not has_supabase_config():
            st.warning("提案・見積りを使うにはSupabase設定が必要です。")
            return
        if load_error:
            st.warning(f"見積りを読み込めませんでした：{load_error}")
            return
        if success_message:
            st.success(success_message)

        if not st.session_state.get(add_key):
            if st.button(
                "＋ 見積りを追加",
                key=f"estimate_add_button_{state_suffix}",
                use_container_width=True,
            ):
                st.session_state[add_key] = True
                st.session_state.pop(edit_key, None)
                st.session_state.pop(delete_key, None)
                st.rerun()
        else:
            st.markdown("**新しい見積り**")
            with st.form(f"estimate_add_form_{state_suffix}"):
                proposal_date = st.date_input(
                    "提案日",
                    value=get_jst_now().date(),
                )
                product_name = st.text_input("商品名")
                manufacturer = st.text_input("メーカー")
                unit_price = st.text_input("単価", placeholder="例：85、3,500")
                remarks = st.text_area("備考", height=110)
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
                st.session_state.pop(add_key, None)
                st.rerun()
            if save:
                if not clean_value(product_name, blank_text="").strip():
                    st.warning("商品名を入力してください。")
                else:
                    after = {
                        "proposal_date": proposal_date,
                        "product_name": product_name,
                        "manufacturer": manufacturer,
                        "unit_price": unit_price,
                        "remarks": remarks,
                    }
                    try:
                        save_customer_estimate(
                            customer_name,
                            customer_key,
                            proposal_date,
                            product_name,
                            manufacturer,
                            unit_price,
                            remarks,
                        )
                        remember_change_history_warning(
                            record_change_history_safely(
                                "顧客",
                                customer_key or "",
                                customer_name,
                                "追加",
                                estimate_history_changes({}, after),
                                section="提案・見積り",
                            )
                        )
                        st.session_state.pop(add_key, None)
                        st.session_state[success_key] = "見積りを保存しました。"
                        st.rerun()
                    except Exception as exc:
                        st.error(f"見積りを保存できませんでした：{exc}")

        if not estimates:
            st.info("提案・見積りはまだありません。")
            return

        active_edit_id = st.session_state.get(edit_key)
        active_delete_id = st.session_state.get(delete_key)

        for estimate in estimates:
            estimate_id = estimate["id"]
            with st.container(border=True):
                if active_edit_id == estimate_id:
                    st.markdown("**見積りを編集**")
                    with st.form(f"estimate_edit_form_{estimate_id}"):
                        proposal_date = st.date_input(
                            "提案日",
                            value=estimate_date_input_value(estimate.get("proposal_date")),
                        )
                        product_name = st.text_input(
                            "商品名", value=estimate.get("product_name", "")
                        )
                        manufacturer = st.text_input(
                            "メーカー", value=estimate.get("manufacturer", "")
                        )
                        unit_price = st.text_input(
                            "単価", value=estimate.get("unit_price", "")
                        )
                        remarks = st.text_area(
                            "備考", value=estimate.get("remarks", ""), height=110
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
                        st.session_state.pop(edit_key, None)
                        st.rerun()
                    if save:
                        if not clean_value(product_name, blank_text="").strip():
                            st.warning("商品名を入力してください。")
                        else:
                            after = {
                                "proposal_date": proposal_date,
                                "product_name": product_name,
                                "manufacturer": manufacturer,
                                "unit_price": unit_price,
                                "remarks": remarks,
                            }
                            changes = estimate_history_changes(estimate, after)
                            if not changes:
                                st.warning("変更された項目がありません。")
                            else:
                                try:
                                    save_customer_estimate(
                                        customer_name,
                                        customer_key,
                                        proposal_date,
                                        product_name,
                                        manufacturer,
                                        unit_price,
                                        remarks,
                                        existing=estimate,
                                    )
                                    remember_change_history_warning(
                                        record_change_history_safely(
                                            "顧客",
                                            customer_key or "",
                                            customer_name,
                                            "変更",
                                            changes,
                                            section="提案・見積り",
                                        )
                                    )
                                    st.session_state.pop(edit_key, None)
                                    st.session_state[success_key] = "見積りを保存しました。"
                                    st.rerun()
                                except Exception as exc:
                                    st.error(f"見積りを保存できませんでした：{exc}")
                    continue

                st.markdown(
                    f"**{html.escape(estimate.get('product_name') or '商品名未入力')}**"
                )
                st.caption(f"提案日：{format_estimate_date(estimate.get('proposal_date'))}")
                info_col, price_col = st.columns(2)
                with info_col:
                    st.caption("メーカー")
                    st.write(estimate.get("manufacturer") or "未入力")
                with price_col:
                    st.caption("単価")
                    st.write(estimate_price_label(estimate))
                if estimate.get("remarks"):
                    st.caption("備考")
                    st.write(estimate["remarks"])

                edit_col, delete_col = st.columns(2)
                with edit_col:
                    if st.button(
                        "編集",
                        key=f"estimate_edit_button_{estimate_id}",
                        use_container_width=True,
                    ):
                        st.session_state[edit_key] = estimate_id
                        st.session_state.pop(add_key, None)
                        st.session_state.pop(delete_key, None)
                        st.rerun()
                with delete_col:
                    if active_delete_id == estimate_id:
                        st.warning("この見積りを削除しますか？")
                        confirm_col, cancel_col = st.columns(2)
                        with confirm_col:
                            if st.button(
                                "削除する",
                                key=f"estimate_delete_confirm_{estimate_id}",
                                use_container_width=True,
                            ):
                                try:
                                    delete_customer_estimate(estimate)
                                    remember_change_history_warning(
                                        record_change_history_safely(
                                            "顧客",
                                            customer_key or "",
                                            customer_name,
                                            "削除",
                                            estimate_history_changes(estimate, {}),
                                            section="提案・見積り",
                                        )
                                    )
                                    st.session_state.pop(delete_key, None)
                                    st.session_state[success_key] = "見積りを削除しました。"
                                    st.rerun()
                                except Exception as exc:
                                    st.error(f"見積りを削除できませんでした：{exc}")
                        with cancel_col:
                            if st.button(
                                "キャンセル",
                                key=f"estimate_delete_cancel_{estimate_id}",
                                use_container_width=True,
                            ):
                                st.session_state.pop(delete_key, None)
                                st.rerun()
                    elif st.button(
                        "削除",
                        key=f"estimate_delete_button_{estimate_id}",
                        use_container_width=True,
                    ):
                        st.session_state[delete_key] = estimate_id
                        st.session_state.pop(add_key, None)
                        st.session_state.pop(edit_key, None)
                        st.rerun()


def estimate_rows_to_dataframe(rows):
    records = []
    for row in rows or []:
        estimate = parse_estimate_item(row)
        if not estimate:
            continue
        records.append(
            {
                "提案日": estimate.get("proposal_date", ""),
                "顧客ID": estimate.get("customer_key", ""),
                "顧客名": estimate.get("customer_name", ""),
                "商品名": estimate.get("product_name", ""),
                "メーカー": estimate.get("manufacturer", ""),
                "単価": estimate.get("unit_price", ""),
                "備考": estimate.get("remarks", ""),
                "保存ID": estimate.get("id", ""),
                "作成日時": estimate.get("created_at", ""),
                "更新日時": estimate.get("updated_at", ""),
            }
        )
    records.sort(
        key=lambda record: (
            estimate_date_text(record.get("提案日")),
            str(record.get("更新日時") or record.get("作成日時") or ""),
        ),
        reverse=True,
    )
    return backup_dataframe(
        records,
        [
            "提案日", "顧客ID", "顧客名", "商品名", "メーカー",
            "単価", "備考", "保存ID", "作成日時", "更新日時",
        ],
    )


def show_estimates_page():
    st.header("📄 提案・見積り")
    st.caption("全顧客の提案・見積りを、提案日の新しい順に表示します。")

    if not has_supabase_config():
        st.warning("提案・見積りを使うにはSupabase設定が必要です。")
        return

    try:
        rows = load_all_estimates_from_supabase()
    except Exception as exc:
        st.error(str(exc))
        return

    estimates = []
    for row in rows:
        parsed = parse_estimate_item(row)
        if parsed:
            estimates.append(parsed)
    estimates.sort(key=estimate_sort_key, reverse=True)

    if not estimates:
        st.info("提案・見積りはまだありません。")
        return

    st.write(f"見積り：{len(estimates)}件")
    for estimate in estimates:
        with st.container(border=True):
            customer_name = estimate.get("customer_name") or "顧客名未設定"
            customer_link = build_customer_detail_link(
                customer_name,
                class_name="dispatch-month-link",
            )
            st.markdown(customer_link, unsafe_allow_html=True)
            st.markdown(
                f"**{html.escape(estimate.get('product_name') or '商品名未入力')}**"
            )
            st.caption(f"提案日：{format_estimate_date(estimate.get('proposal_date'))}")
            info_col, price_col = st.columns(2)
            with info_col:
                st.caption("メーカー")
                st.write(estimate.get("manufacturer") or "未入力")
            with price_col:
                st.caption("単価")
                st.write(estimate_price_label(estimate))
            if estimate.get("remarks"):
                st.caption("備考")
                st.write(estimate["remarks"])


# =========================
# 運送会社の運賃登録・比較（Supabase保存）
# =========================
def make_carrier_freight_field_name():
    """顧客情報テーブル内で、運送会社の運賃を通常項目と分ける内部項目名を作る。"""
    return f"{CARRIER_FREIGHT_PREFIX}{uuid.uuid4()}"


def carrier_freight_storage_key(carrier_id):
    return f"carrier_freight:{clean_value(carrier_id, blank_text='').strip()}"


def is_carrier_freight_item(item):
    """顧客情報テーブル上の運送会社運賃専用レコードか判定する。"""
    return clean_value(item.get("field_name"), blank_text="").startswith(CARRIER_FREIGHT_PREFIX)


def carrier_freight_date_text(value):
    """運賃の適用日をYYYY-MM-DDへそろえる。"""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = clean_value(value, blank_text="").strip()
    match = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", text)
    if not match:
        return text
    year, month, day = (int(part) for part in match.groups())
    try:
        return date(year, month, day).isoformat()
    except ValueError:
        return text


def carrier_freight_date_input_value(value):
    text = carrier_freight_date_text(value)
    try:
        return date.fromisoformat(text)
    except (TypeError, ValueError):
        return get_jst_now().date()


def format_carrier_freight_date(value):
    text = carrier_freight_date_text(value)
    try:
        return date.fromisoformat(text).strftime("%Y/%m/%d")
    except (TypeError, ValueError):
        return text or "未入力"


def parse_carrier_freight_number(value, label):
    """任意入力の正数をDecimalへ変換する。空欄はNone。"""
    text = unicodedata.normalize("NFKC", clean_value(value, blank_text="")).strip()
    if not text:
        return None
    text = text.replace(",", "").replace("，", "").replace(" ", "")
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        raise ValueError(f"{label}は数字で入力してください。")
    try:
        number = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"{label}は数字で入力してください。") from exc
    if number <= 0:
        raise ValueError(f"{label}は0より大きい数字で入力してください。")
    return number


def carrier_freight_decimal_text(value):
    if value is None:
        return ""
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return "0" if text in {"", "-0"} else text


def normalize_carrier_freight_amounts(truck_freight, quantity_kg, kg_rate):
    """確定した計算ルールに従い、運賃・数量・kg単価を整合させる。"""
    truck = parse_carrier_freight_number(truck_freight, "1車運賃")
    quantity = parse_carrier_freight_number(quantity_kg, "数量")
    rate = parse_carrier_freight_number(kg_rate, "kg単価")
    calculation_source = ""

    if quantity is not None and truck is not None:
        # 3項目すべて入力された場合も、1車運賃と数量を正としてkg単価を計算する。
        rate = (truck / quantity).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        calculation_source = "kg単価を自動計算"
    elif quantity is not None and rate is not None:
        truck = (quantity * rate).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
        calculation_source = "1車運賃を自動計算"

    if truck is None and rate is None:
        raise ValueError("1車運賃またはkg単価のどちらかを入力してください。")

    return {
        "truck_freight": carrier_freight_decimal_text(truck),
        "quantity_kg": carrier_freight_decimal_text(quantity),
        "kg_rate": carrier_freight_decimal_text(rate),
        "calculation_source": calculation_source,
    }


def carrier_freight_route_key(value):
    """比較時の表記揺れを減らすため、全半角と空白をそろえる。"""
    text = unicodedata.normalize("NFKC", clean_value(value, blank_text="")).strip().lower()
    return re.sub(r"\s+", "", text)


def serialize_carrier_freight(record):
    payload = {
        "version": CARRIER_FREIGHT_VERSION,
        "carrier_id": clean_value(record.get("carrier_id"), blank_text="").strip(),
        "carrier_name": clean_value(record.get("carrier_name"), blank_text="").strip(),
        "effective_date": carrier_freight_date_text(record.get("effective_date")),
        "pickup_location": clean_value(record.get("pickup_location"), blank_text="").strip(),
        "delivery_destination": clean_value(record.get("delivery_destination"), blank_text="").strip(),
        "truck_freight": clean_value(record.get("truck_freight"), blank_text="").strip(),
        "quantity_kg": clean_value(record.get("quantity_kg"), blank_text="").strip(),
        "kg_rate": clean_value(record.get("kg_rate"), blank_text="").strip(),
        "calculation_source": clean_value(record.get("calculation_source"), blank_text="").strip(),
        "remarks": clean_value(record.get("remarks"), blank_text="").strip(),
    }
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def parse_carrier_freight_item(item):
    """Supabase内部レコードを画面表示用の運賃辞書へ変換する。"""
    if not is_carrier_freight_item(item):
        return None
    try:
        payload = json.loads(str(item.get("content") or "{}"))
    except Exception:
        payload = {}
    if not isinstance(payload, dict):
        payload = {}

    carrier_id = clean_value(payload.get("carrier_id"), blank_text="").strip()
    if not carrier_id:
        storage_key = clean_value(item.get("customer_key"), blank_text="")
        prefix = "carrier_freight:"
        if storage_key.startswith(prefix):
            carrier_id = storage_key[len(prefix):].strip()

    return {
        "id": clean_value(item.get("id"), blank_text=""),
        "field_name": clean_value(item.get("field_name"), blank_text=""),
        "carrier_id": carrier_id,
        "carrier_name": clean_value(
            payload.get("carrier_name") or item.get("customer_name"),
            blank_text="",
        ).strip(),
        "effective_date": carrier_freight_date_text(payload.get("effective_date")),
        "pickup_location": clean_value(payload.get("pickup_location"), blank_text="").strip(),
        "delivery_destination": clean_value(payload.get("delivery_destination"), blank_text="").strip(),
        "truck_freight": clean_value(payload.get("truck_freight"), blank_text="").strip(),
        "quantity_kg": clean_value(payload.get("quantity_kg"), blank_text="").strip(),
        "kg_rate": clean_value(payload.get("kg_rate"), blank_text="").strip(),
        "calculation_source": clean_value(payload.get("calculation_source"), blank_text="").strip(),
        "remarks": clean_value(payload.get("remarks"), blank_text="").strip(),
        "created_at": item.get("created_at", ""),
        "updated_at": item.get("updated_at", ""),
        "sort_order": item.get("sort_order", 0),
    }


def carrier_freight_sort_key(item):
    return (
        carrier_freight_date_text(item.get("effective_date")),
        str(item.get("updated_at") or item.get("created_at") or ""),
        str(item.get("id") or ""),
    )


@st.cache_data(ttl=30, show_spinner=False)
def load_carrier_freight_rows_from_supabase(carrier_id=""):
    """運送会社運賃をSupabaseからページ単位で読み込む。"""
    if not has_supabase_config():
        return []

    rows = []
    page_size = 1000
    offset = 0
    while True:
        params = {
            "select": "id,customer_key,customer_name,field_name,content,sort_order,created_at,updated_at",
            "field_name": f"like.{CARRIER_FREIGHT_PREFIX}*",
            "order": "created_at.desc,id.desc",
            "limit": str(page_size),
            "offset": str(offset),
        }
        if carrier_id:
            params["customer_key"] = f"eq.{carrier_freight_storage_key(carrier_id)}"
        try:
            response = requests.get(
                get_supabase_customer_information_url(),
                headers=get_supabase_headers(),
                params=params,
                timeout=30,
            )
        except Exception as exc:
            raise RuntimeError("運賃の読み込み中にSupabaseへ接続できませんでした。") from exc

        check_customer_information_response("読み込み", response, (200,))
        page = response.json()
        if not isinstance(page, list):
            raise RuntimeError("Supabaseから返った運賃の形式が正しくありません。")

        rows.extend(item for item in page if is_carrier_freight_item(item))
        if len(page) < page_size:
            break
        offset += page_size

    return rows


def get_carrier_freights(carrier_id):
    freights = []
    for item in load_carrier_freight_rows_from_supabase(carrier_id):
        parsed = parse_carrier_freight_item(item)
        if parsed:
            freights.append(parsed)
    freights.sort(key=carrier_freight_sort_key, reverse=True)
    return freights


def get_all_carrier_freights():
    freights = []
    for item in load_carrier_freight_rows_from_supabase(""):
        parsed = parse_carrier_freight_item(item)
        if parsed:
            freights.append(parsed)
    freights.sort(key=carrier_freight_sort_key, reverse=True)
    return freights


def clear_carrier_freight_cache():
    try:
        load_carrier_freight_rows_from_supabase.clear()
    except Exception:
        pass


def carrier_freight_values(item):
    return {
        "適用日": carrier_freight_date_text(item.get("effective_date")),
        "引取場所": clean_value(item.get("pickup_location"), blank_text=""),
        "納品先": clean_value(item.get("delivery_destination"), blank_text=""),
        "1車運賃": clean_value(item.get("truck_freight"), blank_text=""),
        "数量kg": clean_value(item.get("quantity_kg"), blank_text=""),
        "kg単価": clean_value(item.get("kg_rate"), blank_text=""),
        "備考": clean_value(item.get("remarks"), blank_text=""),
    }


def carrier_freight_history_changes(before, after):
    before_values = carrier_freight_values(before or {})
    after_values = carrier_freight_values(after or {})
    return {
        field_name: (before_values.get(field_name, ""), after_values.get(field_name, ""))
        for field_name in after_values
        if before_values.get(field_name, "") != after_values.get(field_name, "")
    }


def build_carrier_freight_record(
    carrier_id,
    carrier_name,
    effective_date,
    pickup_location,
    delivery_destination,
    truck_freight,
    quantity_kg,
    kg_rate,
    remarks,
):
    pickup = clean_value(pickup_location, blank_text="").strip()
    destination = clean_value(delivery_destination, blank_text="").strip()
    if not pickup:
        raise ValueError("引取場所を入力してください。")
    if not destination:
        raise ValueError("納品先を入力してください。")

    amounts = normalize_carrier_freight_amounts(truck_freight, quantity_kg, kg_rate)
    return {
        "carrier_id": clean_value(carrier_id, blank_text="").strip(),
        "carrier_name": clean_value(carrier_name, blank_text="").strip(),
        "effective_date": carrier_freight_date_text(effective_date),
        "pickup_location": pickup,
        "delivery_destination": destination,
        "truck_freight": amounts["truck_freight"],
        "quantity_kg": amounts["quantity_kg"],
        "kg_rate": amounts["kg_rate"],
        "calculation_source": amounts["calculation_source"],
        "remarks": clean_value(remarks, blank_text="").strip(),
    }


def save_carrier_freight(record, existing=None):
    content = serialize_carrier_freight(record)
    carrier_id = record["carrier_id"]
    if existing:
        update_customer_information(existing["id"], existing["field_name"], content)
    else:
        rows = load_carrier_freight_rows_from_supabase(carrier_id)
        next_order = max(
            (int(row.get("sort_order", 0)) for row in rows),
            default=0,
        ) + 10
        insert_customer_information(
            record["carrier_name"],
            carrier_freight_storage_key(carrier_id),
            make_carrier_freight_field_name(),
            content,
            next_order,
        )
    clear_carrier_freight_cache()


def delete_carrier_freight(item):
    item_id = clean_value(item.get("id"), blank_text="")
    if not item_id:
        raise RuntimeError("削除する運賃が見つかりません。")
    delete_customer_information(item_id)
    clear_carrier_freight_cache()


def carrier_freight_decimal_display(value, maximum_decimals=2):
    text = clean_value(value, blank_text="").strip()
    if not text:
        return ""
    try:
        number = Decimal(text)
    except InvalidOperation:
        return text
    if maximum_decimals == 0:
        return f"{number:,.0f}"
    formatted = f"{number:,.{maximum_decimals}f}"
    return formatted.rstrip("0").rstrip(".")


def carrier_freight_truck_label(item):
    value = carrier_freight_decimal_display(item.get("truck_freight"), 0)
    return f"{value}円" if value else "未入力"


def carrier_freight_quantity_label(item):
    value = carrier_freight_decimal_display(item.get("quantity_kg"), 2)
    return f"{value}kg" if value else "未入力"


def carrier_freight_rate_label(item):
    value = carrier_freight_decimal_display(item.get("kg_rate"), 4)
    return f"{value}円/kg" if value else "未入力"


def carrier_freight_success_message(action, record):
    message = f"運賃を{action}しました。"
    calculation = clean_value(record.get("calculation_source"), blank_text="")
    if calculation:
        message += f" {calculation}しました。"
    return message


def render_carrier_freight_display(item):
    route = (
        f"{clean_value(item.get('pickup_location'), blank_text='未入力')}"
        f" → {clean_value(item.get('delivery_destination'), blank_text='未入力')}"
    )
    st.markdown(f"**{html.escape(route)}**")
    st.caption(f"適用日：{format_carrier_freight_date(item.get('effective_date'))}")
    truck_col, quantity_col, rate_col = st.columns(3)
    with truck_col:
        st.caption("1車運賃")
        st.write(carrier_freight_truck_label(item))
    with quantity_col:
        st.caption("数量")
        st.write(carrier_freight_quantity_label(item))
    with rate_col:
        st.caption("kg単価")
        st.write(carrier_freight_rate_label(item))
    if item.get("remarks"):
        st.caption("備考")
        st.write(item["remarks"])


def render_carrier_freight_form(form_key, existing=None):
    existing = existing or {}
    with st.form(form_key):
        effective_date = st.date_input(
            "適用日",
            value=carrier_freight_date_input_value(existing.get("effective_date")),
            key=f"{form_key}_effective_date",
        )
        pickup_location = st.text_input(
            "引取場所",
            value=existing.get("pickup_location", ""),
            placeholder="例：○○工場",
            key=f"{form_key}_pickup_location",
        )
        delivery_destination = st.text_input(
            "納品先",
            value=existing.get("delivery_destination", ""),
            placeholder="例：△△牧場",
            key=f"{form_key}_delivery_destination",
        )
        truck_col, quantity_col, rate_col = st.columns(3)
        with truck_col:
            truck_freight = st.text_input(
                "1車運賃（円）",
                value=existing.get("truck_freight", ""),
                placeholder="例：200000",
                key=f"{form_key}_truck_freight",
            )
        with quantity_col:
            quantity_kg = st.text_input(
                "数量（kg）",
                value=existing.get("quantity_kg", ""),
                placeholder="例：20000",
                key=f"{form_key}_quantity_kg",
            )
        with rate_col:
            kg_rate = st.text_input(
                "kg単価（円）",
                value=existing.get("kg_rate", ""),
                placeholder="例：10",
                key=f"{form_key}_kg_rate",
            )
        st.caption(
            "数量と1車運賃がある場合はkg単価を自動計算します。"
            "数量とkg単価があり1車運賃が空欄の場合は、1車運賃を自動計算します。"
        )
        remarks = st.text_area(
            "備考",
            value=existing.get("remarks", ""),
            height=100,
            placeholder="例：高速代込み、冬季料金 など",
            key=f"{form_key}_remarks",
        )
        save_col, cancel_col = st.columns(2)
        with save_col:
            submitted = st.form_submit_button(
                "自動計算して保存",
                type="primary",
                use_container_width=True,
            )
        with cancel_col:
            cancelled = st.form_submit_button("キャンセル", use_container_width=True)
    return submitted, cancelled, {
        "effective_date": effective_date,
        "pickup_location": pickup_location,
        "delivery_destination": delivery_destination,
        "truck_freight": truck_freight,
        "quantity_kg": quantity_kg,
        "kg_rate": kg_rate,
        "remarks": remarks,
    }


def render_carrier_freight_section(carrier_id, company_name):
    """運送会社詳細に、折りたたみ式の運賃登録・履歴を表示する。"""
    state_suffix = hashlib.sha256(
        f"carrier-freight|{carrier_id}".encode("utf-8")
    ).hexdigest()[:16]
    add_key = f"carrier_freight_add_{state_suffix}"
    edit_key = f"carrier_freight_edit_{state_suffix}"
    delete_key = f"carrier_freight_delete_{state_suffix}"
    success_key = f"carrier_freight_success_{state_suffix}"

    try:
        freights = get_carrier_freights(carrier_id)
    except Exception as exc:
        freights = []
        load_error = str(exc)
    else:
        load_error = ""

    success_message = st.session_state.pop(success_key, None)
    expanded = bool(
        success_message
        or st.session_state.get(add_key)
        or st.session_state.get(edit_key)
        or st.session_state.get(delete_key)
    )

    st.markdown("---")
    with st.expander(f"💰 運賃登録・履歴　{len(freights)}件", expanded=expanded):
        if success_message:
            st.success(success_message)
        if load_error:
            st.warning(load_error)
            return
        if not has_supabase_config():
            st.warning("運賃登録を使うにはSupabase設定が必要です。")
            return

        st.caption(
            "新しい運賃は既存記録を上書きせず追加し、過去分を履歴として残します。"
            "既存記録の編集・削除は入力ミスの訂正用です。"
        )

        if not st.session_state.get(add_key):
            if st.button(
                "＋ 運賃を追加",
                key=f"carrier_freight_add_button_{state_suffix}",
                use_container_width=True,
            ):
                st.session_state[add_key] = True
                st.session_state.pop(edit_key, None)
                st.session_state.pop(delete_key, None)
                st.rerun()
        else:
            st.markdown("#### 新しい運賃")
            submitted, cancelled, values = render_carrier_freight_form(
                f"carrier_freight_add_form_{state_suffix}"
            )
            if cancelled:
                st.session_state.pop(add_key, None)
                st.rerun()
            if submitted:
                try:
                    record = build_carrier_freight_record(
                        carrier_id,
                        company_name,
                        **values,
                    )
                    save_carrier_freight(record)
                    remember_change_history_warning(
                        record_change_history_safely(
                            "運送会社",
                            carrier_id,
                            company_name,
                            "追加",
                            carrier_freight_history_changes({}, record),
                            section="運賃",
                        )
                    )
                    st.session_state.pop(add_key, None)
                    st.session_state[success_key] = carrier_freight_success_message("追加", record)
                    st.rerun()
                except Exception as exc:
                    st.error(f"運賃を保存できませんでした：{exc}")

        if not freights:
            st.info("登録されている運賃はありません。")
            return

        st.markdown("#### 運賃履歴")
        active_edit_id = st.session_state.get(edit_key)
        active_delete_id = st.session_state.get(delete_key)

        for freight in freights:
            freight_id = freight["id"]
            with st.container(border=True):
                if active_edit_id == freight_id:
                    st.markdown("**入力ミスを訂正**")
                    submitted, cancelled, values = render_carrier_freight_form(
                        f"carrier_freight_edit_form_{freight_id}",
                        existing=freight,
                    )
                    if cancelled:
                        st.session_state.pop(edit_key, None)
                        st.rerun()
                    if submitted:
                        try:
                            record = build_carrier_freight_record(
                                carrier_id,
                                company_name,
                                **values,
                            )
                            changes = carrier_freight_history_changes(freight, record)
                            if changes:
                                save_carrier_freight(record, existing=freight)
                                remember_change_history_warning(
                                    record_change_history_safely(
                                        "運送会社",
                                        carrier_id,
                                        company_name,
                                        "変更",
                                        changes,
                                        section="運賃",
                                    )
                                )
                            st.session_state.pop(edit_key, None)
                            st.session_state[success_key] = carrier_freight_success_message("保存", record)
                            st.rerun()
                        except Exception as exc:
                            st.error(f"運賃を保存できませんでした：{exc}")
                    continue

                render_carrier_freight_display(freight)
                edit_col, delete_col = st.columns(2)
                with edit_col:
                    if st.button(
                        "編集",
                        key=f"carrier_freight_edit_button_{freight_id}",
                        use_container_width=True,
                    ):
                        st.session_state[edit_key] = freight_id
                        st.session_state.pop(add_key, None)
                        st.session_state.pop(delete_key, None)
                        st.rerun()
                with delete_col:
                    if active_delete_id == freight_id:
                        st.warning("この運賃記録を削除しますか？")
                        if st.button(
                            "削除する",
                            key=f"carrier_freight_delete_confirm_{freight_id}",
                            use_container_width=True,
                        ):
                            try:
                                delete_carrier_freight(freight)
                                remember_change_history_warning(
                                    record_change_history_safely(
                                        "運送会社",
                                        carrier_id,
                                        company_name,
                                        "削除",
                                        carrier_freight_history_changes(freight, {}),
                                        section="運賃",
                                    )
                                )
                                st.session_state.pop(delete_key, None)
                                st.session_state[success_key] = "運賃を削除しました。"
                                st.rerun()
                            except Exception as exc:
                                st.error(f"運賃を削除できませんでした：{exc}")
                        if st.button(
                            "キャンセル",
                            key=f"carrier_freight_delete_cancel_{freight_id}",
                            use_container_width=True,
                        ):
                            st.session_state.pop(delete_key, None)
                            st.rerun()
                    elif st.button(
                        "削除",
                        key=f"carrier_freight_delete_button_{freight_id}",
                        use_container_width=True,
                    ):
                        st.session_state[delete_key] = freight_id
                        st.session_state.pop(edit_key, None)
                        st.rerun()


def carrier_freight_numeric_value(item, field_name):
    text = clean_value(item.get(field_name), blank_text="").strip()
    if not text:
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def render_carrier_freight_ranking(title, records, field_name, current_names):
    available = [
        record for record in records
        if carrier_freight_numeric_value(record, field_name) is not None
    ]
    available.sort(key=lambda record: carrier_freight_numeric_value(record, field_name))

    st.subheader(title)
    if not available:
        st.info("比較できる運賃はありません。")
        return

    for index, record in enumerate(available):
        with st.container(border=True):
            carrier_id = record.get("carrier_id", "")
            company_name = (
                current_names.get(carrier_id)
                or record.get("carrier_name")
                or "運送会社名未設定"
            )
            company_link = render_page_link(
                company_name,
                page="partner_detail",
                partner_id=carrier_id,
                partner_type="carrier",
                class_name="dispatch-month-link",
            )
            if index == 0:
                st.markdown(f"**最安**　{company_link}", unsafe_allow_html=True)
            else:
                st.markdown(company_link, unsafe_allow_html=True)

            if field_name == "truck_freight":
                st.markdown(f"### {carrier_freight_truck_label(record)}")
            else:
                st.markdown(f"### {carrier_freight_rate_label(record)}")

            st.caption(f"適用日：{format_carrier_freight_date(record.get('effective_date'))}")
            detail_col1, detail_col2 = st.columns(2)
            with detail_col1:
                st.caption("1車運賃")
                st.write(carrier_freight_truck_label(record))
            with detail_col2:
                st.caption("数量・kg単価")
                st.write(
                    f"{carrier_freight_quantity_label(record)} ／ "
                    f"{carrier_freight_rate_label(record)}"
                )
            if record.get("remarks"):
                st.caption("備考")
                st.write(record["remarks"])


def show_carrier_freight_compare():
    """同じ引取場所・納品先の最新運賃を運送会社別に比較する。"""
    show_trade_partner_home_link("carrier")
    st.header("💰 運賃比較")
    st.caption(
        "同じ引取場所と納品先について、各運送会社の適用日が最も新しい記録を比較します。"
    )

    if not has_supabase_config():
        st.warning("運賃比較を使うにはSupabase設定が必要です。")
        return

    try:
        freights = get_all_carrier_freights()
    except Exception as exc:
        st.error(str(exc))
        return

    if not freights:
        st.info("比較できる運賃はまだ登録されていません。")
        return

    # 新しい記録の表記を候補名として優先する。
    pickup_names = {}
    for record in freights:
        pickup_key = carrier_freight_route_key(record.get("pickup_location"))
        if pickup_key:
            pickup_names.setdefault(pickup_key, record.get("pickup_location", ""))

    pickup_keys = sorted(pickup_names, key=lambda key: pickup_names[key])
    selected_pickup_key = st.selectbox(
        "引取場所",
        pickup_keys,
        format_func=lambda key: pickup_names.get(key, key),
        key="carrier_freight_compare_pickup",
    )

    destination_names = {}
    for record in freights:
        if carrier_freight_route_key(record.get("pickup_location")) != selected_pickup_key:
            continue
        destination_key = carrier_freight_route_key(record.get("delivery_destination"))
        if destination_key:
            destination_names.setdefault(
                destination_key,
                record.get("delivery_destination", ""),
            )

    destination_keys = sorted(destination_names, key=lambda key: destination_names[key])
    selected_destination_key = st.selectbox(
        "納品先",
        destination_keys,
        format_func=lambda key: destination_names.get(key, key),
        key="carrier_freight_compare_destination",
    )

    route_records = [
        record for record in freights
        if carrier_freight_route_key(record.get("pickup_location")) == selected_pickup_key
        and carrier_freight_route_key(record.get("delivery_destination")) == selected_destination_key
    ]
    route_records.sort(key=carrier_freight_sort_key, reverse=True)

    latest_by_carrier = {}
    for record in route_records:
        carrier_id = clean_value(record.get("carrier_id"), blank_text="")
        if carrier_id and carrier_id not in latest_by_carrier:
            latest_by_carrier[carrier_id] = record
    latest_records = list(latest_by_carrier.values())

    try:
        partner_data = load_trade_partner_data()
        current_names = {
            trade_partner_text(row.get("取引先ID")): trade_partner_text(row.get("会社名"))
            for row in get_trade_partner_master_rows(partner_data, "carrier")
        }
    except Exception:
        current_names = {}

    st.markdown("---")
    st.markdown(
        f"**{html.escape(pickup_names[selected_pickup_key])}"
        f" → {html.escape(destination_names[selected_destination_key])}**"
    )
    st.caption(f"比較対象：{len(latest_records)}社（各社の最新記録）")

    render_carrier_freight_ranking(
        "1車運賃が安い順",
        latest_records,
        "truck_freight",
        current_names,
    )
    st.markdown("---")
    render_carrier_freight_ranking(
        "kg単価が安い順",
        latest_records,
        "kg_rate",
        current_names,
    )


def carrier_freight_rows_to_dataframe(rows):
    records = []
    for row in rows or []:
        freight = parse_carrier_freight_item(row)
        if not freight:
            continue
        records.append(
            {
                "運送会社ID": freight.get("carrier_id", ""),
                "運送会社": freight.get("carrier_name", ""),
                "適用日": freight.get("effective_date", ""),
                "引取場所": freight.get("pickup_location", ""),
                "納品先": freight.get("delivery_destination", ""),
                "1車運賃": freight.get("truck_freight", ""),
                "数量kg": freight.get("quantity_kg", ""),
                "kg単価": freight.get("kg_rate", ""),
                "計算方法": freight.get("calculation_source", ""),
                "備考": freight.get("remarks", ""),
                "保存ID": freight.get("id", ""),
                "作成日時": freight.get("created_at", ""),
                "更新日時": freight.get("updated_at", ""),
            }
        )
    records.sort(
        key=lambda record: (
            carrier_freight_date_text(record.get("適用日")),
            str(record.get("更新日時") or record.get("作成日時") or ""),
        ),
        reverse=True,
    )
    return backup_dataframe(
        records,
        [
            "運送会社ID", "運送会社", "適用日", "引取場所", "納品先",
            "1車運賃", "数量kg", "kg単価", "計算方法", "備考",
            "保存ID", "作成日時", "更新日時",
        ],
    )


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
            remember_change_history_warning(
                record_change_history_safely(
                    "顧客",
                    customer_key or "",
                    customer_name,
                    "追加",
                    {str(field_name).strip(): ("", content)},
                    section="顧客情報",
                )
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
            items = [
                item for item in items
                if not is_past_product_note_item(item)
                and not is_estimate_item(item)
                and not is_carrier_freight_item(item)
                and not is_onedrive_attachment_item(item)
            ]
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
                            history_changes = {}
                            if str(edited_name).strip() != field_name:
                                history_changes["項目名"] = (field_name, str(edited_name).strip())
                            if str(edited_content) != content:
                                history_changes[str(edited_name).strip() or field_name] = (
                                    content,
                                    edited_content,
                                )
                            remember_change_history_warning(
                                record_change_history_safely(
                                    "顧客",
                                    customer_key or "",
                                    customer_name,
                                    "変更",
                                    history_changes,
                                    section="顧客情報",
                                )
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
                            remember_change_history_warning(
                                record_change_history_safely(
                                    "顧客",
                                    customer_key or "",
                                    customer_name,
                                    "削除",
                                    {field_name: (content, "")},
                                    section="顧客情報",
                                )
                            )
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
                            other_item = items[index - 1]
                            reorder_customer_information(item, other_item)
                            remember_change_history_warning(
                                record_change_history_safely(
                                    "顧客",
                                    customer_key or "",
                                    customer_name,
                                    "並び替え",
                                    {
                                        "表示順": (
                                            f"{clean_value(other_item.get('field_name'), blank_text='')} → {field_name}",
                                            f"{field_name} → {clean_value(other_item.get('field_name'), blank_text='')}",
                                        )
                                    },
                                    section="顧客情報",
                                )
                            )
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
                            other_item = items[index + 1]
                            reorder_customer_information(item, other_item)
                            remember_change_history_warning(
                                record_change_history_safely(
                                    "顧客",
                                    customer_key or "",
                                    customer_name,
                                    "並び替え",
                                    {
                                        "表示順": (
                                            f"{field_name} → {clean_value(other_item.get('field_name'), blank_text='')}",
                                            f"{clean_value(other_item.get('field_name'), blank_text='')} → {field_name}",
                                        )
                                    },
                                    section="顧客情報",
                                )
                            )
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
def calculate_delivery_values(delivery_row_values):
    """
    最新の入力値から、次回配達予定と残数をアプリ表示用に計算する。

    Excel内の数式セルは変更しない。PCでExcelを編集した場合も、保存された
    配達日・配達数量・使用数量/日などの最新入力値を読み、同じ計算を行う。
    """
    def column_value(column_number):
        index = column_number - 1
        return delivery_row_values[index] if index < len(delivery_row_values) else None

    usage = column_value(7)
    kg_per_bottle = column_value(9)
    delivery_date = column_value(10)
    stored_delivery_quantity = column_value(11)
    remaining = column_value(15)

    # ExcelのL列と同じく、K列「配達数量」を使って次回配達予定を計算する。
    effective_delivery_quantity = stored_delivery_quantity

    next_delivery = None
    try:
        if delivery_date is not None:
            next_delivery = delivery_date + timedelta(
                days=math.floor(float(effective_delivery_quantity) / float(usage))
            )
    except Exception:
        next_delivery = None

    if isinstance(remaining, str) and remaining.startswith("="):
        remaining = None

    try:
        if remaining is None and next_delivery is not None:
            target_date = next_delivery.date() if isinstance(next_delivery, datetime) else next_delivery
            remaining = (target_date - date.today()).days * float(usage) / float(kg_per_bottle)
    except Exception:
        remaining = None

    return next_delivery, remaining


def rebuild_sheet1_from_formula_references(excel_source):
    """数式参照元を1回だけ走査し、顧客検索用のSheet1相当データを復元する。"""
    if isinstance(excel_source, BytesIO):
        content = excel_source.getvalue()
    else:
        content = Path(excel_source).read_bytes()

    # 顧客検索時に同じExcelを二重に開かない。read_onlyシートはcell()で
    # 飛び飛びに読むと非常に遅いため、iter_rows()で各シートを1回だけ走査する。
    workbook = load_workbook(
        BytesIO(content),
        keep_vba=True,
        data_only=False,
        read_only=True,
    )
    try:
        if SHEET_NAME not in workbook.sheetnames or DELIVERY_SHEET_NAME not in workbook.sheetnames:
            return pd.DataFrame()

        sheet1 = workbook[SHEET_NAME]
        delivery = workbook[DELIVERY_SHEET_NAME]

        sheet1_records = []
        sheet1_max_column = max(
            2,
            SHEET1_HIRAGANA_COLUMN,
            SHEET1_ADDRESS_COLUMN,
            SHEET1_MAP_COLUMN,
        )
        for values in sheet1.iter_rows(
            min_row=2,
            max_col=sheet1_max_column,
            values_only=True,
        ):
            source_row = None
            for formula in values[:2]:
                if isinstance(formula, str) and formula.startswith("="):
                    match = re.search(r"(\d+)\s*$", formula.strip())
                    if match:
                        source_row = int(match.group(1))
                        break
            if source_row is None:
                continue

            sheet1_records.append(
                {
                    "source_row": source_row,
                    "ひらがな": values[SHEET1_HIRAGANA_COLUMN - 1]
                    if len(values) >= SHEET1_HIRAGANA_COLUMN
                    else None,
                    "住所": values[SHEET1_ADDRESS_COLUMN - 1]
                    if len(values) >= SHEET1_ADDRESS_COLUMN
                    else None,
                    "マップ位置": values[SHEET1_MAP_COLUMN - 1]
                    if len(values) >= SHEET1_MAP_COLUMN
                    else None,
                }
            )

        if not sheet1_records:
            return pd.DataFrame()

        required_source_rows = {record["source_row"] for record in sheet1_records}
        delivery_rows = {}
        for row_number, values in enumerate(
            delivery.iter_rows(min_row=1, max_col=15, values_only=True),
            start=1,
        ):
            if row_number in required_source_rows:
                delivery_rows[row_number] = values
                if len(delivery_rows) == len(required_source_rows):
                    break

        rows = []
        for sheet1_record in sheet1_records:
            source_row = sheet1_record["source_row"]
            delivery_values = delivery_rows.get(source_row)
            if delivery_values is None:
                continue

            customer_name = delivery_values[1] if len(delivery_values) >= 2 else None
            product_name = delivery_values[4] if len(delivery_values) >= 5 else None
            if not normalize_match_value(customer_name) or not normalize_match_value(product_name):
                continue

            next_delivery, remaining = calculate_delivery_values(delivery_values)
            rows.append(
                {
                    "ID": delivery_values[0] if len(delivery_values) >= 1 else None,
                    "顧客名": customer_name,
                    "地域": delivery_values[2] if len(delivery_values) >= 3 else None,
                    "商品名": product_name,
                    "使用数量/日": delivery_values[6] if len(delivery_values) >= 7 else None,
                    "次回配達予定": next_delivery,
                    "残数": remaining,
                    "ひらがな": sheet1_record["ひらがな"],
                    "住所": sheet1_record["住所"],
                    "マップ位置": sheet1_record["マップ位置"],
                    "メーカー": delivery_values[5] if len(delivery_values) >= 6 else None,
                    "本数": delivery_values[7] if len(delivery_values) >= 8 else None,
                    "kg/本": delivery_values[8] if len(delivery_values) >= 9 else None,
                    "配達日": delivery_values[9] if len(delivery_values) >= 10 else None,
                }
            )
        return pd.DataFrame(rows)
    finally:
        workbook.close()

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



def make_app_url(
    page="home",
    customer=None,
    customer_search=None,
    region_search=None,
    product_search=None,
    partner_id=None,
    partner_type=None,
    partner_search=None,
):
    """ブラウザの戻るボタンで戻れるように、通常リンク用URLを作る。"""
    params = {"logged_in": "1", "page": page}
    if customer:
        params["customer"] = str(customer)
    if customer_search:
        params["customer_search"] = str(customer_search)
    if region_search:
        params["region_search"] = str(region_search)
    if product_search:
        params["product_search"] = str(product_search)
    if partner_id:
        params["partner_id"] = str(partner_id)
    if partner_type:
        params["partner_type"] = str(partner_type)
    if partner_search:
        params["partner_search"] = str(partner_search)
    return "?" + urllib.parse.urlencode(params)


def render_page_link(
    label,
    page="home",
    customer=None,
    customer_search=None,
    region_search=None,
    product_search=None,
    partner_id=None,
    partner_type=None,
    partner_search=None,
    class_name="app-nav-link",
):
    """st.buttonではなくHTMLリンクで画面遷移する。これによりブラウザ戻るが効く。"""
    url = make_app_url(
        page=page,
        customer=customer,
        customer_search=customer_search,
        region_search=region_search,
        product_search=product_search,
        partner_id=partner_id,
        partner_type=partner_type,
        partner_search=partner_search,
    )
    return f'<a class="{class_name}" href="{url}" target="_self">{html.escape(str(label))}</a>'

def sync_page_from_query_params():
    """URLの画面情報を読み、ブラウザ戻る・進むに追従する。"""
    page = str(get_query_value("page", "home")).strip() or "home"
    customer = str(get_query_value("customer", "")).strip()
    partner_id = str(get_query_value("partner_id", "")).strip()
    partner_type = str(get_query_value("partner_type", "")).strip()

    valid_pages = {
        "home",
        "customer_home",
        "customer_list",
        "customer",
        "region",
        "product",
        "calendar",
        "dispatch_table",
        "soluble_inventory",
        "water_it_test",
        "notes",
        "trade_notes",
        "detail",
        "supplier_home",
        "supplier_list",
        "supplier_search",
        "supplier_product",
        "supplier_register",
        "carrier_home",
        "carrier_list",
        "carrier_search",
        "carrier_freight_compare",
        "carrier_register",
        "partner_detail",
        "change_history",
        "estimates",
        "data_backup",
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

    if page == "partner_detail" and partner_id:
        st.session_state["selected_partner_id"] = partner_id
        st.session_state["selected_partner_type"] = partner_type
    elif page != "partner_detail":
        st.session_state["selected_partner_id"] = None
        st.session_state["selected_partner_type"] = None


def set_page(page_name, rerun=False):
    st.session_state["page"] = page_name

    if page_name != "detail":
        st.session_state["selected_customer"] = None
    if page_name != "partner_detail":
        st.session_state["selected_partner_id"] = None
        st.session_state["selected_partner_type"] = None

    update_query_params(
        page=page_name,
        customer=None,
        partner_id=None,
        partner_type=None,
    )

    if rerun:
        st.rerun()


def select_customer(customer_name, page_name="detail"):
    st.session_state["selected_customer"] = customer_name
    st.session_state["page"] = page_name
    update_query_params(page=page_name, customer=customer_name)


def show_back_home_button(key):
    """既存の顧客画面から顧客メニューへ戻る共通リンク。"""
    st.markdown(
        render_page_link("← 顧客メニューへ戻る", page="customer_home"),
        unsafe_allow_html=True,
    )


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
        st.error("同じ顧客名・商品名の行が複数見つかりました。確認してください。")
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
                        result["history_warning"] = record_change_history_safely(
                            "顧客",
                            "",
                            customer_name,
                            "変更",
                            pending["changes"],
                            section=f"商品：{product_name}",
                        )
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
                    result["history_warning"] = record_change_history_safely(
                        "顧客",
                        "",
                        customer_name,
                        "変更",
                        changes,
                        section=f"商品：{product_name}",
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
                        result["history_warning"] = record_change_history_safely(
                            "顧客",
                            "",
                            customer_name,
                            "変更",
                            pending["changes"],
                            section="住所・マップ位置",
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
                    result["history_warning"] = record_change_history_safely(
                        "顧客",
                        "",
                        customer_name,
                        "変更",
                        changes,
                        section="住所・マップ位置",
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
        if success.get("history_warning"):
            st.warning(success["history_warning"])

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

    # WATER itのポイント名と顧客名が一致する場合だけ、最新値を読み取り専用で表示する。
    render_customer_water_it_card(customer_name)

    if visible_detail.empty:
        st.info("表示対象の商品はありません。使用数量/日が0または空白の商品は非表示にしています。")

    # 同じ商品に使用中行が複数あっても、商品カードは1つだけ表示する。
    # 複数の使用中行がある場合はカード内で警告し、編集を停止する。
    visible_products = []
    seen_products = set()
    for _, candidate_row in visible_detail.iterrows():
        candidate_product = clean_value(candidate_row["商品名"], blank_text="").strip()
        if not candidate_product or candidate_product in seen_products:
            continue
        seen_products.add(candidate_product)
        visible_products.append(candidate_row)

    for row in visible_products:
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

            product_match_count = int(
                (
                    (detail["商品名"].astype(str).str.strip() == product_name)
                    & (~detail["使用数量/日"].apply(is_blank_or_zero))
                ).sum()
            )
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

    render_customer_estimates_section(customer_name, customer_key)
    render_customer_attachments_section(customer_name, customer_key)

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

    page_name = "customer" if show_home_link else "customer_home"

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
# 商品検索
# =========================
def get_product_search_rows(df):
    """商品名・顧客名がある全行を、現在・過去を問わず商品検索用に返す。"""
    required_columns = {
        "顧客名",
        "地域",
        "商品名",
        "使用数量/日",
    }
    if not required_columns.issubset(df.columns):
        return pd.DataFrame(columns=list(required_columns))

    rows = df.copy()
    rows["_商品名検索"] = rows["商品名"].apply(
        lambda value: clean_value(value, blank_text="").strip()
    )
    rows["_顧客名検索"] = rows["顧客名"].apply(
        lambda value: clean_value(value, blank_text="").strip()
    )
    rows = rows[
        (rows["_商品名検索"] != "")
        & (rows["_顧客名検索"] != "")
    ].copy()
    return rows


def get_product_search_candidates(product_rows, keyword):
    """入力文字を含む、現在または過去に登録された商品名を候補タブ用に返す。"""
    keyword = str(keyword or "").strip()
    if product_rows.empty or not keyword:
        return []

    matches = product_rows[
        product_rows["_商品名検索"].str.contains(
            keyword,
            case=False,
            na=False,
            regex=False,
        )
    ]
    candidates = matches["_商品名検索"].drop_duplicates().tolist()
    keyword_folded = keyword.casefold()
    return sorted(
        candidates,
        key=lambda product_name: (
            product_name.casefold() != keyword_folded,
            not product_name.casefold().startswith(keyword_folded),
            product_name.casefold().find(keyword_folded),
            len(product_name),
            product_name,
        ),
    )


def build_exact_product_search_results(product_rows, product_name):
    """商品名の完全一致結果を顧客単位にまとめ、現在使用中と過去使用に分ける。"""
    exact = product_rows[product_rows["_商品名検索"] == product_name].copy()
    if exact.empty:
        return [], []

    current_results = []
    past_results = []

    for customer_name, group in exact.groupby("_顧客名検索", sort=False):
        group = group.copy()
        active_group = group[
            ~group["使用数量/日"].apply(is_blank_or_zero)
        ].copy()

        region_values = [
            clean_value(value, blank_text="").strip()
            for value in group["地域"].tolist()
        ]
        region = next((value for value in region_values if value), "未設定")

        if not active_group.empty:
            duplicate_count = len(active_group)
            usage_text = (
                format_number(active_group.iloc[0]["使用数量/日"])
                if duplicate_count == 1
                else "複数行（確認が必要）"
            )
            current_results.append(
                {
                    "顧客名": customer_name,
                    "地域": region,
                    "使用数量/日": usage_text,
                    "重複件数": duplicate_count,
                }
            )
        else:
            past_results.append(
                {
                    "顧客名": customer_name,
                    "地域": region,
                }
            )

    current_results.sort(key=lambda item: item["顧客名"])
    past_results.sort(key=lambda item: item["顧客名"])
    return current_results, past_results


def render_product_search_customer(item, keyword, current):
    """商品検索の顧客カードを、現在使用中・過去使用の共通形式で表示する。"""
    customer_name = item["顧客名"]
    with st.container(border=True):
        st.markdown(
            render_page_link(
                f"👤 {customer_name}",
                page="detail",
                customer=customer_name,
                product_search=keyword,
            ),
            unsafe_allow_html=True,
        )

        st.caption("地域")
        st.markdown(f"**{html.escape(item['地域'])}**")

        if current:
            st.caption("使用数量/日")
            st.markdown(f"**{html.escape(item['使用数量/日'])}**")
            if item["重複件数"] > 1:
                st.warning(
                    "同じ顧客名・商品名の使用中行が複数見つかりました。"
                    "顧客詳細で確認してください。"
                )
        else:
            st.caption("この商品を過去に使用")


def render_product_search_results(product_rows, product_name, keyword):
    """選択した商品の顧客を、現在使用中と過去使用に分けて表示する。"""
    current_results, past_results = build_exact_product_search_results(
        product_rows,
        product_name,
    )

    st.markdown(
        f"**現在使用中 {len(current_results)}件 ／ "
        f"過去に使用 {len(past_results)}件**"
    )

    st.markdown(f"### 🟢 現在使用中　{len(current_results)}件")
    if not current_results:
        st.info("この商品を現在使用している顧客はいません。")
    else:
        for item in current_results:
            render_product_search_customer(item, keyword, current=True)

    with st.expander(
        f"⚪ 過去に使用　{len(past_results)}件",
        expanded=False,
    ):
        if not past_results:
            st.info("この商品を過去に使用した顧客はいません。")
        else:
            for item in past_results:
                render_product_search_customer(item, keyword, current=False)


def show_product_search(df=None):
    st.subheader("🔎 商品検索")
    show_back_home_button("product_back_home")
    st.caption(
        "商品名の一部を入力し、候補タブを選んでください。"
        "次回配達日や残数には関係なく、現在使用中の顧客と過去に使用した顧客を分けて表示します。"
    )

    default_keyword = str(get_query_value("product_search", "")).strip()
    if st_keyup is not None:
        keyword = str(
            st_keyup(
                "商品名で検索",
                value=default_keyword,
                placeholder="例：酒 と入力すると酒粕などが候補に出ます",
                debounce=250,
                key="product_search_live",
            )
            or ""
        ).strip()
    else:
        keyword = st.text_input(
            "商品名で検索",
            value=default_keyword,
            placeholder="例：酒 と入力すると酒粕などが候補に出ます",
            key="product_search_input",
            help=VOICE_INPUT_HELP,
        ).strip()

    if keyword:
        update_query_params(page="product", product_search=keyword)
    else:
        update_query_params(page="product", product_search=None)

    if not keyword:
        st.info("商品名を入力してください。")
        return

    if df is None:
        with st.spinner("商品データを読み込んでいます…"):
            df = load_data()

    product_rows = get_product_search_rows(df)
    candidates = get_product_search_candidates(product_rows, keyword)

    if not candidates:
        st.warning("該当する商品名がありません。")
        return

    st.write(f"商品候補：{len(candidates)}件")
    tabs = st.tabs([f"📦 {product_name}" for product_name in candidates])
    for product_name, tab in zip(candidates, tabs):
        with tab:
            st.markdown(f"#### {html.escape(product_name)}")
            render_product_search_results(
                product_rows,
                product_name,
                keyword,
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


def render_dispatch_responsive_list(display_df, customer_names=None):
    """PCはExcel風一覧、スマホは横スクロール不要の縦型カードで表示する。"""
    customer_names = set(customer_names or [])
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
        .dispatch-excel-table .dispatch-teshikaga-text,
        .dispatch-excel-table .dispatch-teshikaga-text a { color: #dc2626 !important; }
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
            .dispatch-route-box.dispatch-teshikaga-destination { background: #fee2e2; }
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

    def is_teshikaga_destination(value):
        return normalize_dispatch_text(value) == "弟子屈"

    def destination_value(value, highlight_text=False):
        destination = normalize_dispatch_text(value) or "未入力"
        if destination in customer_names:
            rendered = build_customer_detail_link(
                destination,
                class_name="dispatch-month-link",
            )
        else:
            rendered = html.escape(destination)
        if highlight_text and destination == "弟子屈":
            return f'<span class="dispatch-teshikaga-text">{rendered}</span>'
        return rendered

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
            cell_value = (
                destination_value(row.get(column), highlight_text=True)
                if column == "納品先"
                else safe_value(row.get(column))
            )
            desktop_parts.append(f'<td class="{css_class}">{cell_value}</td>')
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
                    f'<div class="dispatch-route-box{" dispatch-teshikaga-destination" if is_teshikaga_destination(row.get("納品先")) else ""}"><span class="dispatch-route-label">納品先</span><span class="dispatch-route-value">{destination_value(row.get("納品先"))}</span></div>',
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

    customer_names = set()
    try:
        customer_df = load_data()
        if "顧客名" in customer_df.columns:
            customer_names = {
                clean_value(value, blank_text="").strip()
                for value in customer_df["顧客名"].tolist()
                if clean_value(value, blank_text="").strip()
            }
    except Exception:
        # 顧客データを確認できない場合も、配車表は従来どおり文字表示で続行する。
        customer_names = set()

    render_dispatch_responsive_list(display_df, customer_names)


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
                field_map = {
                    "delivery_date": ("配達日", current.get("配達日")),
                    "delivery_quantity": ("配達数量", current.get("配達数量")),
                    "usage": ("使用数量/日", current.get("使用数量/日")),
                }
                history_changes = {
                    field_map[field][0]: (field_map[field][1], value)
                    for field, value in updates.items()
                    if field in field_map
                }
                remember_change_history_warning(
                    record_change_history_safely(
                        "顧客",
                        "",
                        customer_name,
                        "変更",
                        history_changes,
                        section="ソリュブル",
                    )
                )
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



def get_soluble_water_it_history(dataframe, location):
    """選択会社に対応するWATER itのタンク履歴を、kgの数値行だけ返す。"""
    point_name = SOLUBLE_WATER_IT_POINT_NAMES.get(location)
    if not point_name or dataframe is None or dataframe.empty:
        return pd.DataFrame()

    history = get_water_it_customer_rows(dataframe, point_name)
    if history.empty:
        return history

    # ノベルズ・コスモは保管タンク1基。別の測定項目を誤って混ぜない。
    tank_rows = history[
        history["測定項目"].astype(str).str.contains("保管タンク", regex=False, na=False)
    ].copy()
    tank_rows = tank_rows[
        tank_rows["測定値_数値"].notna()
        & tank_rows["単位_表示"].map(normalize_water_it_unit).eq("kg")
    ].copy()
    if tank_rows.empty:
        return tank_rows

    # 測定項目が複数ある場合は、最新時刻を持つ1項目だけに限定する。
    latest_by_item = tank_rows.groupby("測定項目")["測定日時_解析"].max()
    selected_item = latest_by_item.idxmax()
    tank_rows = tank_rows[tank_rows["測定項目"] == selected_item].copy()
    tank_rows.sort_values("測定日時_解析", inplace=True)
    tank_rows.reset_index(drop=True, inplace=True)
    return tank_rows


def get_soluble_water_it_daily_actuals(history):
    """WATER it履歴から、各日の9:00実測値を日付ごとに返す。

    CSVを数日取り込まなかった場合でも、次回CSVに含まれる過去日の9:00値を
    アプリ表示へまとめて反映する。Excel本体は変更しない。
    """
    if history is None or history.empty:
        return {}

    daily = history.dropna(subset=["測定日時_解析", "測定値_数値"]).copy()
    if daily.empty:
        return {}

    measured_at = pd.to_datetime(daily["測定日時_解析"], errors="coerce")
    daily = daily[
        measured_at.notna()
        & measured_at.dt.hour.eq(9)
        & measured_at.dt.minute.eq(0)
    ].copy()
    if daily.empty:
        return {}

    daily["_water_it_measured_at"] = pd.to_datetime(
        daily["測定日時_解析"],
        errors="coerce",
    )
    daily["_water_it_date"] = daily["_water_it_measured_at"].dt.date
    daily.sort_values("_water_it_measured_at", inplace=True)

    # 同じ日の9:00行が重複していても、CSV内で最後の1件だけを採用する。
    daily = daily.groupby("_water_it_date", sort=True, as_index=False).tail(1)
    result = {}
    for _, row in daily.iterrows():
        value = float(row["測定値_数値"])
        if not math.isfinite(value):
            continue
        if value.is_integer():
            value = int(value)
        result[row["_water_it_date"]] = {
            "value": value,
            "measured_at": row["_water_it_measured_at"],
            "source": "09:00",
        }
    return result


def estimate_soluble_water_it_daily_usage(history, days):
    """WATER it履歴から1日平均使用量を参考値として推定する。

    1時間ごとの中央値で短時間の揺れをならし、大きな上昇は納品として区切る。
    各区間の開始値と終了値の減少分だけを合計するため、Excelの使用量/日や
    将来予測へは一切反映しない。
    """
    result = {
        "days": int(days),
        "average": None,
        "available_days": 0.0,
        "enough_data": False,
    }
    if history is None or history.empty:
        return result

    series = (
        history.dropna(subset=["測定日時_解析", "測定値_数値"])
        .sort_values("測定日時_解析")
        .set_index("測定日時_解析")["測定値_数値"]
        .astype(float)
    )
    series = series[~series.index.duplicated(keep="last")]
    if len(series) < 2:
        return result

    latest_time = series.index.max()
    oldest_time = series.index.min()
    available_days = max(
        0.0,
        (latest_time - oldest_time).total_seconds() / 86400.0,
    )
    result["available_days"] = available_days

    # 10分程度の端数は許容するが、期間が足りない時は無理に期間平均を出さない。
    if available_days < float(days) - 0.1:
        return result

    cutoff = latest_time - pd.Timedelta(days=int(days))
    hourly = series.resample("1h").median().interpolate(limit=2)
    hourly = hourly[(hourly.index >= cutoff) & (hourly.index <= latest_time)].dropna()
    if len(hourly) < 2:
        return result

    median_level = float(hourly.median())
    # 小さな測定揺れは納品扱いにしない。実タンクでは納品上昇が数千kg単位になる。
    delivery_jump = max(1000.0, abs(median_level) * 0.04)
    differences = hourly.diff()
    split_positions = [
        hourly.index.get_loc(timestamp)
        for timestamp in hourly.index[differences > delivery_jump]
    ]

    starts = [0] + split_positions
    ends = [position - 1 for position in split_positions] + [len(hourly) - 1]
    total_decrease = 0.0
    for start_position, end_position in zip(starts, ends):
        if end_position <= start_position:
            continue
        decrease = float(hourly.iloc[start_position] - hourly.iloc[end_position])
        if math.isfinite(decrease) and decrease > 0:
            total_decrease += decrease

    elapsed_days = (
        hourly.index[-1] - hourly.index[0]
    ).total_seconds() / 86400.0
    if elapsed_days <= 0:
        return result

    average = total_decrease / elapsed_days
    if not math.isfinite(average) or average < 0:
        return result

    result["average"] = average
    result["enough_data"] = True
    return result


def get_soluble_water_it_context(location, rows):
    """ソリュブル画面用の実測値・差額・参考平均をまとめる。失敗時はNone。"""
    if location not in SOLUBLE_WATER_IT_POINT_NAMES:
        return None
    try:
        dataframe, source = get_active_water_it_data()
        history = get_soluble_water_it_history(dataframe, location)
    except Exception:
        return None
    if history.empty:
        return None

    latest = history.iloc[-1]
    measured_at = latest["測定日時_解析"]
    if pd.isna(measured_at):
        return None
    actual_value = float(latest["測定値_数値"])
    if not math.isfinite(actual_value):
        return None
    if actual_value.is_integer():
        actual_value = int(actual_value)

    measured_date = measured_at.date()
    excel_row = next((row for row in rows if row.get("date") == measured_date), None)
    excel_value = (
        excel_row.get(f"{location}_inventory")
        if excel_row is not None
        else None
    )
    excel_usage = (
        excel_row.get(f"{location}_usage")
        if excel_row is not None
        else None
    )
    difference = None
    if isinstance(excel_value, (int, float)):
        difference = float(actual_value) - float(excel_value)
        if difference.is_integer():
            difference = int(difference)

    usage_averages = {
        days: estimate_soluble_water_it_daily_usage(history, days)
        for days in SOLUBLE_WATER_IT_USAGE_WINDOWS
    }

    # 日付ごとの在庫欄は、最新値ではなく各日の9:00実測だけを使う。
    # 上部の「現在の実測在庫」は従来どおりCSV内の最新値を表示する。
    daily_actuals = get_soluble_water_it_daily_actuals(history)

    today = get_jst_now().date()
    today_actual = daily_actuals.get(today)
    today_excel_row = next((row for row in rows if row.get("date") == today), None)
    today_excel_value = (
        today_excel_row.get(f"{location}_inventory")
        if today_excel_row is not None
        else None
    )

    return {
        "source": source,
        "history": history,
        "actual_value": actual_value,
        "measured_at": measured_at,
        "measured_date": measured_date,
        "unit": "kg",
        "excel_row": excel_row,
        "excel_value": excel_value,
        "excel_usage": excel_usage,
        "difference": difference,
        "usage_averages": usage_averages,
        "daily_actuals": daily_actuals,
        "today_9am_actual": today_actual,
        "today_9am_excel_row": today_excel_row,
        "today_9am_excel_value": today_excel_value,
    }


def apply_soluble_water_it_forecast(rows, location, context):
    """WATER it実測を日付ごとに反映し、その間だけアプリ予測を行う。

    CSVに各日の9:00実測値が含まれる場合は、過去日も含めてその日の在庫表示を
    実測値へ置き換える。実測値がない日は、直前の実測または黄色い手入力基準から
    Excelの使用量・納品を使って計算する。Excel本体・元のrows・使用量/日・納品は
    変更しない。
    """
    display_rows = [dict(row) for row in rows]
    if not context:
        return display_rows

    daily_actuals = context.get("daily_actuals") or {}
    if not daily_actuals:
        return display_rows

    previous_inventory = None
    started = False
    for row in sorted(display_rows, key=lambda item: item["date"]):
        row_date = row["date"]
        display_key = f"{location}_inventory_display"
        source_key = f"{location}_inventory_display_source"

        # CSV内にその日の9:00実測があれば、最新日の直近値ではなく、
        # 必ず9:00の値をExcel計算値より優先して緑の実測値として表示する。
        actual = daily_actuals.get(row_date)
        if actual is not None:
            actual_value = actual.get("value")
            if isinstance(actual_value, (int, float)) and math.isfinite(float(actual_value)):
                previous_inventory = actual_value
                row[display_key] = actual_value
                row[source_key] = "water_it_actual"
                row[f"{location}_inventory_measured_at"] = actual.get("measured_at")
                started = True
                continue

        if not started:
            row[display_key] = row.get(f"{location}_inventory")
            row[source_key] = "excel"
            continue

        # 実測値がない日の黄色い在庫は、人が指定した次の基準値として尊重する。
        # 後日の9:00実測が現れた時点で、その実測値へ再び補正される。
        if row.get(f"{location}_inventory_manual"):
            manual_value = row.get(f"{location}_inventory")
            if isinstance(manual_value, (int, float)):
                previous_inventory = manual_value
                row[display_key] = manual_value
                row[source_key] = "excel_manual_baseline"
                continue

        usage = row.get(f"{location}_usage")
        delivery = row.get(f"{location}_delivery")
        usage = usage if isinstance(usage, (int, float)) else 0
        delivery = delivery if isinstance(delivery, (int, float)) else 0
        if isinstance(previous_inventory, (int, float)):
            previous_inventory = previous_inventory - usage + delivery
            row[display_key] = previous_inventory
            row[source_key] = "water_it_forecast"
        else:
            row[display_key] = row.get(f"{location}_inventory")
            row[source_key] = "excel"

    return display_rows


def find_soluble_row_by_date(content, target_date):
    matches = [row for row in read_soluble_rows(content) if row.get("date") == target_date]
    if len(matches) != 1:
        if not matches:
            raise RuntimeError(
                f"Excelのソリュブルシートに{target_date.strftime('%Y/%m/%d')}の行がありません。"
            )
        raise RuntimeError("同じ日付の行が複数あるため、安全のため更新を中止しました。")
    return matches[0]


def verify_soluble_water_it_baseline(content, location, target_date, expected_value, next_formula):
    """Dropbox保存後に実測値・黄色・翌日の式が保たれていることを確認する。"""
    row = find_soluble_row_by_date(content, target_date)
    workbook = load_workbook(BytesIO(content), data_only=False, read_only=False)
    try:
        if SOLUBLE_SHEET_NAME not in workbook.sheetnames:
            raise RuntimeError("保存後のExcelにソリュブルシートがありません。")
        ws = workbook[SOLUBLE_SHEET_NAME]
        inventory_column = SOLUBLE_LOCATIONS[location]["inventory"]
        cell = ws.cell(row["row"], inventory_column)
        if not same_soluble_value(cell.value, expected_value):
            raise RuntimeError("保存後のExcelで実測値が一致しません。")
        if not soluble_cell_is_manual(cell):
            raise RuntimeError("保存後のExcelで実測値のセルが黄色になっていません。")
        if next_formula is not None and row["row"] < ws.max_row:
            actual_next_formula = ws.cell(row["row"] + 1, inventory_column).value
            if actual_next_formula != next_formula:
                raise RuntimeError("保存後のExcelで翌日の在庫計算式が変わっています。")
    finally:
        workbook.close()


def save_soluble_water_it_baseline(location, context):
    """選択会社の今日の実測値だけを、バックアップ付きでExcelの黄色い基準値にする。"""
    if location not in SOLUBLE_WATER_IT_POINT_NAMES:
        raise RuntimeError("WATER it実測値を反映できる会社ではありません。")
    if not context:
        raise RuntimeError("WATER itの実測値を確認できません。")

    today_actual = context.get("today_9am_actual") or {}
    measured_at = today_actual.get("measured_at")
    target_date = get_jst_now().date()
    actual_value = today_actual.get("value")
    if measured_at is None or not isinstance(actual_value, (int, float)):
        raise RuntimeError("今日9:00の実測値がないため、Excelへの反映を中止しました。")
    if measured_at.date() != target_date or measured_at.hour != 9 or measured_at.minute != 0:
        raise RuntimeError("今日9:00の実測値を確認できないため、Excelへの反映を中止しました。")
    if normalize_water_it_unit(context.get("unit")) != "kg":
        raise RuntimeError("kg以外の実測値はExcelへ反映しません。")
    if actual_value < 0:
        raise RuntimeError("実測値がマイナスのため、Excelへの反映を中止しました。")

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

        latest_row = find_soluble_row_by_date(original_content, target_date)
        inventory_column = SOLUBLE_LOCATIONS[location]["inventory"]
        formula_book = load_workbook(BytesIO(original_content), data_only=False, read_only=False)
        try:
            ws = formula_book[SOLUBLE_SHEET_NAME]
            next_formula = (
                ws.cell(latest_row["row"] + 1, inventory_column).value
                if latest_row["row"] < ws.max_row
                else None
            )
        finally:
            formula_book.close()

        saved_content, changed = build_soluble_updated_workbook(
            original_content,
            latest_row["row"],
            location,
            {"inventory": actual_value},
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
                "対象Excelを更新できませんでした。\n" + dropbox_error_text(upload_response)
            )

        confirmed_content, confirmed_response = download_dropbox_file(target_path, access_token)
        if confirmed_content is None:
            raise RuntimeError(
                "保存後のExcelを再取得できませんでした。\n"
                + dropbox_error_text(confirmed_response)
            )
        verify_soluble_water_it_baseline(
            confirmed_content,
            location,
            target_date,
            actual_value,
            next_formula,
        )
    else:
        local_path = Path(str(SOLUBLE_LOCAL_FILE))
        if not local_path.exists():
            raise FileNotFoundError(f"対象ファイルが見つかりません：{local_path}")
        original_content = local_path.read_bytes()
        latest_row = find_soluble_row_by_date(original_content, target_date)
        inventory_column = SOLUBLE_LOCATIONS[location]["inventory"]
        formula_book = load_workbook(BytesIO(original_content), data_only=False, read_only=False)
        try:
            ws = formula_book[SOLUBLE_SHEET_NAME]
            next_formula = (
                ws.cell(latest_row["row"] + 1, inventory_column).value
                if latest_row["row"] < ws.max_row
                else None
            )
        finally:
            formula_book.close()
        saved_content, changed = build_soluble_updated_workbook(
            original_content,
            latest_row["row"],
            location,
            {"inventory": actual_value},
        )
        backup_dir = local_path.parent / "Backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup_path = backup_dir / f"aoベンチャーグレイン配車表_{timestamp}.xlsx"
        backup_path.write_bytes(original_content)
        local_path.write_bytes(saved_content)
        verify_soluble_water_it_baseline(
            local_path.read_bytes(),
            location,
            target_date,
            actual_value,
            next_formula,
        )

    st.cache_data.clear()
    return changed


def render_soluble_water_it_summary(location, context):
    """ソリュブル画面に実測値、Excelとの差、参考平均、任意反映ボタンを表示する。"""
    if not context:
        return

    success = st.session_state.pop("soluble_water_it_excel_success", None)
    if success and success.get("location") == location:
        st.success(
            f"{success['date']}の実測値 {soluble_number_label(success['value'])} kg を"
            "Excelの基準値として保存しました。対象セルは黄色です。"
        )

    display_name = SOLUBLE_LOCATION_DISPLAY_NAMES.get(location, location)
    actual_value = context["actual_value"]
    measured_at = context["measured_at"]
    excel_value = context.get("excel_value")
    excel_usage = context.get("excel_usage")
    difference = context.get("difference")

    with st.container(border=True):
        st.subheader(f"💧 {display_name}のWATER it実測")
        st.caption(
            f"最終受信：{measured_at.strftime('%Y/%m/%d %H:%M')}　｜　参照：{context['source']}"
        )
        def summary_card(label, value, tone=""):
            tone_class = f" {tone}" if tone else ""
            return (
                f'<div class="soluble-waterit-stat{tone_class}">'
                f'<span class="soluble-waterit-stat-label">{html.escape(str(label))}</span>'
                f'<span class="soluble-waterit-stat-value">{html.escape(str(value))}</span>'
                '</div>'
            )

        st.markdown(
            """
            <style>
            .soluble-waterit-summary-grid {
                display: grid;
                grid-template-columns: repeat(3, minmax(0, 1fr));
                gap: 0.75rem;
                margin: 0.4rem 0 1.1rem;
            }
            .soluble-waterit-usage-grid {
                display: grid;
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 0.75rem;
                margin: 0.45rem 0 0.8rem;
            }
            .soluble-waterit-stat {
                min-width: 0;
                box-sizing: border-box;
                padding: 0.9rem 1rem;
                border: 1px solid rgba(15, 23, 42, 0.12);
                border-radius: 14px;
                background: rgba(255, 255, 255, 0.9);
            }
            .soluble-waterit-stat.actual {
                background: #dcfce7;
                border-color: #4ade80;
            }
            .soluble-waterit-stat.excel {
                background: #f8fafc;
                border-color: #cbd5e1;
            }
            .soluble-waterit-stat.difference {
                background: #eff6ff;
                border-color: #93c5fd;
            }
            .soluble-waterit-stat.average {
                background: #f0fdfa;
                border-color: #5eead4;
            }
            .soluble-waterit-stat-label {
                display: block;
                color: #667085;
                font-size: 0.88rem;
                font-weight: 700;
                line-height: 1.35;
                margin-bottom: 0.35rem;
                overflow-wrap: anywhere;
            }
            .soluble-waterit-stat-value {
                display: block;
                color: #172033;
                font-size: clamp(1.55rem, 3.2vw, 2.2rem);
                font-weight: 800;
                line-height: 1.15;
                letter-spacing: -0.02em;
                white-space: normal;
                overflow: visible;
                text-overflow: clip;
                overflow-wrap: anywhere;
            }
            @media (max-width: 640px) {
                .soluble-waterit-summary-grid,
                .soluble-waterit-usage-grid {
                    grid-template-columns: 1fr;
                    gap: 0.55rem;
                }
                .soluble-waterit-stat {
                    padding: 0.78rem 0.85rem;
                }
                .soluble-waterit-stat-label {
                    font-size: 0.82rem;
                }
                .soluble-waterit-stat-value {
                    font-size: 1.45rem;
                    white-space: nowrap;
                    overflow-wrap: normal;
                }
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

        actual_label = f"{soluble_number_label(actual_value)} kg"
        excel_label = (
            f"{soluble_number_label(excel_value)} kg"
            if excel_value is not None
            else "—"
        )
        difference_label = "—" if difference is None else f"{difference:+,.0f} kg"
        st.markdown(
            '<div class="soluble-waterit-summary-grid">'
            + summary_card(f"現在の実測在庫/{measured_at.strftime('%H:%M')}", actual_label, "actual")
            + summary_card("同日のExcel計算在庫", excel_label, "excel")
            + summary_card("実測 − Excel", difference_label, "difference")
            + '</div>',
            unsafe_allow_html=True,
        )

        st.markdown("#### 実績から見た1日平均使用量（参考）")
        excel_usage_label = (
            f"{soluble_number_label(excel_usage)} kg/日"
            if isinstance(excel_usage, (int, float))
            else "—"
        )
        st.markdown(
            '<div class="soluble-waterit-usage-grid">'
            + summary_card("Excel設定使用量", excel_usage_label, "excel")
            + '</div>',
            unsafe_allow_html=True,
        )
        st.caption(
            "WATER it履歴を1時間単位でならし、大きな在庫増加は納品として区切った推定値です。"
            "Excelの使用量/日や予測計算へは自動反映しません。"
        )

        # st.tabsは先頭タブが初期表示になるため、7日を先頭にして標準表示にする。
        usage_tab_specs = (
            (7, "7日（標準）"),
            (3, "3日"),
            (20, "20日"),
            (30, "30日"),
        )
        usage_tabs = st.tabs([label for _, label in usage_tab_specs])
        for tab, (days, _) in zip(usage_tabs, usage_tab_specs):
            estimate = context["usage_averages"][days]
            with tab:
                if estimate.get("enough_data") and estimate.get("average") is not None:
                    average_label = f"{estimate['average']:,.0f} kg/日"
                    detail_label = f"直近{days}日の実績平均"
                else:
                    average_label = "データ不足"
                    detail_label = (
                        f"直近{days}日分には不足しています（現在 約"
                        f"{estimate.get('available_days', 0):.1f}日分）"
                    )
                st.markdown(
                    '<div class="soluble-waterit-usage-grid">'
                    + summary_card(detail_label, average_label, "average")
                    + '</div>',
                    unsafe_allow_html=True,
                )

        st.markdown("#### Excelへ反映（任意）")
        st.caption(
            "アプリは各日の9:00実測値を在庫に使います。Excelは自動変更せず、ここで確認して押した場合だけ、"
            "今日9:00の在庫を基準値として黄色セルへ保存します。保存前バックアップと保存後確認を行います。"
        )

        today = get_jst_now().date()
        today_actual = context.get("today_9am_actual") or {}
        baseline_value = today_actual.get("value")
        baseline_measured_at = today_actual.get("measured_at")
        excel_row = context.get("today_9am_excel_row")
        baseline_excel_value = context.get("today_9am_excel_value")
        has_today_9am = (
            isinstance(baseline_value, (int, float))
            and baseline_measured_at is not None
            and baseline_measured_at.date() == today
            and baseline_measured_at.hour == 9
            and baseline_measured_at.minute == 0
        )
        same_value = (
            has_today_9am
            and baseline_excel_value is not None
            and same_soluble_value(baseline_excel_value, baseline_value)
        )
        if not has_today_9am:
            st.warning("今日9:00の実測値がCSVにないため、Excelへの反映ボタンは使えません。")
            return
        if excel_row is None:
            st.warning("今日の行がExcelにないため、反映できません。")
            return
        if same_value:
            st.info("今日のExcel在庫は、すでに9:00実測値と同じです。")
            return

        confirm_key = f"soluble_water_it_confirm_{location}_{today.isoformat()}"
        confirmed = st.checkbox(
            f"{today.strftime('%Y/%m/%d')}のExcel在庫を "
            f"9:00実測の {soluble_number_label(baseline_value)} kg に変更する",
            key=confirm_key,
        )
        if st.button(
            "今日の実測値をExcelの基準値にする",
            key=f"soluble_water_it_save_{location}_{today.isoformat()}",
            type="primary",
            use_container_width=True,
            disabled=not confirmed,
        ):
            try:
                with st.spinner("元ファイルをバックアップし、9:00実測値を保存・確認しています…"):
                    changed = save_soluble_water_it_baseline(location, context)
                    remember_change_history_warning(
                        record_change_history_safely(
                            "顧客",
                            "",
                            SOLUBLE_LOCATION_DISPLAY_NAMES.get(location, location),
                            "変更",
                            {"在庫基準値": (baseline_excel_value, baseline_value)},
                            section="ソリュブル在庫（WATER it実測反映）",
                        )
                    )
                st.session_state["soluble_water_it_excel_success"] = {
                    "location": location,
                    "date": today.strftime("%Y/%m/%d"),
                    "value": baseline_value,
                    "changed_count": len(changed),
                }
                st.rerun()
            except Exception as exc:
                st.error(f"Excelへ反映できませんでした：{exc}")

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
        format_func=lambda name: SOLUBLE_LOCATION_DISPLAY_NAMES.get(name, name),
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
    water_it_context = get_soluble_water_it_context(location, rows)
    if water_it_context is not None:
        render_soluble_water_it_summary(location, water_it_context)
        active_rows = apply_soluble_water_it_forecast(rows, location, water_it_context)
    else:
        # WATER itを読めない時も、既存のExcel表示・編集ルールは従来どおり動かす。
        active_rows = list(rows)
        if location in SOLUBLE_WATER_IT_POINT_NAMES:
            st.info("WATER itの保存済みデータを確認できないため、Excelの値だけを表示しています。")
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
        .soluble-value.waterit-actual {background:#dcfce7; border:1px solid #4ade80;}
        .soluble-value.waterit-forecast {background:#e0f2fe; border:1px solid #7dd3fc;}
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
        f'<div class="soluble-legend"><span class="soluble-yellow-chip"></span><span>黄色はExcelの手入力　｜　緑はWATER it実測　｜　水色は実測起点のアプリ予測　｜　参照：{html.escape(source)}</span></div>',
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
        inventory_display = row.get(f"{location}_inventory_display", inventory)
        inventory_display_source = row.get(f"{location}_inventory_display_source", "excel")
        inventory_label = {
            "water_it_actual": "在庫（実測）",
            "water_it_forecast": "在庫（実測起点予測）",
            "excel_manual_baseline": "在庫（Excel手入力基準）",
        }.get(inventory_display_source, "在庫")
        cells = []
        for label, field, value in (
            ("使用量/日", "usage", usage),
            ("納品", "delivery", delivery),
            (inventory_label, "inventory", inventory_display),
        ):
            classes = ["soluble-value"]
            if field == "inventory" and inventory_display_source == "water_it_actual":
                classes.append("waterit-actual")
            elif field == "inventory" and inventory_display_source == "water_it_forecast":
                classes.append("waterit-forecast")
            elif row.get(f"{location}_{field}_manual"):
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
                        history_changes = {}
                        if "usage" in updates:
                            history_changes["使用量/日"] = (usage, new_usage)
                        if "delivery" in updates:
                            history_changes["納品"] = (delivery, new_delivery)
                        if "inventory" in updates:
                            before_inventory = "自動計算" if current_formula else inventory
                            after_inventory = "自動計算" if auto_inventory else new_inventory
                            history_changes["在庫"] = (before_inventory, after_inventory)
                        remember_change_history_warning(
                            record_change_history_safely(
                                "顧客",
                                "",
                                SOLUBLE_LOCATION_DISPLAY_NAMES.get(location, location),
                                "変更",
                                history_changes,
                                section=f"ソリュブル在庫 {day.strftime('%Y/%m/%d')}",
                            )
                        )
                    st.success(f"保存しました（{len(changed)}セル更新）。黄色は手入力値です。")
                    st.rerun()
                except Exception as error:
                    st.error(str(error))






# =========================
# WATER it接続（読み取り専用）
# =========================
def resolve_water_it_csv_path():
    """data.csvの場所を、このPythonファイル基準で解決する。"""
    configured = str(WATER_IT_CSV_PATH).strip() or "data.csv"
    path = Path(configured).expanduser()
    if not path.is_absolute():
        path = Path(__file__).resolve().parent / path
    return path


def read_water_it_source_bytes():
    """WATER itのCSVを読み取る。書き込み処理は行わない。"""
    csv_url = str(WATER_IT_CSV_URL).strip()
    if csv_url:
        try:
            response = requests.get(
                csv_url,
                timeout=WATER_IT_REQUEST_TIMEOUT,
                headers={"User-Agent": "Aoyama-WATER-it-readonly-test/1.0"},
            )
            response.raise_for_status()
        except requests.RequestException as exc:
            raise RuntimeError(f"WATER it CSV URLから取得できませんでした：{exc}") from exc

        content_type = str(response.headers.get("Content-Type", "")).lower()
        if "text/html" in content_type:
            raise RuntimeError(
                "WATER_IT_CSV_URLからCSVではなくHTMLが返されました。CSVを直接取得できるURLを設定してください。"
            )
        hostname = urllib.parse.urlparse(csv_url).hostname or "設定URL"
        return response.content, f"WATER_IT_CSV_URL（{hostname}）"

    path = resolve_water_it_csv_path()
    if not path.exists():
        raise RuntimeError(
            f"{path.name} が見つかりません。このPythonファイルと同じフォルダに data.csv を置いてください。"
        )
    if not path.is_file():
        raise RuntimeError(f"WATER_IT_CSV_PATH がファイルではありません：{path.name}")
    return path.read_bytes(), path.name


def normalize_water_it_unit(value):
    text = clean_value(value, blank_text="").strip()
    replacements = {
        "㎏": "kg",
        "ＫＧ": "kg",
        "ｋｇ": "kg",
        "Ｋｇ": "kg",
        "Ｌ": "L",
        "ℓ": "L",
        "㍑": "L",
    }
    return replacements.get(text, text)


def water_it_nonblank(value):
    if value is None or pd.isna(value):
        return False
    text = str(value).strip()
    if not text:
        return False
    return text.lower() not in {
        "nan",
        "none",
        "false",
        "0",
        "0.0",
        "-",
        "なし",
        "正常",
        "異常なし",
    }


def parse_water_it_csv(content):
    """WATER itのCSVを画面表示用に整形する。"""
    dataframe = None
    errors = []
    for encoding in ("utf-8-sig", "utf-8", "cp932", "shift_jis"):
        try:
            candidate = pd.read_csv(BytesIO(content), encoding=encoding)
            if len(candidate.columns) <= 1:
                raise ValueError("CSVの列を分割できませんでした。")
            dataframe = candidate
            break
        except Exception as exc:
            errors.append(f"{encoding}: {exc}")

    if dataframe is None:
        detail = " / ".join(errors[:2])
        raise RuntimeError(f"data.csvを読み込めませんでした。{detail}")

    dataframe.columns = [str(column).replace("\ufeff", "").strip() for column in dataframe.columns]
    missing = [column for column in WATER_IT_REQUIRED_COLUMNS if column not in dataframe.columns]
    if missing:
        raise RuntimeError(
            "data.csvに必要な列がありません：" + "、".join(missing)
        )

    dataframe = dataframe.copy()
    dataframe["測定日時_解析"] = pd.to_datetime(
        dataframe["測定日時"],
        errors="coerce",
    )

    number_translation = str.maketrans(
        "０１２３４５６７８９．，－＋",
        "0123456789.,-+",
    )
    number_text = (
        dataframe["測定値"]
        .astype(str)
        .str.translate(number_translation)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    dataframe["測定値_数値"] = pd.to_numeric(number_text, errors="coerce")
    dataframe["単位_表示"] = dataframe["単位"].apply(normalize_water_it_unit)
    dataframe["エリア"] = dataframe["エリア"].fillna("").astype(str).str.strip()
    dataframe["ポイント"] = dataframe["ポイント"].fillna("").astype(str).str.strip()
    dataframe["測定項目"] = dataframe["測定項目"].fillna("").astype(str).str.strip()

    dataframe = dataframe[
        dataframe["測定日時_解析"].notna()
        & dataframe["ポイント"].ne("")
        & dataframe["測定項目"].ne("")
    ].copy()
    dataframe.sort_values("測定日時_解析", ascending=False, inplace=True)
    dataframe.reset_index(drop=True, inplace=True)
    return dataframe


@st.cache_data(ttl=60, show_spinner=False)
def load_water_it_data():
    content, source = read_water_it_source_bytes()
    return parse_water_it_csv(content), source


def make_water_it_snapshot_payload(content, filename, dataframe):
    """検証済みCSVを圧縮し、Supabaseへ保存できるJSON文字列にする。"""
    latest_time = dataframe["測定日時_解析"].max()
    oldest_time = dataframe["測定日時_解析"].min()
    return json.dumps(
        {
            "version": WATER_IT_STORAGE_VERSION,
            "filename": str(filename or "data.csv"),
            "sha256": hashlib.sha256(content).hexdigest(),
            "csv_gzip_base64": base64.b64encode(gzip.compress(content)).decode("ascii"),
            "row_count": int(len(dataframe)),
            "point_count": int(dataframe["ポイント"].nunique()),
            "oldest_time": oldest_time.isoformat() if pd.notna(oldest_time) else None,
            "latest_time": latest_time.isoformat() if pd.notna(latest_time) else None,
            "imported_at": get_jst_now().isoformat(),
        },
        ensure_ascii=False,
        separators=(",", ":"),
    )


def decode_water_it_snapshot_payload(payload_text):
    """Supabaseに保存したスナップショットから元のCSVバイト列を復元する。"""
    try:
        payload = json.loads(str(payload_text or ""))
    except Exception as exc:
        raise RuntimeError("保存済みWATER itデータの形式が正しくありません。") from exc
    if not isinstance(payload, dict):
        raise RuntimeError("保存済みWATER itデータの形式が正しくありません。")
    if int(payload.get("version", 0)) != WATER_IT_STORAGE_VERSION:
        raise RuntimeError("保存済みWATER itデータのバージョンが対応外です。")
    encoded = str(payload.get("csv_gzip_base64") or "")
    if not encoded:
        raise RuntimeError("保存済みWATER itデータにCSV本体がありません。")
    try:
        content = gzip.decompress(base64.b64decode(encoded.encode("ascii")))
    except Exception as exc:
        raise RuntimeError("保存済みWATER itデータを復元できませんでした。") from exc
    expected_hash = str(payload.get("sha256") or "")
    if expected_hash and hashlib.sha256(content).hexdigest() != expected_hash:
        raise RuntimeError("保存済みWATER itデータの検証に失敗しました。")
    return content, str(payload.get("filename") or "data.csv"), payload


def save_water_it_snapshot_to_supabase(content, filename, dataframe):
    """選択したCSVを既存Supabaseへ保存する。WATER itやExcelには書き込まない。"""
    if not has_supabase_config():
        raise RuntimeError("Supabase設定がないため、CSVを永続保存できません。")
    now = get_jst_now().isoformat()
    payload = {
        "id": WATER_IT_STORAGE_ID,
        "customer_key": None,
        "customer_name": WATER_IT_STORAGE_CUSTOMER,
        "field_name": WATER_IT_STORAGE_FIELD,
        "content": make_water_it_snapshot_payload(content, filename, dataframe),
        "sort_order": 0,
        "created_at": now,
        "updated_at": now,
    }
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
        raise RuntimeError("WATER itデータをSupabaseへ保存できませんでした。") from exc
    if response.status_code not in (200, 201):
        detail = str(response.text or "").strip()[:500]
        raise RuntimeError(
            f"WATER itデータをSupabaseへ保存できませんでした（{response.status_code}）。"
            + (f" {detail}" if detail else "")
        )
    load_saved_water_it_snapshot.clear()


@st.cache_data(ttl=30, show_spinner=False)
def load_saved_water_it_snapshot():
    """Supabaseから最後に取り込んだCSVを取得する。未保存ならNoneを返す。"""
    if not has_supabase_config():
        return None
    try:
        response = requests.get(
            get_supabase_customer_information_url(),
            headers=get_supabase_headers(),
            params={
                "select": "content,updated_at",
                "id": f"eq.{WATER_IT_STORAGE_ID}",
                "limit": "1",
            },
            timeout=20,
        )
    except Exception:
        return None
    if response.status_code != 200:
        return None
    try:
        rows = response.json()
    except Exception:
        return None
    if not isinstance(rows, list) or not rows:
        return None
    try:
        content, filename, metadata = decode_water_it_snapshot_payload(rows[0].get("content"))
        dataframe = parse_water_it_csv(content)
    except Exception:
        return None
    return {
        "content": content,
        "filename": filename,
        "metadata": metadata,
        "dataframe": dataframe,
        "updated_at": rows[0].get("updated_at"),
    }


@st.cache_resource(show_spinner=False)
def get_water_it_temporary_store():
    """アプリ再起動まで、選択CSVをサーバーの一時メモリに保持する。"""
    return {"content": None, "name": None, "hash": None}


def get_active_water_it_data():
    """選択中CSV、Supabase保存済みCSV、同梱data.csvの順で読み込む。"""
    uploaded_content = st.session_state.get(WATER_IT_UPLOAD_BYTES_KEY)
    uploaded_name = st.session_state.get(WATER_IT_UPLOAD_NAME_KEY)
    persisted = bool(st.session_state.get(WATER_IT_UPLOAD_PERSISTED_KEY))
    if not uploaded_content:
        temporary_store = get_water_it_temporary_store()
        uploaded_content = temporary_store.get("content")
        uploaded_name = temporary_store.get("name")
        persisted = bool(temporary_store.get("persisted"))
    if uploaded_content:
        uploaded_name = str(uploaded_name or "選択したCSV")
        label = "スマホから選択・Supabase保存済み" if persisted else "スマホから選択（一時）"
        return parse_water_it_csv(uploaded_content), f"{label}：{uploaded_name}"

    saved = load_saved_water_it_snapshot()
    if saved:
        return saved["dataframe"].copy(), f"Supabase保存：{saved['filename']}"

    return load_water_it_data()


def remember_uploaded_water_it_csv(uploaded_file):
    """選択されたCSVを検証し、Supabaseへ自動保存する。"""
    content = uploaded_file.getvalue()
    if not content:
        raise RuntimeError("選択したCSVが空です。")
    dataframe = parse_water_it_csv(content)
    digest = hashlib.sha256(content).hexdigest()
    uploaded_name = uploaded_file.name or "data.csv"

    persisted = False
    st.session_state.pop("water_it_persist_warning_message", None)
    try:
        save_water_it_snapshot_to_supabase(content, uploaded_name, dataframe)
        persisted = True
    except Exception as exc:
        st.session_state["water_it_persist_warning_message"] = str(exc)

    st.session_state[WATER_IT_UPLOAD_BYTES_KEY] = content
    st.session_state[WATER_IT_UPLOAD_NAME_KEY] = uploaded_name
    st.session_state[WATER_IT_UPLOAD_HASH_KEY] = digest
    st.session_state[WATER_IT_UPLOAD_PERSISTED_KEY] = persisted
    temporary_store = get_water_it_temporary_store()
    temporary_store.update(
        {
            "content": content,
            "name": uploaded_name,
            "hash": digest,
            "persisted": persisted,
        }
    )
    label = "スマホから選択・Supabase保存済み" if persisted else "スマホから選択（一時）"
    return dataframe, f"{label}：{uploaded_name}"


def clear_uploaded_water_it_csv():
    for key in (
        WATER_IT_UPLOAD_BYTES_KEY,
        WATER_IT_UPLOAD_NAME_KEY,
        WATER_IT_UPLOAD_HASH_KEY,
        WATER_IT_UPLOAD_PERSISTED_KEY,
        "water_it_upload_success_message",
        "water_it_persist_warning_message",
        "water_it_upload_error_message",
    ):
        st.session_state.pop(key, None)
    temporary_store = get_water_it_temporary_store()
    temporary_store.update({"content": None, "name": None, "hash": None, "persisted": False})


def handle_water_it_mobile_upload(widget_key):
    """スマホのファイル選択完了直後にCSVを検証して保持する。"""
    st.session_state.pop("water_it_upload_success_message", None)
    st.session_state.pop("water_it_upload_error_message", None)
    uploaded_file = st.session_state.get(widget_key)
    if uploaded_file is None:
        return
    try:
        dataframe, source = remember_uploaded_water_it_csv(uploaded_file)
        latest_time = dataframe["測定日時_解析"].max()
        saved_text = (
            " Supabaseへ保存しました。"
            if st.session_state.get(WATER_IT_UPLOAD_PERSISTED_KEY)
            else " 一時表示には反映しました。"
        )
        st.session_state["water_it_upload_success_message"] = (
            f"{uploaded_file.name or '選択したファイル'} を受け取りました。"
            f" 最新測定日時：{latest_time.strftime('%Y/%m/%d %H:%M')}。"
            + saved_text
        )
    except Exception as exc:
        st.session_state["water_it_upload_error_message"] = str(exc)


def get_water_it_latest_rows(dataframe):
    if dataframe.empty:
        return dataframe.copy()
    keys = ["エリア", "ポイント", "測定項目"]
    return (
        dataframe.sort_values("測定日時_解析", ascending=False)
        .drop_duplicates(subset=keys, keep="first")
        .reset_index(drop=True)
    )


def normalize_water_it_customer_key(value):
    """WATER itのポイント名と顧客名を安全に照合するための最小正規化。

    表記ゆれを広く吸収すると別顧客を誤って結び付ける可能性があるため、
    Unicode正規化と空白除去だけを行う。
    """
    text = clean_value(value, blank_text="").strip()
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[\s\u3000]+", "", text)
    return text.casefold()


def water_it_display_name(value):
    """WATER itの元データを変えず、画面上だけ名称を統一する。"""
    text = clean_value(value, blank_text="").strip()
    return WATER_IT_POINT_DISPLAY_NAMES.get(text, text)


def canonical_water_it_customer_key(value):
    """明示した別名だけを同一顧客として扱う。曖昧な部分一致は行わない。"""
    display_name = water_it_display_name(value)
    return normalize_water_it_customer_key(display_name)


def get_water_it_customer_rows(dataframe, customer_name):
    """顧客名と対応するWATER itポイントだけを返す（読み取り専用）。"""
    if dataframe is None or dataframe.empty:
        return dataframe.iloc[0:0].copy() if dataframe is not None else pd.DataFrame()
    target = canonical_water_it_customer_key(customer_name)
    if not target:
        return dataframe.iloc[0:0].copy()
    point_keys = dataframe["ポイント"].map(canonical_water_it_customer_key)
    return dataframe[point_keys == target].copy()


def render_customer_water_it_card(customer_name):
    """顧客詳細にWATER itの最新値を読み取り専用で表示する。

    ポイント名と顧客名が一致しない顧客には何も表示しない。
    ExcelやWATER itへの書き込み処理は行わない。
    """
    try:
        dataframe, source = get_active_water_it_data()
    except Exception:
        # WATER it側が一時的に読めなくても、既存の顧客詳細は通常どおり表示する。
        return

    customer_rows = get_water_it_customer_rows(dataframe, customer_name)
    if customer_rows.empty:
        return

    latest_rows = get_water_it_latest_rows(customer_rows)
    if latest_rows.empty:
        return

    newest = latest_rows["測定日時_解析"].max()
    point_names = [
        water_it_display_name(value)
        for value in latest_rows["ポイント"].drop_duplicates().tolist()
    ]
    areas = [
        clean_value(value, blank_text="")
        for value in latest_rows["エリア"].drop_duplicates().tolist()
        if clean_value(value, blank_text="")
    ]

    st.markdown("---")
    with st.container(border=True):
        st.subheader("💧 WATER it タンク情報")
        st.caption(
            "読み取り専用表示です。ここからExcelやWATER itへの書き込みは行いません。"
        )
        st.caption(
            f"ポイント：{' / '.join(point_names)}"
            + (f"　｜　エリア：{' / '.join(areas)}" if areas else "")
            + f"　｜　最終受信：{newest.strftime('%Y/%m/%d %H:%M')}"
        )

        rows = list(latest_rows.iterrows())
        for start in range(0, len(rows), 3):
            group = rows[start:start + 3]
            columns = st.columns(len(group))
            for display_column, (_, row) in zip(columns, group):
                with display_column:
                    label = clean_value(row.get("測定項目"))
                    value = format_water_it_value(row.get("測定値_数値"))
                    unit = normalize_water_it_unit(row.get("単位_表示"))
                    st.metric(label, f"{value} {unit}".strip())
                    st.caption(row["測定日時_解析"].strftime("%m/%d %H:%M"))

        alert_messages = []
        for _, row in latest_rows.iterrows():
            item_name = clean_value(row.get("測定項目"))
            for alert in get_water_it_alerts(row):
                alert_messages.append(f"{item_name}｜{alert}")
        if alert_messages:
            st.warning(" / ".join(alert_messages))
        else:
            st.caption(f"状態：異常表示なし　｜　参照：{source}")


def format_water_it_value(value):
    if value is None or pd.isna(value):
        return "未設定"
    try:
        number = float(value)
    except Exception:
        return clean_value(value)
    if not math.isfinite(number):
        return "未設定"
    if number.is_integer():
        return f"{int(number):,}"
    return f"{number:,.2f}".rstrip("0").rstrip(".")


def get_water_it_alerts(row):
    alerts = []
    for column in WATER_IT_ALERT_COLUMNS:
        if column not in row.index:
            continue
        value = row.get(column)
        if water_it_nonblank(value):
            alerts.append(f"{column}: {clean_value(value)}")
    return alerts


def show_water_it_latest_cards(latest_rows):
    for point in latest_rows["ポイント"].drop_duplicates().tolist():
        point_rows = latest_rows[latest_rows["ポイント"] == point].copy()
        if point_rows.empty:
            continue
        first = point_rows.iloc[0]
        area = clean_value(first.get("エリア"), blank_text="未設定")
        newest = point_rows["測定日時_解析"].max()

        with st.container(border=True):
            st.subheader(f"💧 {water_it_display_name(point)}")
            st.caption(f"エリア：{area}　｜　最新：{newest.strftime('%Y/%m/%d %H:%M')}")

            rows = list(point_rows.iterrows())
            for start in range(0, len(rows), 3):
                group = rows[start:start + 3]
                columns = st.columns(len(group))
                for display_column, (_, row) in zip(columns, group):
                    with display_column:
                        label = clean_value(row.get("測定項目"))
                        value = format_water_it_value(row.get("測定値_数値"))
                        unit = normalize_water_it_unit(row.get("単位_表示"))
                        st.metric(label, f"{value} {unit}".strip())
                        st.caption(row["測定日時_解析"].strftime("%m/%d %H:%M"))

            alert_messages = []
            for _, row in point_rows.iterrows():
                item_name = clean_value(row.get("測定項目"))
                for alert in get_water_it_alerts(row):
                    alert_messages.append(f"{item_name}｜{alert}")
            if alert_messages:
                st.warning(" / ".join(alert_messages))
            else:
                st.caption("状態：異常表示なし")


def show_water_it_history(dataframe):
    st.markdown("### 測定履歴")
    points = dataframe["ポイント"].drop_duplicates().tolist()
    selected_point = st.selectbox(
        "ポイント",
        points,
        key="water_it_history_point",
        format_func=water_it_display_name,
    )
    point_rows = dataframe[dataframe["ポイント"] == selected_point].copy()
    items = point_rows["測定項目"].drop_duplicates().tolist()
    item_key_suffix = hashlib.sha1(selected_point.encode("utf-8")).hexdigest()[:12]
    selected_item = st.selectbox(
        "測定項目",
        items,
        key=f"water_it_history_item_{item_key_suffix}",
    )
    history = point_rows[point_rows["測定項目"] == selected_item].copy()
    history.sort_values("測定日時_解析", inplace=True)

    period = st.radio(
        "表示期間",
        ["24時間", "3日間", "7日間", "すべて"],
        horizontal=True,
        key="water_it_history_period",
    )
    if not history.empty and period != "すべて":
        hours = {"24時間": 24, "3日間": 72, "7日間": 168}[period]
        cutoff = history["測定日時_解析"].max() - timedelta(hours=hours)
        history = history[history["測定日時_解析"] >= cutoff].copy()

    chart_data = history.dropna(subset=["測定値_数値"]).set_index("測定日時_解析")[["測定値_数値"]]
    if chart_data.empty:
        st.info("グラフに表示できる数値データがありません。")
    else:
        chart_data = chart_data.rename(columns={"測定値_数値": selected_item})
        st.line_chart(chart_data, use_container_width=True)

    with st.expander("直近の測定値を表示"):
        display = history.sort_values("測定日時_解析", ascending=False).head(100).copy()
        display["測定日時"] = display["測定日時_解析"].dt.strftime("%Y/%m/%d %H:%M")
        display["測定値"] = display["測定値_数値"].apply(format_water_it_value)
        display["単位"] = display["単位_表示"]
        display["ポイント"] = display["ポイント"].map(water_it_display_name)
        st.dataframe(
            display[["測定日時", "エリア", "ポイント", "測定項目", "測定値", "単位"]],
            use_container_width=True,
            hide_index=True,
        )


def show_water_it_test_page():
    st.markdown("---")
    st.header("💧 WATER it CSV取込・保存")
    show_back_home_button("water_it_back_home")
    st.markdown(
        render_page_link("🧪 ソリュブル在庫", page="soluble_inventory"),
        unsafe_allow_html=True,
    )
    st.caption(
        "スマホでWATER itからCSVを手動ダウンロードし、そのCSVを選ぶだけで画面へ反映し、既存のSupabaseへ自動保存します。Excel・WATER it・Dropboxへの書き込みは行いません。"
    )

    st.link_button(
        "🌐 WATER itを開く",
        WATER_IT_LOGIN_URL,
        use_container_width=True,
    )

    with st.container(border=True):
        st.markdown("#### スマホでの手順")
        st.write("1. 上のボタンからWATER itを開いてログインします。")
        st.write("2. リスト画面の『ダウンロード』を押します。")
        st.write("3. このカルテへ戻り、下の『CSVを選ぶ』を押します。")
        st.write("4. ファイル画面で『最近使用したファイル』または『ダウンロード』を開き、一番新しいCSVを選びます。")

    uploader_version = int(st.session_state.get("water_it_uploader_version", 0))
    uploader_key = f"water_it_mobile_csv_uploader_{uploader_version}"

    temporary_store_before_upload = get_water_it_temporary_store()
    has_selected_water_it_csv = bool(
        st.session_state.get(WATER_IT_UPLOAD_BYTES_KEY)
        or temporary_store_before_upload.get("content")
    )
    if not has_selected_water_it_csv:
        if st.button(
            "CSV選択をリセット",
            key=f"water_it_reset_uploader_{uploader_version}",
            use_container_width=True,
        ):
            clear_uploaded_water_it_csv()
            st.session_state["water_it_uploader_version"] = uploader_version + 1
            st.rerun()

    uploaded_file = st.file_uploader(
        "ダウンロードしたファイルを選ぶ",
        type=None,
        accept_multiple_files=False,
        key=uploader_key,
        on_change=handle_water_it_mobile_upload,
        args=(uploader_key,),
        help=(
            "Androidでは『最近使用したファイル』または『ダウンロード』から選びます。"
            "端末によっては、ファイルをタップしたあとに『開く』『選択』『完了』または右上のチェックを押します。"
        ),
    )
    st.caption(
        "CSVだけに絞るとAndroidで選択が戻らない場合があるため、この版ではファイル種類を絞っていません。"
        "選択後に中身を確認し、WATER it形式のCSVだけを反映します。"
    )

    success_message = st.session_state.pop("water_it_upload_success_message", None)
    error_message = st.session_state.pop("water_it_upload_error_message", None)
    persist_warning = st.session_state.pop("water_it_persist_warning_message", None)
    if success_message:
        st.success(success_message)
    if persist_warning:
        st.warning("CSVは画面へ反映しましたが、Supabaseへの保存は完了していません。")
        st.write(persist_warning)
    if error_message:
        st.error("選択したファイルを読み込めませんでした。")
        st.write(error_message)
        st.info("WATER itのリスト画面からダウンロードしたCSVを選んでください。")

    dataframe = None
    source = ""

    if uploaded_file is not None:
        try:
            uploaded_hash = hashlib.sha256(uploaded_file.getvalue()).hexdigest()
            if uploaded_hash != st.session_state.get(WATER_IT_UPLOAD_HASH_KEY):
                with st.spinner("選択したファイルを確認しています…"):
                    dataframe, source = remember_uploaded_water_it_csv(uploaded_file)
            else:
                dataframe, source = get_active_water_it_data()
            st.caption(
                f"選択済み：{uploaded_file.name or '名前なし'}　"
                f"{len(uploaded_file.getvalue()):,} bytes"
            )
        except Exception as exc:
            st.error("選択したファイルを読み込めませんでした。")
            st.write(str(exc))
            st.info("WATER itのリスト画面からダウンロードしたCSVを選んでください。")
            return
    else:
        try:
            dataframe, source = get_active_water_it_data()
        except Exception as exc:
            st.info("まだCSVが選択されていません。まずWATER itからCSVをダウンロードし、上の欄から選んでください。")
            st.caption(str(exc))
            return

    temporary_store = get_water_it_temporary_store()
    if st.session_state.get(WATER_IT_UPLOAD_BYTES_KEY) or temporary_store.get("content"):
        button_col, note_col = st.columns([1, 2])
        with button_col:
            if st.button("選択中のCSVを解除", key="water_it_clear_upload"):
                clear_uploaded_water_it_csv()
                st.session_state["water_it_uploader_version"] = uploader_version + 1
                st.rerun()
        with note_col:
            st.caption("選択したCSVはSupabaseへ保存されます。ここで選択状態を解除しても、最後に保存したデータは残り、顧客詳細から引き続き確認できます。")

    if dataframe is None or dataframe.empty:
        st.warning("CSVに表示できる測定データがありません。")
        return

    latest_rows = get_water_it_latest_rows(dataframe)
    latest_time = dataframe["測定日時_解析"].max()
    oldest_time = dataframe["測定日時_解析"].min()

    st.success(f"読込OK　｜　参照：{source}")
    metric1, metric2, metric3 = st.columns(3)
    with metric1:
        st.metric("最新測定日時", latest_time.strftime("%Y/%m/%d %H:%M"))
    with metric2:
        st.metric("ポイント数", f"{dataframe['ポイント'].nunique()}件")
    with metric3:
        st.metric("読込行数", f"{len(dataframe):,}行")
    st.caption(
        f"データ期間：{oldest_time.strftime('%Y/%m/%d %H:%M')} ～ {latest_time.strftime('%Y/%m/%d %H:%M')}"
    )

    show_water_it_latest_cards(latest_rows)
    show_water_it_history(dataframe)


# =========================
# 仕入先・運送会社（取引先カルテ.xlsx）
# =========================
TRADE_PARTNER_HEADER_ALIASES = {
    "連絡方法レンラクホウホウ": "連絡方法",
    "納品先ノウヒンサキ": "納品先",
    "運賃ウンチン": "運賃",
    "地域チイキ": "地域",
}
TRADE_PARTNER_REQUIRED_SHEETS = (
    TRADE_PARTNER_MASTER_SHEET,
    TRADE_PARTNER_CONTACT_SHEET,
    TRADE_PARTNER_PRODUCT_SHEET,
    TRADE_PARTNER_TRANSPORT_SHEET,
)
TRADE_PARTNER_ID_FIELDS = {
    TRADE_PARTNER_MASTER_SHEET: "取引先ID",
    TRADE_PARTNER_CONTACT_SHEET: "担当者ID",
    TRADE_PARTNER_PRODUCT_SHEET: "仕入商品ID",
    TRADE_PARTNER_TRANSPORT_SHEET: "運送条件ID",
}
TRADE_PARTNER_PRIMARY_FIELDS = {
    TRADE_PARTNER_MASTER_SHEET: "会社名",
    TRADE_PARTNER_CONTACT_SHEET: "担当者名",
    TRADE_PARTNER_PRODUCT_SHEET: "商品名",
    TRADE_PARTNER_TRANSPORT_SHEET: "納品先",
}
TRADE_PARTNER_NOTE_PREFIXES = {
    "supplier": "【仕入先】",
    "carrier": "【運送会社】",
}


def normalize_trade_partner_header(value):
    text = clean_value(value, blank_text="").strip()
    return TRADE_PARTNER_HEADER_ALIASES.get(text, text)


def trade_partner_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y/%m/%d")
    text = str(value).strip()
    if text.startswith("="):
        return ""
    return text


def is_trade_partner_marked(value):
    text = trade_partner_text(value).strip().lower()
    return text in {"○", "〇", "1", "true", "yes", "有", "あり"}


def trade_partner_type_label(partner_type):
    return "仕入先" if partner_type == "supplier" else "運送会社"


def trade_partner_home_page(partner_type):
    return "supplier_home" if partner_type == "supplier" else "carrier_home"


def trade_partner_list_page(partner_type):
    return "supplier_list" if partner_type == "supplier" else "carrier_list"


def trade_partner_search_page(partner_type):
    return "supplier_search" if partner_type == "supplier" else "carrier_search"


def trade_partner_category_field(partner_type):
    return "仕入先区分" if partner_type == "supplier" else "運送会社区分"


def get_trade_partner_file_path():
    path = str(TRADE_PARTNER_DROPBOX_FILE_PATH or "").strip()
    return path or TRADE_PARTNER_DROPBOX_DEFAULT_FILE_PATH


XLSX_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XLSX_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XLSX_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XLSX_XML_NS = "http://www.w3.org/XML/1998/namespace"

# ElementTreeは、要素名・属性名で直接使われていない名前空間宣言を
# 再シリアライズ時に省略する。Excelはmc:Ignorableに記載された接頭辞の
# xmlns宣言が欠けるとワークシートを破損扱いにするため、宣言一覧を明示的に保持する。
XLSX_NAMESPACE_DECLARATIONS = {
    "": XLSX_MAIN_NS,
    "r": XLSX_REL_NS,
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "xm": "http://schemas.microsoft.com/office/excel/2006/main",
}

# Excelの拡張データ検証やリビジョン情報を壊さないよう、元と同じ名前空間接頭辞を保つ。
for _prefix, _namespace in XLSX_NAMESPACE_DECLARATIONS.items():
    try:
        ET.register_namespace(_prefix, _namespace)
    except Exception:
        pass


def xlsx_tag(local_name):
    return f"{{{XLSX_MAIN_NS}}}{local_name}"


def xlsx_column_name(column_number):
    result = ""
    number = int(column_number)
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def xlsx_column_number(cell_reference):
    match = re.match(r"^([A-Z]+)", str(cell_reference or "").upper())
    if not match:
        return 0
    result = 0
    for char in match.group(1):
        result = result * 26 + (ord(char) - 64)
    return result


def extract_worksheet_namespace_declarations(xml_content):
    """worksheetルートにあるxmlns宣言を接頭辞別に取り出す。"""
    try:
        text = xml_content.decode("utf-8")
    except Exception:
        return {}
    root_match = re.search(r"<worksheet\b[^>]*>", text, re.DOTALL)
    if not root_match:
        return {}
    declarations = {}
    for prefix, uri in re.findall(
        r"\bxmlns(?::([A-Za-z_][\w.-]*))?=[\"']([^\"']+)[\"']",
        root_match.group(0),
    ):
        declarations[prefix or ""] = uri
    return declarations


def ensure_worksheet_namespace_declarations(xml_content, original_declarations=None):
    """
    元のxmlns宣言とmc:Ignorableが要求する宣言をworksheetへ戻す。

    ElementTreeはmc:Ignorableの値にだけ登場するx14ac/xr2/xr3などを
    未使用と判断して削るため、Excelで開く前に必ず宣言を復元する。
    """
    try:
        text = xml_content.decode("utf-8")
    except Exception as error:
        raise ValueError("ワークシートXMLをUTF-8として確認できません。") from error

    root_match = re.search(r"<worksheet\b[^>]*>", text, re.DOTALL)
    if not root_match:
        raise ValueError("ワークシートXMLのルート要素を確認できません。")

    root_tag = root_match.group(0)
    required = dict(original_declarations or {})
    ignorable_match = re.search(
        r"\bmc:Ignorable=[\"']([^\"']*)[\"']",
        root_tag,
    )
    if ignorable_match:
        for prefix in ignorable_match.group(1).split():
            uri = XLSX_NAMESPACE_DECLARATIONS.get(prefix)
            if uri:
                required.setdefault(prefix, uri)

    additions = []
    for prefix, uri in required.items():
        attribute_name = "xmlns" if not prefix else f"xmlns:{prefix}"
        if re.search(
            rf"\b{re.escape(attribute_name)}=[\"']",
            root_tag,
        ):
            continue
        additions.append(f' {attribute_name}="{uri}"')

    if additions:
        root_tag = root_tag[:-1] + "".join(additions) + ">"
        text = text[:root_match.start()] + root_tag + text[root_match.end():]

    return text.encode("utf-8")


def missing_worksheet_ignorable_namespaces(xml_content):
    """mc:Ignorableにあるのにxmlns宣言がない接頭辞を返す。"""
    try:
        text = xml_content.decode("utf-8")
    except Exception:
        return ["XMLをUTF-8として読めません"]
    root_match = re.search(r"<worksheet\b[^>]*>", text, re.DOTALL)
    if not root_match:
        return ["worksheetルートがありません"]
    root_tag = root_match.group(0)
    ignorable_match = re.search(
        r"\bmc:Ignorable=[\"']([^\"']*)[\"']",
        root_tag,
    )
    if not ignorable_match:
        return []
    declared = extract_worksheet_namespace_declarations(xml_content)
    return [
        prefix
        for prefix in ignorable_match.group(1).split()
        if prefix not in declared
    ]


def remove_calc_chain_relationship(xml_content):
    """workbook.xml.relsから古いcalcChain参照だけを取り除く。"""
    try:
        text = xml_content.decode("utf-8")
    except Exception as error:
        raise ValueError("Excelの計算関係情報を読み取れませんでした。") from error

    relationship_pattern = re.compile(
        r"<(?:[A-Za-z_][\w.-]*:)?Relationship\b"
        r"(?=[^>]*(?:"
        r"\bType\s*=\s*[\"'][^\"']*/calcChain[\"']"
        r"|\bTarget\s*=\s*[\"'][^\"']*calcChain\.xml[\"']"
        r"))[^>]*(?:/>|>\s*</(?:[A-Za-z_][\w.-]*:)?Relationship\s*>)",
        re.IGNORECASE | re.DOTALL,
    )
    return relationship_pattern.sub("", text).encode("utf-8")


def remove_calc_chain_content_type(xml_content):
    """[Content_Types].xmlからcalcChainの登録だけを取り除く。"""
    try:
        text = xml_content.decode("utf-8")
    except Exception as error:
        raise ValueError("Excelのコンテンツ種類情報を読み取れませんでした。") from error

    override_pattern = re.compile(
        r"<(?:[A-Za-z_][\w.-]*:)?Override\b"
        r"(?=[^>]*\bPartName\s*=\s*[\"']/xl/calcChain\.xml[\"'])"
        r"[^>]*(?:/>|>\s*</(?:[A-Za-z_][\w.-]*:)?Override\s*>)",
        re.IGNORECASE | re.DOTALL,
    )
    return override_pattern.sub("", text).encode("utf-8")


def set_xml_tag_attribute(tag_text, attribute_name, value):
    """XML開始タグの既存属性を更新し、なければ末尾へ追加する。"""
    pattern = re.compile(
        rf"(\s{re.escape(attribute_name)}\s*=\s*)([\"'])(.*?)(\2)",
        re.IGNORECASE | re.DOTALL,
    )
    if pattern.search(tag_text):
        return pattern.sub(
            lambda match: match.group(1) + match.group(2) + str(value) + match.group(2),
            tag_text,
            count=1,
        )

    closing = "/>" if tag_text.rstrip().endswith("/>") else ">"
    position = tag_text.rfind(closing)
    if position < 0:
        raise ValueError("Excelの再計算設定を更新できませんでした。")
    return tag_text[:position] + f' {attribute_name}="{value}"' + tag_text[position:]


def force_workbook_recalculation(xml_content):
    """Excelを開いた時に数式を自動で全再計算する設定へ更新する。"""
    try:
        text = xml_content.decode("utf-8")
    except Exception as error:
        raise ValueError("Excelのブック設定を読み取れませんでした。") from error

    calc_pattern = re.compile(
        r"<(?:[A-Za-z_][\w.-]*:)?calcPr\b[^>]*>",
        re.IGNORECASE | re.DOTALL,
    )
    match = calc_pattern.search(text)
    if match:
        tag = match.group(0)
        for name, value in (
            ("calcMode", "auto"),
            ("fullCalcOnLoad", "1"),
            ("forceFullCalc", "1"),
        ):
            tag = set_xml_tag_attribute(tag, name, value)
        text = text[:match.start()] + tag + text[match.end():]
    else:
        closing_match = re.search(
            r"</(?:[A-Za-z_][\w.-]*:)?workbook\s*>",
            text,
            re.IGNORECASE,
        )
        if not closing_match:
            raise ValueError("Excelのブック設定にworkbook要素が見つかりません。")
        calc_tag = (
            '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
        )
        text = text[:closing_match.start()] + calc_tag + text[closing_match.start():]

    return text.encode("utf-8")


def workbook_recalculation_is_forced(xml_content):
    """保存後のworkbook.xmlに全再計算設定があるか確認する。"""
    try:
        text = xml_content.decode("utf-8")
    except Exception:
        return False
    match = re.search(
        r"<(?:[A-Za-z_][\w.-]*:)?calcPr\b[^>]*>",
        text,
        re.IGNORECASE | re.DOTALL,
    )
    if not match:
        return False
    tag = match.group(0)
    required = {
        "calcMode": "auto",
        "fullCalcOnLoad": "1",
        "forceFullCalc": "1",
    }
    for name, expected in required.items():
        attribute = re.search(
            rf"\b{re.escape(name)}\s*=\s*[\"']([^\"']*)[\"']",
            tag,
            re.IGNORECASE,
        )
        if not attribute or attribute.group(1).lower() != expected.lower():
            return False
    return True


class TradePartnerXlsxEditor:
    """セル値だけをXMLで差し替え、入力規則・書式・数式をそのまま保持する。"""

    def __init__(self, content):
        self.original_infos = []
        self.parts = {}
        with zipfile.ZipFile(BytesIO(content), "r") as archive:
            self.original_infos = archive.infolist()
            for info in self.original_infos:
                self.parts[info.filename] = archive.read(info.filename)

        self.shared_strings = self._read_shared_strings()
        self.sheet_paths = self._read_sheet_paths()
        self.sheet_namespace_declarations = {
            path: extract_worksheet_namespace_declarations(self.parts[path])
            for path in self.sheet_paths.values()
            if path in self.parts
        }
        self.sheet_roots = {}
        self.changed_sheet_names = set()

    def _read_shared_strings(self):
        path = "xl/sharedStrings.xml"
        if path not in self.parts:
            return []
        root = ET.fromstring(self.parts[path])
        result = []
        for item in root.findall(xlsx_tag("si")):
            texts = []
            direct_text = item.find(xlsx_tag("t"))
            if direct_text is not None:
                texts.append(direct_text.text or "")
            for run in item.findall(xlsx_tag("r")):
                run_text = run.find(xlsx_tag("t"))
                if run_text is not None:
                    texts.append(run_text.text or "")
            result.append("".join(texts))
        return result

    def _read_sheet_paths(self):
        workbook_root = ET.fromstring(self.parts["xl/workbook.xml"])
        relation_root = ET.fromstring(self.parts["xl/_rels/workbook.xml.rels"])
        relation_map = {
            relation.attrib.get("Id"): relation.attrib.get("Target", "")
            for relation in relation_root.findall(f"{{{XLSX_PACKAGE_REL_NS}}}Relationship")
        }
        result = {}
        sheets = workbook_root.find(xlsx_tag("sheets"))
        if sheets is None:
            return result
        relation_attribute = f"{{{XLSX_REL_NS}}}id"
        for sheet in sheets.findall(xlsx_tag("sheet")):
            name = sheet.attrib.get("name", "")
            target = relation_map.get(sheet.attrib.get(relation_attribute), "")
            if not target:
                continue
            if target.startswith("/"):
                path = target.lstrip("/")
            else:
                path = posixpath.normpath(posixpath.join("xl", target))
            result[name] = path
        return result

    def has_sheet(self, sheet_name):
        return sheet_name in self.sheet_paths

    def get_sheet_root(self, sheet_name):
        if sheet_name not in self.sheet_paths:
            raise ValueError(f"{sheet_name}シートがありません。")
        if sheet_name not in self.sheet_roots:
            self.sheet_roots[sheet_name] = ET.fromstring(self.parts[self.sheet_paths[sheet_name]])
        return self.sheet_roots[sheet_name]

    def get_sheet_data(self, sheet_name):
        root = self.get_sheet_root(sheet_name)
        sheet_data = root.find(xlsx_tag("sheetData"))
        if sheet_data is None:
            sheet_data = ET.SubElement(root, xlsx_tag("sheetData"))
        return sheet_data

    def _display_inline_string(self, container):
        if container is None:
            return ""
        texts = []
        direct_text = container.find(xlsx_tag("t"))
        if direct_text is not None:
            texts.append(direct_text.text or "")
        for run in container.findall(xlsx_tag("r")):
            run_text = run.find(xlsx_tag("t"))
            if run_text is not None:
                texts.append(run_text.text or "")
        return "".join(texts)

    def cell_value_from_element(self, cell):
        if cell is None:
            return None
        formula = cell.find(xlsx_tag("f"))
        if formula is not None:
            return "=" + (formula.text or "")
        cell_type = cell.attrib.get("t", "")
        if cell_type == "inlineStr":
            return self._display_inline_string(cell.find(xlsx_tag("is")))
        value_element = cell.find(xlsx_tag("v"))
        raw_value = value_element.text if value_element is not None else None
        if raw_value is None:
            return None
        if cell_type == "s":
            try:
                return self.shared_strings[int(raw_value)]
            except Exception:
                return ""
        if cell_type == "b":
            return raw_value == "1"
        if cell_type in {"str", "e"}:
            return raw_value
        try:
            number = float(raw_value)
            return int(number) if number.is_integer() else number
        except Exception:
            return raw_value

    def get_row_element(self, sheet_name, row_number, create=False):
        sheet_data = self.get_sheet_data(sheet_name)
        target = int(row_number)
        rows = list(sheet_data.findall(xlsx_tag("row")))
        for row in rows:
            if int(row.attrib.get("r", "0") or 0) == target:
                return row
        if not create:
            return None
        new_row = ET.Element(xlsx_tag("row"), {"r": str(target)})
        inserted = False
        for index, row in enumerate(rows):
            current = int(row.attrib.get("r", "0") or 0)
            if current > target:
                sheet_data.insert(index, new_row)
                inserted = True
                break
        if not inserted:
            sheet_data.append(new_row)
        self.changed_sheet_names.add(sheet_name)
        return new_row

    def get_cell_element(self, sheet_name, row_number, column_number, create=False):
        row = self.get_row_element(sheet_name, row_number, create=create)
        if row is None:
            return None
        reference = f"{xlsx_column_name(column_number)}{int(row_number)}"
        cells = list(row.findall(xlsx_tag("c")))
        for cell in cells:
            if cell.attrib.get("r") == reference:
                return cell
        if not create:
            return None
        new_cell = ET.Element(xlsx_tag("c"), {"r": reference})
        inserted = False
        target_column = int(column_number)
        for index, cell in enumerate(cells):
            if xlsx_column_number(cell.attrib.get("r")) > target_column:
                row.insert(index, new_cell)
                inserted = True
                break
        if not inserted:
            row.append(new_cell)
        self.changed_sheet_names.add(sheet_name)
        return new_cell

    def get_cell_value(self, sheet_name, row_number, column_number):
        return self.cell_value_from_element(
            self.get_cell_element(sheet_name, row_number, column_number, create=False)
        )

    def set_cell_value(self, sheet_name, row_number, column_number, value):
        cell = self.get_cell_element(sheet_name, row_number, column_number, create=True)
        for child_name in ("f", "v", "is"):
            child = cell.find(xlsx_tag(child_name))
            if child is not None:
                cell.remove(child)
        if value is None or value == "":
            cell.attrib.pop("t", None)
        elif isinstance(value, bool):
            cell.attrib["t"] = "b"
            ET.SubElement(cell, xlsx_tag("v")).text = "1" if value else "0"
        elif isinstance(value, (int, float)) and not isinstance(value, bool):
            cell.attrib.pop("t", None)
            ET.SubElement(cell, xlsx_tag("v")).text = str(value)
        else:
            text = value.strftime("%Y/%m/%d") if isinstance(value, (datetime, date)) else str(value)
            cell.attrib["t"] = "inlineStr"
            inline = ET.SubElement(cell, xlsx_tag("is"))
            text_element = ET.SubElement(inline, xlsx_tag("t"))
            if text != text.strip() or "\n" in text:
                text_element.attrib[f"{{{XLSX_XML_NS}}}space"] = "preserve"
            text_element.text = text
        self.changed_sheet_names.add(sheet_name)

    def get_max_row(self, sheet_name):
        maximum = 1
        for row in self.get_sheet_data(sheet_name).findall(xlsx_tag("row")):
            try:
                maximum = max(maximum, int(row.attrib.get("r", "0") or 0))
            except Exception:
                pass
        return maximum

    def get_header_map(self, sheet_name):
        result = {}
        row = self.get_row_element(sheet_name, 1, create=False)
        if row is None:
            return result
        for cell in row.findall(xlsx_tag("c")):
            column = xlsx_column_number(cell.attrib.get("r"))
            header = normalize_trade_partner_header(self.cell_value_from_element(cell))
            if header and header not in result:
                result[header] = column
        return result

    def read_sheet(self, sheet_name):
        header_map = self.get_header_map(sheet_name)
        headers = list(header_map.keys())
        rows = []
        for row_number in range(2, self.get_max_row(sheet_name) + 1):
            row = {
                header: self.get_cell_value(sheet_name, row_number, column)
                for header, column in header_map.items()
            }
            row["_row_number"] = row_number
            rows.append(row)
        return {"headers": headers, "rows": rows}

    def validate_worksheet_namespaces(self):
        problems = []
        for sheet_name, path in self.sheet_paths.items():
            if path not in self.parts:
                problems.append(f"{sheet_name}: XMLがありません")
                continue
            missing = missing_worksheet_ignorable_namespaces(self.parts[path])
            if missing:
                problems.append(f"{sheet_name}: " + "、".join(missing))
        if problems:
            raise ValueError(
                "Excel互換性の確認で名前空間宣言が不足しています。保存を中止しました。\n"
                + "\n".join(problems)
            )

    def remove_stale_calculation_chain(self):
        """セル変更前のcalcChainを除去し、Excelへ安全に再計算させる。"""
        self.parts.pop("xl/calcChain.xml", None)

        relationships_path = "xl/_rels/workbook.xml.rels"
        if relationships_path in self.parts:
            self.parts[relationships_path] = remove_calc_chain_relationship(
                self.parts[relationships_path]
            )

        content_types_path = "[Content_Types].xml"
        if content_types_path in self.parts:
            self.parts[content_types_path] = remove_calc_chain_content_type(
                self.parts[content_types_path]
            )

        workbook_path = "xl/workbook.xml"
        if workbook_path not in self.parts:
            raise ValueError("Excelのworkbook.xmlが見つかりません。")
        self.parts[workbook_path] = force_workbook_recalculation(
            self.parts[workbook_path]
        )

    def validate_calculation_state(self):
        """古いcalcChain参照が残らず、全再計算設定が有効か確認する。"""
        if "xl/calcChain.xml" in self.parts:
            raise ValueError("Excelの古い計算順序情報が残っています。保存を中止しました。")

        relationships = self.parts.get("xl/_rels/workbook.xml.rels", b"")
        if re.search(
            rb'(?:relationships/calcChain|Target\s*=\s*["\'][^"\']*calcChain\.xml)',
            relationships,
            re.IGNORECASE,
        ):
            raise ValueError("Excelの計算順序への参照が残っています。保存を中止しました。")

        content_types = self.parts.get("[Content_Types].xml", b"")
        if re.search(rb"/xl/calcChain\.xml", content_types, re.IGNORECASE):
            raise ValueError("Excelの計算順序の種類登録が残っています。保存を中止しました。")

        workbook = self.parts.get("xl/workbook.xml", b"")
        if not workbook_recalculation_is_forced(workbook):
            raise ValueError("Excelの自動再計算設定を確認できません。保存を中止しました。")

    def to_bytes(self):
        for sheet_name in self.changed_sheet_names:
            path = self.sheet_paths[sheet_name]
            serialized = ET.tostring(
                self.get_sheet_root(sheet_name),
                encoding="utf-8",
                xml_declaration=True,
            )
            self.parts[path] = ensure_worksheet_namespace_declarations(
                serialized,
                self.sheet_namespace_declarations.get(path),
            )

        # 以前のアプリ保存で名前空間宣言が欠けたファイルも、次の保存時に
        # 全ワークシートを安全な状態へ戻す。セル値や書式・数式は変更しない。
        for path in set(self.sheet_paths.values()):
            if path not in self.parts:
                continue
            self.parts[path] = ensure_worksheet_namespace_declarations(
                self.parts[path],
                self.sheet_namespace_declarations.get(path),
            )

        self.validate_worksheet_namespaces()

        # セルを書き換えた後に古いcalcChainを残すと、Excelが修復画面を出す。
        # 本体・関連付け・Content Typesの3か所をそろえて除去し、開いた時に
        # Excel自身が数式を全再計算する設定へ更新する。
        self.remove_stale_calculation_chain()
        self.validate_calculation_state()

        output = BytesIO()
        with zipfile.ZipFile(output, "w") as archive:
            for info in self.original_infos:
                # calcChain.xmlは意図的に削除しているため、元のZIP一覧にあっても書き戻さない。
                if info.filename not in self.parts:
                    continue
                archive.writestr(info, self.parts[info.filename])
        return output.getvalue()


@st.cache_data(ttl=60, show_spinner=False)
def load_trade_partner_data():
    access_token = get_dropbox_access_token()
    path = get_trade_partner_file_path()
    content, response = download_dropbox_file(path, access_token)
    if content is None:
        raise RuntimeError(
            "取引先カルテ.xlsxをDropboxから取得できませんでした。\n"
            + dropbox_error_text(response)
        )

    editor = TradePartnerXlsxEditor(content)
    missing = [name for name in TRADE_PARTNER_REQUIRED_SHEETS if not editor.has_sheet(name)]
    if missing:
        raise RuntimeError("取引先カルテ.xlsxに必要なシートがありません：" + "、".join(missing))
    return {name: editor.read_sheet(name) for name in TRADE_PARTNER_REQUIRED_SHEETS}

def get_trade_partner_master_rows(data, partner_type=None):
    rows = []
    for row in data[TRADE_PARTNER_MASTER_SHEET]["rows"]:
        if not trade_partner_text(row.get("会社名")):
            continue
        if partner_type and not is_trade_partner_marked(row.get(trade_partner_category_field(partner_type))):
            continue
        rows.append(row)
    return rows


def get_trade_partner_by_id(data, partner_id):
    target = str(partner_id or "").strip()
    for row in get_trade_partner_master_rows(data):
        if trade_partner_text(row.get("取引先ID")) == target:
            return row
    return None


def get_trade_partner_related_rows(data, sheet_name, partner_id):
    target = str(partner_id or "").strip()
    primary_field = TRADE_PARTNER_PRIMARY_FIELDS[sheet_name]
    result = []
    for row in data[sheet_name]["rows"]:
        if trade_partner_text(row.get("取引先ID")) != target:
            continue
        if not trade_partner_text(row.get(primary_field)):
            continue
        result.append(row)
    return result


def trade_partner_sort_key(row):
    kana = trade_partner_text(row.get("会社名かな"))
    company = trade_partner_text(row.get("会社名"))
    return (kana or company, company)


def make_trade_partner_note_key(partner_type, partner_id, company_name=None):
    """会社名が変わってもメモが外れないよう、区分と取引先IDだけで紐づける。"""
    prefix = TRADE_PARTNER_NOTE_PREFIXES.get(partner_type, "【取引先】")
    return f"{prefix}{partner_id}"


def parse_trade_partner_note_key(value):
    text = trade_partner_text(value)
    for partner_type, prefix in TRADE_PARTNER_NOTE_PREFIXES.items():
        if text.startswith(prefix):
            body = text[len(prefix):]
            partner_id, separator, company = body.partition("|")
            return {
                "partner_type": partner_type,
                "partner_id": partner_id.strip(),
                "company_name": company.strip() if separator else "",
            }
    return None


def ensure_trade_partner_backup_folder(access_token):
    response = call_dropbox_rpc(
        "files/create_folder_v2",
        {"path": TRADE_PARTNER_BACKUP_FOLDER, "autorename": False},
        access_token,
    )
    if response.status_code == 200:
        return
    try:
        summary = str(response.json().get("error_summary", "")).lower()
    except Exception:
        summary = str(getattr(response, "text", "")).lower()
    if "conflict" in summary and "folder" in summary:
        return
    raise RuntimeError(
        "取引先カルテのバックアップフォルダを作成できませんでした。\n"
        + dropbox_error_text(response)
    )


def create_trade_partner_backup(target_path, backup_path, original_content, access_token):
    ensure_trade_partner_backup_folder(access_token)
    copy_response = copy_dropbox_file(target_path, backup_path, access_token)
    if copy_response.status_code == 200:
        metadata = get_dropbox_response_metadata(copy_response)
        if not metadata.get("content_hash") or metadata.get("size") is None:
            metadata = get_dropbox_file_metadata(backup_path, access_token)
        try:
            verify_dropbox_file_metadata(metadata, original_content)
            return
        except Exception:
            call_dropbox_rpc("files/delete_v2", {"path": backup_path}, access_token)
            raise RuntimeError(
                "取引先カルテ.xlsxが別の端末で更新された可能性があります。再読み込みしてやり直してください。"
            )

    backup_response = upload_dropbox_file(
        backup_path,
        original_content,
        access_token,
        mode="add",
    )
    if backup_response.status_code != 200:
        raise RuntimeError(
            "取引先カルテのバックアップを作成できないため、本番ファイルは更新しません。\n"
            + dropbox_error_text(backup_response)
        )
    metadata = get_dropbox_response_metadata(backup_response)
    if not metadata.get("content_hash") or metadata.get("size") is None:
        metadata = get_dropbox_file_metadata(backup_path, access_token)
    verify_dropbox_file_metadata(metadata, original_content)


def trim_trade_partner_backups(access_token, keep=30):
    response = call_dropbox_rpc(
        "files/list_folder",
        {"path": TRADE_PARTNER_BACKUP_FOLDER, "recursive": False, "include_deleted": False},
        access_token,
    )
    if response.status_code != 200:
        return
    try:
        entries = list(response.json().get("entries", []))
    except Exception:
        return
    pattern = re.compile(r"^取引先カルテ_\d{8}_\d{6}_\d+\.xlsx$")
    files = [item for item in entries if pattern.match(str(item.get("name", "")))]
    files.sort(key=lambda item: str(item.get("server_modified", "")), reverse=True)
    for item in files[keep:]:
        path = item.get("path_lower") or item.get("path_display")
        if path:
            call_dropbox_rpc("files/delete_v2", {"path": path}, access_token)


def save_trade_partner_workbook(mutator):
    access_token = get_dropbox_access_token()
    target_path = get_trade_partner_file_path()
    original_content, download_response = download_dropbox_file(target_path, access_token)
    if original_content is None:
        raise RuntimeError(
            "最新の取引先カルテ.xlsxを取得できませんでした。\n"
            + dropbox_error_text(download_response)
        )
    revision = get_download_revision(download_response)
    if not revision:
        raise RuntimeError("Dropboxの更新番号を確認できないため、安全のため保存を中止しました。")

    timestamp = get_jst_now().strftime("%Y%m%d_%H%M%S_%f")
    backup_path = f"{TRADE_PARTNER_BACKUP_FOLDER}/取引先カルテ_{timestamp}.xlsx"
    create_trade_partner_backup(
        target_path,
        backup_path,
        original_content,
        access_token,
    )

    editor = TradePartnerXlsxEditor(original_content)
    missing = [name for name in TRADE_PARTNER_REQUIRED_SHEETS if not editor.has_sheet(name)]
    if missing:
        raise ValueError("必要なシートがありません：" + "、".join(missing))
    result = mutator(editor)
    saved_content = editor.to_bytes()

    # XML更新後もブック構造と入力規則の拡張部分が残っていることを確認する。
    verified = TradePartnerXlsxEditor(saved_content)
    missing = [name for name in TRADE_PARTNER_REQUIRED_SHEETS if not verified.has_sheet(name)]
    if missing:
        raise ValueError("保存後の検証で必要なシートがありません：" + "、".join(missing))
    for sheet_name in TRADE_PARTNER_REQUIRED_SHEETS:
        if not verified.get_header_map(sheet_name):
            raise ValueError(f"保存後の検証で{sheet_name}の見出しを確認できません。")

    upload_response = upload_dropbox_file(
        target_path,
        saved_content,
        access_token,
        mode="update",
        rev=revision,
    )
    if upload_response.status_code != 200:
        raise RuntimeError(
            "取引先カルテ.xlsxを更新できませんでした。\n"
            + dropbox_error_text(upload_response)
        )
    metadata = get_dropbox_response_metadata(upload_response)
    if not metadata.get("content_hash") or metadata.get("size") is None:
        metadata = get_dropbox_file_metadata(target_path, access_token)
    verify_dropbox_file_metadata(metadata, saved_content, previous_revision=revision)
    trim_trade_partner_backups(access_token, keep=30)
    st.cache_data.clear()
    return result

def trade_partner_input_value(header, value):
    text = str(value or "").strip()
    if not text:
        return None
    if header in {"単価", "運賃"}:
        normalized = text.replace(",", "").translate(
            str.maketrans("０１２３４５６７８９．－", "0123456789.-")
        )
        try:
            number = float(normalized)
            return int(number) if number.is_integer() else number
        except Exception:
            return text
    return text


def update_trade_partner_row(sheet_name, record_id, values):
    id_field = TRADE_PARTNER_ID_FIELDS[sheet_name]
    target_id = str(record_id or "").strip()

    def mutator(editor):
        header_map = editor.get_header_map(sheet_name)
        if id_field not in header_map:
            raise ValueError(f"{sheet_name}に{id_field}列がありません。")
        target_row = None
        for row_number in range(2, editor.get_max_row(sheet_name) + 1):
            current_id = trade_partner_text(
                editor.get_cell_value(sheet_name, row_number, header_map[id_field])
            )
            if current_id == target_id:
                target_row = row_number
                break
        if target_row is None:
            raise ValueError(f"{sheet_name}で対象IDが見つかりません。")
        changes = {}
        for header, value in values.items():
            if header not in header_map or header == id_field or header == "会社名（確認用）":
                continue
            old_value = editor.get_cell_value(sheet_name, target_row, header_map[header])
            new_value = trade_partner_input_value(header, value)
            if not same_excel_value(old_value, new_value):
                editor.set_cell_value(sheet_name, target_row, header_map[header], new_value)
                changes[header] = (old_value, new_value)
        if not changes:
            raise ValueError("変更された項目がありません。")
        partner_id = target_id
        if sheet_name != TRADE_PARTNER_MASTER_SHEET and "取引先ID" in header_map:
            partner_id = trade_partner_text(
                editor.get_cell_value(sheet_name, target_row, header_map["取引先ID"])
            )
        return {
            "record_id": target_id,
            "partner_id": partner_id,
            "changed": len(changes),
            "changes": changes,
        }

    return save_trade_partner_workbook(mutator)


def find_trade_partner_empty_row(editor, sheet_name, header_map, primary_field):
    if primary_field not in header_map:
        raise ValueError(f"{sheet_name}に{primary_field}列がありません。")
    id_field = TRADE_PARTNER_ID_FIELDS[sheet_name]
    if id_field not in header_map:
        raise ValueError(f"{sheet_name}に{id_field}列がありません。")
    for row_number in range(2, editor.get_max_row(sheet_name) + 1):
        primary = trade_partner_text(
            editor.get_cell_value(sheet_name, row_number, header_map[primary_field])
        )
        record_id = trade_partner_text(
            editor.get_cell_value(sheet_name, row_number, header_map[id_field])
        )
        if not primary and record_id:
            return row_number, record_id
    raise ValueError(
        f"{sheet_name}に登録用の空き行がありません。ExcelでID付きの空き行を追加してください。"
    )


def create_trade_partner_record(sheet_name, values):
    primary_field = TRADE_PARTNER_PRIMARY_FIELDS[sheet_name]
    id_field = TRADE_PARTNER_ID_FIELDS[sheet_name]

    def mutator(editor):
        header_map = editor.get_header_map(sheet_name)
        row_number, record_id = find_trade_partner_empty_row(
            editor,
            sheet_name,
            header_map,
            primary_field,
        )
        if sheet_name == TRADE_PARTNER_MASTER_SHEET:
            company = trade_partner_text(values.get("会社名"))
            if not company:
                raise ValueError("会社名を入力してください。")
            if "会社名" not in header_map:
                raise ValueError("取引先マスターに会社名列がありません。")
            for check_row in range(2, editor.get_max_row(sheet_name) + 1):
                existing = trade_partner_text(
                    editor.get_cell_value(sheet_name, check_row, header_map["会社名"])
                )
                if existing and normalize_match_value(existing) == normalize_match_value(company):
                    raise ValueError("同じ会社名がすでに登録されています。")
        changes = {}
        for header, value in values.items():
            if header not in header_map or header in {id_field, "会社名（確認用）"}:
                continue
            new_value = trade_partner_input_value(header, value)
            editor.set_cell_value(
                sheet_name,
                row_number,
                header_map[header],
                new_value,
            )
            if new_value not in (None, ""):
                changes[header] = ("", new_value)
        partner_id = record_id
        if sheet_name != TRADE_PARTNER_MASTER_SHEET:
            partner_id = trade_partner_text(values.get("取引先ID"))
        return {
            "record_id": record_id,
            "partner_id": partner_id,
            "row_number": row_number,
            "changes": changes,
        }

    return save_trade_partner_workbook(mutator)

def show_top_home_link():
    st.markdown(render_page_link("← トップへ戻る", page="home"), unsafe_allow_html=True)


def show_trade_partner_home_link(partner_type):
    st.markdown(
        render_page_link(
            f"← {trade_partner_type_label(partner_type)}メニューへ戻る",
            page=trade_partner_home_page(partner_type),
        ),
        unsafe_allow_html=True,
    )


def trade_partner_detail_link(row, partner_type, label=None, class_name="dispatch-month-link"):
    partner_id = trade_partner_text(row.get("取引先ID"))
    company = trade_partner_text(row.get("会社名"))
    return render_page_link(
        label or company,
        page="partner_detail",
        partner_id=partner_id,
        partner_type=partner_type,
        class_name=class_name,
    )


def render_trade_partner_directory_cards(rows, partner_type):
    """仕入先一覧・検索結果を、顧客名一覧と同じ押しやすいカードで表示する。"""
    parts = ['<div class="customer-directory">']
    for row in rows:
        partner_id = trade_partner_text(row.get("取引先ID"))
        company = trade_partner_text(row.get("会社名")) or "名称未設定"
        region = trade_partner_text(row.get("地域")) or "未設定"
        url = html.escape(
            make_app_url(
                page="partner_detail",
                partner_id=partner_id,
                partner_type=partner_type,
            ),
            quote=True,
        )
        parts.append(
            (
                f'<a class="customer-directory-item" href="{url}" target="_self">'
                f'<span class="customer-directory-name">{html.escape(company)}</span>'
                f'<span class="customer-directory-meta">地域：{html.escape(region)}</span>'
                '</a>'
            )
        )
    parts.append("</div>")
    st.markdown("".join(parts), unsafe_allow_html=True)


def show_trade_partner_home(partner_type):
    show_top_home_link()
    label = trade_partner_type_label(partner_type)
    icon = "🏢" if partner_type == "supplier" else "🚚"
    st.header(f"{icon} {label}")

    register_page = "supplier_register" if partner_type == "supplier" else "carrier_register"
    st.markdown(
        render_page_link(f"＋ 新しい{label}を登録", page=register_page),
        unsafe_allow_html=True,
    )
    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(
            render_page_link(f"📋 {label}一覧", page=trade_partner_list_page(partner_type)),
            unsafe_allow_html=True,
        )
    with col2:
        st.markdown(
            render_page_link(f"🔍 {label}検索", page=trade_partner_search_page(partner_type)),
            unsafe_allow_html=True,
        )

    if partner_type == "supplier":
        st.markdown(
            render_page_link("📦 商品検索", page="supplier_product"),
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            render_page_link("💰 運賃比較", page="carrier_freight_compare"),
            unsafe_allow_html=True,
        )


def show_trade_partner_directory(partner_type):
    show_trade_partner_home_link(partner_type)
    label = trade_partner_type_label(partner_type)
    st.header(f"📋 {label}一覧")
    data = load_trade_partner_data()
    rows = sorted(get_trade_partner_master_rows(data, partner_type), key=trade_partner_sort_key)
    if not rows:
        st.info(f"登録されている{label}はありません。")
        return
    st.caption(f"{len(rows)}件")
    render_trade_partner_directory_cards(rows, partner_type)


def show_trade_partner_search(partner_type):
    show_trade_partner_home_link(partner_type)
    label = trade_partner_type_label(partner_type)
    st.header(f"🔍 {label}検索")
    default_keyword = str(get_query_value("partner_search", "")).strip()
    keyword = st.text_input(
        "会社名・会社名かな・地域で検索",
        value=default_keyword,
        placeholder="入力すると候補を表示します",
        key=f"{partner_type}_partner_search_input",
    ).strip()
    update_query_params(
        page=trade_partner_search_page(partner_type),
        partner_search=keyword or None,
    )
    if not keyword:
        st.info("検索文字を入力してください。")
        return
    data = load_trade_partner_data()
    target = normalize_match_value(keyword).lower()
    matches = []
    for row in get_trade_partner_master_rows(data, partner_type):
        haystack = " ".join(
            trade_partner_text(row.get(field))
            for field in ("会社名", "会社名かな", "地域")
        ).lower()
        if target in normalize_match_value(haystack).lower():
            matches.append(row)
    matches.sort(key=trade_partner_sort_key)
    if not matches:
        st.warning("一致する会社が見つかりません。")
        return
    st.caption(f"{len(matches)}件")
    render_trade_partner_directory_cards(matches, partner_type)


def show_supplier_product_search():
    show_trade_partner_home_link("supplier")
    st.header("📦 仕入商品の検索")
    data = load_trade_partner_data()
    supplier_ids = {
        trade_partner_text(row.get("取引先ID"))
        for row in get_trade_partner_master_rows(data, "supplier")
    }
    products = [
        row for row in data[TRADE_PARTNER_PRODUCT_SHEET]["rows"]
        if trade_partner_text(row.get("取引先ID")) in supplier_ids
        and trade_partner_text(row.get("商品名"))
    ]
    keyword = st.text_input(
        "商品名で検索",
        placeholder="例：酒、醤油粕",
        key="supplier_product_search_input",
    ).strip()
    if not keyword:
        st.info("商品名を入力してください。")
        return
    candidates = sorted({
        trade_partner_text(row.get("商品名"))
        for row in products
        if normalize_match_value(keyword).lower()
        in normalize_match_value(trade_partner_text(row.get("商品名"))).lower()
    })
    if not candidates:
        st.warning("一致する商品名が見つかりません。")
        return
    st.caption("候補の商品名を選んでください。")
    selected = st.session_state.get("selected_supplier_product", "")
    columns = st.columns(min(3, max(1, len(candidates))))
    for index, product_name in enumerate(candidates):
        with columns[index % len(columns)]:
            if st.button(product_name, key=f"supplier_product_candidate_{index}", use_container_width=True):
                st.session_state["selected_supplier_product"] = product_name
                selected = product_name
                st.rerun()
    if selected not in candidates:
        return
    st.markdown(f"### {selected}")
    master_by_id = {
        trade_partner_text(row.get("取引先ID")): row
        for row in get_trade_partner_master_rows(data, "supplier")
    }
    exact_rows = [row for row in products if trade_partner_text(row.get("商品名")) == selected]
    exact_rows.sort(key=lambda row: trade_partner_sort_key(master_by_id.get(trade_partner_text(row.get("取引先ID")), {})))
    for product in exact_rows:
        master = master_by_id.get(trade_partner_text(product.get("取引先ID")))
        if not master:
            continue
        with st.container(border=True):
            st.markdown(trade_partner_detail_link(master, "supplier"), unsafe_allow_html=True)
            for field in data[TRADE_PARTNER_PRODUCT_SHEET]["headers"]:
                if field in {"仕入商品ID", "取引先ID", "会社名（確認用）", "商品名"}:
                    continue
                value = trade_partner_text(product.get(field))
                if value:
                    st.write(f"**{field}：** {value}")


def show_carrier_condition_search():
    show_trade_partner_home_link("carrier")
    st.header("🗺 運送条件検索")
    keyword = st.text_input(
        "納品先・地域・運賃などで検索",
        placeholder="例：帯広、釧路",
        key="carrier_condition_search_input",
    ).strip()
    if not keyword:
        st.info("検索文字を入力してください。")
        return
    data = load_trade_partner_data()
    carrier_ids = {
        trade_partner_text(row.get("取引先ID"))
        for row in get_trade_partner_master_rows(data, "carrier")
    }
    target = normalize_match_value(keyword).lower()
    matches = []
    for row in data[TRADE_PARTNER_TRANSPORT_SHEET]["rows"]:
        partner_id = trade_partner_text(row.get("取引先ID"))
        if partner_id not in carrier_ids:
            continue
        values = [
            trade_partner_text(row.get(header))
            for header in data[TRADE_PARTNER_TRANSPORT_SHEET]["headers"]
            if header not in {"運送条件ID", "取引先ID", "会社名（確認用）"}
        ]
        if target in normalize_match_value(" ".join(values)).lower():
            matches.append(row)
    if not matches:
        st.warning("一致する運送条件が見つかりません。")
        return
    master_by_id = {
        trade_partner_text(row.get("取引先ID")): row
        for row in get_trade_partner_master_rows(data, "carrier")
    }
    for condition in matches:
        master = master_by_id.get(trade_partner_text(condition.get("取引先ID")))
        if not master:
            continue
        with st.container(border=True):
            st.markdown(trade_partner_detail_link(master, "carrier"), unsafe_allow_html=True)
            for field in data[TRADE_PARTNER_TRANSPORT_SHEET]["headers"]:
                if field in {"運送条件ID", "取引先ID", "会社名（確認用）"}:
                    continue
                value = trade_partner_text(condition.get(field))
                if value:
                    st.write(f"**{field}：** {value}")


def render_trade_partner_fields(row, headers, excluded=None):
    excluded = set(excluded or [])
    visible = []
    for header in headers:
        if header in excluded or header.startswith("_"):
            continue
        value = trade_partner_text(row.get(header))
        if value:
            visible.append((header, value))
    if not visible:
        st.caption("入力済みの情報はありません。")
        return
    for start in range(0, len(visible), 2):
        cols = st.columns(2)
        for offset, item in enumerate(visible[start:start + 2]):
            header, value = item
            with cols[offset]:
                st.caption(header)
                st.markdown(f"**{html.escape(value)}**")


def trade_partner_history_section(sheet_name):
    return {
        TRADE_PARTNER_MASTER_SHEET: "基本情報",
        TRADE_PARTNER_CONTACT_SHEET: "担当者",
        TRADE_PARTNER_PRODUCT_SHEET: "仕入商品",
        TRADE_PARTNER_TRANSPORT_SHEET: "運送条件",
    }.get(sheet_name, sheet_name)


def render_trade_partner_row_editor(
    sheet_name,
    row,
    headers,
    key_prefix,
    partner_type,
    company_name,
):
    id_field = TRADE_PARTNER_ID_FIELDS[sheet_name]
    record_id = trade_partner_text(row.get(id_field))
    excluded_headers = {id_field, "取引先ID", "会社名（確認用）"}
    if sheet_name == TRADE_PARTNER_MASTER_SHEET:
        excluded_headers.update({"仕入先区分", "運送会社区分"})
    if sheet_name == TRADE_PARTNER_CONTACT_SHEET and partner_type == "carrier":
        excluded_headers.add("会社名")
    editable_headers = [header for header in headers if header not in excluded_headers]
    with st.expander("編集"):
        with st.form(f"edit_{key_prefix}_{record_id}"):
            inputs = {}
            for header in editable_headers:
                inputs[header] = st.text_input(
                    header,
                    value=trade_partner_text(row.get(header)),
                    key=f"edit_{key_prefix}_{record_id}_{header}",
                )
            submitted = st.form_submit_button("バックアップして保存", use_container_width=True)
        if submitted:
            try:
                with st.spinner("バックアップを作成して保存しています…"):
                    result = update_trade_partner_row(sheet_name, record_id, inputs)
                    new_company_name = (
                        trade_partner_text(inputs.get("会社名"))
                        if sheet_name == TRADE_PARTNER_MASTER_SHEET
                        else company_name
                    ) or company_name
                    remember_change_history_warning(
                        record_change_history_safely(
                            trade_partner_type_label(partner_type),
                            result.get("partner_id") or trade_partner_text(row.get("取引先ID")),
                            new_company_name,
                            "変更",
                            result.get("changes", {}),
                            section=trade_partner_history_section(sheet_name),
                        )
                    )
                st.success("保存しました。")
                st.rerun()
            except Exception as error:
                st.error(str(error))


def render_trade_partner_related_section(
    data,
    sheet_name,
    partner_id,
    title,
    add_label,
    partner_type,
    company_name,
):
    headers = data[sheet_name]["headers"]
    rows = get_trade_partner_related_rows(data, sheet_name, partner_id)
    id_field = TRADE_PARTNER_ID_FIELDS[sheet_name]
    primary_field = TRADE_PARTNER_PRIMARY_FIELDS[sheet_name]
    st.markdown("---")
    st.subheader(title)
    if not rows:
        st.info(f"{title}はまだ登録されていません。")
    for row in rows:
        with st.container(border=True):
            heading = trade_partner_text(row.get(primary_field))
            st.markdown(f"**{html.escape(heading)}**")
            display_excluded = {id_field, "取引先ID", "会社名（確認用）", primary_field}
            if sheet_name == TRADE_PARTNER_CONTACT_SHEET and partner_type == "carrier":
                display_excluded.add("会社名")
            render_trade_partner_fields(
                row,
                headers,
                excluded=display_excluded,
            )
            render_trade_partner_row_editor(
                sheet_name,
                row,
                headers,
                key_prefix=f"{sheet_name}_{partner_id}",
                partner_type=partner_type,
                company_name=company_name,
            )

    with st.expander(f"＋ {add_label}"):
        add_excluded_headers = {id_field, "取引先ID", "会社名（確認用）"}
        if sheet_name == TRADE_PARTNER_CONTACT_SHEET and partner_type == "carrier":
            add_excluded_headers.add("会社名")
        editable_headers = [
            header for header in headers
            if header not in add_excluded_headers
        ]
        with st.form(f"add_{sheet_name}_{partner_id}"):
            values = {"取引先ID": partner_id}
            if sheet_name == TRADE_PARTNER_CONTACT_SHEET and partner_type == "carrier":
                values["会社名"] = company_name
            for header in editable_headers:
                values[header] = st.text_input(
                    header,
                    key=f"add_{sheet_name}_{partner_id}_{header}",
                )
            submitted = st.form_submit_button("バックアップして追加", use_container_width=True)
        if submitted:
            try:
                if not trade_partner_text(values.get(primary_field)):
                    raise ValueError(f"{primary_field}を入力してください。")
                with st.spinner("バックアップを作成して保存しています…"):
                    result = create_trade_partner_record(sheet_name, values)
                    remember_change_history_warning(
                        record_change_history_safely(
                            trade_partner_type_label(partner_type),
                            result.get("partner_id") or partner_id,
                            company_name,
                            "追加",
                            result.get("changes", {}),
                            section=trade_partner_history_section(sheet_name),
                        )
                    )
                st.success("追加しました。")
                st.rerun()
            except Exception as error:
                st.error(str(error))


def show_trade_partner_notes(partner_type, partner_id, company_name):
    st.markdown("---")
    st.subheader(f"📝 この{trade_partner_type_label(partner_type)}のメモ")
    note_key = make_trade_partner_note_key(partner_type, partner_id, company_name)
    input_key = f"trade_partner_note_{partner_type}_{partner_id}"
    clear_key = f"clear_{input_key}"
    if st.session_state.pop(clear_key, False):
        st.session_state[input_key] = ""
    note_text = st.text_area(
        "メモ本文",
        key=input_key,
        height=120,
        help=VOICE_INPUT_HELP,
    )
    if st.button("メモを保存", key=f"save_{input_key}"):
        if add_note(note_key, note_text):
            st.session_state[clear_key] = True
            st.rerun()
    notes = get_notes_for_customer(note_key)
    if not notes:
        st.info("メモはまだありません。")
        return
    st.markdown("#### メモ履歴")
    for note in notes:
        render_note_card(note, show_customer=False)
        render_note_delete_controls(note)


def show_trade_partner_detail(partner_type, partner_id):
    show_trade_partner_home_link(partner_type)
    data = load_trade_partner_data()
    master = get_trade_partner_by_id(data, partner_id)
    if not master or not is_trade_partner_marked(master.get(trade_partner_category_field(partner_type))):
        st.warning("選択した会社の情報が見つかりません。")
        return
    label = trade_partner_type_label(partner_type)
    company = trade_partner_text(master.get("会社名"))
    st.title(f"{'🏢' if partner_type == 'supplier' else '🚚'} {company}")
    st.caption(f"{label}ID：{trade_partner_text(master.get('取引先ID'))}")

    map_value = trade_partner_text(master.get("マップ位置")) or trade_partner_text(master.get("住所"))
    if map_value:
        map_url = build_google_maps_url(map_value)
        if map_url:
            show_google_maps_button(map_url)

    st.subheader("基本情報")
    master_headers = data[TRADE_PARTNER_MASTER_SHEET]["headers"]
    render_trade_partner_fields(
        master,
        master_headers,
        excluded={"取引先ID", "仕入先区分", "運送会社区分", "会社名", "会社名かな"},
    )
    render_trade_partner_row_editor(
        TRADE_PARTNER_MASTER_SHEET,
        master,
        master_headers,
        key_prefix=f"master_{partner_type}",
        partner_type=partner_type,
        company_name=company,
    )

    render_trade_partner_related_section(
        data,
        TRADE_PARTNER_CONTACT_SHEET,
        partner_id,
        "担当者",
        "担当者を追加",
        partner_type,
        company,
    )
    if partner_type == "supplier":
        render_trade_partner_related_section(
            data,
            TRADE_PARTNER_PRODUCT_SHEET,
            partner_id,
            "取扱商品",
            "商品を追加",
            partner_type,
            company,
        )
    else:
        render_carrier_freight_section(partner_id, company)
    show_trade_partner_notes(partner_type, partner_id, company)


def show_trade_partner_register(partner_type):
    show_trade_partner_home_link(partner_type)
    label = trade_partner_type_label(partner_type)
    st.header(f"＋ 新しい{label}を登録")
    data = load_trade_partner_data()
    headers = data[TRADE_PARTNER_MASTER_SHEET]["headers"]
    editable_headers = [
        header for header in headers
        if header not in {"取引先ID", "仕入先区分", "運送会社区分"}
    ]
    other_label = "運送会社でもある" if partner_type == "supplier" else "仕入先でもある"
    with st.form(f"register_{partner_type}"):
        values = {}
        for header in editable_headers:
            values[header] = st.text_input(
                header,
                key=f"register_{partner_type}_{header}",
            )
        also_other = st.checkbox(other_label, key=f"register_{partner_type}_also_other")
        submitted = st.form_submit_button("バックアップして登録", use_container_width=True)
    if submitted:
        try:
            if not trade_partner_text(values.get("会社名")):
                raise ValueError("会社名を入力してください。")
            values[trade_partner_category_field(partner_type)] = "○"
            if also_other:
                values[trade_partner_category_field("carrier" if partner_type == "supplier" else "supplier")] = "○"
            with st.spinner("バックアップを作成して登録しています…"):
                result = create_trade_partner_record(TRADE_PARTNER_MASTER_SHEET, values)
                remember_change_history_warning(
                    record_change_history_safely(
                        trade_partner_type_label(partner_type),
                        result.get("partner_id") or result.get("record_id"),
                        trade_partner_text(values.get("会社名")),
                        "登録",
                        result.get("changes", {}),
                        section="基本情報",
                    )
                )
            partner_id = result["record_id"]
            st.session_state["selected_partner_id"] = partner_id
            st.session_state["selected_partner_type"] = partner_type
            st.session_state["page"] = "partner_detail"
            update_query_params(
                page="partner_detail",
                partner_id=partner_id,
                partner_type=partner_type,
            )
            st.rerun()
        except Exception as error:
            st.error(str(error))


def render_trade_note_card(note, category, partner_names=None):
    parsed = parse_trade_partner_note_key(note.get("customer_name"))
    created_at = format_note_datetime(note.get("created_at", ""))
    body = html.escape(clean_value(note.get("body"), blank_text="")).replace("\n", "<br>")
    if parsed:
        partner_names = partner_names or {}
        company_name = (
            partner_names.get((parsed["partner_type"], parsed["partner_id"]))
            or parsed.get("company_name")
            or parsed["partner_id"]
        )
        company_link = render_page_link(
            company_name,
            page="partner_detail",
            partner_id=parsed["partner_id"],
            partner_type=parsed["partner_type"],
            class_name="dispatch-month-link",
        )
        meta = f"{html.escape(created_at)}　{company_link}"
    else:
        customer_name = clean_value(note.get("customer_name"), blank_text="未設定")
        customer_link = build_customer_detail_link(customer_name, class_name="dispatch-month-link")
        meta = f"{html.escape(created_at)}　{customer_link}"
    st.markdown(
        f'<div class="note-card"><div class="note-meta">{meta}</div><div class="note-body">{body}</div></div>',
        unsafe_allow_html=True,
    )


def show_trade_notes_page():
    show_top_home_link()
    st.header("📝 取引先メモ")
    notes = load_notes_from_supabase()
    try:
        partner_data = load_trade_partner_data()
        partner_names = {}
        for partner_type in ("supplier", "carrier"):
            for row in get_trade_partner_master_rows(partner_data, partner_type):
                partner_names[(partner_type, trade_partner_text(row.get("取引先ID")))] = trade_partner_text(row.get("会社名"))
    except Exception:
        partner_names = {}

    # st.tabs は削除確認などで再実行されるたびに先頭の「顧客」へ戻るため、
    # 選択値を session_state に保持できる横並びラジオで区分を切り替える。
    # これにより、仕入先や運送会社のメモを続けて削除しても同じ区分を維持する。
    category_labels = ["顧客", "仕入先", "運送会社"]
    category_by_label = {
        "顧客": "customer",
        "仕入先": "supplier",
        "運送会社": "carrier",
    }
    selected_label = st.radio(
        "表示する区分",
        category_labels,
        horizontal=True,
        key="trade_notes_selected_category",
        label_visibility="collapsed",
    )
    category = category_by_label[selected_label]

    filtered = []
    for note in notes:
        parsed = parse_trade_partner_note_key(note.get("customer_name"))
        if category == "customer" and parsed is None:
            filtered.append(note)
        elif parsed and parsed["partner_type"] == category:
            filtered.append(note)

    if category == "customer":
        try:
            past_product_items = load_all_past_product_notes_from_supabase()
        except Exception as exc:
            st.warning(f"過去商品メモを読み込めませんでした：{exc}")
            past_product_items = []

        for item in past_product_items:
            product_name = extract_past_product_name(item.get("field_name"))
            customer_name = clean_value(item.get("customer_name"), blank_text="")
            content = clean_value(item.get("content"), blank_text="").strip()
            if not product_name or not customer_name or not content:
                continue
            filtered.append(
                {
                    "id": item.get("id"),
                    "customer_name": customer_name,
                    "body": f"過去に使用した商品：{product_name}\n{content}",
                    "created_at": item.get("updated_at") or item.get("created_at") or "",
                    "_past_product_note": True,
                }
            )

    filtered.sort(
        key=lambda note: str(note.get("created_at") or ""),
        reverse=True,
    )

    if not filtered:
        st.info("メモはまだありません。")
        return

    for note in filtered:
        render_trade_note_card(note, category, partner_names=partner_names)
        if not note.get("_past_product_note"):
            render_note_delete_controls(note)


def show_top_home():
    st.subheader("取引先を選択")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(render_page_link("👥 顧客", page="customer_home"), unsafe_allow_html=True)
    with col2:
        st.markdown(render_page_link("🏢 仕入先", page="supplier_home"), unsafe_allow_html=True)
    col3, col4 = st.columns(2)
    with col3:
        st.markdown(render_page_link("🚚 運送会社", page="carrier_home"), unsafe_allow_html=True)
    with col4:
        st.markdown(render_page_link("📝 取引先メモ", page="trade_notes"), unsafe_allow_html=True)
    col5, col6 = st.columns(2)
    with col5:
        st.markdown(render_page_link("🕘 変更確認", page="change_history"), unsafe_allow_html=True)
    with col6:
        st.markdown(render_page_link("📄 商品見積り履歴", page="estimates"), unsafe_allow_html=True)


# =========================
# ホームメニュー
# =========================
def show_home_menu():
    show_top_home_link()
    st.subheader("顧客メニュー")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(render_page_link("👥 顧客名一覧", page="customer_list"), unsafe_allow_html=True)
    with col2:
        st.markdown(render_page_link("🔍 顧客検索", page="customer"), unsafe_allow_html=True)

    col3, col4 = st.columns(2)
    with col3:
        st.markdown(render_page_link("📍 地域検索", page="region"), unsafe_allow_html=True)
    with col4:
        st.markdown(render_page_link("🔎 商品検索", page="product"), unsafe_allow_html=True)

    col5, col6 = st.columns(2)
    with col5:
        st.markdown(render_page_link("🗓 在庫カレンダー", page="calendar"), unsafe_allow_html=True)
    with col6:
        st.markdown(render_page_link("🚚 配車表", page="dispatch_table"), unsafe_allow_html=True)

    col7, col8 = st.columns(2)
    with col7:
        st.markdown(render_page_link("🧪 ソリュブル在庫", page="soluble_inventory"), unsafe_allow_html=True)
    with col8:
        st.markdown(render_page_link("📝 メモ帳", page="notes"), unsafe_allow_html=True)

    col9, _ = st.columns(2)
    with col9:
        st.markdown(render_page_link("💧 WATER it接続", page="water_it_test"), unsafe_allow_html=True)

    st.markdown("---")

# =========================
# 全データバックアップ（読み取り専用）
# =========================
def backup_safe_value(value):
    """CSVへ安全に書き出せる値へ変換する。"""
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    if isinstance(value, (dict, list, tuple, set)):
        return json.dumps(value, ensure_ascii=False, default=str)
    return value


def backup_csv_bytes(dataframe):
    """Excelで文字化けしにくいUTF-8 BOM付きCSVを返す。"""
    export_df = dataframe.copy()
    for column in export_df.columns:
        export_df[column] = export_df[column].map(backup_safe_value)
    return b"\xef\xbb\xbf" + export_df.to_csv(
        index=False,
        lineterminator="\n",
    ).encode("utf-8")


def backup_dataframe(records, columns=None):
    """空の一覧でも必要な見出しを残す。"""
    dataframe = pd.DataFrame(records or [])
    if columns:
        for column in columns:
            if column not in dataframe.columns:
                dataframe[column] = ""
        ordered = list(columns) + [
            column for column in dataframe.columns if column not in columns
        ]
        dataframe = dataframe[ordered]
    return dataframe


def backup_read_all_supabase_rows(url, label):
    """Supabaseのテーブルを読み取り専用で全件取得する。"""
    if not has_supabase_config():
        raise RuntimeError(f"{label}を取得するためのSupabase設定がありません。")

    rows = []
    page_size = 1000
    offset = 0
    while True:
        try:
            response = requests.get(
                url,
                headers=get_supabase_headers(),
                params={
                    "select": "*",
                    "limit": str(page_size),
                    "offset": str(offset),
                },
                timeout=30,
            )
        except Exception as exc:
            raise RuntimeError(f"{label}の取得中にSupabaseへ接続できませんでした。") from exc

        if response.status_code != 200:
            detail = str(response.text or "").strip()[:500]
            raise RuntimeError(
                f"{label}を取得できませんでした（{response.status_code}）。"
                + (f" {detail}" if detail else "")
            )

        try:
            page = response.json()
        except Exception as exc:
            raise RuntimeError(f"{label}の応答形式が正しくありません。") from exc
        if not isinstance(page, list):
            raise RuntimeError(f"{label}の応答形式が正しくありません。")

        rows.extend(page)
        if len(page) < page_size:
            break
        offset += page_size
    return rows


def backup_get_main_excel_bytes():
    """顧客・在庫の元Excelを読み取り専用で取得する。"""
    if has_dropbox_auth_config():
        content = get_cached_dropbox_excel_content()
        path = get_dropbox_file_path()
        return content, Path(path).name or "配車予定 次郎.xlsm", f"Dropbox: {path}"

    path = Path(EXCEL_FILE)
    if not path.exists():
        raise FileNotFoundError(f"顧客・在庫の元Excelが見つかりません：{path}")
    return path.read_bytes(), path.name, f"ローカル: {path}"


def backup_get_dispatch_excel_bytes():
    """配車表の元Excelを画面と同じ優先順位で取得する。"""
    dropbox_error = None
    if has_dropbox_auth_config():
        try:
            content = get_cached_dispatch_dropbox_content()
            path = str(DISPATCH_DROPBOX_FILE_PATH or DISPATCH_DROPBOX_DEFAULT_FILE_PATH).strip()
            return content, Path(path).name or "配車表1.xlsm", f"Dropbox: {path}"
        except Exception as exc:
            dropbox_error = exc

    path = Path(str(DISPATCH_LOCAL_FILE or "").strip())
    if path.exists():
        return path.read_bytes(), path.name, f"ローカル: {path}"

    message = f"配車表の元Excelが見つかりません：{path}"
    if dropbox_error is not None:
        message += f"\nDropbox取得エラー：{dropbox_error}"
    raise FileNotFoundError(message)


def backup_get_soluble_excel_bytes():
    """ソリュブル在庫の元Excelを読み取り専用で取得する。"""
    content, source = load_soluble_workbook_content()
    return content, SOLUBLE_FILE_NAME, source


def backup_get_trade_partner_excel_bytes():
    """仕入先・運送会社の元Excelを読み取り専用で取得する。"""
    access_token = get_dropbox_access_token()
    path = get_trade_partner_file_path()
    content, response = download_dropbox_file(path, access_token)
    if content is None:
        raise RuntimeError(
            "取引先カルテ.xlsxをDropboxから取得できませんでした。\n"
            + dropbox_error_text(response)
        )
    return content, TRADE_PARTNER_FILE_NAME, f"Dropbox: {path}"


def backup_build_product_usage(customer_df):
    """商品検索と同じ基準で現在使用中・過去使用を一覧化する。"""
    product_rows = get_product_search_rows(customer_df)
    records = []
    if product_rows.empty:
        return backup_dataframe(
            [],
            ["商品名", "利用区分", "顧客名", "地域", "使用数量/日", "使用中行数"],
        )

    for product_name, product_group in product_rows.groupby("_商品名検索", sort=True):
        for customer_name, group in product_group.groupby("_顧客名検索", sort=True):
            current_group = group[
                ~group["使用数量/日"].apply(is_blank_or_zero)
            ].copy()
            regions = [
                clean_value(value, blank_text="").strip()
                for value in group["地域"].tolist()
            ]
            region = next((value for value in regions if value), "")
            if current_group.empty:
                status = "過去に使用"
                usage = ""
                current_count = 0
            else:
                status = "現在使用中"
                current_count = len(current_group)
                usage_values = [
                    clean_value(value, blank_text="").strip()
                    for value in current_group["使用数量/日"].tolist()
                ]
                usage = " / ".join(dict.fromkeys(value for value in usage_values if value))
            records.append(
                {
                    "商品名": product_name,
                    "利用区分": status,
                    "顧客名": customer_name,
                    "地域": region,
                    "使用数量/日": usage,
                    "使用中行数": current_count,
                }
            )
    return pd.DataFrame(records)


def backup_build_calendar_export(customer_df):
    """在庫カレンダーで使う基本項目を元データから書き出す。"""
    candidates = [
        "ID", "顧客名", "地域", "商品名", "メーカー",
        "使用数量/日", "次回配達予定", "残数",
    ]
    columns = [column for column in candidates if column in customer_df.columns]
    if not columns:
        return pd.DataFrame()
    result = customer_df[columns].copy()
    if "次回配達予定" in result.columns:
        result["次回配達予定_解析"] = pd.to_datetime(
            result["次回配達予定"], errors="coerce"
        )
        result.sort_values(
            ["次回配達予定_解析", "顧客名"],
            na_position="last",
            inplace=True,
        )
    return result.reset_index(drop=True)


def backup_build_note_exports(raw_notes):
    """notesテーブルを通常メモとLINE状態へ分ける。"""
    normal_notes = []
    line_statuses = []
    for row in raw_notes:
        row_id = clean_value(row.get("id"), blank_text="")
        if row_id.startswith(LINE_STATUS_NOTE_PREFIX):
            line_statuses.append(
                {
                    "顧客名": row.get("customer_name", ""),
                    "LINE状態": "接続中",
                    "保存ID": row_id,
                    "作成日時": row.get("created_at", ""),
                }
            )
        else:
            normal_notes.append(row)
    return backup_dataframe(normal_notes), backup_dataframe(
        line_statuses,
        ["顧客名", "LINE状態", "保存ID", "作成日時"],
    )


def backup_add_entry(entries, path, content, description, count=""):
    if not isinstance(content, (bytes, bytearray)) or not content:
        raise RuntimeError(f"バックアップ内容が空です：{path}")
    entries.append(
        {
            "path": path,
            "content": bytes(content),
            "description": description,
            "count": count,
        }
    )


def ensure_full_data_backup_dropbox_folder(access_token):
    """全データバックアップ専用フォルダを作る。既存フォルダは成功扱いにする。"""
    folder = str(FULL_DATA_BACKUP_DROPBOX_FOLDER or "").strip().rstrip("/")
    if not folder:
        raise RuntimeError("全データバックアップのDropbox保存先が設定されていません。")

    response = call_dropbox_rpc(
        "files/create_folder_v2",
        {"path": folder, "autorename": False},
        access_token,
    )
    if response.status_code == 200:
        return folder
    if response.status_code == 409:
        try:
            summary = str(response.json().get("error_summary", ""))
            if "conflict" in summary and "folder" in summary:
                return folder
        except Exception:
            pass
    raise RuntimeError(
        "Dropboxに全データバックアップ用フォルダを作成できませんでした。\n"
        + dropbox_error_text(response)
    )


def save_full_data_backup_to_dropbox(filename, content):
    """作成済みZIPを専用Dropboxフォルダへ追加保存し、内容を検証する。"""
    if not filename or not isinstance(content, (bytes, bytearray)) or not content:
        raise RuntimeError("Dropboxへ保存するバックアップZIPが空です。")

    access_token = get_dropbox_access_token()
    folder = ensure_full_data_backup_dropbox_folder(access_token)
    target_path = f"{folder}/{filename}"
    response = upload_dropbox_file(
        target_path,
        bytes(content),
        access_token,
        mode="add",
    )
    if response.status_code == 409:
        raise RuntimeError(
            "同じ名前のバックアップがDropboxに存在します。もう一度作成してください。"
        )
    if response.status_code != 200:
        raise RuntimeError(
            "Dropboxへ全データバックアップを保存できませんでした。\n"
            + dropbox_error_text(response)
        )

    metadata = get_dropbox_response_metadata(response)
    if not metadata.get("content_hash") or metadata.get("size") is None:
        metadata = get_dropbox_file_metadata(target_path, access_token)
    verify_dropbox_file_metadata(metadata, bytes(content))
    return target_path


def create_full_data_backup_zip():
    """現在の保存処理を変更せず、読み取りだけで全データZIPを作る。"""
    created_at = get_jst_now()
    timestamp = created_at.strftime("%Y%m%d_%H%M%S")
    entries = []
    sources = []

    main_excel, main_name, main_source = backup_get_main_excel_bytes()
    dispatch_excel, dispatch_name, dispatch_source = backup_get_dispatch_excel_bytes()
    soluble_excel, soluble_name, soluble_source = backup_get_soluble_excel_bytes()
    trade_excel, trade_name, trade_source = backup_get_trade_partner_excel_bytes()
    sources.extend([main_source, dispatch_source, soluble_source, trade_source])

    customer_df = normalize_excel_table(BytesIO(main_excel))
    dispatch_df = read_dispatch_month_sheets(BytesIO(dispatch_excel))
    soluble_rows_df = backup_dataframe(read_soluble_rows(soluble_excel))
    soluble_summary_df = backup_dataframe(
        list(read_soluble_customer_summaries(soluble_excel).values())
    )

    trade_editor = TradePartnerXlsxEditor(trade_excel)
    missing_sheets = [
        name for name in TRADE_PARTNER_REQUIRED_SHEETS
        if not trade_editor.has_sheet(name)
    ]
    if missing_sheets:
        raise RuntimeError(
            "取引先カルテ.xlsxに必要なシートがありません："
            + "、".join(missing_sheets)
        )
    trade_data = {
        name: trade_editor.read_sheet(name)
        for name in TRADE_PARTNER_REQUIRED_SHEETS
    }

    water_df, water_source = get_active_water_it_data()
    sources.append(f"WATER it表示データ: {water_source}")

    raw_notes = backup_read_all_supabase_rows(
        get_supabase_notes_url(),
        "メモ・LINE状態",
    )
    raw_customer_information = backup_read_all_supabase_rows(
        get_supabase_customer_information_url(),
        "顧客情報",
    )
    change_history_rows = [
        row for row in raw_customer_information
        if clean_value(row.get("customer_name"), blank_text="") == CHANGE_HISTORY_CUSTOMER
    ]
    water_storage_rows = [
        row for row in raw_customer_information
        if clean_value(row.get("id"), blank_text="") == WATER_IT_STORAGE_ID
        or clean_value(row.get("customer_name"), blank_text="") == WATER_IT_STORAGE_CUSTOMER
        or clean_value(row.get("field_name"), blank_text="") == WATER_IT_STORAGE_FIELD
    ]
    customer_information = [
        row for row in raw_customer_information
        if row not in water_storage_rows and row not in change_history_rows
    ]
    water_metadata = []
    for row in water_storage_rows:
        metadata = {
            "id": row.get("id", ""),
            "customer_name": row.get("customer_name", ""),
            "field_name": row.get("field_name", ""),
            "updated_at": row.get("updated_at", ""),
        }
        try:
            payload = json.loads(str(row.get("content") or "{}"))
            metadata.update(
                {
                    "保存ファイル名": payload.get("filename", ""),
                    "SHA256": payload.get("sha256", ""),
                    "行数": payload.get("row_count", ""),
                    "ポイント数": payload.get("point_count", ""),
                    "最古測定日時": payload.get("oldest_time", ""),
                    "最新測定日時": payload.get("latest_time", ""),
                    "取込日時": payload.get("imported_at", ""),
                    "保存形式バージョン": payload.get("version", ""),
                }
            )
        except Exception:
            metadata["解析結果"] = "保存メタデータを解析できませんでした"
        water_metadata.append(metadata)

    notes_df, line_df = backup_build_note_exports(raw_notes)
    product_usage_df = backup_build_product_usage(customer_df)
    calendar_df = backup_build_calendar_export(customer_df)
    estimates_df = estimate_rows_to_dataframe(customer_information)
    carrier_freights_df = carrier_freight_rows_to_dataframe(customer_information)
    onedrive_attachments_df = onedrive_attachment_rows_to_dataframe(customer_information)

    backup_add_entry(entries, f"元Excel/{main_name}", main_excel, "顧客・在庫の元Excel")
    backup_add_entry(entries, f"元Excel/{dispatch_name}", dispatch_excel, "配車表の元Excel")
    backup_add_entry(entries, f"元Excel/{soluble_name}", soluble_excel, "ソリュブル在庫の元Excel")
    backup_add_entry(entries, f"元Excel/{trade_name}", trade_excel, "仕入先・運送会社の元Excel")

    csv_exports = [
        ("CSV/01_顧客商品在庫_全行.csv", customer_df, "アプリが読み取る顧客・商品・在庫の全行"),
        ("CSV/02_商品利用状況.csv", product_usage_df, "現在使用中・過去使用の分類"),
        ("CSV/03_在庫カレンダー.csv", calendar_df, "在庫カレンダーの基本情報"),
        ("CSV/04_配車表.csv", dispatch_df.drop(columns=["_引取日", "_着日"], errors="ignore"), "配車表1月～12月"),
        ("CSV/05_ソリュブル在庫履歴.csv", soluble_rows_df, "ソリュブル在庫の日別情報"),
        ("CSV/06_ソリュブル顧客概要.csv", soluble_summary_df, "ソリュブル顧客の概要"),
        ("CSV/07_WATER_it表示データ.csv", water_df, "アプリで表示できるWATER it情報（元CSVは含まない）"),
        ("CSV/08_WATER_it保存メタデータ.csv", backup_dataframe(water_metadata), "WATER it保存データの概要のみ"),
        ("CSV/09_顧客情報_Supabase生データ.csv", backup_dataframe(customer_information), "顧客情報・過去商品メモ・提案見積り・運送会社運賃・写真資料メタデータの生データ"),
        ("CSV/10_メモ_Supabase生データ.csv", notes_df, "通常メモ・取引先メモの生データ"),
        ("CSV/11_LINE状態.csv", line_df, "LINE接続状態"),
    ]

    next_number = 12
    for sheet_name in TRADE_PARTNER_REQUIRED_SHEETS:
        sheet = trade_data[sheet_name]
        sheet_df = backup_dataframe(
            sheet["rows"],
            list(sheet["headers"]) + ["_row_number"],
        )
        csv_exports.append(
            (
                f"CSV/{next_number:02d}_{sheet_name}.csv",
                sheet_df,
                f"取引先カルテ.xlsxの{sheet_name}シート",
            )
        )
        next_number += 1

    csv_exports.extend(
        [
            (
                f"CSV/{next_number:02d}_変更履歴.csv",
                change_history_rows_to_dataframe(change_history_rows),
                "アプリから保存した変更履歴（メモ帳は対象外）",
            ),
            (
                f"CSV/{next_number + 1:02d}_変更履歴_Supabase生データ.csv",
                backup_dataframe(change_history_rows),
                "変更履歴のSupabase生データ",
            ),
            (
                f"CSV/{next_number + 2:02d}_提案見積り.csv",
                estimates_df,
                "顧客ごとの提案・見積り一覧",
            ),
            (
                f"CSV/{next_number + 3:02d}_運送会社運賃履歴.csv",
                carrier_freights_df,
                "運送会社ごとの運賃履歴",
            ),
            (
                f"CSV/{next_number + 4:02d}_写真資料メタデータ.csv",
                onedrive_attachments_df,
                "OneDriveに保存した顧客の写真・PDFの管理情報（ファイル本体はOneDrive）",
            ),
        ]
    )

    for path, dataframe, description in csv_exports:
        backup_add_entry(
            entries,
            path,
            backup_csv_bytes(dataframe),
            description,
            len(dataframe),
        )

    manifest = pd.DataFrame(
        [
            {
                "ファイル": entry["path"],
                "内容": entry["description"],
                "件数": entry["count"],
                "バイト数": len(entry["content"]),
                "SHA256": hashlib.sha256(entry["content"]).hexdigest(),
            }
            for entry in entries
        ]
    )
    backup_add_entry(
        entries,
        "バックアップ一覧.csv",
        backup_csv_bytes(manifest),
        "ZIP内ファイルの一覧・件数・SHA256",
        len(manifest),
    )

    info_lines = [
        f"作成日時: {created_at.strftime('%Y/%m/%d %H:%M:%S %z')}",
        f"アプリ名: {APP_TITLE}",
        "",
        "取得元:",
        *[f"- {source}" for source in sources],
        "",
        "注意:",
        "- 元のExcel・Supabaseデータは変更しません。作成したZIPだけをDropboxへ追加保存します。",
        "- パスワード、接続キー、secrets.tomlは含みません。",
        "- WATER itの元CSV本体および圧縮保存本文は含みません。",
        "- CSVはUTF-8 BOM付きです。",
        "- 元ExcelとCSVの重複は、データ保全のため意図的に残しています。",
    ]
    backup_add_entry(
        entries,
        "バックアップ情報.txt",
        ("\ufeff" + "\n".join(info_lines) + "\n").encode("utf-8"),
        "バックアップの作成日時・取得元・注意事項",
    )

    output = BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for entry in entries:
            archive.writestr(entry["path"], entry["content"])
    zip_content = output.getvalue()

    with zipfile.ZipFile(BytesIO(zip_content), "r") as archive:
        bad_file = archive.testzip()
        if bad_file:
            raise RuntimeError(f"ZIP内の検証に失敗しました：{bad_file}")
        archived_names = set(archive.namelist())
    expected_names = {entry["path"] for entry in entries}
    if archived_names != expected_names:
        raise RuntimeError("ZIP内のファイル一覧が作成前と一致しません。")

    return {
        "content": zip_content,
        "filename": f"取引先カルテ全データ_{timestamp}.zip",
        "file_count": len(entries),
        "customer_count": int(customer_df["顧客名"].nunique()) if "顧客名" in customer_df.columns else 0,
    }


@st.fragment
def show_full_data_backup_download_button():
    """ダウンロード操作ではアプリ全体を再実行せず、ZIPをブラウザへ渡す。"""
    zip_content = st.session_state.get("full_data_backup_zip_bytes")
    zip_name = st.session_state.get("full_data_backup_zip_name")
    if not zip_content or not zip_name:
        return

    st.download_button(
        "⬇ ZIPをダウンロード",
        data=BytesIO(zip_content),
        file_name=zip_name,
        mime="application/zip",
        key="download_full_data_backup",
        use_container_width=True,
    )


def show_full_data_backup_page():
    """ボタンを押した時だけZIPを作り、Dropboxへ追加保存する。"""
    st.header("📦 全データバックアップ")
    st.write("現在のExcel保存ルールは変更せず、元データを変更せずにバックアップを作成します。")
    st.caption("作成したZIPはDropboxの専用フォルダへ自動保存し、端末にもダウンロードできます。")
    st.caption(f"Dropbox保存先：{FULL_DATA_BACKUP_DROPBOX_FOLDER}")

    bytes_key = "full_data_backup_zip_bytes"
    name_key = "full_data_backup_zip_name"
    summary_key = "full_data_backup_summary"
    dropbox_path_key = "full_data_backup_dropbox_path"
    dropbox_error_key = "full_data_backup_dropbox_error"

    if st.button(
        "📦 全データバックアップを作成",
        key="create_full_data_backup",
        type="primary",
        use_container_width=True,
    ):
        for key in (
            bytes_key,
            name_key,
            summary_key,
            dropbox_path_key,
            dropbox_error_key,
        ):
            st.session_state.pop(key, None)
        try:
            with st.spinner("全データを読み取り、ZIPを作成しています…"):
                result = create_full_data_backup_zip()
            st.session_state[bytes_key] = result["content"]
            st.session_state[name_key] = result["filename"]
            st.session_state[summary_key] = (
                f"{result['file_count']}ファイル・"
                f"{result['customer_count']}顧客のZIPを作成しました。"
            )

            try:
                with st.spinner("作成したZIPをDropboxへ保存しています…"):
                    dropbox_path = save_full_data_backup_to_dropbox(
                        result["filename"],
                        result["content"],
                    )
                st.session_state[dropbox_path_key] = dropbox_path
            except Exception as exc:
                st.session_state[dropbox_error_key] = str(exc)
        except Exception as exc:
            st.error(f"完全バックアップを作成できませんでした：{exc}")

    zip_content = st.session_state.get(bytes_key)
    zip_name = st.session_state.get(name_key)
    if zip_content and zip_name:
        summary = st.session_state.get(summary_key)
        if summary:
            st.success(summary)

        dropbox_path = st.session_state.get(dropbox_path_key)
        dropbox_error = st.session_state.get(dropbox_error_key)
        if dropbox_path:
            st.success(f"Dropboxへ保存しました：{dropbox_path}")
        elif dropbox_error:
            st.warning(
                "Dropboxへの保存に失敗しました。下のボタンから端末へダウンロードしてください。"
            )
            st.error(dropbox_error)

        show_full_data_backup_download_button()


# =========================
# メイン
# =========================
if "page" not in st.session_state:
    st.session_state["page"] = "home"

if "selected_customer" not in st.session_state:
    st.session_state["selected_customer"] = None
if "selected_partner_id" not in st.session_state:
    st.session_state["selected_partner_id"] = None
if "selected_partner_type" not in st.session_state:
    st.session_state["selected_partner_type"] = None

# URLに画面情報がある場合は、ブラウザの戻る・進むに合わせて復元する。
handle_customer_query_param()

current_page = st.session_state.get("page", "home")
customer_pages = {
    "customer_home", "customer_list", "customer", "region", "product", "calendar",
    "dispatch_table", "soluble_inventory", "water_it_test", "notes", "detail",
}
supplier_pages = {
    "supplier_home", "supplier_list", "supplier_search", "supplier_product",
    "supplier_register",
}
carrier_pages = {
    "carrier_home", "carrier_list", "carrier_search",
    "carrier_freight_compare", "carrier_register",
}

with st.sidebar:
    st.title(f"🚚 {APP_TITLE}")
    st.markdown("### メニュー")
    if st.button("🔄 更新", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.markdown("---")

    st.markdown(render_page_link("🏠 トップ", page="home"), unsafe_allow_html=True)
    st.markdown(render_page_link("👥 顧客", page="customer_home"), unsafe_allow_html=True)
    st.markdown(render_page_link("🏢 仕入先", page="supplier_home"), unsafe_allow_html=True)
    st.markdown(render_page_link("🚚 運送会社", page="carrier_home"), unsafe_allow_html=True)
    st.markdown(render_page_link("📝 取引先メモ", page="trade_notes"), unsafe_allow_html=True)

    if current_page in customer_pages:
        st.markdown("---")
        st.markdown("#### 顧客メニュー")
        st.markdown(render_page_link("👥 顧客名一覧", page="customer_list"), unsafe_allow_html=True)
        st.markdown(render_page_link("🔍 顧客検索", page="customer"), unsafe_allow_html=True)
        st.markdown(render_page_link("📍 地域検索", page="region"), unsafe_allow_html=True)
        st.markdown(render_page_link("🔎 商品検索", page="product"), unsafe_allow_html=True)
        st.markdown(render_page_link("🗓 在庫カレンダー", page="calendar"), unsafe_allow_html=True)
        st.markdown(render_page_link("🚚 配車表", page="dispatch_table"), unsafe_allow_html=True)
        st.markdown(render_page_link("🧪 ソリュブル在庫", page="soluble_inventory"), unsafe_allow_html=True)
        st.markdown(render_page_link("💧 WATER it接続", page="water_it_test"), unsafe_allow_html=True)
        st.markdown(render_page_link("📝 顧客メモ", page="notes"), unsafe_allow_html=True)
    elif current_page in supplier_pages or (
        current_page == "partner_detail" and st.session_state.get("selected_partner_type") == "supplier"
    ):
        st.markdown("---")
        st.markdown("#### 仕入先メニュー")
        st.markdown(render_page_link("📋 仕入先一覧", page="supplier_list"), unsafe_allow_html=True)
        st.markdown(render_page_link("🔍 仕入先検索", page="supplier_search"), unsafe_allow_html=True)
        st.markdown(render_page_link("📦 商品検索", page="supplier_product"), unsafe_allow_html=True)
    elif current_page in carrier_pages or (
        current_page == "partner_detail" and st.session_state.get("selected_partner_type") == "carrier"
    ):
        st.markdown("---")
        st.markdown("#### 運送会社メニュー")
        st.markdown(render_page_link("📋 運送会社一覧", page="carrier_list"), unsafe_allow_html=True)
        st.markdown(render_page_link("🔍 運送会社検索", page="carrier_search"), unsafe_allow_html=True)
        st.markdown(render_page_link("💰 運賃比較", page="carrier_freight_compare"), unsafe_allow_html=True)

    st.markdown("---")
    st.markdown(
        render_page_link("📦 全データバックアップ", page="data_backup"),
        unsafe_allow_html=True,
    )


col_title, col_logout = st.columns([3, 1])

with col_title:
    st.title(f"🚚 {APP_TITLE}")
    st.caption("顧客・仕入先・運送会社の情報を確認・編集します。")

with col_logout:
    st.write("")
    if st.button("ログアウト"):
        st.session_state.authenticated = False
        st.session_state.page = "home"
        st.session_state.selected_customer = None
        st.session_state.selected_partner_id = None
        st.session_state.selected_partner_type = None
        clear_onedrive_auth_state()
        try:
            st.query_params.clear()
        except Exception:
            pass
        st.rerun()

history_warning = st.session_state.pop("change_history_warning", None)
if history_warning:
    st.warning(history_warning)

# 各機能ページから、区分メニューを経由せずトップへ直接戻れるようにする。
# 顧客・仕入先・運送会社の各ホームと取引先メモは、従来から同じリンクを
# 表示しているため、二重表示にならないようここでは除外する。
pages_with_existing_top_link = {
    "customer_home",
    "supplier_home",
    "carrier_home",
    "trade_notes",
}
if current_page != "home" and current_page not in pages_with_existing_top_link:
    show_top_home_link()

try:
    if st.session_state["page"] == "home":
        show_top_home()

    elif st.session_state["page"] == "customer_home":
        show_home_menu()
        show_customer_search()

    elif st.session_state["page"] == "customer":
        show_customer_search(show_home_link=True)

    elif st.session_state["page"] == "customer_list":
        show_customer_directory()

    elif st.session_state["page"] == "region":
        df = load_data()
        show_region_search(df)

    elif st.session_state["page"] == "product":
        show_product_search()

    elif st.session_state["page"] == "calendar":
        df = load_data()
        show_dispatch_calendar(df)

    elif st.session_state["page"] == "dispatch_table":
        show_dispatch_board()

    elif st.session_state["page"] == "soluble_inventory":
        show_soluble_inventory_page()

    elif st.session_state["page"] == "water_it_test":
        show_water_it_test_page()

    elif st.session_state["page"] == "notes":
        show_notes_page(None)

    elif st.session_state["page"] == "trade_notes":
        show_trade_notes_page()

    elif st.session_state["page"] == "change_history":
        show_change_history_page()

    elif st.session_state["page"] == "estimates":
        show_estimates_page()

    elif st.session_state["page"] == "data_backup":
        show_full_data_backup_page()

    elif st.session_state["page"] == "detail":
        selected = st.session_state.get("selected_customer")
        if selected:
            immediate_df = st.session_state.pop("customer_excel_immediate_df", None)
            if isinstance(immediate_df, pd.DataFrame) and not immediate_df.empty:
                df = immediate_df
            else:
                df = load_data()
            show_customer_detail(df, selected)
        else:
            set_page("customer_home")
            st.rerun()

    elif st.session_state["page"] == "supplier_home":
        show_trade_partner_home("supplier")
    elif st.session_state["page"] == "supplier_list":
        show_trade_partner_directory("supplier")
    elif st.session_state["page"] == "supplier_search":
        show_trade_partner_search("supplier")
    elif st.session_state["page"] == "supplier_product":
        show_supplier_product_search()
    elif st.session_state["page"] == "supplier_register":
        show_trade_partner_register("supplier")

    elif st.session_state["page"] == "carrier_home":
        show_trade_partner_home("carrier")
    elif st.session_state["page"] == "carrier_list":
        show_trade_partner_directory("carrier")
    elif st.session_state["page"] == "carrier_search":
        show_trade_partner_search("carrier")
    elif st.session_state["page"] == "carrier_freight_compare":
        show_carrier_freight_compare()
    elif st.session_state["page"] == "carrier_register":
        show_trade_partner_register("carrier")

    elif st.session_state["page"] == "partner_detail":
        partner_id = st.session_state.get("selected_partner_id")
        partner_type = st.session_state.get("selected_partner_type")
        if partner_id and partner_type in {"supplier", "carrier"}:
            show_trade_partner_detail(partner_type, partner_id)
        else:
            set_page("home")
            st.rerun()
except Exception as e:
    st.error("画面表示中にエラーが発生しました。")
    st.write("原因確認のため、エラー内容を表示しています。")
    st.exception(e)
    st.stop()

st.caption(
    "※ 顧客情報は配車予定 次郎.xlsm、仕入先・運送会社は取引先カルテ.xlsxを読み込んで表示しています。"
)

