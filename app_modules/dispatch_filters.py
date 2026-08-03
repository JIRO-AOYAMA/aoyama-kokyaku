"""配車表の絞り込みに使う純粋な処理。"""

import re
from datetime import date, timedelta

import pandas as pd


def normalize_dispatch_text(value):
    """絞り込み用に前後空白と連続空白をそろえる。"""
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u3000", " ")).strip()


def dispatch_filter_options(series):
    values = sorted(
        {
            normalize_dispatch_text(value)
            for value in series
            if normalize_dispatch_text(value)
        }
    )
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
        return df[
            values.map(lambda value: pd.notna(value) and start <= value <= end)
        ]
    if mode == "未入力":
        return df[values.isna()]
    if (
        mode == "期間指定"
        and isinstance(range_value, (tuple, list))
        and len(range_value) == 2
    ):
        start, end = range_value
        return df[
            values.map(lambda value: pd.notna(value) and start <= value <= end)
        ]
    return df
