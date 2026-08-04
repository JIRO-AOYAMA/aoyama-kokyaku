import streamlit as st


st.set_page_config(
    page_title="起動確認",
    page_icon="✅",
    layout="centered",
)

st.title("起動確認")
st.success("テストアプリは正常に画面を表示できています。")
st.info("顧客データ、Excel、Dropbox、Supabaseには接続していません。")
