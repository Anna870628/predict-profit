import streamlit as st
import pandas as pd
from datetime import datetime
import calendar

# --- 網頁基本設定 ---
st.set_page_config(page_title="滾動式業務分析系統", page_icon="📊", layout="wide")
st.title("📊 滾動式業務分析系統")

# --- 側邊欄：固定參數設定 ---
st.sidebar.header("⚙️ 參數與目標設定")
TARGET_REVENUE = st.sidebar.number_input("業績目標 (未稅)", value=190476, step=1000)
UNIT_PRICE = st.sidebar.number_input("客單價", value=200, step=10)
TAX_RATE = st.sidebar.number_input("稅率 (例：1.05)", value=1.05, format="%.2f")
ADJUSTMENT_FACTOR = st.sidebar.slider("日均增調整係數", min_value=0.1, max_value=1.5, value=0.8, step=0.1)

# --- 0. 自動滾動時間設定 ---
now = datetime.now()
this_year = now.year
this_month = now.month
rolling_day = now.day

_, total_days_in_month = calendar.monthrange(this_year, this_month)
remaining_days = total_days_in_month - rolling_day
cutoff_date = datetime(this_year, this_month, rolling_day)
this_month_1st = datetime(this_year, this_month, 1)

st.sidebar.info(f"📅 **基準日**：{this_year}/{this_month}/{rolling_day}\n\n本月剩餘天數：{remaining_days} 天")

# --- 1. 檔案上傳區 ---
st.header("📂 資料匯入")
col1, col2 = st.columns(2)
with col1:
    file_curr = st.file_uploader("上傳「本月」報表 (Excel)", type=['xlsx'])
with col2:
    file_prev = st.file_uploader("上傳「上月」報表 (Excel)", type=['xlsx'])

# --- 關鍵修正：更嚴格的退款過濾 ---
def is_valid_order(val):
    if pd.isna(val): return True
    s = str(val).strip().lower()
    if s in ['0', '0.0', '0.00', '無', 'nan', '', 'none']: return True
    try:
        if float(val) == 0: return True
    except: pass
    return False

st.divider()

# --- 2. 執行分析邏輯 ---
if file_curr is not None and file_prev is not None:
    if st.button("🚀 執行分析", type="primary", use_container_width=True):
        try:
            with st.spinner('資料處理中...'):
                # 讀取數據
                df_c_raw = pd
