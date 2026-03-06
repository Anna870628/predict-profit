import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import calendar

# --- 網頁基本設定 ---
st.set_page_config(page_title="滾動式業務分析系統", page_icon="📊", layout="wide")
st.title("📊 滾動式業務分析系統")

# --- 建立兩個分頁標籤 ---
tab1, tab2 = st.tabs(["🚗 洗車業務分析", "📺 LiTV 訂閱分析"])

# ==========================================
# TAB 1: 洗車業務分析 (保留原有邏輯)
# ==========================================
with tab1:
    st.header("⚙️ 洗車參數與目標設定")
    col_p1, col_p2, col_p3, col_p4 = st.columns(4)
    TARGET_REVENUE = col_p1.number_input("業績目標 (未稅)", value=190476, step=1000, key="wash_target")
    UNIT_PRICE = col_p2.number_input("客單價", value=200, step=10, key="wash_price")
    TAX_RATE = col_p3.number_input("稅率 (例：1.05)", value=1.05, format="%.2f", key="wash_tax")
    ADJUSTMENT_FACTOR = col_p4.slider("日均增調整係數", min_value=0.1, max_value=1.5, value=0.8, step=0.1, key="wash_adj")

    # 自動滾動時間設定
    now = datetime.now()
    this_year = now.year
    this_month = now.month
    rolling_day = now.day

    _, total_days_in_month = calendar.monthrange(this_year, this_month)
    remaining_days = total_days_in_month - rolling_day
    cutoff_date = datetime(this_year, this_month, rolling_day)
    this_month_1st = datetime(this_year, this_month, 1)

    st.info(f"📅 **基準日**：{this_year}/{this_month}/{rolling_day} | 本月剩餘天數：{remaining_days} 天")

    st.header("📂 資料匯入")
    col1, col2 = st.columns(2)
    with col1:
        file_curr = st.file_uploader("上傳「本月」報表 (Excel)", type=['xlsx'], key="wash_curr")
    with col2:
        file_prev = st.file_uploader("上傳「上月」報表 (Excel)", type=['xlsx'], key="wash_prev")

    def is_valid_order(val):
        if pd.isna(val): return True
        s = str(val).strip().lower()
        if s in ['0', '0.0', '0.00', '無', 'nan', '', 'none']: return True
        try:
            if float(val) == 0: return True
        except: pass
        return False

    st.divider()

    if file_curr is not None and file_prev is not None:
        if st.button("🚀 執行洗車分析", type="primary", use_container_width=True, key="btn_wash"):
            try:
                with st.spinner('洗車資料處理中...'):
                    df_c_raw = pd.read_excel(file_curr, skiprows=2)
                    df_p_raw = pd.read_excel(file_prev, skiprows=2)

                    for df in [df_c_raw, df_p_raw]:
                        df.columns = [str(c).strip() for c in df.columns]

                    date_col = [c for c in df_c_raw.columns if '首次服務起始日' in c][0]
                    refund_col = [c for c in df_c_raw.columns if '退款' in c][0]

                    df_c_raw[date_col] = pd.to_datetime(df_c_raw[date_col], errors='coerce')
                    df_p_raw[date_col] = pd.to_datetime(df_p_raw[date_col], errors='coerce')

                    df_c_valid = df_c_raw[(df_c_raw[date_col] <= cutoff_date) & 
                                          (df_c_raw[refund_col].apply(is_valid_order))].copy()

                    df_c_valid['類型'] = df_c_valid[date_col].apply(lambda x: '新增' if x >= this_month_1st else '續訂')

                    count_new = len(df_c_valid[df_c_valid['類型'] == '新增'])
                    count_renew = len(df_c_valid[df_c_valid['類型'] == '續訂'])
                    count_total = count_new + count_renew

                    daily_new = df_c_valid[df_c_valid['類型'] == '新增'].groupby(df_c_valid[date_col].dt.date).size()
                    if not daily_new.empty:
                        if len(daily_new) > 1:
                            filtered_daily = daily_new.sort_values()[:-1]
                            raw_daily_avg = filtered_daily.mean()
                        else:
                            raw_daily_avg = daily_new.mean()
                    else:
                        raw_daily_avg = 0
                    
                    final_daily_avg = int(raw_daily_avg * ADJUSTMENT_FACTOR)

                    df_p_all_valid = df_p_raw[df_p_raw[refund_col].apply(is_valid_order)]
                    last_month_full_total = len(df_p_all_valid)

                    df_p_sync = df_p_raw[(df_p_raw[date_col].dt.day <= rolling_day) & (df_p_raw[refund_col].apply(is_valid_order))]
                    renewal_rate = (count_renew / len(df_p_sync)) if len(df_p_sync) > 0 else 0

                    est_future_new = final_daily_avg * remaining_days
                    est_future_renew = round(last_month_full_total * 0.79) - count_renew
                    est_future_renew = max(0, est_future_renew)

                    current_rev_gross = count_total * UNIT_PRICE
                    future_rev_est = (est_future_new + est_future_renew) * UNIT_PRICE
                    total_rev_gross_est = current_rev_gross + future_rev_est
                    total_rev_net_est = total_rev_gross_est / TAX_RATE
                    final_achievement_rate = (total_rev_net_est / TARGET_REVENUE) * 100

                # --- 洗車 Dashboard ---
                st.header("📈 洗車業務分析報告")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("✅ 當月總訂閱數", f"{count_total} 單", f"新增 {count_new} / 續訂 {count_renew}", delta_color="off")
                m2.metric("🔄 滾動續訂率", f"{renewal_rate*100:.1f}%")
                m3.metric("🏆 目前達成 (未稅)", f"${current_rev_gross/TAX_RATE:,.0f}")
                m4.metric("🎯 業績達成率", f"{(current_rev_gross/TAX_RATE)/TARGET_REVENUE*100:.1f}%")

                st.divider()
                st.subheader("📝 執行結果明細 (可直接複製報告)")
                adj_str = f"{raw_daily_avg:.1f} * {ADJUSTMENT_FACTOR} = {raw_daily_avg*ADJUSTMENT_FACTOR:.1f} -> 取 {final_daily_avg}"
                report_text = f"""{this_month}月營收預算  ${TARGET_REVENUE:,}
截至{this_month}/{rolling_day}營收 $ {current_rev_gross:,}(已扣款)
{this_month}/1~{rolling_day} (均增 {adj_str} 位)
【日均增 {final_daily_avg}*剩{remaining_days}天*${UNIT_PRICE}】，推估{this_month}/{rolling_day+1}~{total_days_in_month} 累積新增 {est_future_new}單、續訂約 {est_future_renew}單
${current_rev_gross:,} + ${future_rev_est:,}= ${total_rev_gross_est:,}
回除{TAX_RATE}稅率={total_rev_gross_est:,}/{TAX_RATE}={total_rev_net_est:,.0f}
預估達成率 {final_achievement_rate:.0f}%"""
                st.code(report_text, language="text")

            except Exception as e:
                st.error(f"分析過程中發生錯誤：{e}")
    else:
        st.info("請上傳「本月」與「上月」的報表檔案，以解鎖洗車分析按鈕。")


# ==========================================
# TAB 2: LiTV 訂閱分析
# ==========================================
with tab2:
    st.header("⚙️ LiTV 參數與目標設定")
    TARGET_REVENUE_LITV = st.number_input("LiTV 業績目標 (未稅)", value=6262, step=100, key="litv_target")
    
    # 隱藏的常數設定 (如果未來需要調整可以移出來)
    PRICE_M = 250
    PRICE_Y = 2290

    st.header("📂 資料匯入")
    col_b, col_c = st.columns(2)
    with col_b:
        file_b = st.file_uploader("上傳 B表：本月供應商報表 (實績)", type=['xlsx'], key="litv_b")
    with col_c:
        file_c = st.file_uploader("上傳 C表：車輛狀態表 (預估母體)", type=['xlsx'], key="litv_c")

    st.divider()

    if file_b is not None and file_c is not None:
        if st.button("🚀 執行 LiTV 分析", type="primary", use_container_width=True, key="btn_litv"):
            try:
                with st.spinner('LiTV 資料處理中...'):
                    # 讀取資料 (Header=2 表示標題在第3列)
                    df_b = pd.read_excel(file_b, header=2)
                    df_c = pd.read_excel(file_c, header=2)

                    # 欄位清理
                    for df in [df_b, df_c]:
                        df.columns = [str(c).strip() for c in df.columns]

                    # 關鍵欄位設定
                    COL_PHONE = '手機號碼'
                    COL_STATUS_C = '服務狀態(訂閱狀態)'
                    COL_EXPIRY_C = '當前服務到期日'
                    COL_SKU_C = '當前訂閱方案(SKU)' if '當前訂閱方案(SKU)' in df_c.columns else '方案(SKU)'

                    # 資料前處理
                    def clean_df(df):
                        if 'VIN' in df.columns:
                            df = df[~df['VIN'].astype(str).str.contains('總計|nan', case=False, na=True)].copy()
                        if COL_PHONE in df.columns:
                            df[COL_PHONE] = df[COL_PHONE].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', np.nan)
                        return df

                    df_b = clean_df(df_b)
                    df_c = clean_df(df_c)

                    # 日期轉換
                    if '訂單建立時間' in df_b.columns: 
                        df_b['訂單建立時間'] = pd.to_datetime(df_b['訂單建立時間'], errors='coerce')
                    if COL_EXPIRY_C in df_c.columns: 
                        df_c[COL_EXPIRY_C] = pd.to_datetime(df_c[COL_EXPIRY_C], errors='coerce')

                    # 確定統計日期範圍
                    if not df_b.empty and df_b['訂單建立時間'].notna().any():
                        latest = df_b['訂單建立時間'].max()
                        curr_y, curr_m, curr_d = latest.year, latest.month, latest.day
                    else:
                        now_litv = datetime.now()
                        curr_y, curr_m, curr_d = now_litv.year, now_litv.month, now_litv.day

                    last_day_of_month_litv = calendar.monthrange(curr_y, curr_m)[1]

                    # 實績 (Actuals)
                    df_act = df_b[df_b['金額'] > 0].copy()
                    act_cnt_m = df_act['方案(SKU)'].str.contains('1M', case=False, na=False).sum()
                    act_cnt_y = df_act['方案(SKU)'].str.contains('1Y', case=False, na=False).sum()
                    act_rev_m = df_act[df_act['方案(SKU)'].str.contains('1M', case=False, na=False)]['金額'].sum()
                    act_rev_y = df_act[df_act['方案(SKU)'].str.contains('1Y', case=False, na=False)]['金額'].sum()
                    total_act = df_act['金額'].sum()

                    paid_phones = df_act[COL_PHONE].dropna().unique()

                    # 預估 (Forecast)
                    mask_c_date = (df_c[COL_EXPIRY_C].dt.year == curr_y) & (df_c[COL_EXPIRY_C].dt.month == curr_m)
                    mask_status = ~df_c[COL_STATUS_C].astype(str).str.contains('暫停繼續訂閱|取消訂閱', na=False)
                    mask_sku = df_c[COL_SKU_C].astype(str).str.contains('LiTV', case=False, na=False)

                    candidates = df_c[mask_c_date & mask_status & mask_sku].copy()
                    final_fc = candidates[~candidates[COL_PHONE].isin(paid_phones)].copy()

                    fc_cnt_m = final_fc[COL_SKU_C].str.contains('1M', case=False, na=False).sum()
                    fc_cnt_y = final_fc[COL_SKU_C].str.contains('1Y', case=False, na=False).sum()
                    est_rev = (fc_cnt_m * PRICE_M) + (fc_cnt_y * PRICE_Y)

                    # 總計
                    total_gross = total_act + est_rev
                    total_net = total_gross / 1.05
                    achievement_rate = (total_net / TARGET_REVENUE_LITV) * 100

                    # 日期字串
                    date_range_act = f"{curr_m}/1~{curr_m}/{curr_d}"
                    fc_start_day = curr_d if curr_d < last_day_of_month_litv else curr_d
                    date_range_fc = f"{curr_m}/{fc_start_day}~{last_day_of_month_litv}"

                # --- LiTV Dashboard ---
                st.header("📈 LiTV 業務分析報告")
                lm1, lm2, lm3 = st.columns(3)
                lm1.metric("💰 目前實績 (含稅)", f"${total_act:,.0f}")
                lm2.metric("🔮 月底預估總計 (含稅)", f"${total_gross:,.0f}")
                lm3.metric("🎯 預估達成率", f"{achievement_rate:.1f}%")

                st.divider()
                st.subheader("📝 執行結果明細 (可直接複製報告)")

                fc_parts = []
                if fc_cnt_y > 0: fc_parts.append(f"年訂閱為{fc_cnt_y}人")
                if fc_cnt_m > 0: fc_parts.append(f"月訂閱為{fc_cnt_m}人")
                if not fc_parts: fc_parts.append("無預計續訂人數")
                fc_str = "，".join(fc_parts)

                math_parts = []
                if act_cnt_m > 0: math_parts.append(f"${PRICE_M}*{act_cnt_m}")
                if act_cnt_y > 0: math_parts.append(f"${PRICE_Y}*{act_cnt_y}")
                if not math_parts: math_parts.append("$0")
                math_str = "+".join(math_parts)

                report_text_litv = f"""LiTV訂閱服務
1.預估營收與實績
{curr_m}月營收 ${TARGET_REVENUE_LITV:,}
{date_range_act} 訂閱實績營收為${act_rev_m:,.0f}，年訂閱營收為${act_rev_y:,.0f}
推估營收為{date_range_fc} 續訂人數：{fc_str}
{math_str}+預估營收${est_rev:,.0f}=${total_act:,.0f}+${est_rev:,.0f}=${total_gross:,.0f}
回除1.05稅率=(${total_act:,.0f}+${est_rev:,.0f})/1.05=${total_net:,.0f}
預估達成率=${achievement_rate:.1f}%"""

                st.code(report_text_litv, language="text")

            except Exception as e:
                st.error(f"分析過程中發生錯誤：{e}")
                st.warning("請確認上傳的 Excel 檔案格式與欄位名稱正確。")
    else:
        st.info("請上傳「B表」與「C表」，以解鎖 LiTV 分析按鈕。")
