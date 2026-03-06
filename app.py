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
                df_c_raw = pd.read_excel(file_curr, skiprows=2)
                df_p_raw = pd.read_excel(file_prev, skiprows=2)

                for df in [df_c_raw, df_p_raw]:
                    df.columns = [str(c).strip() for c in df.columns]

                # 取得時間與退款欄位
                date_col = [c for c in df_c_raw.columns if '首次服務起始日' in c][0]
                refund_col = [c for c in df_c_raw.columns if '退款' in c][0]

                df_c_raw[date_col] = pd.to_datetime(df_c_raw[date_col], errors='coerce')
                df_p_raw[date_col] = pd.to_datetime(df_p_raw[date_col], errors='coerce')

                # 數據篩選
                df_c_valid = df_c_raw[(df_c_raw[date_col] <= cutoff_date) & 
                                      (df_c_raw[refund_col].apply(is_valid_order))].copy()

                df_c_valid['類型'] = df_c_valid[date_col].apply(lambda x: '新增' if x >= this_month_1st else '續訂')

                count_new = len(df_c_valid[df_c_valid['類型'] == '新增'])
                count_renew = len(df_c_valid[df_c_valid['類型'] == '續訂'])
                count_total = count_new + count_renew

                # 指標計算
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

            # --- 3. 輸出結果 Dashboard ---
            st.header("📈 滾動式業務分析報告")
            
            # 第一排指標：當前狀況
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("✅ 當月總訂閱數", f"{count_total} 單", f"新增 {count_new} / 續訂 {count_renew}", delta_color="off")
            m2.metric("🔄 滾動續訂率 (vs上月同期)", f"{renewal_rate*100:.1f}%")
            m3.metric("🏆 目前達成 (未稅)", f"${current_rev_gross/TAX_RATE:,.0f}")
            m4.metric("🎯 業績達成率", f"{(current_rev_gross/TAX_RATE)/TARGET_REVENUE*100:.1f}%")

            st.write(f"*(註：上月整月有效單數為 **{last_month_full_total}** 單)*")
            st.divider()

            # 推估明細與最終預測
            st.subheader("🔮 月底預估與推算")
            
            col_info, col_result = st.columns([1, 1])
            
            with col_info:
                st.markdown(f"""
                **📌 數據推算邏輯**
                * **營收預算 (未稅)**： **${TARGET_REVENUE:,}**
                * **日均增推算**： {this_month}/1 ~ {rolling_day} (均增 {raw_daily_avg:.1f} * {ADJUSTMENT_FACTOR} = {raw_daily_avg*ADJUSTMENT_FACTOR:.1f} ➔ 取 **{final_daily_avg}**)
                * **未來單數預估**： 日均增 {final_daily_avg} * 剩 {remaining_days} 天 * 客單價 ${UNIT_PRICE}。
                * 推估 {this_month}/{rolling_day+1} ~ {total_days_in_month} 將累積新增 **{est_future_new}** 單、續訂約 **{est_future_renew}** 單。
                """)

            with col_result:
                st.info(f"""
                **🏆 最終預估達成**
                
                * 已扣款： **${current_rev_gross:,}**
                * 未來預估： **${future_rev_est:,}**
                * 總額： **${total_rev_gross_est:,}**
                
                **回除 {TAX_RATE} 稅率預估未稅營收： ${total_rev_net_est:,.0f}**
                
                ### 👉 預估達成率： {final_achievement_rate:.0f}%
                """)

        except Exception as e:
            st.error(f"分析過程中發生錯誤：{e}")
            st.warning("請確認上傳的 Excel 檔案格式與原本的 report_supplier 結構相符。")
else:
    st.info("請上傳「本月」與「上月」的報表檔案，以上鎖定分析按鈕。")




# --- 4. 文字版執行結果 (方便複製) ---
            st.divider()
            st.subheader("📝 執行結果明細 (可直接複製報告)")
            
            # 組合字串
            adj_str = f"{raw_daily_avg:.1f} * {ADJUSTMENT_FACTOR} = {raw_daily_avg*ADJUSTMENT_FACTOR:.1f} -> 取 {final_daily_avg}"
            
            # 將所有 print 內容整理成一個多行字串
            report_text = f"""{this_month}月營收預算  ${TARGET_REVENUE:,}
截至{this_month}/{rolling_day}營收 $ {current_rev_gross:,}(已扣款)
{this_month}/1~{rolling_day} (均增 {adj_str} 位)
【日均增 {final_daily_avg}*剩{remaining_days}天*${UNIT_PRICE}】，推估{this_month}/{rolling_day+1}~{total_days_in_month} 累積新增 {est_future_new}單、續訂約 {est_future_renew}單
${current_rev_gross:,} + ${future_rev_est:,}= ${total_rev_gross_est:,}
回除{TAX_RATE}稅率={total_rev_gross_est:,}/{TAX_RATE}={total_rev_net_est:,.0f}
預估達成率 {final_achievement_rate:.0f}%"""

            # 用 st.code 呈現，右上角會自動自帶一個「複製」按鈕
            st.code(report_text, language="text")
