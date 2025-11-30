import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import time

# --- 1. 頁面設定 ---
st.set_page_config(page_title="生產效能智慧分析系統", layout="centered")

# CSS 優化：調整標題與區塊間距
st.markdown("""
    <style>
    .main { background-color: #fcfcfc; }
    .stButton>button { width: 100%; border-radius: 8px; height: 3.5em; font-weight: bold; font-size: 1.1em; }
    h1 { color: #2c3e50; font-family: 'Microsoft JhengHei'; }
    .step-header { color: #2980b9; font-weight: bold; font-size: 1.3em; margin-top: 20px; border-left: 5px solid #2980b9; padding-left: 10px; }
    </style>
""", unsafe_allow_html=True)

# --- 2. 核心邏輯 ---

def init_session_state():
    if 'input_data' not in st.session_state:
        # 預設範例數據 (方便你第一次看)
        st.session_state.input_data = pd.DataFrame([
            {"日期": "2025-11-17", "廠別": "A廠", "設備": "ACO2", "OEE(%)": 50.1, "產量(雙)": 2009.5, "用電量(kWh)": 6.2},
            {"日期": "2025-11-17", "廠別": "A廠", "設備": "ACO4", "OEE(%)": 55.4, "產量(雙)": 4416.5, "用電量(kWh)": 9.1},
        ])
        st.session_state.input_data['日期'] = pd.to_datetime(st.session_state.input_data['日期']).dt.date

init_session_state()

def smart_load_file(uploaded_file):
    """讀取 Excel 並轉成標準格式"""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # 簡單欄位處理
        if "日期" in df.columns:
            df["日期"] = pd.to_datetime(df["日期"]).dt.date
        if "廠別" not in df.columns:
            df["廠別"] = "匯入廠區"
        return df, "OK"
    except Exception as e:
        return None, str(e)

# --- 3. 介面設計：Step 1 數據輸入 (上方) ---

st.title("🏭 生產效能智慧分析系統")

st.markdown('<div class="step-header">1. 數據輸入 (Data Input)</div>', unsafe_allow_html=True)
st.caption("請在下方表格直接輸入數據，或點擊右上角「Browse files」上傳 Excel。")

# 上傳區塊 (放在表格上方)
uploaded_file = st.file_uploader("批次匯入 Excel (選填)", type=["xlsx", "csv"], label_visibility="collapsed")
if uploaded_file:
    new_df, status = smart_load_file(uploaded_file)
    if status == "OK":
        st.session_state.input_data = new_df # 覆蓋數據
    else:
        st.error(f"檔案讀取錯誤: {status}")

# 核心：可編輯表格 (Data Editor)
# num_rows="dynamic" 讓你可以新增、刪除行
st.info("💡 操作提示：點擊表格可直接修改。若要**刪除單筆**，請點擊該行左側選取後，按 Delete 鍵或表格右上角垃圾桶。")
edited_df = st.data_editor(
    st.session_state.input_data,
    num_rows="dynamic", # 關鍵：允許新增與刪除
    use_container_width=True,
    column_config={
        "日期": st.column_config.DateColumn("日期"),
        "OEE(%)": st.column_config.NumberColumn("OEE(%)", format="%.1f"),
        "產量(雙)": st.column_config.NumberColumn("產量(雙)"),
        "用電量(kWh)": st.column_config.NumberColumn("用電量(kWh)"),
    },
    key="editor" 
)

# 快速清空按鈕
if st.button("🗑️ 清空表格數據", help="點擊後將清除上方所有內容"):
    st.session_state.input_data = pd.DataFrame(columns=["日期", "廠別", "設備", "OEE(%)", "產量(雙)", "用電量(kWh)"])
    st.rerun()

# --- 4. 介面設計：Step 2 參數設定 (下方) ---

st.markdown('<div class="step-header">2. 參數設定 (Parameters)</div>', unsafe_allow_html=True)

col_param1, col_param2 = st.columns(2)
with col_param1:
    elec_price = st.number_input("平均電價 (元/度)", value=3.5, step=0.1)
with col_param2:
    target_oee = st.number_input("目標 OEE 基準 (%)", value=85.0, step=0.5)

st.write("") # 空行

# --- 5. 執行分析 (動態按鈕) ---

# 這裡使用一個 Primary 按鈕作為觸發
start_analysis = st.button("🚀 開始執行分析 (Start Analysis)", type="primary")

if start_analysis:
    # --- 動態分析效果 ---
    with st.spinner('🔄 正在進行 AI 運算與數據建模，請稍候...'):
        time.sleep(1.0) # 模擬運算時間 (讓使用者感覺真的在跑)
        
        # 1. 鎖定數據
        df_clean = edited_df.copy()
        
        # 2. 欄位轉譯 (Mapping)
        rename_map = {"設備": "機台編號", "用電量(kWh)": "耗電量", "產量(雙)": "產量", "OEE(%)": "OEE_RAW"}
        for user_col, sys_col in rename_map.items():
            if user_col in df_clean.columns:
                df_clean = df_clean.rename(columns={user_col: sys_col})

        # 3. 檢查數據完整性
        required = ["機台編號", "耗電量", "產量", "OEE_RAW"]
        if df_clean.empty or not all(col in df_clean.columns for col in required):
            st.error("❌ 數據不足或欄位錯誤，無法進行分析。請檢查上方表格。")
        else:
            # 4. 運算邏輯
            df_clean["OEE"] = df_clean["OEE_RAW"].apply(lambda x: x / 100.0 if x > 1.0 else x)
            df_clean["單位能耗"] = df_clean["耗電量"] / df_clean["產量"]
            best_energy = df_clean["單位能耗"].min()
            df_clean["能源損失(元)"] = (df_clean["單位能耗"] - best_energy) * df_clean["產量"] * elec_price
            df_clean["能源損失(元)"] = df_clean["能源損失(元)"].apply(lambda x: max(x, 0))
            df_clean["效率排名"] = df_clean["OEE"].rank(ascending=False, method='min')

            # 判斷維度
            if "廠別" not in df_clean.columns: df_clean["廠別"] = "預設廠區"
            is_multi_factory = df_clean["廠別"].nunique() > 1
            group_col = "廠別" if is_multi_factory else "機台編號"
            analysis_title = "跨廠總和" if is_multi_factory else "單廠設備"

            # --- 報告產出區 ---
            st.success("✅ 分析完成！")
            st.markdown("---")

            # 標題區
            st.title(f"📊 {analysis_title}效能診斷報告")
            
            # (A) 數據摘要
            st.subheader("1. 數據全貌與排名")
            display_cols = ["日期", "廠別", "機台編號", "OEE", "產量", "耗電量", "單位能耗", "效率排名", "能源損失(元)"]
            final_table = df_clean[display_cols].rename(columns={"機台編號": "設備", "耗電量": "用電量(kWh)", "產量": "產量(雙)"})
            
            def highlight(val):
                return 'background-color: #d4edda' if val >= 0.85 else 'background-color: #f8d7da' if val < 0.70 else ''

            st.dataframe(
                final_table.sort_values("效率排名").style
                .applymap(highlight, subset=['OEE'])
                .format({"OEE": "{:.1%}", "單位能耗": "{:.5f}", "能源損失(元)": "${:,.0f}"}),
                use_container_width=True, hide_index=True
            )

            # (B) 圖表區
            col_chart1, col_chart2 = st.columns(2)
            
            with col_chart1:
                st.subheader("2. 生產穩定度")
                df_trend = df_clean.groupby([group_col, "日期"])[["產量", "OEE"]].mean().reset_index()
                fig_stab = go.Figure()
                for item in df_clean[group_col].unique():
                    subset = df_trend[df_trend[group_col] == item]
                    fig_stab.add_trace(go.Bar(x=subset["日期"], y=subset["產量"], name=f"{item} 產量", opacity=0.3))
                    fig_stab.add_trace(go.Scatter(x=subset["日期"], y=subset["OEE"], name=f"{item} OEE", yaxis="y2", mode='lines+markers'))
                
                fig_stab.update_layout(yaxis=dict(title="產量"), yaxis2=dict(title="OEE", overlaying="y", side="right", range=[0, 1.1]), legend=dict(orientation="h", y=-0.2))
                st.plotly_chart(fig_stab, use_container_width=True)

            with col_chart2:
                st.subheader("3. 能耗矩陣")
                fig_energy = px.scatter(
                    df_clean, x="OEE", y="單位能耗", color=group_col, size="產量",
                    title="OEE vs 單位能耗", labels={"OEE": "OEE", "單位能耗": "能耗"}
                )
                fig_energy.add_vline(x=df_clean["OEE"].mean(), line_dash="dash", line_color="gray")
                fig_energy.add_hline(y=df_clean["單位能耗"].mean(), line_dash="dash", line_color="gray")
                st.plotly_chart(fig_energy, use_container_width=True)

            # (C) 結論區
            st.subheader("4. 智慧結論")
            agg = df_clean.groupby(group_col).agg({"OEE": "mean", "能源損失(元)": "sum"}).reset_index()
            best = agg.loc[agg["OEE"].idxmax()]
            worst = agg.loc[agg["OEE"].idxmin()]
            
            st.info(f"""
            **診斷結果：**
            * 表現最佳：**{best[group_col]}** (OEE {best['OEE']:.1%})
            * 需改善：**{worst[group_col]}** (OEE {worst['OEE']:.1%})
            * 此期間潛在可節省成本： **NT$ {agg['能源損失(元)'].sum():,.0f}**
            """)
