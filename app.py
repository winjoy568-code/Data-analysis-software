import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. 頁面設定 (UI 設計) ---
st.set_page_config(page_title="生產效能智慧分析報告 Pro", layout="centered")

# 自訂 CSS 以符合報告格式 (直式、清晰)
st.markdown("""
    <style>
    .main { background-color: #f9f9f9; }
    h1 { color: #2c3e50; font-family: 'Microsoft JhengHei'; }
    h2 { color: #34495e; border-bottom: 2px solid #3498db; padding-bottom: 10px; margin-top: 30px; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .report-text { font-size: 1.1rem; line-height: 1.6; color: #444; }
    .highlight-good { color: #27ae60; font-weight: bold; }
    .highlight-bad { color: #c0392b; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# --- 2. 數據處理核心邏輯 ---

def init_session_state():
    # 初始化數據庫，預設範例數據 [cite: 12, 13]
    if 'data' not in st.session_state:
        st.session_state.data = pd.DataFrame([
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO2", "OEE": 0.82, "產量": 1150, "耗電量": 155.0},
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO3", "OEE": 0.68, "產量": 920, "耗電量": 148.0},
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO4", "OEE": 0.91, "產量": 1500, "耗電量": 160.2},
            {"日期": "2025-11-18", "廠別": "A廠", "機台編號": "ACO2", "OEE": 0.85, "產量": 1200, "耗電量": 152.0},
            {"日期": "2025-11-18", "廠別": "A廠", "機台編號": "ACO3", "OEE": 0.70, "產量": 950, "耗電量": 146.0},
            {"日期": "2025-11-18", "廠別": "A廠", "機台編號": "ACO4", "OEE": 0.89, "產量": 1480, "耗電量": 158.0},
        ])
        # 確保日期格式正確
        st.session_state.data['日期'] = pd.to_datetime(st.session_state.data['日期']).dt.date

init_session_state()

def calculate_metrics(df, elec_price):
    # 1. 基礎計算 [cite: 18]
    df["單位能耗"] = df["耗電量"] / df["產量"]
    
    # 2. 基準比較 (Benchmarking) 
    # 找出全場最佳 OEE 作為黃金標準
    best_oee = df["OEE"].max()
    # 找出全場最佳能耗 (最低)
    best_energy_unit = df["單位能耗"].min()
    
    # 計算落差損失 (假設每度電費 elec_price 元)
    # 能源損失金額 = (目前能耗 - 最佳能耗) * 產量 * 電價
    df["能源損失(元)"] = (df["單位能耗"] - best_energy_unit) * df["產量"] * elec_price
    df["能源損失(元)"] = df["能源損失(元)"].apply(lambda x: max(x, 0)) # 不會有負的損失
    
    # 3. 排名 [cite: 25]
    df["效率排名"] = df["OEE"].rank(ascending=False, method='min')
    
    return df

# --- 3. 側邊欄：進階數據控制台 ---
st.sidebar.title("⚙️ 控制台")

# 參數設定 (增加專業度)
st.sidebar.subheader("1. 參數設定")
elec_price = st.sidebar.number_input("平均電價 (元/度)", value=3.5, step=0.1)

# 數據管理 
st.sidebar.subheader("2. 數據管理")
input_method = st.sidebar.radio("數據來源", ["手動編輯/檢視", "上傳 Excel"])

if st.sidebar.button("🗑️ 清除所有數據", type="primary"):
    st.session_state.data = pd.DataFrame(columns=["日期", "廠別", "機台編號", "OEE", "產量", "耗電量"])
    st.rerun()

# 數據輸入邏輯
df_input = st.session_state.data.copy()

if input_method == "上傳 Excel":
    uploaded_file = st.sidebar.file_uploader("上傳報表", type=["xlsx", "csv"])
    if uploaded_file:
        try:
            if uploaded_file.name.endswith('.csv'):
                new_data = pd.read_csv(uploaded_file)
            else:
                new_data = pd.read_excel(uploaded_file)
            # 簡單欄位檢查
            required_cols = ["日期", "廠別", "機台編號", "OEE", "產量", "耗電量"]
            if all(col in new_data.columns for col in required_cols):
                st.session_state.data = pd.concat([st.session_state.data, new_data], ignore_index=True)
                st.sidebar.success("上傳成功！")
                st.rerun()
            else:
                st.sidebar.error(f"格式錯誤，需包含: {required_cols}")
        except Exception as e:
            st.sidebar.error(f"讀取錯誤: {e}")

else:
    # 使用 Data Editor 達成單筆新增/刪除/修改 
    st.sidebar.info("👇 在下方表格直接修改，可新增行或勾選刪除")
    edited_df = st.data_editor(
        df_input,
        num_rows="dynamic", # 允許新增
        use_container_width=True,
        column_config={
            "日期": st.column_config.DateColumn("日期"),
            "OEE": st.column_config.NumberColumn("OEE", min_value=0.0, max_value=1.0, format="%.2f"),
        }
    )
    # 同步回 Session State
    if not edited_df.equals(st.session_state.data):
        st.session_state.data = edited_df
        st.rerun()

# --- 4. 報告主內容 (直式輸出) [cite: 33] ---

if not st.session_state.data.empty:
    df_analysis = calculate_metrics(st.session_state.data.copy(), elec_price)
    
    # 智慧判斷分析範圍 (單廠 vs 跨廠) [cite: 27, 29]
    factory_count = df_analysis["廠別"].nunique()
    if factory_count > 1:
        analysis_mode = "跨廠總和"
        group_col = "廠別"
    else:
        analysis_mode = "單廠設備"
        group_col = "機台編號"

    # 標題
    st.title(f"📊 {analysis_mode}效能與能耗診斷報告")
    st.markdown(f"**報告產出時間：** {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}")
    
    # Section 1: 分析範圍與目的 [cite: 7, 11]
    st.header("1. 分析範圍與目的")
    col1, col2 = st.columns([1, 2])
    with col1:
        st.info(f"""
        **🎯 分析對象**
        * **模式**：{analysis_mode}
        * **對象**：{', '.join(df_analysis[group_col].unique())}
        * **期間**：{df_analysis['日期'].min()} ~ {df_analysis['日期'].max()}
        """)
    with col2:
        st.markdown(f"""
        **📌 分析目的**
        1.  **評估效率**：分析 {len(df_analysis[group_col].unique())} 個單位的生產與能源效率，找出熱點。
        2.  **量化損失**：透過基準比較 (Benchmarking)，計算低效造成的產能與能源貨幣損失。
        3.  **改善建議**：提供具體行動方針以提升整體 OEE。
        """)

    # Section 2: 分析處理說明 [cite: 16]
    st.header("2. 分析指標定義")
    st.markdown("""
    > 本報告採用以下關鍵指標進行診斷：
    
    * **⚡ 單位能耗 (Unit Energy Consumption)**：`總用電 ÷ 總產量`。排除規模差異，直接比較每生產單位的電力成本。**[數值越低越好]**
    * **📈 OEE (整體設備效率)**：衡量設備穩定性的核心指標。重點分析「低 OEE 高耗能」的異常空轉。**[數值越高越好]**
    * **🏆 基準比較 (Benchmarking)**：將表現最佳者設為標準，計算其他設備的落差空間。
    """)

    # Section 3: 原始數據與排名 [cite: 24]
    st.header("3. 數據全貌與排名")
    
    # 使用 Pandas Styler 製作有設計感的表格 (Highlighter)
    def highlight_oee(val):
        color = '#d4edda' if val >= 0.85 else '#f8d7da' if val < 0.70 else ''
        return f'background-color: {color}'

    display_cols = ["日期", "廠別", "機台編號", "產量", "耗電量", "OEE", "單位能耗", "效率排名", "能源損失(元)"]
    st.dataframe(
        df_analysis[display_cols].sort_values("效率排名").style
        .applymap(highlight_oee, subset=['OEE'])
        .format({
            "OEE": "{:.2%}", 
            "單位能耗": "{:.4f}", 
            "能源損失(元)": "${:,.0f}"
        }),
        use_container_width=True,
        hide_index=True
    )
    st.caption("* 綠色底色代表優異 (OEE ≥ 85%)，紅色底色代表需改善 (OEE < 70%)")

    # Section 4: 生產穩定度分析 [cite: 26, 34]
    st.header("4. 生產穩定度分析")
    st.markdown(f"針對 **{group_col}** 進行日產量與效率穩定性檢視。")

    # 聚合數據
    df_trend = df_analysis.groupby([group_col, "日期"])[["產量", "OEE"]].mean().reset_index()
    
    fig_stab = go.Figure()
    colors = px.colors.qualitative.Plotly
    
    for i, item in enumerate(df_analysis[group_col].unique()):
        subset = df_trend[df_trend[group_col] == item]
        # 產量 (Bar)
        fig_stab.add_trace(go.Bar(
            x=subset["日期"], y=subset["產量"], name=f"{item} 產量",
            marker_color=colors[i % len(colors)], opacity=0.3
        ))
        # OEE (Line)
        fig_stab.add_trace(go.Scatter(
            x=subset["日期"], y=subset["OEE"], name=f"{item} OEE",
            yaxis="y2", line=dict(color=colors[i % len(colors)], width=3), mode='lines+markers'
        ))

    fig_stab.update_layout(
        title="產能與效率複合趨勢圖",
        yaxis=dict(title="產量 (雙)"),
        yaxis2=dict(title="OEE (%)", overlaying="y", side="right", range=[0, 1.1], tickformat=".0%"),
        legend=dict(orientation="h", y=-0.2),
        hovermode="x unified"
    )
    st.plotly_chart(fig_stab, use_container_width=True)

    # Section 5: 能耗分析 [cite: 28, 35]
    st.header("5. 能耗效率矩陣分析")
    
    # 計算平均線
    avg_oee = df_analysis["OEE"].mean()
    avg_energy = df_analysis["單位能耗"].mean()

    fig_energy = px.scatter(
        df_analysis, x="OEE", y="單位能耗",
        color=group_col, size="產量",
        hover_data=["日期", "能源損失(元)"],
        text=group_col,
        title="OEE vs 單位能耗 矩陣圖 (氣泡大小=產量)",
        labels={"OEE": "OEE (效率)", "單位能耗": "單位能耗 (kWh/雙)"}
    )
    
    # 畫象限分割線
    fig_energy.add_vline(x=avg_oee, line_dash="dash", line_color="gray", annotation_text="平均 OEE")
    fig_energy.add_hline(y=avg_energy, line_dash="dash", line_color="gray", annotation_text="平均能耗")
    
    # 標註象限意義
    fig_energy.add_annotation(x=df_analysis["OEE"].max(), y=df_analysis["單位能耗"].min(), text="🏆 最佳區 (高效節能)", showarrow=False, bgcolor="#d4edda")
    fig_energy.add_annotation(x=df_analysis["OEE"].min(), y=df_analysis["單位能耗"].max(), text="⚠️ 改善區 (低效耗能)", showarrow=False, bgcolor="#f8d7da")

    st.plotly_chart(fig_energy, use_container_width=True)

    # Section 6: 結論與行動建議 [cite: 30, 31]
    st.header("6. 智慧診斷結論與行動建議")

    # 自動生成分析文案
    agg_df = df_analysis.groupby(group_col).agg({
        "OEE": "mean", "單位能耗": "mean", "能源損失(元)": "sum"
    }).reset_index()
    
    best_performer = agg_df.loc[agg_df["OEE"].idxmax()]
    worst_performer = agg_df.loc[agg_df["OEE"].idxmin()]
    total_loss = agg_df["能源損失(元)"].sum()

    st.markdown(f"""
    ### 📊 綜合診斷總結
    1.  **績效排名**：本次分析中，**{best_performer[group_col]}** 表現最佳，平均 OEE 達 **{best_performer['OEE']:.1%}**，單位能耗最低（**{best_performer['單位能耗']:.4f}** kWh/雙）。
    2.  **改善重點**：**{worst_performer[group_col]}** 表現最弱，OEE 僅 **{worst_performer['OEE']:.1%}**。
    3.  **潛在效益**：若所有設備皆達到最佳設備的水準，估計此期間可節省能源成本約 **NT$ {total_loss:,.0f} 元**。

    ### 🚀 具體行動建議
    * **針對 {worst_performer[group_col]}**：
        * 檢視「單位能耗」是否過高？若是，請檢查待機時間是否未關機。
        * 請調閱 {worst_performer[group_col]} 的異常代碼 (Error Code)，確認是否為頻繁短停機造成 OEE 低落。
    * **管理層面**：
        * 建議將 **{best_performer[group_col]}** 的參數設定 (Parameter) 匯出，作為 {worst_performer[group_col]} 的標準化作業參數。
    """)

else:
    st.warning("請在左側輸入數據或上傳 Excel 檔案以開始分析")
