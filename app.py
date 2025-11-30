import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. 頁面設定 ---
st.set_page_config(page_title="生產效能自動化分析報告", layout="centered") # 使用 centered 模擬直式報告

# --- 2. 核心邏輯與數據處理 ---

# 初始化 Session State (用於儲存手動輸入的數據)
if 'manual_data' not in st.session_state:
    # 預設模擬數據 (符合你的 PRD 範例)
    st.session_state.manual_data = pd.DataFrame([
        {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO2", "OEE": 0.85, "產量": 1200, "耗電量": 150.5},
        {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO3", "OEE": 0.72, "產量": 980, "耗電量": 145.0},
        {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO4", "OEE": 0.91, "產量": 1500, "耗電量": 160.2},
        {"日期": "2025-11-18", "廠別": "A廠", "機台編號": "ACO2", "OEE": 0.88, "產量": 1250, "耗電量": 152.0},
        {"日期": "2025-11-18", "廠別": "A廠", "機台編號": "ACO3", "OEE": 0.75, "產量": 1000, "耗電量": 148.0},
        {"日期": "2025-11-18", "廠別": "A廠", "機台編號": "ACO4", "OEE": 0.89, "產量": 1480, "耗電量": 158.0},
        # ... 更多模擬數據可以加在這裡
    ])

def process_data(df):
    # 自動計算：單位能耗 (kWh/雙)
    df["單位能耗"] = df["耗電量"] / df["產量"]
    # 自動計算：效率排名 (根據 OEE 由高到低)
    df["OEE排名"] = df["OEE"].rank(ascending=False, method='min')
    return df

# --- 3. 側邊欄：數據輸入區 ---
st.sidebar.header("📥 數據輸入控制台")
input_mode = st.sidebar.radio("選擇數據來源", ["使用範例/手動輸入", "上傳 Excel 檔案"])

df = pd.DataFrame()

if input_mode == "上傳 Excel 檔案":
    uploaded_file = st.sidebar.file_uploader("上傳 Excel (需包含: 日期, 廠別, 機台編號, OEE, 產量, 耗電量)", type=["xlsx", "csv"])
    if uploaded_file:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.sidebar.error(f"檔案讀取錯誤: {e}")
else:
    # 手動輸入介面
    st.sidebar.subheader("新增單筆數據")
    with st.sidebar.form("add_data_form"):
        col1, col2 = st.columns(2)
        in_date = col1.date_input("日期")
        in_factory = col2.text_input("廠別", "A廠")
        in_machine = st.text_input("機台編號", "ACOX")
        in_oee = st.number_input("OEE (0.0 - 1.0)", 0.0, 1.0, 0.85, 0.01)
        in_output = st.number_input("產量 (雙)", 1, 10000, 1000)
        in_power = st.number_input("耗電量 (kWh)", 0.0, 10000.0, 100.0)
        submitted = st.form_submit_button("💾 加入數據庫")
        
        if submitted:
            new_data = {
                "日期": str(in_date), "廠別": in_factory, "機台編號": in_machine,
                "OEE": in_oee, "產量": in_output, "耗電量": in_power
            }
            st.session_state.manual_data = pd.concat([st.session_state.manual_data, pd.DataFrame([new_data])], ignore_index=True)
            st.sidebar.success("數據已新增！")
    
    df = st.session_state.manual_data

# 確保數據不為空才執行分析
if not df.empty:
    df = process_data(df)
    
    # 判斷分析維度 (單廠 vs 跨廠)
    unique_factories = df["廠別"].nunique()
    analysis_level = "跨廠總和" if unique_factories > 1 else "單機台"
    group_col = "廠別" if analysis_level == "跨廠總和" else "機台編號"

    # --- 4. 報告主體 (直式輸出) ---

    st.title("📊 生產效能與能耗診斷報告")
    st.markdown(f"**分析維度偵測：** {analysis_level}分析模式")
    st.markdown("---")

    # 1. 分析範圍
    st.header("1. 分析範圍與目的")
    st.info(f"""
    **分析目的：**
    * 評估{analysis_level}的生產效率與能源使用效率，找出能耗熱點。
    * 透過對比分析，確立最佳生產模式，量化潛在損失。
    
    **分析範圍：**
    * **對象：** {', '.join(df[group_col].unique())}
    * **時間：** {df['日期'].min()} 至 {df['日期'].max()}
    * **數據來源：** 系統整合日報表（含產量、OEE、用電量）
    """)

    # 2. 分析處理說明
    st.header("2. 分析指標定義")
    col_def1, col_def2 = st.columns(2)
    with col_def1:
        st.markdown("""
        **⚡ 單位能耗 (Unit Energy Consumption)**
        * 公式：`總用電 ÷ 總產量`
        * 意義：每生產一雙鞋的電力成本。**數值越低越好**。
        """)
    with col_def2:
        st.markdown("""
        **📈 OEE (整體設備效率)**
        * 意義：衡量設備穩定性核心指標。
        * 重點：分析是否出現「低OEE、高耗能」的空轉浪費。
        """)

    # 3. 原始數據全貌
    st.header("3. 數據全貌與排名")
    st.write("以下表格已自動計算單位能耗與效率排名，並標示表現優異者。")
    
    # 格式化表格顯示
    st.dataframe(
        df.sort_values(by="OEE", ascending=False),
        column_config={
            "OEE": st.column_config.ProgressColumn("OEE", format="%.2f", min_value=0, max_value=1),
            "單位能耗": st.column_config.NumberColumn("單位能耗 (kWh/雙)", format="%.4f"),
            "耗電量": st.column_config.NumberColumn("總耗電 (kWh)", format="%.1f"),
            "OEE排名": st.column_config.NumberColumn("排名", help="數字越小越好")
        },
        use_container_width=True,
        hide_index=True
    )

    # 4. 生產穩定度分析
    st.header("4. 生產穩定度分析")
    st.markdown(f"針對 **{group_col}** 進行 OEE 趨勢與產量穩定性檢視。")

    # 雙軸圖：Bar(產量) + Line(OEE)
    fig_stab = go.Figure()
    
    # 這裡做一個簡單的平均聚合以便畫圖
    df_agg = df.groupby([group_col, "日期"])[["產量", "OEE"]].mean().reset_index()
    
    for item in df[group_col].unique():
        subset = df_agg[df_agg[group_col] == item]
        fig_stab.add_trace(go.Bar(
            x=subset["日期"], y=subset["產量"], name=f"{item} 產量", opacity=0.5
        ))
        fig_stab.add_trace(go.Scatter(
            x=subset["日期"], y=subset["OEE"], name=f"{item} OEE", yaxis="y2", mode='lines+markers'
        ))

    fig_stab.update_layout(
        title="產量與 OEE 走勢複合圖",
        yaxis=dict(title="產量 (雙)"),
        yaxis2=dict(title="OEE", overlaying="y", side="right", range=[0, 1]),
        legend=dict(orientation="h", y=-0.2),
        height=400
    )
    st.plotly_chart(fig_stab, use_container_width=True)
    
    st.caption("說明：折線代表設備效率(OEE)，長條代表實際產出。若折線高但長條低，可能代表速度慢或小停機多；若兩者皆低則為重大異常。")

    # 5. 能耗分析
    st.header("5. 能耗效率矩陣")
    st.markdown("透過 **OEE (X軸)** 與 **單位能耗 (Y軸)** 的關係，找出「黃金生產區」與「浪費區」。")

    fig_energy = px.scatter(
        df, x="OEE", y="單位能耗", 
        color=group_col, size="產量", hover_data=["日期"],
        title="能耗效率矩陣分析 (氣泡大小=產量)",
        labels={"單位能耗": "單位能耗 (kWh/雙, 越低越好)", "OEE": "OEE (越高越好)"}
    )
    # 畫十字線 (平均值)
    avg_oee = df["OEE"].mean()
    avg_energy = df["單位能耗"].mean()
    fig_energy.add_hline(y=avg_energy, line_dash="dash", annotation_text="平均耗能")
    fig_energy.add_vline(x=avg_oee, line_dash="dash", annotation_text="平均OEE")

    st.plotly_chart(fig_energy, use_container_width=True)
    st.caption("說明：位於**右下角**的點位最佳（高效率、低耗能）；位於**左上角**的點位最差（低效率、高耗能），為優先改善對象。")

    # 6. 結論與建議
    st.header("6. 智慧診斷結論與行動建議")
    
    # 簡單的規則基礎自動化結論
    best_machine = df.groupby(group_col)["OEE"].mean().idxmax()
    worst_machine = df.groupby(group_col)["OEE"].mean().idxmin()
    worst_energy_machine = df.groupby(group_col)["單位能耗"].mean().idxmax()
    
    st.markdown(f"""
    **📊 數據總結：**
    1.  **表現最佳：** **{best_machine}** 在分析期間內平均 OEE 最高，為目前的基準標竿 (Benchmark)。
    2.  **需關注對象：** **{worst_machine}** 的平均效率最低，且 **{worst_energy_machine}** 的單位生產成本最高。
    
    **🚀 行動建議：**
    1.  **複製成功模式：** 請產線主管分析 {best_machine} 的操作參數與排程方式，嘗試將其模式複製到 {worst_machine}。
    2.  **能耗異常排查：** 針對位於能耗矩陣「左上角」的時段/機台，檢查是否在低產量時未執行待機節能（如空壓機空轉）。
    3.  **減少短暫停機：** 若 OEE 低落主因為性能稼動率低，建議優先檢查進料順暢度。
    """)
    
else:
    st.warning("👈 請在左側輸入數據或上傳 Excel 檔案以開始分析")