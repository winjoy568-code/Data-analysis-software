import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. 頁面設定 ---
st.set_page_config(page_title="生產效能智慧分析報告", layout="centered")

st.markdown("""
    <style>
    .main { background-color: #f9f9f9; }
    h1 { color: #2c3e50; font-family: 'Microsoft JhengHei'; }
    h2 { color: #34495e; border-bottom: 2px solid #3498db; padding-bottom: 10px; margin-top: 30px; }
    </style>
""", unsafe_allow_html=True)

# --- 2. 核心邏輯與智慧讀取 ---

def init_session_state():
    if 'data' not in st.session_state:
        # 預設範例數據
        st.session_state.data = pd.DataFrame([
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO2", "OEE": 0.501, "產量": 2009.5, "耗電量": 6.2},
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO4", "OEE": 0.554, "產量": 4416.5, "耗電量": 9.1},
            {"日期": "2025-11-18", "廠別": "A廠", "機台編號": "ACO4", "OEE": 0.605, "產量": 4921.5, "耗電量": 9.5},
        ])
        st.session_state.data['日期'] = pd.to_datetime(st.session_state.data['日期']).dt.date

init_session_state()

def smart_load_file(uploaded_file):
    """智慧讀取並轉換欄位名稱"""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # 1. 欄位對照字典 (左邊是你的Excel欄位，右邊是系統欄位)
        rename_map = {
            "設備": "機台編號",
            "用電量 (kWh)": "耗電量",
            "產量 (雙)": "產量",
            "OEE (%)": "OEE",
            "OEE(%)": "OEE"
        }
        df = df.rename(columns=rename_map)
        
        # 2. 處理必要欄位
        if "日期" in df.columns:
            df["日期"] = pd.to_datetime(df["日期"]).dt.date
            
        # 3. 自動修正 OEE (如果是 76.1 這種百分比格式，除以 100)
        if "OEE" in df.columns:
            if df["OEE"].mean() > 1.0: 
                df["OEE"] = df["OEE"] / 100.0
                
        # 4. 處理缺少的廠別
        if "廠別" not in df.columns:
            df["廠別"] = "匯入廠區" # 預設值
            
        # 5. 過濾出系統需要的欄位
        required_cols = ["日期", "廠別", "機台編號", "OEE", "產量", "耗電量"]
        
        # 檢查是否還有缺少的欄位
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            return None, f"缺少必要欄位: {missing}"
            
        return df[required_cols], "OK"
        
    except Exception as e:
        return None, str(e)

def calculate_metrics(df, elec_price):
    df["單位能耗"] = df["耗電量"] / df["產量"]
    best_energy_unit = df["單位能耗"].min()
    df["能源損失(元)"] = (df["單位能耗"] - best_energy_unit) * df["產量"] * elec_price
    df["能源損失(元)"] = df["能源損失(元)"].apply(lambda x: max(x, 0))
    df["效率排名"] = df["OEE"].rank(ascending=False, method='min')
    return df

# --- 3. 側邊欄：經典輸入介面 ---
st.sidebar.title("⚙️ 數據控制台")

# 參數設定
st.sidebar.subheader("1. 參數設定")
elec_price = st.sidebar.number_input("平均電價 (元/度)", value=3.5, step=0.1)

# 數據輸入切換
st.sidebar.subheader("2. 數據輸入")
input_mode = st.sidebar.radio("選擇模式", ["手動輸入", "上傳 Excel"])

if input_mode == "上傳 Excel":
    uploaded_file = st.sidebar.file_uploader("上傳報表 (支援欄位: 日期, 設備, OEE%, 產量, 用電量)", type=["xlsx", "csv"])
    if uploaded_file:
        new_df, status = smart_load_file(uploaded_file)
        if status == "OK":
            st.session_state.data = pd.concat([st.session_state.data, new_df], ignore_index=True)
            st.sidebar.success(f"成功匯入 {len(new_df)} 筆數據！")
            st.rerun()
        else:
            st.sidebar.error(f"讀取失敗: {status}")
            st.sidebar.info("提示: 請確保 Excel 包含「日期, 設備, OEE (%), 產量 (雙), 用電量 (kWh)」等資訊")

else:
    # 回歸經典：表單輸入模式
    with st.sidebar.form("add_data_form"):
        st.write("📝 新增單筆紀錄")
        col1, col2 = st.columns(2)
        in_date = col1.date_input("日期")
        in_factory = col2.text_input("廠別", "A廠")
        in_machine = st.text_input("設備/機台", "ACO-X")
        
        in_oee = st.number_input("OEE (0.0 - 1.0)", 0.0, 1.0, 0.85, 0.01)
        in_output = st.number_input("產量 (雙)", 1, 10000, 1000)
        in_power = st.number_input("用電量 (kWh)", 0.0, 10000.0, 150.0)
        
        submitted = st.form_submit_button("💾 加入數據庫", type="primary")
        
        if submitted:
            new_row = {
                "日期": in_date, "廠別": in_factory, "機台編號": in_machine,
                "OEE": in_oee, "產量": in_output, "耗電量": in_power
            }
            st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
            st.sidebar.success("已新增！")
            st.rerun()

# 清除按鈕
if st.sidebar.button("🗑️ 清除所有數據"):
    st.session_state.data = pd.DataFrame(columns=["日期", "廠別", "機台編號", "OEE", "產量", "耗電量"])
    st.rerun()

# --- 4. 報告主體 ---

if not st.session_state.data.empty:
    df_analysis = calculate_metrics(st.session_state.data.copy(), elec_price)
    
    # 判斷分析維度
    factory_count = df_analysis["廠別"].nunique()
    analysis_mode = "跨廠總和" if factory_count > 1 else "單廠設備"
    group_col = "廠別" if factory_count > 1 else "機台編號"

    st.title(f"📊 {analysis_mode}效能診斷報告")
    st.markdown(f"**分析時間：** {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}")

    # 1. 分析範圍
    st.header("1. 分析範圍與目的")
    st.info(f"""
    **🎯 分析對象 ({analysis_mode})**
    * **對象**：{', '.join(df_analysis[group_col].unique())}
    * **期間**：{df_analysis['日期'].min()} ~ {df_analysis['日期'].max()}
    * **目的**：評估生產效率與能源使用，計算潛在貨幣損失。
    """)

    # 2. 指標定義
    st.header("2. 分析指標定義")
    c1, c2 = st.columns(2)
    c1.markdown("**⚡ 單位能耗**：每生產一雙鞋的電力成本 (kWh/雙)。")
    c2.markdown("**💰 能源損失**：因效率未達最佳水準而多浪費的電費 (NTD)。")

    # 3. 數據全貌
    st.header("3. 數據全貌與排名")
    
    def highlight_oee(val):
        return 'background-color: #d4edda' if val >= 0.85 else 'background-color: #f8d7da' if val < 0.70 else ''

    # 顯示使用者習慣的欄位名稱
    display_df = df_analysis.rename(columns={
        "機台編號": "設備", "耗電量": "用電量(kWh)", "產量": "產量(雙)"
    })
    
    st.dataframe(
        display_df[["日期", "廠別", "設備", "OEE", "產量(雙)", "用電量(kWh)", "單位能耗", "效率排名", "能源損失(元)"]]
        .sort_values("效率排名").style
        .applymap(highlight_oee, subset=['OEE'])
        .format({"OEE": "{:.2%}", "單位能耗": "{:.5f}", "能源損失(元)": "${:,.0f}"}),
        use_container_width=True,
        hide_index=True
    )
    
    # 4. 生產穩定度
    st.header("4. 生產穩定度分析")
    df_trend = df_analysis.groupby([group_col, "日期"])[["產量", "OEE"]].mean().reset_index()
    
    fig_stab = go.Figure()
    for item in df_analysis[group_col].unique():
        subset = df_trend[df_trend[group_col] == item]
        fig_stab.add_trace(go.Bar(x=subset["日期"], y=subset["產量"], name=f"{item} 產量", opacity=0.3))
        fig_stab.add_trace(go.Scatter(x=subset["日期"], y=subset["OEE"], name=f"{item} OEE", yaxis="y2", mode='lines+markers'))

    fig_stab.update_layout(
        title="產能與 OEE 走勢圖",
        yaxis=dict(title="產量"),
        yaxis2=dict(title="OEE", overlaying="y", side="right", range=[0, 1.1], tickformat=".0%"),
        legend=dict(orientation="h", y=-0.2)
    )
    st.plotly_chart(fig_stab, use_container_width=True)

    # 5. 能耗矩陣
    st.header("5. 能耗效率矩陣")
    fig_energy = px.scatter(
        df_analysis, x="OEE", y="單位能耗", color=group_col, size="產量",
        hover_data=["日期", "能源損失(元)"],
        title="OEE vs 單位能耗 (氣泡=產量)", labels={"OEE": "OEE (效率)", "單位能耗": "單位能耗 (kWh/雙)"}
    )
    st.plotly_chart(fig_energy, use_container_width=True)

    # 6. 結論
    st.header("6. 智慧診斷結論")
    agg = df_analysis.groupby(group_col).agg({"OEE": "mean", "能源損失(元)": "sum"}).reset_index()
    best = agg.loc[agg["OEE"].idxmax()]
    worst = agg.loc[agg["OEE"].idxmin()]
    
    st.markdown(f"""
    * **表現最佳**：**{best[group_col]}** (平均 OEE {best['OEE']:.1%})。
    * **需改善**：**{worst[group_col]}** (平均 OEE {worst['OEE']:.1%})。
    * **潛在效益**：若全廠最佳化，預計可節省電費約 **NT$ {agg['能源損失(元)'].sum():,.0f}**。
    """)

else:
    st.info("👈 請在左側使用「手動輸入」或「上傳 Excel」建立數據")
