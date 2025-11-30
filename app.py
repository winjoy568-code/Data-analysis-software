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

# --- 2. 核心邏輯 ---

def init_session_state():
    if 'data' not in st.session_state:
        # 預設範例數據
        st.session_state.data = pd.DataFrame([
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO2", "OEE": 0.82, "產量": 1150, "耗電量": 155.0},
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO3", "OEE": 0.68, "產量": 920, "耗電量": 148.0},
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO4", "OEE": 0.91, "產量": 1500, "耗電量": 160.2},
        ])
        st.session_state.data['日期'] = pd.to_datetime(st.session_state.data['日期']).dt.date

init_session_state()

def calculate_metrics(df, elec_price):
    # 計算單位能耗與成本損失
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
    uploaded_file = st.sidebar.file_uploader("上傳報表", type=["xlsx", "csv"])
    if uploaded_file:
        try:
            if uploaded_file.name.endswith('.csv'):
                new_data = pd.read_csv(uploaded_file)
            else:
                new_data = pd.read_excel(uploaded_file)
            
            required_cols = ["日期", "廠別", "機台編號", "OEE", "產量", "耗電量"]
            if all(col in new_data.columns for col in required_cols):
                # 確保日期格式一致
                if '日期' in new_data.columns:
                     new_data['日期'] = pd.to_datetime(new_data['日期']).dt.date
                st.session_state.data = pd.concat([st.session_state.data, new_data], ignore_index=True)
                st.sidebar.success(f"成功匯入 {len(new_data)} 筆數據！")
                st.rerun()
            else:
                st.sidebar.error(f"欄位錯誤，需包含: {required_cols}")
        except Exception as e:
            st.sidebar.error(f"讀取錯誤: {e}")

else:
    # 回歸經典：表單輸入模式
    with st.sidebar.form("add_data_form"):
        st.write("📝 新增單筆紀錄")
        col1, col2 = st.columns(2)
        in_date = col1.date_input("日期")
        in_factory = col2.text_input("廠別", "A廠")
        in_machine = st.text_input("機台編號", "ACO-X")
        
        in_oee = st.number_input("OEE (0.0 - 1.0)", 0.0, 1.0, 0.85, 0.01)
        in_output = st.number_input("產量 (雙)", 1, 10000, 1000)
        in_power = st.number_input("耗電量 (kWh)", 0.0, 10000.0, 150.0)
        
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

    st.dataframe(
        df_analysis.sort_values("效率排名").style
        .applymap(highlight_oee, subset=['OEE'])
        .format({"OEE": "{:.2%}", "單位能耗": "{:.4f}", "能源損失(元)": "${:,.0f}"}),
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
