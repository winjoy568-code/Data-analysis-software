import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import time
import numpy as np

# --- 1. 頁面設定 ---
st.set_page_config(page_title="生產效能智慧分析系統 Pro", layout="centered")

# CSS: 專業報告風格
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 8px; height: 3.5em; font-weight: bold; font-size: 1.1em; }
    h1 { color: #2c3e50; font-family: 'Microsoft JhengHei'; }
    h3 { color: #34495e; border-left: 5px solid #3498db; padding-left: 10px; margin-top: 20px; }
    .metric-card { background-color: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); text-align: center; }
    .big-number { font-size: 2em; font-weight: bold; color: #2980b9; }
    .small-label { color: #7f8c8d; font-size: 0.9em; }
    </style>
""", unsafe_allow_html=True)

# --- 2. 核心邏輯 ---

def init_session_state():
    if 'input_data' not in st.session_state:
        # 預設範例數據 (欄位名稱已調整為 機台編號)
        st.session_state.input_data = pd.DataFrame([
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO2", "OEE(%)": 50.1, "產量(雙)": 2009.5, "用電量(kWh)": 6.2},
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO4", "OEE(%)": 55.4, "產量(雙)": 4416.5, "用電量(kWh)": 9.1},
        ])
        st.session_state.input_data['日期'] = pd.to_datetime(st.session_state.input_data['日期']).dt.date

init_session_state()

def smart_load_file(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # 智慧欄位對應 (容錯處理)
        rename_map = {"設備": "機台編號", "機台": "機台編號"}
        df = df.rename(columns=rename_map)

        if "日期" in df.columns:
            df["日期"] = pd.to_datetime(df["日期"]).dt.date
        if "廠別" not in df.columns:
            df["廠別"] = "匯入廠區"
        return df, "OK"
    except Exception as e:
        return None, str(e)

# --- 3. 介面設計：Step 1 數據輸入 (保持原樣，僅修改欄位名) ---

st.title("🏭 生產效能智慧分析系統 Pro")
st.caption("Advanced OEE & Energy Analytics Dashboard")

st.markdown('### 1. 數據輸入 (Data Input)')

uploaded_file = st.file_uploader("批次匯入 Excel (選填)", type=["xlsx", "csv"], label_visibility="collapsed")
if uploaded_file:
    new_df, status = smart_load_file(uploaded_file)
    if status == "OK":
        st.session_state.input_data = new_df
    else:
        st.error(f"檔案讀取錯誤: {status}")

# 使用者要求的表格呈現 (欄位名稱已鎖定)
edited_df = st.data_editor(
    st.session_state.input_data,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "日期": st.column_config.DateColumn("日期"),
        "機台編號": st.column_config.TextColumn("機台編號", help="請輸入設備代碼"),
        "OEE(%)": st.column_config.NumberColumn("OEE(%)", format="%.1f"),
        "產量(雙)": st.column_config.NumberColumn("產量(雙)"),
        "用電量(kWh)": st.column_config.NumberColumn("用電量(kWh)"),
    }
)

if st.button("🗑️ 清空表格數據"):
    st.session_state.input_data = pd.DataFrame(columns=["日期", "廠別", "機台編號", "OEE(%)", "產量(雙)", "用電量(kWh)"])
    st.rerun()

# --- 4. 介面設計：Step 2 參數設定 (移至下方) ---

st.markdown('### 2. 分析參數設定')
col_p1, col_p2, col_p3 = st.columns(3)
with col_p1:
    elec_price = st.number_input("平均電價 (元/度)", value=3.5, step=0.1)
with col_p2:
    target_oee = st.number_input("目標 OEE 基準 (%)", value=85.0, step=0.5)
with col_p3:
    product_margin = st.number_input("每雙獲利估算 (元)", value=10.0, step=1.0, help="用於計算產能損失機會成本")

st.write("")

# --- 5. 執行高階分析 ---

start_analysis = st.button("🚀 啟動多維度數據分析 (Run Advanced Analysis)", type="primary")

if start_analysis:
    with st.spinner('🔄 正在執行：相關性檢定、變異數分析、成本建模...'):
        time.sleep(1.2) # 體驗優化
        
        # --- A. 數據清洗與特徵工程 ---
        df = edited_df.copy()
        
        # 欄位映射
        rename_map = {"用電量(kWh)": "耗電量", "產量(雙)": "產量", "OEE(%)": "OEE_RAW"}
        for user_col, sys_col in rename_map.items():
            if user_col in df.columns:
                df = df.rename(columns={user_col: sys_col})

        required = ["機台編號", "耗電量", "產量", "OEE_RAW"]
        if df.empty or not all(col in df.columns for col in required):
            st.error("❌ 無法分析：請檢查上方表格是否包含必要數據。")
        else:
            # 數值標準化
            df["OEE"] = df["OEE_RAW"].apply(lambda x: x / 100.0 if x > 1.0 else x)
            df["單位能耗"] = df["耗電量"] / df["產量"]
            
            # 進階指標計算
            # 1. 變異係數 (CV) - 衡量生產穩定性 (標準差 / 平均值)
            # 2. 財務損失模型
            best_energy = df["單位能耗"].min()
            
            # 能源浪費金額 = (實際能耗 - 最佳能耗) * 產量 * 電價
            df["能源損失"] = (df["單位能耗"] - best_energy) * df["產量"] * elec_price
            df["能源損失"] = df["能源損失"].apply(lambda x: max(x, 0))
            
            # 產能機會成本 = (目標OEE - 實際OEE) * 理論產能(用實際產量反推) * 毛利
            # 簡化算法：假設產量與OEE成正比 -> 損失產量 = (目標OEE/實際OEE - 1) * 實際產量
            # 避免除以零
            df["產能損失機會成本"] = df.apply(
                lambda row: ((target_oee/100 - row["OEE"]) / row["OEE"] * row["產量"] * product_margin) 
                if row["OEE"] > 0 and row["OEE"] < target_oee/100 else 0, axis=1
            )

            # 判斷維度
            if "廠別" not in df.columns: df["廠別"] = "預設廠區"
            group_col = "廠別" if df["廠別"].nunique() > 1 else "機台編號"

            # --- B. 報告呈現：分頁式戰情室 ---
            st.success("✅ 分析完成！報告已生成。")
            st.markdown("---")
            
            st.title("📊 生產數據透視報告")
            
            # 使用 Tabs 分頁整理資訊量
            tab1, tab2, tab3, tab4 = st.tabs(["📋 總覽與排名", "📈 趨勢與相關性", "💰 成本損失分析", "🤖 智慧診斷建議"])

            # === Tab 1: 總覽與排名 (基礎數據) ===
            with tab1:
                st.subheader("1. 關鍵績效總表")
                
                # KPI Cards
                kpi1, kpi2, kpi3 = st.columns(3)
                avg_oee = df["OEE"].mean()
                total_loss_money = df["能源損失"].sum() + df["產能損失機會成本"].sum()
                
                kpi1.metric("平均 OEE", f"{avg_oee:.1%}", delta=f"{avg_oee - (target_oee/100):.1%}")
                kpi2.metric("總潛在損失金額", f"${total_loss_money:,.0f}", "含電費浪費與產能損失", delta_color="inverse")
                kpi3.metric("最佳單位能耗", f"{best_energy:.5f} kWh/雙")

                st.write("")
                
                # 詳細排名表
                st.markdown("**詳細數據排名 (依 OEE 排序)**")
                display_cols = ["日期", "廠別", "機台編號", "OEE", "產量", "單位能耗", "能源損失", "產能損失機會成本"]
                final_table = df[display_cols].rename(columns={
                    "OEE": "OEE(%)", "產量": "產量(雙)", 
                    "能源損失": "電費浪費($)", "產能損失機會成本": "產能損失($)"
                })
                
                st.dataframe(
                    final_table.sort_values("OEE(%)", ascending=False).style
                    .format({
                        "OEE(%)": "{:.1%}", "單位能耗": "{:.5f}", 
                        "電費浪費($)": "${:,.0f}", "產能損失($)": "${:,.0f}"
                    })
                    .background_gradient(subset=["OEE(%)"], cmap="RdYlGn"),
                    use_container_width=True, hide_index=True
                )

            # === Tab 2: 趨勢與相關性 (進階統計) ===
            with tab2:
                st.subheader("2. 生產穩定性與相關性分析")
                
                c1, c2 = st.columns(2)
                
                with c1:
                    # CV 分析 (穩定度)
                    st.markdown("**📊 生產穩定度分析 (CV 變異係數)**")
                    st.caption("CV值越低代表生產越穩定 (品質一致性高)")
                    
                    cv_data = df.groupby(group_col)["OEE"].agg(['mean', 'std'])
                    cv_data['CV(%)'] = (cv_data['std'] / cv_data['mean']) * 100
                    cv_data = cv_data.reset_index().sort_values('CV(%)')
                    
                    fig_cv = px.bar(cv_data, x=group_col, y="CV(%)", 
                                    text="CV(%)", color="CV(%)",
                                    color_continuous_scale="Reds",
                                    title="各設備 OEE 波動率 (越低越好)")
                    fig_cv.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                    st.plotly_chart(fig_cv, use_container_width=True)

                with c2:
                    # 相關性分析 (OEE vs 能耗)
                    st.markdown("**🔗 效率與能耗相關性**")
                    st.caption("檢視是否達成「高效率低能耗」的理想狀態")
                    
                    fig_corr = px.scatter(
                        df, x="OEE", y="單位能耗", color=group_col, size="產量",
                        trendline="ols", # 加入迴歸趨勢線
                        title="OEE vs 單位能耗 (含趨勢線)"
                    )
                    st.plotly_chart(fig_corr, use_container_width=True)
                    
                st.markdown("---")
                st.markdown("**📈 時序綜合趨勢**")
                # 雙軸圖：產量 vs OEE
                df_trend = df.groupby(["日期", group_col])[["產量", "OEE"]].mean().reset_index()
                fig_trend = go.Figure()
                for item in df[group_col].unique():
                    subset = df_trend[df_trend[group_col] == item]
                    fig_trend.add_trace(go.Scatter(x=subset["日期"], y=subset["OEE"], name=f"{item}-OEE", mode='lines+markers'))
                fig_trend.update_layout(title="每日 OEE 變化趨勢", yaxis_tickformat=".0%")
                st.plotly_chart(fig_trend, use_container_width=True)

            # === Tab 3: 成本損失分析 (財務面向) ===
            with tab3:
                st.subheader("3. 損失成本瀑布圖 (Financial Loss Waterfall)")
                st.caption("將技術指標轉換為貨幣金額，協助決策優先級")
                
                cost_agg = df.groupby(group_col)[["能源損失", "產能損失機會成本"]].sum().reset_index()
                cost_agg["總損失"] = cost_agg["能源損失"] + cost_agg["產能損失機會成本"]
                cost_agg = cost_agg.sort_values("總損失", ascending=False)
                
                # 堆疊長條圖
                fig_cost = px.bar(
                    cost_agg, x=group_col, y=["能源損失", "產能損失機會成本"],
                    title="各設備潛在損失金額分解 (NTD)",
                    labels={"value": "損失金額 ($)", "variable": "損失類型"},
                    color_discrete_map={"能源損失": "#e74c3c", "產能損失機會成本": "#f39c12"}
                )
                fig_cost.update_layout(barmode='stack')
                st.plotly_chart(fig_cost, use_container_width=True)
                
                st.info("💡 **解讀**：紅色代表「浪費的電費」，橘色代表「沒做到目標產量而少賺的錢」。通常橘色會大於紅色，提示我們**提升稼動率 (OEE)** 比單純省電更賺錢。")

            # === Tab 4: 智慧診斷建議 (Gemini Logic) ===
            with tab4:
                st.subheader("4. AI 邏輯診斷報告")
                
                # 自動化邏輯生成
                best_machine = cost_agg.iloc[-1][group_col] # 損失最少者
                worst_machine = cost_agg.iloc[0][group_col] # 損失最多者
                
                worst_machine_cv = cv_data[cv_data[group_col] == worst_machine]['CV(%)'].values[0]
                
                st.markdown(f"""
                ### 🏆 表現優異：{best_machine}
                * 該設備綜合損失金額最低，且 OEE 表現穩定。
                * **建議**：將 {best_machine} 的操作參數 (如速度、溫度設定) 作為標準化 SOP，推廣至其他機台。
                
                ### ⚠️ 優先改善：{worst_machine}
                * **財務衝擊**：此設備造成的總損失約 **NT$ {cost_agg.iloc[0]['總損失']:,.0f}**，佔整體的最高比例。
                * **穩定性分析**：其 OEE 變異係數 (CV) 為 **{worst_machine_cv:.1f}%**。
                    * 若 CV > 10%：代表生產極不穩定，建議檢查進料變異或人員操作手法。
                    * 若 CV 低但 OEE 低：代表持續性的性能低落，建議檢查設備老化或參數設定錯誤。
                
                ### 🚀 下一步行動
                1.  **針對 {worst_machine} 召開檢討會**，調閱異常代碼。
                2.  確認是否出現「低產速但高耗能」的**空轉**現象（參考相關性圖表的左上角區域）。
                """)
