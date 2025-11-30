import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import time
import numpy as np

# --- 1. 頁面設定 ---
st.set_page_config(page_title="生產效能智慧分析系統 Pro", layout="centered")

# CSS 優化：專業報告風格
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 8px; height: 3.5em; font-weight: bold; font-size: 1.1em; }
    h1 { color: #2c3e50; font-family: 'Microsoft JhengHei'; }
    h3 { color: #34495e; border-left: 5px solid #3498db; padding-left: 10px; margin-top: 20px; }
    </style>
""", unsafe_allow_html=True)

# --- 2. 核心邏輯 ---

def init_session_state():
    if 'input_data' not in st.session_state:
        # 預設範例 (內部統一使用 '機台編號')
        st.session_state.input_data = pd.DataFrame([
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO2", "OEE(%)": 50.1, "產量(雙)": 2009.5, "用電量(kWh)": 6.2},
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO4", "OEE(%)": 55.4, "產量(雙)": 4416.5, "用電量(kWh)": 9.1},
        ])
        # 確保日期格式
        st.session_state.input_data['日期'] = pd.to_datetime(st.session_state.input_data['日期']).dt.date

init_session_state()

def smart_load_file(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # 智慧欄位對應 (讓使用者的 Excel 標題 '設備' 也能通)
        rename_map = {"設備": "機台編號", "機台": "機台編號"}
        df = df.rename(columns=rename_map)

        if "日期" in df.columns:
            df["日期"] = pd.to_datetime(df["日期"]).dt.date
        if "廠別" not in df.columns:
            df["廠別"] = "匯入廠區"
        return df, "OK"
    except Exception as e:
        return None, str(e)

# --- 3. 數據輸入介面 (UI Step 1) ---

st.title("🏭 生產效能智慧分析系統 Pro")
st.caption("Advanced OEE & Energy Analytics Dashboard")

st.markdown('### 1. 數據輸入 (Data Input)')

# 上傳區塊
uploaded_file = st.file_uploader("批次匯入 Excel (選填)", type=["xlsx", "csv"], label_visibility="collapsed")
if uploaded_file:
    new_df, status = smart_load_file(uploaded_file)
    if status == "OK":
        st.session_state.input_data = new_df
    else:
        st.error(f"檔案讀取錯誤: {status}")

# 編輯表格
edited_df = st.data_editor(
    st.session_state.input_data,
    num_rows="dynamic", # 允許新增刪除
    use_container_width=True,
    column_config={
        "日期": st.column_config.DateColumn("日期"),
        # 【修正點】：移除了多餘的參數，只保留 label 和 help
        "機台編號": st.column_config.TextColumn(label="設備/機台編號", help="請輸入設備代碼"),
        "OEE(%)": st.column_config.NumberColumn("OEE(%)", format="%.1f"),
        "產量(雙)": st.column_config.NumberColumn("產量(雙)"),
        "用電量(kWh)": st.column_config.NumberColumn("用電量(kWh)"),
    }
)

if st.button("🗑️ 清空表格數據"):
    st.session_state.input_data = pd.DataFrame(columns=["日期", "廠別", "機台編號", "OEE(%)", "產量(雙)", "用電量(kWh)"])
    st.rerun()

# --- 4. 參數設定 (UI Step 2) ---

st.markdown('### 2. 分析參數設定')
col_p1, col_p2, col_p3 = st.columns(3)
with col_p1:
    elec_price = st.number_input("平均電價 (元/度)", value=3.5, step=0.1)
with col_p2:
    target_oee = st.number_input("目標 OEE 基準 (%)", value=85.0, step=0.5)
with col_p3:
    product_margin = st.number_input("每雙獲利估算 (元)", value=10.0, step=1.0)

st.write("")

# --- 5. 執行分析邏輯 (Execution) ---

start_analysis = st.button("🚀 啟動多維度數據分析 (Run Advanced Analysis)", type="primary")

if start_analysis:
    # 顯示載入動畫
    with st.spinner('🔄 正在執行：相關性檢定、變異數分析、成本建模...'):
        time.sleep(1.0) # 模擬運算體驗
        
        # 1. 複製並鎖定數據
        df = edited_df.copy()
        
        # 2. 關鍵修正：確保所有別名都轉回系統標準名稱
        rename_map = {
            "用電量(kWh)": "耗電量", 
            "產量(雙)": "產量", 
            "OEE(%)": "OEE_RAW",
            "設備": "機台編號", # 把 '設備' 轉回 '機台編號'
            "機台": "機台編號"
        }
        for user_col, sys_col in rename_map.items():
            if user_col in df.columns:
                df = df.rename(columns={user_col: sys_col})

        required = ["機台編號", "耗電量", "產量", "OEE_RAW"]
        
        # 3. 欄位檢查 (防呆)
        if df.empty or not all(col in df.columns for col in required):
            missing = [c for c in required if c not in df.columns]
            st.error(f"❌ 無法分析：缺少必要欄位。缺少的欄位: {missing}")
            st.info("💡 請確認上方的表格標題是否包含：日期, 廠別, 設備(或機台編號), OEE(%), 產量(雙), 用電量(kWh)")
        else:
            # 4. 數據運算
            # OEE 轉小數
            df["OEE"] = df["OEE_RAW"].apply(lambda x: x / 100.0 if x > 1.0 else x)
            # 單位能耗
            df["單位能耗"] = df["耗電量"] / df["產量"]
            
            # 成本模型
            best_energy = df["單位能耗"].min()
            df["能源損失"] = (df["單位能耗"] - best_energy) * df["產量"] * elec_price
            df["能源損失"] = df["能源損失"].apply(lambda x: max(x, 0))
            
            # 產能損失 (機會成本)
            df["產能損失機會成本"] = df.apply(
                lambda row: ((target_oee/100 - row["OEE"]) / row["OEE"] * row["產量"] * product_margin) 
                if row["OEE"] > 0 and row["OEE"] < target_oee/100 else 0, axis=1
            )

            # 判斷維度
            if "廠別" not in df.columns: df["廠別"] = "預設廠區"
            group_col = "廠別" if df["廠別"].nunique() > 1 else "機台編號"

            # --- 報告生成區 ---
            st.success("✅ 分析完成！")
            st.markdown("---")
            st.title("📊 生產數據透視報告")
            
            tab1, tab2, tab3, tab4 = st.tabs(["📋 總覽與排名", "📈 趨勢與相關性", "💰 成本損失分析", "🤖 智慧診斷建議"])

            # === Tab 1: 總覽 ===
            with tab1:
                st.subheader("1. 關鍵績效總表")
                kpi1, kpi2, kpi3 = st.columns(3)
                avg_oee = df["OEE"].mean()
                total_loss_money = df["能源損失"].sum() + df["產能損失機會成本"].sum()
                
                kpi1.metric("平均 OEE", f"{avg_oee:.1%}", delta=f"{avg_oee - (target_oee/100):.1%}")
                kpi2.metric("總潛在損失金額", f"${total_loss_money:,.0f}", "含電費浪費與產能損失", delta_color="inverse")
                kpi3.metric("最佳單位能耗", f"{best_energy:.5f} kWh/雙")
                
                st.write("")
                st.markdown("**詳細數據排名 (依 OEE 排序)**")
                
                # 準備顯示表格
                display_cols = ["日期", "廠別", "機台編號", "OEE", "產量", "單位能耗", "能源損失", "產能損失機會成本"]
                final_table = df[display_cols].rename(columns={
                    "OEE": "OEE(%)", "產量": "產量(雙)", 
                    "能源損失": "電費浪費($)", "產能損失機會成本": "產能損失($)"
                })
                
                # 顏色漸層顯示 (需要 jinja2 和 matplotlib)
                try:
                    st.dataframe(
                        final_table.sort_values("OEE(%)", ascending=False).style
                        .format({
                            "OEE(%)": "{:.1%}", "單位能耗": "{:.5f}", 
                            "電費浪費($)": "${:,.0f}", "產能損失($)": "${:,.0f}"
                        })
                        .background_gradient(subset=["OEE(%)"], cmap="RdYlGn"),
                        use_container_width=True, hide_index=True
                    )
                except Exception as e:
                    st.warning("⚠️ 表格顏色渲染失敗 (可能是缺少 jinja2)，顯示為標準表格。")
                    st.dataframe(final_table, use_container_width=True)

            # === Tab 2: 趨勢與相關性 ===
            with tab2:
                st.subheader("2. 生產穩定性與相關性")
                c1, c2 = st.columns(2)
                
                with c1:
                    # CV 圖
                    if len(df) > 1:
                        cv_data = df.groupby(group_col)["OEE"].agg(['mean', 'std'])
                        cv_data['CV(%)'] = (cv_data['std'] / cv_data['mean']) * 100
                        cv_data = cv_data.reset_index().sort_values('CV(%)')
                        fig_cv = px.bar(cv_data, x=group_col, y="CV(%)", text="CV(%)", 
                                      color="CV(%)", color_continuous_scale="Reds", 
                                      title="OEE 波動率 (CV, 越低越穩)")
                        fig_cv.update_traces(texttemplate='%{text:.1f}%')
                        st.plotly_chart(fig_cv, use_container_width=True)
                    else:
                        st.info("ℹ️ 數據量不足，無法計算波動率")

                with c2:
                    # 相關性圖 (加入防護罩)
                    try:
                        fig_corr = px.scatter(
                            df, x="OEE", y="單位能耗", 
                            color=group_col, size="產量", 
                            trendline="ols", # 這裡需要 statsmodels
                            title="OEE vs 能耗相關性 (含趨勢預測)"
                        )
                        st.plotly_chart(fig_corr, use_container_width=True)
                    except Exception as e:
                        st.caption("⚠️ 數據點過少或缺少套件，顯示為標準散佈圖 (無趨勢線)")
                        fig_corr = px.scatter(
                            df, x="OEE", y="單位能耗", 
                            color=group_col, size="產量",
                            title="OEE vs 能耗相關性"
                        )
                        st.plotly_chart(fig_corr, use_container_width=True)

            # === Tab 3: 成本 ===
            with tab3:
                st.subheader("3. 損失成本分析")
                cost_agg = df.groupby(group_col)[["能源損失", "產能損失機會成本"]].sum().reset_index()
                cost_agg["總損失"] = cost_agg["能源損失"] + cost_agg["產能損失機會成本"]
                
                fig_cost = px.bar(
                    cost_agg.sort_values("總損失", ascending=False), 
                    x=group_col, y=["能源損失", "產能損失機會成本"], 
                    title="潛在損失金額分解 (NTD)", 
                    barmode='stack',
                    color_discrete_map={"能源損失": "#e74c3c", "產能損失機會成本": "#f39c12"}
                )
                st.plotly_chart(fig_cost, use_container_width=True)

            # === Tab 4: 診斷 ===
            with tab4:
                st.subheader("4. AI 診斷報告")
                if not cost_agg.empty:
                    worst_machine = cost_agg.iloc[0][group_col]
                    loss_val = cost_agg.iloc[0]['總損失']
                    st.markdown(f"""
                    ### ⚠️ 重點關注對象：{worst_machine}
                    * **財務衝擊**：該設備在此期間造成的總潛在損失達 **NT$ {loss_val:,.0f}**。
                    * **建議行動**：
                        1. 檢查 {worst_machine} 的待機設定，避免空轉浪費電力。
                        2. 檢討該設備是否經常發生短暫停機，導致 OEE 低落進而造成產能損失。
                    """)
