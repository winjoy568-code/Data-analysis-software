import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import time
import numpy as np

# --- 1. 頁面設定 ---
st.set_page_config(page_title="生產效能智慧分析系統 Pro", layout="centered")

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
        # 預設範例 (機台編號)
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
        
        # 讀取時的欄位容錯
        rename_map = {"設備": "機台編號", "機台": "機台編號"}
        df = df.rename(columns=rename_map)

        if "日期" in df.columns:
            df["日期"] = pd.to_datetime(df["日期"]).dt.date
        if "廠別" not in df.columns:
            df["廠別"] = "匯入廠區"
        return df, "OK"
    except Exception as e:
        return None, str(e)

# --- 3. 數據輸入介面 ---

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

# 這裡顯示的標題會優先使用 DataFrame 裡的，如果舊資料是「設備」，這裡就會顯示「設備」
edited_df = st.data_editor(
    st.session_state.input_data,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "日期": st.column_config.DateColumn("日期"),
        # 這裡設定機台編號，但如果資料是「設備」，Data Editor 會自動顯示「設備」
        "機台編號": st.column_config.TextColumn("機台編號", help="請輸入設備代碼"),
        "OEE(%)": st.column_config.NumberColumn("OEE(%)", format="%.1f"),
        "產量(雙)": st.column_config.NumberColumn("產量(雙)"),
        "用電量(kWh)": st.column_config.NumberColumn("用電量(kWh)"),
    }
)

if st.button("🗑️ 清空表格數據"):
    st.session_state.input_data = pd.DataFrame(columns=["日期", "廠別", "機台編號", "OEE(%)", "產量(雙)", "用電量(kWh)"])
    st.rerun()

# --- 4. 參數設定 ---

st.markdown('### 2. 分析參數設定')
col_p1, col_p2, col_p3 = st.columns(3)
with col_p1:
    elec_price = st.number_input("平均電價 (元/度)", value=3.5, step=0.1)
with col_p2:
    target_oee = st.number_input("目標 OEE 基準 (%)", value=85.0, step=0.5)
with col_p3:
    product_margin = st.number_input("每雙獲利估算 (元)", value=10.0, step=1.0)

st.write("")

# --- 5. 執行分析 (修正 Bug 的地方) ---

start_analysis = st.button("🚀 啟動多維度數據分析 (Run Advanced Analysis)", type="primary")

if start_analysis:
    with st.spinner('🔄 正在執行：相關性檢定、變異數分析、成本建模...'):
        time.sleep(1.2)
        
        df = edited_df.copy()
        
        # 【關鍵修正】：這裡多加了 "設備": "機台編號" 的對應
        # 這樣就算表格標題是「設備」，程式也會自動轉成「機台編號」再去算，就不會報錯了
        rename_map = {
            "用電量(kWh)": "耗電量", 
            "產量(雙)": "產量", 
            "OEE(%)": "OEE_RAW",
            "設備": "機台編號" 
        }
        for user_col, sys_col in rename_map.items():
            if user_col in df.columns:
                df = df.rename(columns={user_col: sys_col})

        required = ["機台編號", "耗電量", "產量", "OEE_RAW"]
        
        # 檢查欄位
        if df.empty or not all(col in df.columns for col in required):
            missing = [c for c in required if c not in df.columns]
            st.error(f"❌ 無法分析：缺少必要欄位。系統偵測到的欄位: {list(df.columns)}，缺少的欄位: {missing}")
            st.info("💡 建議點擊上方「🗑️ 清空表格數據」按鈕重置格式。")
        else:
            # --- 正常分析流程 ---
            df["OEE"] = df["OEE_RAW"].apply(lambda x: x / 100.0 if x > 1.0 else x)
            df["單位能耗"] = df["耗電量"] / df["產量"]
            
            best_energy = df["單位能耗"].min()
            df["能源損失"] = (df["單位能耗"] - best_energy) * df["產量"] * elec_price
            df["能源損失"] = df["能源損失"].apply(lambda x: max(x, 0))
            
            df["產能損失機會成本"] = df.apply(
                lambda row: ((target_oee/100 - row["OEE"]) / row["OEE"] * row["產量"] * product_margin) 
                if row["OEE"] > 0 and row["OEE"] < target_oee/100 else 0, axis=1
            )

            if "廠別" not in df.columns: df["廠別"] = "預設廠區"
            group_col = "廠別" if df["廠別"].nunique() > 1 else "機台編號"

            st.success("✅ 分析完成！")
            st.markdown("---")
            st.title("📊 生產數據透視報告")
            
            tab1, tab2, tab3, tab4 = st.tabs(["📋 總覽與排名", "📈 趨勢與相關性", "💰 成本損失分析", "🤖 智慧診斷建議"])

            # Tab 1
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
                display_cols = ["日期", "廠別", "機台編號", "OEE", "產量", "單位能耗", "能源損失", "產能損失機會成本"]
                final_table = df[display_cols].rename(columns={"OEE": "OEE(%)", "產量": "產量(雙)", "能源損失": "電費浪費($)", "產能損失機會成本": "產能損失($)"})
                st.dataframe(final_table.sort_values("OEE(%)", ascending=False).style.format({"OEE(%)": "{:.1%}", "單位能耗": "{:.5f}", "電費浪費($)": "${:,.0f}", "產能損失($)": "${:,.0f}"}).background_gradient(subset=["OEE(%)"], cmap="RdYlGn"), use_container_width=True, hide_index=True)

            # Tab 2
            with tab2:
                st.subheader("2. 生產穩定性與相關性")
                c1, c2 = st.columns(2)
                with c1:
                    cv_data = df.groupby(group_col)["OEE"].agg(['mean', 'std'])
                    cv_data['CV(%)'] = (cv_data['std'] / cv_data['mean']) * 100
                    cv_data = cv_data.reset_index().sort_values('CV(%)')
                    fig_cv = px.bar(cv_data, x=group_col, y="CV(%)", text="CV(%)", color="CV(%)", color_continuous_scale="Reds", title="OEE 波動率 (CV, 越低越穩)")
                    fig_cv.update_traces(texttemplate='%{text:.1f}%')
                    st.plotly_chart(fig_cv, use_container_width=True)
                with c2:
                    fig_corr = px.scatter(df, x="OEE", y="單位能耗", color=group_col, size="產量", trendline="ols", title="OEE vs 能耗相關性")
                    st.plotly_chart(fig_corr, use_container_width=True)

            # Tab 3
            with tab3:
                st.subheader("3. 損失成本分析")
                cost_agg = df.groupby(group_col)[["能源損失", "產能損失機會成本"]].sum().reset_index()
                cost_agg["總損失"] = cost_agg["能源損失"] + cost_agg["產能損失機會成本"]
                fig_cost = px.bar(cost_agg.sort_values("總損失", ascending=False), x=group_col, y=["能源損失", "產能損失機會成本"], title="潛在損失金額分解 (NTD)", barmode='stack')
                st.plotly_chart(fig_cost, use_container_width=True)

            # Tab 4
            with tab4:
                st.subheader("4. AI 診斷報告")
                if not cost_agg.empty:
                    worst_machine = cost_agg.iloc[0][group_col]
                    st.markdown(f"### ⚠️ 重點關注：{worst_machine}")
                    st.markdown(f"該設備總損失達 **NT$ {cost_agg.iloc[0]['總損失']:,.0f}**，建議優先檢查參數設定與停機原因。")
