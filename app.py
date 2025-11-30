import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import time
import numpy as np

# --- 1. 頁面設定 ---
# 改回 centered 模式，模擬 A4 紙張的閱讀體驗
st.set_page_config(page_title="生產效能診斷報告", layout="centered")

# CSS 優化：Word 報告風格 (白底黑字、大字體)
st.markdown("""
    <style>
    .main { background-color: #ffffff; }
    
    /* 字體設定：加大、加深，適合閱讀 */
    html, body, [class*="css"] {
        font-family: 'Microsoft JhengHei', sans-serif;
        color: #000000;
    }
    
    /* 標題設定 */
    h1 { color: #000000; font-weight: 900; font-size: 2.5em; text-align: center; margin-bottom: 30px; }
    h2 { color: #2c3e50; border-bottom: 3px solid #000000; padding-bottom: 10px; margin-top: 60px; font-size: 1.8em; }
    h3 { color: #2980b9; margin-top: 40px; font-size: 1.5em; font-weight: bold; }
    
    /* 內文設定 */
    p, li, .stMarkdown {
        font-size: 18px !important; /* 強制加大內文字體 */
        line-height: 1.8 !important;
        color: #333333 !important;
    }
    
    /* 數據指標卡片 */
    div[data-testid="stMetricValue"] {
        font-size: 36px !important;
        color: #000000 !important;
    }
    
    /* 分析結論段落 */
    .analysis-text {
        font-size: 20px;
        font-weight: 500;
        color: #2c3e50;
        margin-top: 10px;
        margin-bottom: 30px;
        border-left: 5px solid #2980b9;
        padding-left: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. 核心邏輯 ---

def init_session_state():
    if 'input_data' not in st.session_state:
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
        
        rename_map = {"設備": "機台編號", "機台": "機台編號"}
        df = df.rename(columns=rename_map)

        if "日期" in df.columns:
            df["日期"] = pd.to_datetime(df["日期"]).dt.date
        if "廠別" not in df.columns:
            df["廠別"] = "匯入廠區"
        return df, "OK"
    except Exception as e:
        return None, str(e)

# --- 3. 數據輸入介面 (保持原本功能) ---

st.markdown("### 📥 數據輸入控制台")
st.caption("請在此處輸入數據，完成後點擊下方按鈕生成正式報告。")

uploaded_file = st.file_uploader("批次匯入 Excel (選填)", type=["xlsx", "csv"], label_visibility="collapsed")
if uploaded_file:
    new_df, status = smart_load_file(uploaded_file)
    if status == "OK":
        st.session_state.input_data = new_df
    else:
        st.error(f"檔案讀取錯誤: {status}")

edited_df = st.data_editor(
    st.session_state.input_data,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "日期": st.column_config.DateColumn("日期"),
        "機台編號": st.column_config.TextColumn(label="機台編號"),
        "OEE(%)": st.column_config.NumberColumn("OEE(%)", format="%.1f"),
        "產量(雙)": st.column_config.NumberColumn("產量(雙)"),
        "用電量(kWh)": st.column_config.NumberColumn("用電量(kWh)"),
    }
)

if st.button("🗑️ 清空表格數據"):
    st.session_state.input_data = pd.DataFrame(columns=["日期", "廠別", "機台編號", "OEE(%)", "產量(雙)", "用電量(kWh)"])
    st.rerun()

st.markdown("---")
st.markdown("#### ⚙️ 分析參數設定")
c1, c2, c3 = st.columns(3)
with c1:
    elec_price = st.number_input("平均電價 (元/度)", value=3.5, step=0.1)
with c2:
    target_oee = st.number_input("目標 OEE (%)", value=85.0, step=0.5)
with c3:
    product_margin = st.number_input("每雙獲利估算 (元)", value=10.0, step=1.0)

st.write("")
start_analysis = st.button("📄 生成正式分析報告", type="primary")

# --- 4. 報告生成區 (Word 導向) ---

if start_analysis:
    with st.spinner('正在撰寫分析報告...'):
        time.sleep(1.0)
        
        # --- 資料處理 ---
        df = edited_df.copy()
        rename_map = {
            "用電量(kWh)": "耗電量", "產量(雙)": "產量", 
            "OEE(%)": "OEE_RAW", "設備": "機台編號", "機台": "機台編號"
        }
        for user_col, sys_col in rename_map.items():
            if user_col in df.columns:
                df = df.rename(columns={user_col: sys_col})

        required = ["機台編號", "耗電量", "產量", "OEE_RAW"]
        if df.empty or not all(col in df.columns for col in required):
            st.error("資料不足，無法生成報告。")
        else:
            # 計算
            df["OEE"] = df["OEE_RAW"].apply(lambda x: x / 100.0 if x > 1.0 else x)
            df["單位能耗"] = df["耗電量"] / df["產量"]
            best_energy = df["單位能耗"].min()
            df["能源損失"] = (df["單位能耗"] - best_energy) * df["產量"] * elec_price
            df["能源損失"] = df["能源損失"].apply(lambda x: max(x, 0))
            df["產能損失機會成本"] = df.apply(
                lambda row: ((target_oee/100 - row["OEE"]) / row["OEE"] * row["產量"] * product_margin) 
                if row["OEE"] > 0 and row["OEE"] < target_oee/100 else 0, axis=1
            )
            df["總損失"] = df["能源損失"] + df["產能損失機會成本"]
            
            # 聚合
            machine_agg = df.groupby("機台編號").agg({
                "OEE": "mean", "產量": "sum", "耗電量": "sum", 
                "能源損失": "sum", "總損失": "sum"
            }).reset_index()
            machine_agg["平均單位能耗"] = machine_agg["耗電量"] / machine_agg["產量"]
            
            # --- 報告開始 ---
            st.markdown("---")
            st.title("生產效能診斷分析報告")
            st.markdown(f"**報告日期：** {pd.Timestamp.now().strftime('%Y-%m-%d')}")
            
            # ==========================================
            # 第一部分：總體績效概覽
            # ==========================================
            st.header("1. 總體績效概覽 (Executive Summary)")
            
            # KPI (單行排列)
            avg_oee_total = df["OEE"].mean()
            total_loss = df["總損失"].sum()
            
            c_kpi1, c_kpi2, c_kpi3 = st.columns(3)
            c_kpi1.metric("全廠平均 OEE", f"{avg_oee_total:.1%}")
            c_kpi2.metric("總潛在損失 (NTD)", f"${total_loss:,.0f}")
            c_kpi3.metric("總產量 (雙)", f"{df['產量'].sum():,.0f}")
            
            st.write("")
            st.subheader("原始數據明細表")
            display_cols = ["日期", "機台編號", "OEE", "產量", "耗電量", "單位能耗"]
            final_table = df[display_cols].rename(columns={"OEE": "OEE(%)", "產量": "產量(雙)", "耗電量": "用電量(kWh)"})
            st.dataframe(final_table.style.format({"OEE(%)": "{:.1%}", "單位能耗": "{:.5f}"}), use_container_width=True)

            # 機台排行榜 (圖表 1)
            st.subheader("機台綜合實力排名")
            fig_rank = px.bar(
                machine_agg.sort_values("OEE", ascending=True), 
                x="OEE", y="機台編號", orientation='h',
                text="OEE", color="OEE", color_continuous_scale="Blues"
            )
            fig_rank.update_traces(texttemplate='%{text:.1%}', textposition='outside')
            fig_rank.update_layout(height=400, font=dict(size=14))
            st.plotly_chart(fig_rank, use_container_width=True)
            
            # 分析解讀 1
            top_machine = machine_agg.sort_values("OEE", ascending=False).iloc[0]['機台編號']
            last_machine = machine_agg.sort_values("OEE", ascending=False).iloc[-1]['機台編號']
            
            st.markdown(f"""
            <div class="analysis-text">
            <b>📊 排行榜分析：</b><br>
            根據本次分析區間數據，<b>{top_machine}</b> 的平均 OEE 最高，為目前的標竿機台。
            相對而言，<b>{last_machine}</b> 的效率表現敬陪末座，是目前拉低整體產能的主要瓶頸，建議優先列為改善對象。
            </div>
            """, unsafe_allow_html=True)

            # ==========================================
            # 第二部分：趨勢與穩定性分析
            # ==========================================
            st.header("2. 生產趨勢與穩定性分析")
            
            # CV 分析 (圖表 2)
            st.subheader("機台生產穩定度 (CV變異係數)")
            if len(df) > 1:
                cv_data = df.groupby("機台編號")["OEE"].agg(['mean', 'std'])
                cv_data['CV(%)'] = (cv_data['std'] / cv_data['mean']) * 100
                cv_data = cv_data.reset_index().sort_values('CV(%)')
                
                fig_cv = px.bar(cv_data, x="機台編號", y="CV(%)", text="CV(%)", 
                                color="CV(%)", color_continuous_scale="Reds")
                fig_cv.update_traces(texttemplate='%{text:.1f}%')
                fig_cv.update_layout(height=400, font=dict(size=14), title_text="數值越低代表生產越穩定")
                st.plotly_chart(fig_cv, use_container_width=True)
                
                # 分析解讀 2
                most_stable = cv_data.iloc[0]['機台編號']
                most_unstable = cv_data.iloc[-1]['機台編號']
                
                st.markdown(f"""
                <div class="analysis-text">
                <b>📊 穩定度分析：</b><br>
                <b>{most_stable}</b> 的 CV 值最低，顯示其每日生產表現最為一致，製程控制能力佳。
                <b>{most_unstable}</b> 的 CV 值最高，代表該設備容易出現「忽高忽低」的生產狀況，可能原因包含：頻繁換線、人員操作不標準或進料品質不穩。
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("數據量不足，無法分析波動率。")

            # 相關性分析 (圖表 3)
            st.subheader("OEE 與 單位能耗 關聯性")
            try:
                fig_corr = px.scatter(
                    df, x="OEE", y="單位能耗", 
                    color="機台編號", size="產量", 
                    trendline="ols"
                )
                fig_corr.update_layout(height=500, font=dict(size=14))
                st.plotly_chart(fig_corr, use_container_width=True)
            except:
                st.info("數據點不足以繪製趨勢線。")
            
            st.markdown(f"""
            <div class="analysis-text">
            <b>📊 關聯性分析：</b><br>
            圖表顯示了「效率」與「耗電」的關係。位於圖表<b>左上方</b>的點位代表「低效率、高耗能」，這是明顯的能源浪費訊號（通常源於設備空轉或待機時間過長）。
            建議檢查落於左上角區域的機台紀錄，確認當日是否有異常停機未關閉電源之情事。
            </div>
            """, unsafe_allow_html=True)

            # ==========================================
            # 第三部分：電力耗能深度分析
            # ==========================================
            st.header("3. 電力耗能深度分析")

            # 總耗電佔比 (圖表 4)
            st.subheader("各機台總耗電量分佈")
            fig_pie = px.pie(machine_agg, values="耗電量", names="機台編號", hole=0.4)
            fig_pie.update_traces(textinfo='percent+label')
            fig_pie.update_layout(font=dict(size=14))
            st.plotly_chart(fig_pie, use_container_width=True)
            
            st.markdown(f"""
            <div class="analysis-text">
            <b>📊 總用電分析：</b><br>
            上圖呈現了各機台的用電總量佔比。佔比最高的機台若是主力生產設備則屬正常；若非主力設備卻佔比過高，則需檢查是否有漏電或設備老化造成的高負載問題。
            </div>
            """, unsafe_allow_html=True)

            # 單位能耗 (圖表 5)
            st.subheader("平均單位能耗比較 (kWh/雙)")
            fig_unit = px.bar(
                machine_agg.sort_values("平均單位能耗"), 
                x="機台編號", y="平均單位能耗", 
                text="平均單位能耗", color="平均單位能耗", color_continuous_scale="Viridis_r"
            )
            fig_unit.update_traces(texttemplate='%{text:.4f}')
            fig_unit.update_layout(height=400, font=dict(size=14))
            st.plotly_chart(fig_unit, use_container_width=True)
            
            # 分析解讀 3
            best_p = machine_agg.sort_values("平均單位能耗").iloc[0]['機台編號']
            worst_p = machine_agg.sort_values("平均單位能耗").iloc[-1]['機台編號']
            
            st.markdown(f"""
            <div class="analysis-text">
            <b>📊 能耗效率分析：</b><br>
            <b>{best_p}</b> 是目前的節能冠軍，每生產一雙鞋僅消耗最少的電力。
            <b>{worst_p}</b> 的單位生產成本最高，建議工程單位檢查其馬達效率、傳動系統阻力，或加熱系統的保溫效果。
            </div>
            """, unsafe_allow_html=True)

            # ==========================================
            # 第四部分：結論與行動建議
            # ==========================================
            st.header("4. 結論與行動建議 (Conclusion)")
            st.markdown("針對全廠設備之綜合診斷結果：")

            for index, row in machine_agg.iterrows():
                m_name = row['機台編號']
                m_oee = row['OEE']
                m_loss = row['總損失']
                
                if m_oee >= target_oee/100:
                    status = "✅ 優良"
                    action = "維持現狀，將其參數設定作為標準 SOP 推廣至全廠。"
                elif m_oee >= 0.70:
                    status = "⚠️ 普通"
                    action = "需針對短暫停機進行分析，目標提升稼動率 5% 以上。"
                else:
                    status = "❌ 異常"
                    action = "為主要虧損來源，建議立即停機檢修，並審視排程與人員操作。"

                st.markdown(f"""
                ### 🔧 機台：{m_name}
                * **狀態評估**：{status} (平均 OEE: {m_oee:.1%})
                * **財務衝擊**：此期間累計潛在損失 **NT$ {m_loss:,.0f}**。
                * **行動建議**：{action}
                """)
                st.markdown("---")
