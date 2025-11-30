import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import time
import numpy as np

# --- 1. 頁面設定 ---
st.set_page_config(page_title="生產效能診斷報告", layout="centered")

# CSS 優化：Word 報告風格 (高對比、大字體、無捲軸)
st.markdown("""
    <style>
    .main { background-color: #ffffff; }
    
    /* 全域字體設定 */
    html, body, [class*="css"] {
        font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif;
        color: #000000;
    }
    
    /* 標題設定 */
    h1 { color: #000000; font-weight: 900; font-size: 2.6em; text-align: center; margin-bottom: 20px; border-bottom: 4px solid #2c3e50; padding-bottom: 20px; }
    h2 { color: #1a5276; border-left: 8px solid #1a5276; padding-left: 15px; margin-top: 50px; font-size: 2em; font-weight: bold; background-color: #f2f3f4; padding-top: 5px; padding-bottom: 5px;}
    h3 { color: #2e4053; margin-top: 30px; font-size: 1.5em; font-weight: 700; }
    
    /* 內文設定 */
    p, li, .stMarkdown {
        font-size: 18px !important;
        line-height: 1.6 !important;
        color: #212f3d !important;
    }
    
    /* 數據指標卡片 */
    div[data-testid="stMetricValue"] {
        font-size: 32px !important;
        color: #17202a !important;
        font-weight: bold;
    }
    
    /* 分析結論段落框 */
    .analysis-text {
        font-size: 18px;
        font-weight: 500;
        color: #2c3e50;
        margin-top: 15px;
        margin-bottom: 30px;
        border: 2px solid #5d6d7e;
        background-color: #ebf5fb;
        padding: 20px;
        border-radius: 8px;
    }
    
    /* 隱藏表格索引行以節省空間 */
    thead tr th:first-child {display:none}
    tbody th {display:none}
    </style>
""", unsafe_allow_html=True)

# --- 2. 核心邏輯 ---

def init_session_state():
    if 'input_data' not in st.session_state:
        st.session_state.input_data = pd.DataFrame([
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO2", "OEE(%)": 50.1, "產量(雙)": 2009.5, "用電量(kWh)": 6.2},
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO4", "OEE(%)": 55.4, "產量(雙)": 4416.5, "用電量(kWh)": 9.1},
            {"日期": "2025-11-18", "廠別": "A廠", "機台編號": "ACO2", "OEE(%)": 48.5, "產量(雙)": 1950.0, "用電量(kWh)": 6.0},
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

# --- 3. 數據輸入介面 ---

st.markdown("### 📥 數據輸入控制台")
uploaded_file = st.file_uploader("批次匯入 Excel", type=["xlsx", "csv"], label_visibility="collapsed")
if uploaded_file:
    new_df, status = smart_load_file(uploaded_file)
    if status == "OK":
        st.session_state.input_data = new_df
    else:
        st.error(f"錯誤: {status}")

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

if st.button("🗑️ 清空表格"):
    st.session_state.input_data = pd.DataFrame(columns=["日期", "廠別", "機台編號", "OEE(%)", "產量(雙)", "用電量(kWh)"])
    st.rerun()

st.markdown("---")
st.markdown("#### ⚙️ 分析參數")
c1, c2, c3 = st.columns(3)
with c1:
    elec_price = st.number_input("電價 (元/度)", value=3.5, step=0.1)
with c2:
    target_oee = st.number_input("目標 OEE (%)", value=85.0, step=0.5)
with c3:
    product_margin = st.number_input("獲利估算 (元/雙)", value=10.0, step=1.0)

st.write("")
start_analysis = st.button("📄 生成正式分析報告", type="primary")

# --- 4. 報告生成區 ---

if start_analysis:
    with st.spinner('正在彙整數據並生成圖表...'):
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
            # 計算指標
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
            
            # --- 判斷單廠還是多廠，決定彙整邏輯 ---
            if "廠別" not in df.columns: df["廠別"] = "匯入廠區"
            
            is_multi_factory = df["廠別"].nunique() > 1
            if is_multi_factory:
                # 多廠模式：以「廠別」為群組
                group_col = "廠別"
                summary_title = "各廠區生產績效總表"
                analysis_scope = "跨廠區分析"
            else:
                # 單廠模式：以「機台編號」為群組
                group_col = "機台編號"
                summary_title = "各機台生產績效總表"
                analysis_scope = "單廠設備分析"

            # 聚合運算 (Summary Table)
            summary_agg = df.groupby(group_col).agg({
                "OEE": "mean", "產量": "sum", "耗電量": "sum", 
                "能源損失": "sum", "總損失": "sum"
            }).reset_index()
            summary_agg["平均單位能耗"] = summary_agg["耗電量"] / summary_agg["產量"]
            summary_agg = summary_agg.sort_values("OEE", ascending=False) # 依效率排名

            # --- 報告開始 ---
            st.markdown("---")
            st.title("生產效能診斷分析報告")
            st.markdown(f"**分析範圍：** {analysis_scope} &nbsp;&nbsp;&nbsp; **報告日期：** {pd.Timestamp.now().strftime('%Y-%m-%d')}")
            
            # ==========================================
            # 第一部分：總體績效概覽
            # ==========================================
            st.header("1. 總體績效概覽 (Executive Summary)")
            
            # KPI
            avg_oee_total = df["OEE"].mean()
            total_loss = df["總損失"].sum()
            
            c_kpi1, c_kpi2, c_kpi3 = st.columns(3)
            c_kpi1.metric("整體平均 OEE", f"{avg_oee_total:.1%}")
            c_kpi2.metric("總潛在損失 (NTD)", f"${total_loss:,.0f}")
            c_kpi3.metric("總產量 (雙)", f"{df['產量'].sum():,.0f}")
            
            st.write("")
            st.subheader(f"📊 {summary_title}")
            
            # 準備顯示表格 (去除不必要的欄位，只留彙整數據)
            display_cols = [group_col, "OEE", "產量", "耗電量", "平均單位能耗", "總損失"]
            final_table = summary_agg[display_cols].rename(columns={
                "OEE": "平均OEE", "產量": "總產量", "耗電量": "總耗電", "總損失": "潛在損失($)"
            })
            
            # 計算表格高度以取消捲軸: (行數 + 表頭) * 行高
            table_height = (len(final_table) + 1) * 35 + 5
            
            st.dataframe(
                final_table.style.format({
                    "平均OEE": "{:.1%}", "平均單位能耗": "{:.5f}", "潛在損失($)": "${:,.0f}", "總產量": "{:,.0f}", "總耗電": "{:,.1f}"
                }).background_gradient(subset=["平均OEE"], cmap="Blues"),
                use_container_width=True,
                height=table_height # 自動展開所有高度
            )

            # 排行榜 (高對比色)
            st.subheader(f"{group_col} 綜合實力排名")
            
            # 設定顏色：使用深藍色單色，避免淺色看不清
            fig_rank = px.bar(
                summary_agg.sort_values("OEE", ascending=True), 
                x="OEE", y=group_col, orientation='h',
                text="OEE", 
                title=f"依平均 OEE 排序 (數值越高越好)"
            )
            # 強制設定高對比顏色
            fig_rank.update_traces(marker_color='#1f618d', texttemplate='%{text:.1%}', textposition='outside', textfont=dict(size=14, color='black'))
            fig_rank.update_layout(
                plot_bgcolor='white', 
                xaxis=dict(showgrid=True, gridcolor='#eee'),
                height=400, font=dict(size=14, color='black')
            )
            st.plotly_chart(fig_rank, use_container_width=True)
            
            # 分析解讀
            top_performer = summary_agg.iloc[0][group_col]
            last_performer = summary_agg.iloc[-1][group_col]
            
            st.markdown(f"""
            <div class="analysis-text">
            <b>📈 數據解讀：</b><br>
            根據數據彙整結果，<b>{top_performer}</b> 在本次分析區間內的綜合效率 (OEE) 表現最佳，為績效標竿。<br>
            <b>{last_performer}</b> 的平均效率最低，建議優先檢查該單位的異常停機狀況或作業流程。
            </div>
            """, unsafe_allow_html=True)

            # ==========================================
            # 第二部分：趨勢與穩定性分析
            # ==========================================
            st.header("2. 生產趨勢與穩定性分析")
            
            # CV 分析 (如果有多筆資料才做)
            st.subheader("生產穩定度 (CV變異係數)")
            if len(df) > 1:
                # 計算每個群組的 CV
                cv_data = df.groupby(group_col)["OEE"].agg(['mean', 'std'])
                cv_data['CV(%)'] = (cv_data['std'] / cv_data['mean']) * 100
                cv_data = cv_data.fillna(0).reset_index().sort_values('CV(%)')
                
                fig_cv = px.bar(cv_data, x=group_col, y="CV(%)", text="CV(%)", title="OEE 波動率 (數值越低代表生產越穩定)")
                # 使用深紅色強調
                fig_cv.update_traces(marker_color='#922b21', texttemplate='%{text:.1f}%', textposition='outside', textfont=dict(size=14, color='black'))
                fig_cv.update_layout(plot_bgcolor='white', yaxis=dict(showgrid=True, gridcolor='#eee'), height=400, font=dict(size=14, color='black'))
                st.plotly_chart(fig_cv, use_container_width=True)
                
                most_stable = cv_data.iloc[0][group_col]
                most_unstable = cv_data.iloc[-1][group_col]
                
                st.markdown(f"""
                <div class="analysis-text">
                <b>📈 數據解讀：</b><br>
                <b>{most_stable}</b> 的 CV 值最低，顯示其每日生產表現最為穩定。<br>
                <b>{most_unstable}</b> 的 CV 值最高，代表生產過程容易忽快忽慢，品質與產出較難預測，建議進行標準化作業輔導。
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("數據量不足，無法分析波動率。")

            # 相關性分析
            st.subheader("效率 vs 能耗 關聯分析")
            try:
                # 使用深色點位
                fig_corr = px.scatter(
                    df, x="OEE", y="單位能耗", 
                    color=group_col, size="產量", 
                    trendline="ols",
                    title="X軸:效率(越高越好) / Y軸:能耗(越低越好)"
                )
                fig_corr.update_layout(
                    plot_bgcolor='white', 
                    xaxis=dict(showgrid=True, gridcolor='#eee'),
                    yaxis=dict(showgrid=True, gridcolor='#eee'),
                    height=500, font=dict(size=14, color='black')
                )
                st.plotly_chart(fig_corr, use_container_width=True)
            except:
                fig_corr = px.scatter(df, x="OEE", y="單位能耗", color=group_col, size="產量")
                st.plotly_chart(fig_corr, use_container_width=True)
            
            st.markdown(f"""
            <div class="analysis-text">
            <b>📈 數據解讀：</b><br>
            圖表呈現了生產效率與電力消耗的關係。理想狀態應位於<b>右下角</b>（高 OEE、低單位能耗）。
            若發現有數據點落於<b>左上角</b>（低 OEE、高單位能耗），代表該時段設備可能處於「空轉」或「低速運轉但全功率耗電」的異常狀態。
            </div>
            """, unsafe_allow_html=True)

            # ==========================================
            # 第三部分：電力耗能深度分析
            # ==========================================
            st.header("3. 電力耗能深度分析")

            col_p1, col_p2 = st.columns(2)

            with col_p1:
                st.subheader("總耗電量佔比")
                # 使用簡單配色
                fig_pie = px.pie(summary_agg, values="耗電量", names=group_col, hole=0.4)
                fig_pie.update_traces(textinfo='percent+label', textfont=dict(size=14, color='black'), marker=dict(colors=px.colors.qualitative.Safe))
                st.plotly_chart(fig_pie, use_container_width=True)

            with col_p2:
                st.subheader("平均單位能耗 (kWh/雙)")
                fig_unit = px.bar(
                    summary_agg.sort_values("平均單位能耗"), 
                    x=group_col, y="平均單位能耗", 
                    text="平均單位能耗",
                    title="生產每雙產品之平均耗電 (越低越好)"
                )
                # 使用深綠色表示節能
                fig_unit.update_traces(marker_color='#145a32', texttemplate='%{text:.4f}', textposition='outside', textfont=dict(size=14, color='black'))
                fig_unit.update_layout(plot_bgcolor='white', height=400, font=dict(size=14, color='black'))
                st.plotly_chart(fig_unit, use_container_width=True)
            
            best_p = summary_agg.sort_values("平均單位能耗").iloc[0][group_col]
            worst_p = summary_agg.sort_values("平均單位能耗").iloc[-1][group_col]
            
            st.markdown(f"""
            <div class="analysis-text">
            <b>📈 數據解讀：</b><br>
            <b>{best_p}</b> 的能源轉換效率最高，每生產一單位的產品耗電量最少。<br>
            <b>{worst_p}</b> 的單位能耗最高，建議工程部門檢查其馬達效率、傳動系統阻力，或加熱系統的保溫效果是否老化。
            </div>
            """, unsafe_allow_html=True)

            # ==========================================
            # 第四部分：結論與行動建議
            # ==========================================
            st.header("4. 綜合診斷結論 (Conclusion)")
            st.markdown(f"針對 {analysis_scope} 之綜合診斷結果：")

            for index, row in summary_agg.iterrows():
                target_name = row[group_col]
                m_oee = row['OEE']
                m_loss = row['總損失']
                
                if m_oee >= target_oee/100:
                    status = "✅ 優良"
                    action = "維持現狀，將其運作模式標準化，並作為其他單位的學習標竿。"
                    color = "#2ecc71" # Green
                elif m_oee >= 0.70:
                    status = "⚠️ 尚可"
                    action = "需針對短暫停機進行分析，目標提升稼動率 5% 以上。"
                    color = "#f1c40f" # Yellow
                else:
                    status = "❌ 異常"
                    action = "為主要虧損來源，建議立即檢修設備，並審視排程規劃與人員操作手法。"
                    color = "#e74c3c" # Red

                st.markdown(f"""
                ### 🔧 {group_col}：{target_name}
                * **狀態評估**：<span style='color:{color}; font-weight:bold'>{status}</span> (平均 OEE: {m_oee:.1%})
                * **財務衝擊**：此期間累計潛在損失 **NT$ {m_loss:,.0f}**。
                * **行動建議**：{action}
                """)
                st.markdown("---")
