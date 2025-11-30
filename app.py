import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import time
import numpy as np
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 1. 頁面設定 (Ver 12.0 原始設定) ---
st.set_page_config(page_title="生產效能診斷報告", layout="centered")

st.markdown("""
    <style>
    .main { background-color: #ffffff; }
    
    html, body, [class*="css"] {
        font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif;
        color: #000000;
    }
    
    h1 { color: #000000; font-weight: 900; font-size: 2.6em; text-align: center; margin-bottom: 20px; border-bottom: 4px solid #2c3e50; padding-bottom: 20px; }
    h2 { color: #1a5276; border-left: 8px solid #1a5276; padding-left: 15px; margin-top: 50px; font-size: 2em; font-weight: bold; background-color: #f2f3f4; padding-top: 5px; padding-bottom: 5px;}
    h3 { color: #2e4053; margin-top: 30px; font-size: 1.5em; font-weight: 700; }
    
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

    /* 分析觀點框 */
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
    
    /* 結論總結框 */
    .summary-box {
        border: 2px solid #333;
        padding: 20px;
        border-radius: 5px;
        background-color: #fafafa;
        margin-bottom: 20px;
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

# --- Word 匯出功能 (背景執行，不影響介面) ---
def create_word_doc(df, summary_agg, figures_map, texts, analysis_scope):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Microsoft JhengHei'
    style.font.size = Pt(12)
    
    # 標題
    head = doc.add_heading('生產效能診斷分析報告', 0)
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"分析範圍：{analysis_scope}")
    doc.add_paragraph(f"數據期間：{df['日期'].min()} 至 {df['日期'].max()}")
    doc.add_paragraph("-" * 50)

    # 1. 總覽
    doc.add_heading('1. 總體績效概覽', level=1)
    doc.add_paragraph(texts['summary'])
    
    # 插入彙整表格
    doc.add_heading('績效總表', level=2)
    table = doc.add_table(rows=1, cols=len(summary_agg.columns))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(summary_agg.columns): hdr_cells[i].text = str(col_name)
    for index, row in summary_agg.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = f"{val:.2f}" if isinstance(val, float) else str(val)

    # 安全插入圖片 (防止報錯)
    def add_fig_safe(key, title):
        doc.add_heading(title, level=2)
        if key in figures_map:
            try:
                img_bytes = figures_map[key].to_image(format="png", width=800, height=400, scale=1.5)
                doc.add_picture(BytesIO(img_bytes), width=Inches(6))
            except:
                doc.add_paragraph("[註：此圖表無法在當前環境生成]")

    add_fig_safe('rank', '綜合實力排名')
    doc.add_paragraph(texts['rank_insight'])

    # 2. 趨勢
    doc.add_heading('2. 生產趨勢與穩定性', level=1)
    add_fig_safe('cv', '生產穩定度 (CV)')
    doc.add_paragraph(texts.get('cv_insight', ''))
    add_fig_safe('corr', '效率 vs 能耗')
    doc.add_paragraph(texts.get('corr_insight', ''))

    # 3. 能耗
    doc.add_heading('3. 電力耗能分析', level=1)
    add_fig_safe('pie', '總耗電量佔比')
    add_fig_safe('unit', '平均單位能耗')
    doc.add_paragraph(texts['unit_insight'])

    # 4. 結論
    doc.add_heading('4. 綜合診斷結論', level=1)
    doc.add_paragraph(texts['conclusion'])
    doc.add_heading('策略行動建議', level=2)
    doc.add_paragraph(texts['actions'])

    bio = BytesIO()
    doc.save(bio)
    return bio

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
    with st.spinner('正在進行深度數據洞察...'):
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
            
            # 取得分析日期區間
            start_date = df["日期"].min()
            end_date = df["日期"].max()
            
            # --- 判斷單廠還是多廠 ---
            if "廠別" not in df.columns: df["廠別"] = "匯入廠區"
            
            is_multi_factory = df["廠別"].nunique() > 1
            if is_multi_factory:
                group_col = "廠別"
                summary_title = "各廠區生產績效總表"
                analysis_scope = "跨廠區分析"
            else:
                group_col = "機台編號"
                summary_title = "各機台生產績效總表"
                analysis_scope = "單廠設備分析"

            # 聚合運算
            summary_agg = df.groupby(group_col).agg({
                "OEE": "mean", "產量": "sum", "耗電量": "sum", 
                "能源損失": "sum", "總損失": "sum"
            }).reset_index()
            summary_agg["平均單位能耗"] = summary_agg["耗電量"] / summary_agg["產量"]
            summary_agg = summary_agg.sort_values("OEE", ascending=False)

            # 收集數據給 Word
            figures_map = {}
            texts_map = {}

            # --- 報告開始 (Ver 12.0 介面) ---
            st.markdown("---")
            st.title("生產效能診斷分析報告")
            st.markdown(f"**分析範圍：** {analysis_scope} &nbsp;&nbsp; **數據期間：** {start_date} 至 {end_date} &nbsp;&nbsp; **生成日期：** {pd.Timestamp.now().strftime('%Y-%m-%d')}")
            
            # ==========================================
            # 1. 總體績效
            # ==========================================
            st.header("1. 總體績效概覽")
            
            avg_oee_total = df["OEE"].mean()
            total_loss = df["總損失"].sum()
            texts_map['summary'] = f"整體平均 OEE: {avg_oee_total:.1%}, 總潛在損失: NT$ {total_loss:,.0f}"
            
            c_kpi1, c_kpi2, c_kpi3 = st.columns(3)
            c_kpi1.metric("整體平均 OEE", f"{avg_oee_total:.1%}")
            c_kpi2.metric("總潛在損失 (NTD)", f"${total_loss:,.0f}")
            c_kpi3.metric("總產量 (雙)", f"{df['產量'].sum():,.0f}")
            
            st.write("")
            st.subheader(f"📊 {summary_title}")
            
            display_cols = [group_col, "OEE", "產量", "耗電量", "平均單位能耗", "總損失"]
            final_table = summary_agg[display_cols].rename(columns={
                "OEE": "平均OEE", "產量": "總產量", "耗電量": "總耗電", "總損失": "潛在損失($)"
            })
            
            table_height = (len(final_table) + 1) * 35 + 5
            
            st.dataframe(
                final_table.style.format({
                    "平均OEE": "{:.1%}", "平均單位能耗": "{:.5f}", "潛在損失($)": "${:,.0f}", "總產量": "{:,.0f}", "總耗電": "{:,.1f}"
                }).background_gradient(subset=["平均OEE"], cmap="Blues"),
                use_container_width=True,
                height=table_height
            )

            # 排行榜
            st.subheader(f"{group_col} 綜合實力排名")
            max_oee = summary_agg["OEE"].max()
            fig_rank = px.bar(
                summary_agg.sort_values("OEE", ascending=True), 
                x="OEE", y=group_col, orientation='h',
                text="OEE", 
                title=f"依平均 OEE 排序"
            )
            fig_rank.update_traces(marker_color='#1f618d', texttemplate='%{text:.1%}', textposition='outside', textfont=dict(size=14, color='black'))
            fig_rank.update_layout(
                plot_bgcolor='white', 
                xaxis=dict(showgrid=True, gridcolor='#eee', range=[0, max_oee * 1.25]),
                height=400, font=dict(size=14, color='black')
            )
            st.plotly_chart(fig_rank, use_container_width=True)
            figures_map['rank'] = fig_rank
            
            top_p = summary_agg.iloc[0][group_col]
            last_p = summary_agg.iloc[-1][group_col]
            texts_map['rank_insight'] = f"{top_p} 表現最佳，{last_p} 效率最低建議優先檢查。"

            # ==========================================
            # 2. 趨勢與穩定性
            # ==========================================
            st.header("2. 生產趨勢與穩定性分析")
            
            st.subheader("生產穩定度 (CV變異係數)")
            if len(df) > 1:
                cv_data = df.groupby(group_col)["OEE"].agg(['mean', 'std'])
                cv_data['CV(%)'] = (cv_data['std'] / cv_data['mean']) * 100
                cv_data = cv_data.fillna(0).reset_index().sort_values('CV(%)')
                max_cv = cv_data['CV(%)'].max()

                fig_cv = px.bar(cv_data, x=group_col, y="CV(%)", text="CV(%)", title="OEE 波動率 (數值越低代表生產越穩定)")
                fig_cv.update_traces(marker_color='#922b21', texttemplate='%{text:.1f}%', textposition='outside', textfont=dict(size=14, color='black'))
                fig_cv.update_layout(
                    plot_bgcolor='white', 
                    yaxis=dict(showgrid=True, gridcolor='#eee', range=[0, max_cv * 1.2]),
                    height=400, font=dict(size=14, color='black')
                )
                st.plotly_chart(fig_cv, use_container_width=True)
                figures_map['cv'] = fig_cv
                texts_map['cv_insight'] = "CV 值越低代表生產節奏越穩定。若過高建議檢查進料或人員操作。"
                
                st.markdown("""
                <div class="analysis-text">
                <b>📈 分析觀點：</b><br>
                CV 值越低代表該設備的生產節奏越穩定，品質控制能力越好。若 CV 值過高 (>15%)，建議優先檢查該設備的進料狀況或操作人員是否頻繁更換。
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("數據量不足，無法分析波動率。")

            st.subheader("效率 vs 能耗 關聯分析")
            try:
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
                figures_map['corr'] = fig_corr
                texts_map['corr_insight'] = "理想落點為右下角。左上角異常點代表可能處於空轉浪費狀態。"
                
                st.markdown("""
                <div class="analysis-text">
                <b>📈 分析觀點：</b><br>
                此圖表用於檢視「高效率是否伴隨低能耗」。理想落點為<b>右下角</b>。若出現位於<b>左上角</b>的異常點（低效率、高耗能），通常代表設備處於「空轉浪費」狀態，應查核當日日誌。
                </div>
                """, unsafe_allow_html=True)
            except:
                fig_corr = px.scatter(df, x="OEE", y="單位能耗", color=group_col, size="產量")
                st.plotly_chart(fig_corr, use_container_width=True)

            # ==========================================
            # 3. 電力耗能
            # ==========================================
            st.header("3. 電力耗能深度分析")

            col_p1, col_p2 = st.columns(2)
            with col_p1:
                st.subheader("總耗電量佔比")
                fig_pie = px.pie(summary_agg, values="耗電量", names=group_col, hole=0.4)
                fig_pie.update_traces(textinfo='percent+label', textfont=dict(size=14, color='black'), marker=dict(colors=px.colors.qualitative.Safe))
                st.plotly_chart(fig_pie, use_container_width=True)
                figures_map['pie'] = fig_pie

            with col_p2:
                st.subheader("平均單位能耗 (kWh/雙)")
                max_unit = summary_agg["平均單位能耗"].max()
                fig_unit = px.bar(
                    summary_agg.sort_values("平均單位能耗"), 
                    x=group_col, y="平均單位能耗", 
                    text="平均單位能耗",
                    title="生產每雙產品之平均耗電 (越低越好)"
                )
                fig_unit.update_traces(marker_color='#145a32', texttemplate='%{text:.4f}', textposition='outside', textfont=dict(size=14, color='black'))
                fig_unit.update_layout(
                    plot_bgcolor='white', 
                    yaxis=dict(range=[0, max_unit * 1.2]),
                    height=400, font=dict(size=14, color='black')
                )
                st.plotly_chart(fig_unit, use_container_width=True)
                figures_map['unit'] = fig_unit
                texts_map['unit_insight'] = f"{summary_agg.sort_values('平均單位能耗').iloc[0][group_col]} 能源效率最高。"
            
            st.markdown("""
            <div class="analysis-text">
            <b>📈 分析觀點：</b><br>
            單位能耗反映了設備的能源轉換效率。數值過高的設備，可能存在馬達老化、傳動阻力過大或保溫失效等硬體問題，建議列入年度歲修重點。
            </div>
            """, unsafe_allow_html=True)

            # ==========================================
            # 4. 綜合診斷結論
            # ==========================================
            st.header("4. 綜合診斷結論 (Executive Conclusion)")

            # --- A. 分類運算 ---
            excellent_machines = []
            average_machines = []
            critical_machines = []
            
            for index, row in summary_agg.iterrows():
                name = row[group_col]
                oee = row['OEE']
                loss = row['總損失']
                info = f"**{name}** (OEE: {oee:.1%}, 損失: ${loss:,.0f})"
                
                if oee >= target_oee/100:
                    excellent_machines.append(info)
                elif oee >= 0.70:
                    average_machines.append(info)
                else:
                    critical_machines.append(info)
            
            # --- B. 診斷內容生成 ---
            st.markdown("### 📌 現況總結")
            status_summary = f"本次分析區間內 ({start_date} 至 {end_date})，全廠平均 OEE 為 **{avg_oee_total:.1%}**。"
            if avg_oee_total < 0.7:
                status_summary += " 整體生產效率偏低，存在顯著改善空間，主要虧損來源於產能未達標造成的機會成本。"
            elif avg_oee_total >= target_oee/100:
                status_summary += " 整體生產效率優異，已達世界級水準。"
            else:
                status_summary += " 生產效率維持在一般水準，部分設備表現優異，但仍有落後設備拉低平均。"
            
            texts_map['conclusion'] = f"{status_summary}\n累計潛在財務損失總額：NT$ {total_loss:,.0f}。"
            
            st.markdown(f"""
            <div class="summary-box">
            {status_summary}
            <br><br>
            累計潛在財務損失總額： <b>NT$ {total_loss:,.0f}</b>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("### 🚦 分級診斷與矩陣表")
            
            # 準備矩陣表格資料
            matrix_data = []
            for m in summary_agg.to_dict('records'):
                oee = m['OEE']
                if oee >= target_oee/100:
                    grade = "🟢 優良"
                elif oee >= 0.70:
                    grade = "🟡 尚可"
                else:
                    grade = "🔴 異常"
                matrix_data.append({
                    "設備名稱": m[group_col],
                    "平均 OEE": f"{m['OEE']:.1%}",
                    "評級": grade,
                    "財務損失佔比": f"{(m['總損失']/total_loss):.1%}" if total_loss > 0 else "0%"
                })
            
            st.dataframe(pd.DataFrame(matrix_data), use_container_width=True, hide_index=True)

            st.markdown("### 🚀 策略行動建議")

            # 針對異常設備的建議
            action_text = ""
            if critical_machines:
                names = ", ".join([m.split(' ')[0].replace('*','') for m in critical_machines])
                text = f"**1. 優先改善對象 (Priority Action):**\n* 目標設備：{names}\n* 問題診斷：OEE 低於 70%。\n* 行動方案：建議工程部門立即調閱這些設備的「異常停機代碼」。\n\n"
                st.markdown(text)
                action_text += text
            
            # 針對普通設備的建議
            if average_machines:
                names = ", ".join([m.split(' ')[0].replace('*','') for m in average_machines])
                text = f"**2. 效能提升計畫 (Improvement Plan):**\n* 目標設備：{names}\n* 行動方案：表現平穩但未達標竿。建議對照優良設備的參數設定 (Parameter)，進行參數優化微調。\n\n"
                st.markdown(text)
                action_text += text

            # 針對優良設備的建議
            if excellent_machines:
                names = ", ".join([m.split(' ')[0].replace('*','') for m in excellent_machines])
                text = f"**3. 標竿管理 (Benchmark):**\n* 目標設備：{names}\n* 行動方案：運作狀況極佳。建議將其操作標準書 (SOP) 與保養模式標準化。\n"
                st.markdown(text)
                action_text += text
            
            texts_map['actions'] = action_text

            # --- Word 下載按鈕 (靜靜地放在最後) ---
            st.markdown("---")
            
            # 生成 Word
            doc_file = create_word_doc(df, summary_agg, figures_map, texts_map, analysis_scope)
            
            st.download_button(
                label="📥 下載 Word 報表 (.docx)",
                data=doc_file.getvalue(),
                file_name=f"生產效能報告_{pd.Timestamp.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
