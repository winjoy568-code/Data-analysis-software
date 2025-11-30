import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import time
import numpy as np
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 1. 頁面設定 (Ver 12.0 風格) ---
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
    
    /* 結論區塊樣式 */
    .summary-box {
        border: 2px solid #333;
        padding: 20px;
        border-radius: 5px;
        background-color: #fafafa;
        margin-bottom: 20px;
    }
    
    /* 按鈕樣式調整 */
    div.stButton > button:first-child {
        width: 100%;
        height: 3em;
        font-size: 18px;
        font-weight: bold;
    }
    
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
        if uploaded_file.name.endswith('.csv'): df = pd.read_csv(uploaded_file)
        else: df = pd.read_excel(uploaded_file)
        rename_map = {"設備": "機台編號", "機台": "機台編號"}
        df = df.rename(columns=rename_map)
        if "日期" in df.columns: df["日期"] = pd.to_datetime(df["日期"]).dt.date
        if "廠別" not in df.columns: df["廠別"] = "匯入廠區"
        return df, "OK"
    except Exception as e: return None, str(e)

# --- Word 生成引擎 (穩定版：文字+表格) ---
def generate_word_report(df, summary_agg, texts, analysis_scope):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Microsoft JhengHei'
    style.font.size = Pt(12)
    
    # 標題
    head = doc.add_heading('生產效能診斷分析報告', 0)
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"分析範圍：{analysis_scope}")
    doc.add_paragraph(f"數據期間：{df['日期'].min()} 至 {df['日期'].max()}")
    doc.add_paragraph(f"生成日期：{pd.Timestamp.now().strftime('%Y-%m-%d')}")
    doc.add_paragraph("-" * 50)

    # 1. 總覽
    doc.add_heading('1. 總體績效概覽', level=1)
    doc.add_paragraph(texts.get('summary_kpi', ''))
    
    doc.add_heading('績效總表', level=2)
    table = doc.add_table(rows=1, cols=len(summary_agg.columns))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(summary_agg.columns): hdr_cells[i].text = str(col_name)
    for index, row in summary_agg.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            if isinstance(val, float): row_cells[i].text = f"{val:.2f}"
            else: row_cells[i].text = str(val)

    # 2. 結論
    doc.add_heading('2. 綜合診斷結論', level=1)
    doc.add_paragraph("現況總結：")
    doc.add_paragraph(texts.get('conclusion_summary', ''))
    
    doc.add_heading('策略行動建議', level=2)
    doc.add_paragraph(texts.get('conclusion_action', ''))
    
    doc.add_paragraph("-" * 20)
    doc.add_paragraph("註：詳細圖表請參閱線上分析系統。")

    bio = BytesIO()
    doc.save(bio)
    return bio

# --- 3. 介面 (Ver 12.0) ---

st.markdown("### 📥 數據輸入控制台")
uploaded_file = st.file_uploader("批次匯入 Excel", type=["xlsx", "csv"], label_visibility="collapsed")
if uploaded_file:
    new_df, status = smart_load_file(uploaded_file)
    if status == "OK": st.session_state.input_data = new_df
    else: st.error(f"錯誤: {status}")

edited_df = st.data_editor(
    st.session_state.input_data, num_rows="dynamic", use_container_width=True,
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
with c1: elec_price = st.number_input("電價 (元/度)", value=3.5, step=0.1)
with c2: target_oee = st.number_input("目標 OEE (%)", value=85.0, step=0.5)
with c3: product_margin = st.number_input("獲利估算 (元/雙)", value=10.0, step=1.0)

st.write("")

# --- 4. 數據運算 (提前至按鈕前，確保匯出資料就緒) ---

# 資料處理與變數準備
df = edited_df.copy()
rename_map = {"用電量(kWh)": "耗電量", "產量(雙)": "產量", "OEE(%)": "OEE_RAW", "設備": "機台編號", "機台": "機台編號"}
for user_col, sys_col in rename_map.items():
    if user_col in df.columns: df = df.rename(columns={user_col: sys_col})

data_ready = False
if not df.empty and all(col in df.columns for col in ["機台編號", "耗電量", "產量", "OEE_RAW"]):
    data_ready = True
    # 計算邏輯
    df["OEE"] = df["OEE_RAW"].apply(lambda x: x / 100.0 if x > 1.0 else x)
    df["單位能耗"] = df["耗電量"] / df["產量"]
    best_energy = df["單位能耗"].min()
    df["能源損失"] = (df["單位能耗"] - best_energy) * df["產量"] * elec_price
    df["能源損失"] = df["能源損失"].apply(lambda x: max(x, 0))
    df["產能損失機會成本"] = df.apply(lambda row: ((target_oee/100 - row["OEE"]) / row["OEE"] * row["產量"] * product_margin) if row["OEE"] > 0 and row["OEE"] < target_oee/100 else 0, axis=1)
    df["總損失"] = df["能源損失"] + df["產能損失機會成本"]
    
    # 判斷範圍
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

    summary_agg = df.groupby(group_col).agg({"OEE": "mean", "產量": "sum", "耗電量": "sum", "能源損失": "sum", "總損失": "sum"}).reset_index()
    summary_agg["平均單位能耗"] = summary_agg["耗電量"] / summary_agg["產量"]
    summary_agg = summary_agg.sort_values("OEE", ascending=False)

    # 準備結論文字 (給 Word 和 網頁共用)
    avg_oee_total = df["OEE"].mean()
    total_loss = df["總損失"].sum()
    texts = {}
    texts['summary_kpi'] = f"整體平均 OEE: {avg_oee_total:.1%}\n總潛在損失: NT$ {total_loss:,.0f}\n總產量: {df['產量'].sum():,.0f} 雙"
    
    # 結論文字邏輯
    status_summary = f"本次分析區間內，全廠平均 OEE 為 **{avg_oee_total:.1%}**。"
    if avg_oee_total < 0.7: status_summary += " 整體效率偏低，存在改善空間。"
    else: status_summary += " 整體效率表現尚可。"
    texts['conclusion_summary'] = f"{status_summary}\n累計潛在財務損失總額：NT$ {total_loss:,.0f}。"

    crit_list, avg_list, good_list = [], [], []
    for m in summary_agg.to_dict('records'):
        if m['OEE'] >= target_oee/100: good_list.append(m[group_col])
        elif m['OEE'] >= 0.7: avg_list.append(m[group_col])
        else: crit_list.append(m[group_col])
    
    action_text = ""
    if crit_list: action_text += f"【優先改善】{', '.join(crit_list)}：效率偏低，請檢查異常停機。\n"
    if avg_list: action_text += f"【效能提升】{', '.join(avg_list)}：建議微調參數，提升稼動率。\n"
    if good_list: action_text += f"【標竿管理】{', '.join(good_list)}：運作優良，建議標準化SOP。"
    texts['conclusion_action'] = action_text

# --- 5. 按鈕區 (Ver 15.0: 並排顯示) ---

col_btn1, col_btn2 = st.columns([1, 1])

with col_btn1:
    start_analysis = st.button("🚀 啟動全方位分析", type="primary")

with col_btn2:
    if data_ready:
        doc_file = generate_word_report(df, summary_agg, texts, analysis_scope)
        st.download_button(
            label="📥 下載 Word 報告",
            data=doc_file.getvalue(),
            file_name=f"生產效能報告_{pd.Timestamp.now().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.button("📥 下載 Word 報告", disabled=True, help="請先輸入數據")

# --- 6. 分析畫面顯示 (只有按下啟動才顯示) ---

if start_analysis and data_ready:
    with st.spinner('正在分析數據...'):
        time.sleep(0.5)
        st.markdown("---")
        st.title("生產效能診斷分析報告")
        
        # 1. 總體
        st.header("1. 總體績效概覽")
        c1, c2, c3 = st.columns(3)
        c1.metric("整體平均 OEE", f"{avg_oee_total:.1%}")
        c2.metric("總潛在損失 (NTD)", f"${total_loss:,.0f}")
        c3.metric("總產量 (雙)", f"{df['產量'].sum():,.0f}")
        
        st.subheader(f"📊 {summary_title}")
        display_cols = [group_col, "OEE", "產量", "耗電量", "平均單位能耗", "總損失"]
        final_table = summary_agg[display_cols].rename(columns={"OEE": "平均OEE", "產量": "總產量", "耗電量": "總耗電", "總損失": "潛在損失($)"})
        st.dataframe(final_table.style.format({"平均OEE": "{:.1%}", "平均單位能耗": "{:.5f}", "潛在損失($)": "${:,.0f}", "總產量": "{:,.0f}", "總耗電": "{:,.1f}"}).background_gradient(subset=["平均OEE"], cmap="Blues"), use_container_width=True)

        # 排行榜
        st.subheader(f"{group_col} 綜合實力排名")
        max_oee = summary_agg["OEE"].max()
        fig_rank = px.bar(summary_agg.sort_values("OEE", ascending=True), x="OEE", y=group_col, orientation='h', text="OEE", title=f"依平均 OEE 排序")
        fig_rank.update_traces(marker_color='#1f618d', texttemplate='%{text:.1%}', textposition='outside', textfont=dict(size=14, color='black'))
        fig_rank.update_layout(plot_bgcolor='white', xaxis=dict(showgrid=True, gridcolor='#eee', range=[0, max_oee * 1.25]), height=400, font=dict(size=14, color='black'))
        st.plotly_chart(fig_rank, use_container_width=True)

        # 2. 趨勢
        st.header("2. 生產趨勢與穩定性分析")
        st.subheader("生產穩定度 (CV變異係數)")
        if len(df) > 1:
            cv_data = df.groupby(group_col)["OEE"].agg(['mean', 'std'])
            cv_data['CV(%)'] = (cv_data['std'] / cv_data['mean']) * 100
            cv_data = cv_data.fillna(0).reset_index().sort_values('CV(%)')
            max_cv = cv_data['CV(%)'].max()
            fig_cv = px.bar(cv_data, x=group_col, y="CV(%)", text="CV(%)", title="OEE 波動率")
            fig_cv.update_traces(marker_color='#922b21', texttemplate='%{text:.1f}%', textposition='outside', textfont=dict(size=14, color='black'))
            fig_cv.update_layout(plot_bgcolor='white', yaxis=dict(showgrid=True, gridcolor='#eee', range=[0, max_cv * 1.2]), height=400, font=dict(size=14, color='black'))
            st.plotly_chart(fig_cv, use_container_width=True)
            st.markdown("""<div class="analysis-text"><b>📈 分析觀點：</b><br>CV 值越低代表該設備的生產節奏越穩定。</div>""", unsafe_allow_html=True)
        else:
            st.info("數據量不足，無法分析波動率。")

        st.subheader("效率 vs 能耗 關聯分析")
        try:
            fig_corr = px.scatter(df, x="OEE", y="單位能耗", color=group_col, size="產量", trendline="ols", title="效率(X) vs 能耗(Y)")
            fig_corr.update_layout(plot_bgcolor='white', xaxis=dict(showgrid=True, gridcolor='#eee'), yaxis=dict(showgrid=True, gridcolor='#eee'), height=500, font=dict(size=14, color='black'))
            st.plotly_chart(fig_corr, use_container_width=True)
            st.markdown("""<div class="analysis-text"><b>📈 分析觀點：</b><br>理想狀態為落點於右下角。若出現左上角異常點，代表設備可能處於空轉浪費狀態。</div>""", unsafe_allow_html=True)
        except:
            fig_corr = px.scatter(df, x="OEE", y="單位能耗", color=group_col, size="產量")
            st.plotly_chart(fig_corr, use_container_width=True)

        # 3. 能耗
        st.header("3. 電力耗能深度分析")
        cp1, cp2 = st.columns(2)
        with cp1:
            st.subheader("總耗電量佔比")
            fig_pie = px.pie(summary_agg, values="耗電量", names=group_col, hole=0.4)
            fig_pie.update_traces(textinfo='percent+label', textfont=dict(size=14, color='black'), marker=dict(colors=px.colors.qualitative.Safe))
            st.plotly_chart(fig_pie, use_container_width=True)
        with cp2:
            st.subheader("平均單位能耗")
            max_unit = summary_agg["平均單位能耗"].max()
            fig_unit = px.bar(summary_agg.sort_values("平均單位能耗"), x=group_col, y="平均單位能耗", text="平均單位能耗", title="平均耗電")
            fig_unit.update_traces(marker_color='#145a32', texttemplate='%{text:.4f}', textposition='outside', textfont=dict(size=14, color='black'))
            fig_unit.update_layout(plot_bgcolor='white', yaxis=dict(range=[0, max_unit * 1.2]), height=400, font=dict(size=14, color='black'))
            st.plotly_chart(fig_unit, use_container_width=True)
        st.markdown(f"""<div class="analysis-text"><b>📈 分析觀點：</b><br>單位能耗反映了設備的能源轉換效率。**{summary_agg.sort_values('平均單位能耗').iloc[0][group_col]}** 表現最佳。</div>""", unsafe_allow_html=True)

        # 4. 結論
        st.header("4. 綜合診斷結論")
        st.markdown("### 📌 現況總結")
        st.markdown(f'<div class="summary-box">{texts["conclusion_summary"]}</div>', unsafe_allow_html=True)

        st.markdown("### 🚦 分級診斷與矩陣表")
        matrix_data = []
        for m in summary_agg.to_dict('records'):
            oee = m['OEE']
            grade = "🟢 優良" if oee >= target_oee/100 else "🟡 尚可" if oee >= 0.70 else "🔴 異常"
            matrix_data.append({"設備名稱": m[group_col], "平均 OEE": f"{m['OEE']:.1%}", "評級": grade, "財務損失佔比": f"{(m['總損失']/total_loss):.1%}" if total_loss > 0 else "0%"})
        st.dataframe(pd.DataFrame(matrix_data), use_container_width=True, hide_index=True)

        st.markdown("### 🚀 策略行動建議")
        st.markdown(texts['conclusion_action'])
