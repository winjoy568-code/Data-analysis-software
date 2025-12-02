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
import re

# ==========================================
# 0. 系統設定 (System Config)
# ==========================================
st.set_page_config(page_title="生產效能智慧分析系統 Pro", layout="centered")

st.markdown("""
    <style>
    .main { background-color: #ffffff; }
    html, body, [class*="css"] { font-family: Arial, sans-serif; color: #000000; }
    
    h1 { color: #000000; font-weight: 900; font-size: 2.4em; text-align: center; border-bottom: 4px solid #2c3e50; padding-bottom: 15px; margin-bottom: 30px; }
    h2 { color: #1a5276; border-left: 7px solid #1a5276; padding-left: 12px; margin-top: 40px; font-size: 1.6em; font-weight: bold; background-color: #f8f9fa; padding-top: 5px; padding-bottom: 5px; }
    h3 { color: #2e4053; margin-top: 25px; font-size: 1.3em; font-weight: 700; }
    
    p, li, .stMarkdown { font-size: 16px !important; line-height: 1.7 !important; color: #212f3d !important; }
    div[data-testid="stMetricValue"] { font-size: 28px !important; color: #17202a !important; font-weight: bold; }
    
    .insight-box { border: 1px solid #d6eaf8; background-color: #ebf5fb; padding: 15px; border-radius: 5px; margin-top: 10px; margin-bottom: 20px; }
    .summary-box { border: 2px solid #566573; background-color: #fdfefe; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
    
    thead tr th:first-child {display:none} tbody th {display:none}
    div.stButton > button:first-child { width: 100%; height: 3em; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. Data Engine (數據處理核心) - 保持穩定
# ==========================================
class DataEngine:
    @staticmethod
    def clean_and_process(df_raw, params):
        try:
            df = df_raw.copy()
            rename_map = {"用電量(kWh)": "耗電量", "產量(雙)": "產量", "OEE(%)": "OEE_RAW", "設備": "機台編號", "機台": "機台編號"}
            for user_col, sys_col in rename_map.items():
                if user_col in df.columns: df = df.rename(columns={user_col: sys_col})
                
            required_cols = ["機台編號", "耗電量", "產量", "OEE_RAW"]
            if not all(col in df.columns for col in required_cols):
                missing = [c for c in required_cols if c not in df.columns]
                return None, None, f"資料表缺少必要欄位: {missing}"
                
            if "日期" in df.columns: df["日期"] = pd.to_datetime(df["日期"]).dt.date
            if "廠別" not in df.columns: df["廠別"] = "匯入廠區"

            df["OEE"] = df["OEE_RAW"].apply(lambda x: x / 100.0 if x > 1.0 else x)
            df["單位能耗"] = df.apply(lambda row: row["耗電量"] / row["產量"] if row["產量"] > 0 else 0, axis=1)
            
            valid_energies = df[df["單位能耗"] > 0]["單位能耗"]
            best_energy = valid_energies.min() if not valid_energies.empty else 0
            
            elec_price = params['elec_price']
            target_oee = params['target_oee'] / 100.0
            margin = params['product_margin']
            
            df["能源損失"] = df.apply(lambda row: max(0, (row["單位能耗"] - best_energy) * row["產量"] * elec_price), axis=1)
            df["產能損失機會成本"] = df.apply(
                lambda row: ((target_oee - row["OEE"]) / row["OEE"] * row["產量"] * margin) 
                if 0 < row["OEE"] < target_oee else 0, axis=1
            )
            df["總損失"] = df["能源損失"] + df["產能損失機會成本"]
            
            group_col = "廠別" if df["廠別"].nunique() > 1 else "機台編號"
            analysis_scope = "跨廠區分析" if group_col == "廠別" else "單廠設備分析"
            
            summary_agg = df.groupby(group_col).agg({
                "OEE": "mean", "產量": "sum", "耗電量": "sum", 
                "能源損失": "sum", "產能損失機會成本": "sum", "總損失": "sum"
            }).reset_index()
            
            summary_agg["平均單位能耗"] = summary_agg.apply(
                lambda row: row["耗電量"] / row["產量"] if row["產量"] > 0 else 0, axis=1
            )
            summary_agg = summary_agg.sort_values("OEE", ascending=False)
            
            return df, summary_agg, analysis_scope
        except Exception as e:
            return None, None, f"數據處理發生錯誤: {str(e)}"

# ==========================================
# 2. Insight Engine (診斷分析大腦) - 升級：倍數與潛力
# ==========================================
class InsightEngine:
    @staticmethod
    def generate_narrative(df, summary_agg, group_col, params):
        texts = {}
        target_oee = params['target_oee'] / 100.0
        margin = params['product_margin']
        
        # 1. 總體 KPI
        avg_oee = df["OEE"].mean()
        total_loss = df["總損失"].sum()
        best_machine = summary_agg.iloc[0]
        worst_machine = summary_agg.iloc[-1]
        
        texts['kpi_summary'] = f"本次分析區間內，整體平均 OEE 為 **{avg_oee:.1%}**。其中 **{best_machine[group_col]}** 表現最佳，為全廠標竿；而 **{worst_machine[group_col]}** 效率敬陪末座，是造成全廠 **NT$ {total_loss:,.0f}** 潛在損失的主要原因。"
        
        # 2. 標竿與落差分析 (加入倍數比較)
        eff_gap_pct = 0
        multiplier_msg = ""
        if best_machine['平均單位能耗'] > 0 and worst_machine['平均單位能耗'] > 0:
            eff_gap_pct = ((worst_machine['平均單位能耗'] - best_machine['平均單位能耗']) / best_machine['平均單位能耗']) * 100
            multiplier = worst_machine['平均單位能耗'] / best_machine['平均單位能耗']
            multiplier_msg = f"換算下來，生產同樣數量的產品，**{worst_machine[group_col]}** 消耗了 **{multiplier:.1f} 倍** 的電力。"
            
        texts['benchmark_analysis'] = f"""
        * **標竿設備 ({best_machine[group_col]})**：表現最佳，平均 OEE 達 **{best_machine['OEE']:.1%}**，單位能耗最低 ({best_machine['平均單位能耗']:.5f} kWh/雙)。
        * **瓶頸設備 ({worst_machine[group_col]})**：表現最弱，單位生產成本比標竿高出 **{eff_gap_pct:.1f}%**。{multiplier_msg}
        """
        
        # 3. 潛力預估 (新增：獨立段落)
        potential_prod = 0
        if worst_machine['OEE'] > 0:
            potential_prod = (best_machine['OEE'] - worst_machine['OEE']) / worst_machine['OEE'] * worst_machine['產量']
        potential_revenue = potential_prod * margin
        
        texts['opportunity_analysis'] = f"""
        若能將 **{worst_machine[group_col]}** 的效率提升至標竿水準：
        1. 預計本期間可額外生產 **{potential_prod:,.0f} 雙**。
        2. 相當於挽回 **NT$ {potential_revenue:,.0f}** 的營收損失。
        這顯示該設備具有極大的改善投資回報率 (ROI)。
        """

        # 4. 穩定性分析
        cv_text = "數據量不足以計算波動率。"
        if len(df) > 1:
            cv_series = df.groupby(group_col)["OEE"].std() / df.groupby(group_col)["OEE"].mean()
            most_stable = cv_series.idxmin()
            most_unstable = cv_series.idxmax()
            cv_text = f"**{most_stable}** 生產節奏最穩定 (CV最低)；**{most_unstable}** 波動最大，顯示製程或人員操作存在變異。"
        texts['stability_analysis'] = cv_text
        
        # 5. 策略行動建議
        crit_list, avg_list, good_list = [], [], []
        for _, row in summary_agg.iterrows():
            name = row[group_col]
            if row['OEE'] >= target_oee: good_list.append(name)
            elif row['OEE'] >= 0.7: avg_list.append(name)
            else: crit_list.append(name)
            
        action_text = ""
        if crit_list:
            action_text += f"🔴 **優先改善 (Priority)**：{', '.join(crit_list)}\n   * 問題：OEE 低於 70%，可能存在空轉浪費。\n   * 行動：立即調閱異常停機代碼，檢查是否有「待機未關機」情況。\n\n"
        if avg_list:
            action_text += f"🟡 **效能提升 (Improvement)**：{', '.join(avg_list)}\n   * 問題：表現平穩但未達標竿。\n   * 行動：微調參數 (速度/溫度)，目標提升 5-10% 稼動率。\n\n"
        if good_list:
            action_text += f"🟢 **標竿管理 (Benchmark)**：{', '.join(good_list)}\n   * 表現：運作優異。\n   * 行動：將其操作參數標準化 (SOP)，推廣至其他設備。"
            
        texts['action_plan'] = action_text
        return texts

# ==========================================
# 3. Viz Engine (視覺化中心) - 升級：補齊圖表
# ==========================================
class VizEngine:
    @staticmethod
    def _common_layout():
        return dict(
            plot_bgcolor='white',
            font=dict(family='Arial, sans-serif', color='black', size=12),
            xaxis=dict(showgrid=True, gridcolor='#f0f0f0'),
            yaxis=dict(showgrid=True, gridcolor='#f0f0f0'),
            margin=dict(l=40, r=40, t=40, b=40)
        )

    @staticmethod
    def create_rank_chart(summary_agg, group_col):
        fig = px.bar(
            summary_agg.sort_values("OEE", ascending=True),
            x="OEE", y=group_col, orientation='h', text="OEE",
            title="綜合實力排名 (依 OEE 排序)"
        )
        fig.update_traces(marker_color='#2E86C1', texttemplate='%{text:.1%}', textposition='outside')
        fig.update_layout(VizEngine._common_layout())
        fig.update_layout(xaxis=dict(range=[0, summary_agg['OEE'].max() * 1.25])) 
        return fig

    @staticmethod
    def create_cv_chart(df, group_col):
        cv_data = df.groupby(group_col)["OEE"].agg(['mean', 'std'])
        cv_data['CV'] = (cv_data['std'] / cv_data['mean']) * 100
        cv_data = cv_data.fillna(0).reset_index()
        fig = px.bar(cv_data, x=group_col, y="CV", text="CV", title="生產穩定度 (CV變異係數，越低越好)")
        fig.update_traces(marker_color='#C0392B', texttemplate='%{text:.1f}%', textposition='outside')
        fig.update_layout(VizEngine._common_layout())
        fig.update_layout(yaxis=dict(range=[0, cv_data['CV'].max() * 1.2]))
        return fig

    @staticmethod
    def create_scatter_chart(df, group_col):
        try:
            fig = px.scatter(
                df, x="OEE", y="單位能耗", color=group_col, size="產量",
                trendline="ols", title="效率 vs 能耗 關聯分析",
                color_discrete_sequence=px.colors.qualitative.Set1
            )
        except:
            fig = px.scatter(
                df, x="OEE", y="單位能耗", color=group_col, size="產量",
                title="效率 vs 能耗 關聯分析 (無趨勢線)",
                color_discrete_sequence=px.colors.qualitative.Set1
            )
        fig.update_layout(VizEngine._common_layout())
        return fig

    @staticmethod
    def create_dual_axis_chart(df, group_col):
        df_sorted = df.sort_values(["日期", group_col])
        x_label = df_sorted["日期"].astype(str) + " " + df_sorted[group_col]
        fig = go.Figure()
        fig.add_trace(go.Bar(x=x_label, y=df_sorted["產量"], name="產量", marker_color='#BDC3C7', opacity=0.7))
        fig.add_trace(go.Scatter(x=x_label, y=df_sorted["耗電量"], name="耗電量", yaxis="y2", line=dict(color='#E74C3C', width=3)))
        layout = VizEngine._common_layout()
        layout.update(dict(
            title="產量與耗電量趨勢對比",
            yaxis2=dict(title="耗電量(kWh)", overlaying="y", side="right", showgrid=False),
            xaxis=dict(tickangle=45)
        ))
        fig.update_layout(layout)
        return fig

    @staticmethod
    def create_pie_chart(summary_agg, group_col):
        # 新增：能耗佔比圓餅圖
        fig = px.pie(summary_agg, values="耗電量", names=group_col, hole=0.4, title="總耗電量佔比")
        fig.update_traces(textinfo='percent+label', textfont=dict(size=14, color='black'), marker=dict(colors=px.colors.qualitative.Safe))
        layout = VizEngine._common_layout()
        fig.update_layout(layout)
        return fig

    @staticmethod
    def create_unit_energy_chart(summary_agg, group_col):
        # 新增：單位能耗長條圖
        sorted_agg = summary_agg.sort_values("平均單位能耗")
        fig = px.bar(
            sorted_agg, x=group_col, y="平均單位能耗", text="平均單位能耗",
            title="平均單位能耗 (越低越好)"
        )
        fig.update_traces(marker_color='#145a32', texttemplate='%{text:.5f}', textposition='outside')
        layout = VizEngine._common_layout()
        layout.update(yaxis=dict(range=[0, sorted_agg['平均單位能耗'].max() * 1.2]))
        fig.update_layout(layout)
        return fig

# ==========================================
# 4. Report Engine (匯出中心) - 升級：整合新圖表
# ==========================================
class ReportEngine:
    @staticmethod
    def clean_markdown(text):
        if not isinstance(text, str): return str(text)
        return re.sub(r'(\*\*|\*|🔴|🟡|🟢)', '', text).strip()

    @staticmethod
    def generate_docx(df, summary_agg, texts, figures, analysis_scope):
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)
        
        head = doc.add_heading('生產效能診斷分析報告', 0)
        head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"分析範圍：{analysis_scope}")
        doc.add_paragraph(f"期間：{df['日期'].min()} ~ {df['日期'].max()}")
        doc.add_paragraph("-" * 60)
        
        doc.add_heading('1. 總體績效概覽', level=1)
        doc.add_paragraph(ReportEngine.clean_markdown(texts['kpi_summary']))
        
        table = doc.add_table(rows=1, cols=len(summary_agg.columns))
        table.style = 'Table Grid'
        hdr = table.rows[0].cells
        for i, col in enumerate(summary_agg.columns): hdr[i].text = str(col)
        
        for _, row in summary_agg.iterrows():
            cells = table.add_row().cells
            for i, val in enumerate(row):
                col_name = summary_agg.columns[i]
                if "OEE" in col_name: cells[i].text = f"{val:.1%}"
                elif "能耗" in col_name: cells[i].text = f"{val:.5f}"
                elif "損失" in col_name or "產量" in col_name: cells[i].text = f"{val:,.0f}"
                elif isinstance(val, float): cells[i].text = f"{val:.1f}"
                else: cells[i].text = str(val)
        
        doc.add_heading('2. 深度診斷分析', level=1)
        doc.add_paragraph(ReportEngine.clean_markdown(texts['benchmark_analysis']))
        doc.add_paragraph(ReportEngine.clean_markdown(texts['opportunity_analysis'])) # 新增：潛力預估
        
        def add_chart(key, title):
            doc.add_heading(title, level=2)
            if key in figures:
                try:
                    img = figures[key].to_image(format="png", width=800, height=400, scale=1.5)
                    doc.add_picture(BytesIO(img), width=Inches(6.5))
                except:
                    doc.add_paragraph("[註：圖表無法自動生成，請參考網頁版]")

        add_chart('rank', '綜合實力排名')
        add_chart('dual', '產量與能耗趨勢')
        
        doc.add_heading('3. 電力耗能深度分析', level=1) # 新增：獨立章節
        add_chart('pie', '總耗電量佔比')
        add_chart('unit', '平均單位能耗排名')

        doc.add_heading('4. 生產穩定性', level=1)
        doc.add_paragraph(ReportEngine.clean_markdown(texts['stability_analysis']))
        add_chart('cv', 'CV 變異係數')
        add_chart('scatter', '效率能耗矩陣')
        
        doc.add_heading('5. 策略行動建議', level=1)
        doc.add_paragraph(ReportEngine.clean_markdown(texts['action_plan']))
        
        bio = BytesIO()
        doc.save(bio)
        return bio

# ==========================================
# 5. Main App (主程式邏輯)
# ==========================================
def main():
    st.markdown("### 📥 數據輸入控制台")
    uploaded_file = st.file_uploader("匯入生產報表 (Excel/CSV)", type=["xlsx", "csv"], label_visibility="collapsed")
    
    if 'input_data' not in st.session_state:
        st.session_state.input_data = pd.DataFrame([
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO2", "OEE(%)": 50.1, "產量(雙)": 2009.5, "用電量(kWh)": 6.2},
            {"日期": "2025-11-17", "廠別": "A廠", "機台編號": "ACO4", "OEE(%)": 55.4, "產量(雙)": 4416.5, "用電量(kWh)": 9.1},
            {"日期": "2025-11-18", "廠別": "A廠", "機台編號": "ACO2", "OEE(%)": 48.5, "產量(雙)": 1950.0, "用電量(kWh)": 6.0},
        ])
        st.session_state.input_data['日期'] = pd.to_datetime(st.session_state.input_data['日期']).dt.date
    
    if uploaded_file:
        try:
            if uploaded_file.name.endswith('.csv'): 
                df_new = pd.read_csv(uploaded_file)
            else: 
                df_new = pd.read_excel(uploaded_file)
            
            rename_map = {"用電量(kWh)": "耗電量", "產量(雙)": "產量", "OEE(%)": "OEE_RAW", "設備": "機台編號", "機台": "機台編號"}
            for user_col, sys_col in rename_map.items():
                if user_col in df_new.columns: df_new = df_new.rename(columns={user_col: sys_col})
            
            st.session_state.input_data = df_new
        except:
            st.error("檔案讀取失敗，請檢查格式。")

    edited_df = st.data_editor(st.session_state.input_data, num_rows="dynamic", use_container_width=True)
    
    if st.button("🗑️ 清空所有數據"):
        st.session_state.input_data = pd.DataFrame(columns=["日期", "廠別", "機台編號", "OEE(%)", "產量(雙)", "用電量(kWh)"])
        st.rerun()

    st.markdown("---")
    st.markdown("#### ⚙️ 分析參數設定")
    c1, c2, c3 = st.columns(3)
    params = {
        'elec_price': c1.number_input("電價 (元/度)", value=3.5, step=0.1),
        'target_oee': c2.number_input("目標 OEE (%)", value=85.0, step=0.5),
        'product_margin': c3.number_input("獲利估算 (元/雙)", value=10.0, step=1.0)
    }
    
    st.write("")
    
    col_run, col_export = st.columns([1, 1])
    
    data_ready = False
    # 預先初始化
    df_res, summary_res, scope_res, texts_res, figs_res = None, None, None, None, {}

    if not edited_df.empty:
        try:
            df_res, summary_res, scope_res = DataEngine.clean_and_process(edited_df, params)
            if df_res is not None and summary_res is not None:
                data_ready = True
                texts_res = InsightEngine.generate_narrative(df_res, summary_res, 
                                                           "廠別" if scope_res=="跨廠區分析" else "機台編號", 
                                                           params)
                # 擴充圖表集
                figs_res = {
                    'rank': VizEngine.create_rank_chart(summary_res, "廠別" if scope_res=="跨廠區分析" else "機台編號"),
                    'cv': VizEngine.create_cv_chart(df_res, "廠別" if scope_res=="跨廠區分析" else "機台編號"),
                    'scatter': VizEngine.create_scatter_chart(df_res, "廠別" if scope_res=="跨廠區分析" else "機台編號"),
                    'dual': VizEngine.create_dual_axis_chart(df_res, "廠別" if scope_res=="跨廠區分析" else "機台編號"),
                    'pie': VizEngine.create_pie_chart(summary_res, "廠別" if scope_res=="跨廠區分析" else "機台編號"),
                    'unit': VizEngine.create_unit_energy_chart(summary_res, "廠別" if scope_res=="跨廠區分析" else "機台編號")
                }
            elif isinstance(scope_res, str):
                st.warning(scope_res)
        except Exception as e:
            st.error(f"數據處理錯誤: {e}")

    with col_run:
        start_btn = st.button("🚀 啟動全方位分析", type="primary")
        
    with col_export:
        if data_ready:
            try:
                docx = ReportEngine.generate_docx(df_res, summary_res, texts_res, figs_res, scope_res)
                st.download_button("📥 下載 Word 報告", docx.getvalue(), 
                                 f"生產效能報告_{pd.Timestamp.now().strftime('%Y%m%d')}.docx",
                                 "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"匯出失敗: {e}")
        else:
            st.button("📥 下載 Word 報告", disabled=True)

    if start_btn and data_ready:
        with st.spinner('正在進行深度診斷...'):
            time.sleep(0.5)
            st.markdown("---")
            st.title("生產效能診斷分析報告")
            
            st.header("1. 總體績效概覽")
            st.markdown(f'<div class="insight-box">{texts_res["kpi_summary"]}</div>', unsafe_allow_html=True)
            
            st.subheader("績效總表")
            st.dataframe(summary_res.style.format({
                "OEE": "{:.1%}", "平均單位能耗": "{:.5f}", "總損失": "${:,.0f}"
            }).background_gradient(subset=["OEE"], cmap="Blues"), use_container_width=True)
            
            st.plotly_chart(figs_res['rank'], use_container_width=True)
            
            # 深度分析
            st.header("2. 深度診斷分析")
            st.markdown(f'<div class="analysis-text">{texts_res["benchmark_analysis"]}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="insight-box">{texts_res["opportunity_analysis"]}</div>', unsafe_allow_html=True)
            st.subheader("產量與能耗趨勢")
            st.plotly_chart(figs_res['dual'], use_container_width=True)
            
            # 能耗分析 (新章節)
            st.header("3. 電力耗能深度分析")
            c1, c2 = st.columns(2)
            with c1: st.plotly_chart(figs_res['pie'], use_container_width=True)
            with c2: st.plotly_chart(figs_res['unit'], use_container_width=True)
            
            # 穩定性
            st.header("4. 生產趨勢與穩定性")
            c3, c4 = st.columns(2)
            with c3: 
                st.plotly_chart(figs_res['cv'], use_container_width=True)
                st.markdown(f'<div class="analysis-text">{texts_res["stability_analysis"]}</div>', unsafe_allow_html=True)
            with c4: 
                st.plotly_chart(figs_res['scatter'], use_container_width=True)
                st.markdown('<div class="analysis-text">理想落點為<b>右下角</b> (高效率低能耗)。</div>', unsafe_allow_html=True)
            
            st.header("5. 綜合診斷與建議")
            st.markdown(texts_res['action_plan'])

if __name__ == "__main__":
    main()
