import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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
    
    /* 規格 5.1: 全域通用字體 */
    html, body, [class*="css"] {
        font-family: Arial, sans-serif;
        color: #000000;
    }
    
    h1 { color: #000000; font-weight: 900; font-size: 2.4em; text-align: center; border-bottom: 4px solid #2c3e50; padding-bottom: 15px; margin-bottom: 30px; }
    h2 { color: #1a5276; border-left: 7px solid #1a5276; padding-left: 12px; margin-top: 40px; font-size: 1.6em; font-weight: bold; background-color: #f8f9fa; padding-top: 5px; padding-bottom: 5px; }
    h3 { color: #2e4053; margin-top: 25px; font-size: 1.3em; font-weight: 700; }
    
    p, li, .stMarkdown { font-size: 16px !important; line-height: 1.7 !important; color: #212f3d !important; }
    div[data-testid="stMetricValue"] { font-size: 28px !important; color: #17202a !important; font-weight: bold; }
    
    /* 規格 5.2: 視覺元件樣式 */
    .insight-box { border: 1px solid #d6eaf8; background-color: #ebf5fb; padding: 15px; border-radius: 5px; margin-top: 10px; margin-bottom: 20px; }
    .summary-box { border: 2px solid #566573; background-color: #fdfefe; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
    
    /* 表格優化 */
    thead tr th:first-child {display:none} tbody th {display:none}
    div.stButton > button:first-child { width: 100%; height: 3em; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. Data Engine (數據處理核心) - 規格 4.1 & 4.2
# ==========================================
class DataEngine:
    @staticmethod
    def clean_and_process(df_raw, params):
        df = df_raw.copy()
        
        # 1. 智慧欄位映射 (Smart Mapping)
        rename_map = {
            "用電量(kWh)": "耗電量", "產量(雙)": "產量", 
            "OEE(%)": "OEE_RAW", "設備": "機台編號", "機台": "機台編號"
        }
        for user_col, sys_col in rename_map.items():
            if user_col in df.columns: df = df.rename(columns={user_col: sys_col})
            
        # 2. 完整性檢查
        required_cols = ["機台編號", "耗電量", "產量", "OEE_RAW"]
        if not all(col in df.columns for col in required_cols):
            return None, None, f"缺少必要欄位: {[c for c in required_cols if c not in df.columns]}"
            
        if "日期" in df.columns: df["日期"] = pd.to_datetime(df["日期"]).dt.date
        if "廠別" not in df.columns: df["廠別"] = "匯入廠區"

        # 3. 核心指標運算
        # OEE 正規化
        df["OEE"] = df["OEE_RAW"].apply(lambda x: x / 100.0 if x > 1.0 else x)
        
        # 單位能耗 (防除以0)
        df["單位能耗"] = df.apply(lambda row: row["耗電量"] / row["產量"] if row["產量"] > 0 else 0, axis=1)
        
        # 基準運算 (Benchmark Energy - 全局最低)
        valid_energies = df[df["單位能耗"] > 0]["單位能耗"]
        best_energy = valid_energies.min() if not valid_energies.empty else 0
        
        # 財務損失運算
        elec_price = params['elec_price']
        target_oee = params['target_oee'] / 100.0
        margin = params['product_margin']
        
        # 能源損失計算
        df["能源損失"] = df.apply(lambda row: max(0, (row["單位能耗"] - best_energy) * row["產量"] * elec_price), axis=1)
        
        # 產能機會成本計算 (規格 4.2.4)
        df["產能損失機會成本"] = df.apply(
            lambda row: ((target_oee - row["OEE"]) / row["OEE"] * row["產量"] * margin) 
            if 0 < row["OEE"] < target_oee else 0, axis=1
        )
        
        df["總損失"] = df["能源損失"] + df["產能損失機會成本"]
        
        # 4. 聚合運算
        group_col = "廠別" if df["廠別"].nunique() > 1 else "機台編號"
        analysis_scope = "跨廠區分析" if group_col == "廠別" else "單廠設備分析"
        
        summary_agg = df.groupby(group_col).agg({
            "OEE": "mean", "產量": "sum", "耗電量": "sum", 
            "能源損失": "sum", "產能損失機會成本": "sum", "總損失": "sum"
        }).reset_index()
        
        # 重新計算聚合後的平均單位能耗
        summary_agg["平均單位能耗"] = summary_agg.apply(
            lambda row: row["耗電量"] / row["產量"] if row["產量"] > 0 else 0, axis=1
        )
        summary_agg = summary_agg.sort_values("OEE", ascending=False)
        
        return df, summary_agg, analysis_scope

# ==========================================
# 2. Insight Engine (診斷分析大腦) - 規格 3.0
# ==========================================
class InsightEngine:
    @staticmethod
    def generate_narrative(df, summary_agg, group_col, params):
        texts = {}
        target_oee = params['target_oee'] / 100.0
        margin = params['product_margin']
        
        # 1. 總體 KPI (Executive Summary)
        avg_oee = df["OEE"].mean()
        total_loss = df["總損失"].sum()
        best_name = summary_agg.iloc[0][group_col]
        worst_name = summary_agg.iloc[-1][group_col]
        
        texts['kpi_summary'] = f"本次分析區間內，全廠平均 OEE 為 **{avg_oee:.1%}**。其中 **{best_name}** 表現最佳，為全廠標竿；而 **{worst_name}** 效率敬陪末座，是造成全廠 **NT$ {total_loss:,.0f}** 潛在損失的主要原因。"
        
        # 2. 標竿與落差分析 (Benchmark Analysis)
        best_machine = summary_agg.iloc[0]
        worst_machine = summary_agg.iloc[-1]
        
        # 計算倍數落差
        energy_gap_msg = ""
        if best_machine['平均單位能耗'] > 0 and worst_machine['平均單位能耗'] > 0:
            ratio = worst_machine['平均單位能耗'] / best_machine['平均單位能耗']
            energy_gap_msg = f"比冠軍機台多消耗了 **{ratio:.1f} 倍** 的電力"
        
        # 3. 產能潛力預估 (Opportunity Estimation)
        # 估算若最差機台達到標竿機台的 OEE，能多生產多少
        potential_prod = 0
        if worst_machine['OEE'] > 0:
            potential_prod = (best_machine['OEE'] - worst_machine['OEE']) / worst_machine['OEE'] * worst_machine['產量']
        potential_revenue = potential_prod * margin
        
        texts['benchmark_analysis'] = f"""
        * **標竿設備 ({best_machine[group_col]})**：表現最佳，平均 OEE 達 **{best_machine['OEE']:.1%}**，單位能耗最低 ({best_machine['平均單位能耗']:.5f} kWh/雙)。
        * **瓶頸設備 ({worst_machine[group_col]})**：{energy_gap_msg}。若能將其效率提升至標竿水準，本期間預計可額外生產 **{potential_prod:,.0f} 雙**，相當於挽回 **NT$ {potential_revenue:,.0f}** 的營收損失。
        """
        
        # 4. 穩定性分析
        cv_text = "數據量不足以計算波動率。"
        if len(df) > 1:
            cv_series = df.groupby(group_col)["OEE"].std() / df.groupby(group_col)["OEE"].mean()
            most_stable = cv_series.idxmin()
            most_unstable = cv_series.idxmax()
            cv_text = f"**{most_stable}** 生產節奏最穩定 (CV最低)；**{most_unstable}** 波動最大，顯示製程或人員操作存在變異。"
        texts['stability_analysis'] = cv_text
        
        # 5. 策略行動建議 (Strategic Action)
        crit_list, avg_list, good_list = [], [], []
        for _, row in summary_agg.iterrows():
            name = row[group_col]
            if row['OEE'] >= target_oee: good_list.append(name)
            elif row['OEE'] >= 0.7: avg_list.append(name)
            else: crit_list.append(name)
            
        action_text = ""
        if crit_list:
            action_text += f"🔴 **優先改善 (Priority)**：{', '.join(crit_list)}\n   * 問題：OEE 低於 70%，效率偏低。\n   * 行動：立即調閱異常停機代碼，檢查是否有「待機未關機」或「頻繁短停機」。\n\n"
        if avg_list:
            action_text += f"🟡 **效能提升 (Improvement)**：{', '.join(avg_list)}\n   * 問題：表現平穩但未達標竿。\n   * 行動：微調參數 (速度/溫度)，目標提升 5-10% 稼動率。\n\n"
        if good_list:
            action_text += f"🟢 **標竿管理 (Benchmark)**：{', '.join(good_list)}\n   * 表現：運作優異。\n   * 行動：將其操作參數標準化 (SOP)，推廣至其他設備。"
            
        texts['action_plan'] = action_text
        return texts

# ==========================================
# 3. Viz Engine (視覺化中心) - 規格 3.1
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
        fig.update_traces(marker_
