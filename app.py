import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import time
import numpy as np

# --- 1. 頁面設定 ---
st.set_page_config(page_title="生產效能深度診斷報告", layout="centered")

# CSS 優化：模擬專業顧問報告格式 (黑底標題、清晰內文)
st.markdown("""
    <style>
    .main { background-color: #ffffff; }
    
    html, body, [class*="css"] {
        font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif;
        color: #1a1a1a;
    }
    
    /* 標題層級 */
    h1 { color: #000000; font-weight: 900; font-size: 2.4em; text-align: left; margin-bottom: 30px; border-bottom: 3px solid #000; padding-bottom: 10px; }
    h2 { color: #333333; font-weight: 800; font-size: 1.6em; margin-top: 50px; margin-bottom: 20px; border-left: 6px solid #e74c3c; padding-left: 15px; }
    h3 { color: #555555; font-weight: 700; font-size: 1.3em; margin-top: 30px; }
    
    /* 內文文字 */
    p, li, .stMarkdown {
        font-size: 16px !important;
        line-height: 1.8 !important;
        color: #333333 !important;
    }
    
    /* 重點強調字 */
    .highlight { font-weight: bold; color: #e74c3c; }
    .good { font-weight: bold; color: #27ae60; }
    
    /* 模擬圖片中的黑色表格風格 */
    .stDataFrame { border: 1px solid #ccc; }
    </style>
""", unsafe_allow_html=True)

# --- 2. 核心邏輯 ---

def init_session_state():
    if 'input_data' not in st.session_state:
        # 預設範例 (依照您的圖片邏輯模擬數據)
        st.session_state.input_data = pd.DataFrame([
            {"日期": "2025-11-17", "廠別": "S工廠", "機台編號": "ACO4", "OEE(%)": 60.5, "產量(雙)": 4400, "用電量(kWh)": 8.5},
            {"日期": "2025-11-17", "廠別": "S工廠", "機台編號": "ACO2", "OEE(%)": 45.2, "產量(雙)": 2100, "用電量(kWh)": 7.2},
            {"日期": "2025-11-17", "廠別": "S工廠", "機台編號": "ACO3", "OEE(%)": 28.5, "產量(雙)": 2150, "用電量(kWh)": 8.1},
            {"日期": "2025-11-18", "廠別": "S工廠", "機台編號": "ACO4", "OEE(%)": 62.1, "產量(雙)": 4500, "用電量(kWh)": 8.4},
            {"日期": "2025-11-18", "廠別": "S工廠", "機台編號": "ACO2", "OEE(%)": 46.5, "產量(雙)": 2200, "用電量(kWh)": 7.5},
            {"日期": "2025-11-18", "廠別": "S工廠", "機台編號": "ACO3", "OEE(%)": 29.0, "產量(雙)": 2100, "用電量(kWh)": 8.3},
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
        if "日期" in df.columns: df["日期"] = pd.to_datetime(df["日期"]).dt.date
        if "廠別" not in df.columns: df["廠別"] = "匯入廠區"
        return df, "OK"
    except Exception as e:
        return None, str(e)

# --- 3. 數據輸入介面 (保持不變) ---

st.markdown("### 📥 數據輸入")
col_in1, col_in2 = st.columns([3, 1])
with col_in1:
    uploaded_file = st.file_uploader("批次匯入 Excel", type=["xlsx", "csv"], label_visibility="collapsed")
    if uploaded_file:
        new_df, status = smart_load_file(uploaded_file)
        if status == "OK": st.session_state.input_data = new_df
with col_in2:
    if st.button("🗑️ 清空表格"):
        st.session_state.input_data = pd.DataFrame(columns=["日期", "廠別", "機台編號", "OEE(%)", "產量(雙)", "用電量(kWh)"])
        st.rerun()

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

# 參數設定
with st.expander("⚙️ 分析參數設定 (點擊展開)", expanded=False):
    c1, c2 = st.columns(2)
    elec_price = c1.number_input("電價 (元/度)", value=3.5, step=0.1)
    target_oee = c2.number_input("目標 OEE (%)", value=85.0, step=0.5)

st.write("")
start_analysis = st.button("📄 生成深度分析報告", type="primary")

# --- 4. 深度報告生成區 ---

if start_analysis:
    with st.spinner('正在進行深度數據洞察...'):
        time.sleep(1.0)
        
        # --- A. 數據清洗與計算 ---
        df = edited_df.copy()
        rename_map = {"用電量(kWh)": "耗電量", "產量(雙)": "產量", "OEE(%)": "OEE_RAW", "設備": "機台編號", "機台": "機台編號"}
        for user_col, sys_col in rename_map.items():
            if user_col in df.columns: df = df.rename(columns={user_col: sys_col})

        required = ["機台編號", "耗電量", "產量", "OEE_RAW"]
        if df.empty or not all(col in df.columns for col in required):
            st.error("❌ 資料不足，無法分析。請檢查必要欄位。")
        else:
            # 計算核心指標
            df["OEE"] = df["OEE_RAW"].apply(lambda x: x / 100.0 if x > 1.0 else x)
            df["單位能耗"] = df["耗電量"] / df["產量"]
            
            if "廠別" not in df.columns: df["廠別"] = "匯入廠區"
            factory_name = df["廠別"].iloc[0]
            start_date = df["日期"].min()
            end_date = df["日期"].max()

            # 機台彙整表 (Aggregation)
            agg = df.groupby("機台編號").agg({
                "產量": "sum", "耗電量": "sum", "OEE": "mean"
            }).reset_index()
            agg["單位能耗"] = agg["耗電量"] / agg["產量"]
            agg["排名"] = agg["單位能耗"].rank(ascending=True) # 單位能耗越低排名越前
            agg = agg.sort_values("排名")
            
            # --- 找出關鍵角色 ---
            best_m = agg.iloc[0] # 冠軍
            worst_m = agg.iloc[-1] # 問題
            middle_m = agg.iloc[1] if len(agg) > 2 else None
            
            # 計算比較倍率
            output_ratio = best_m["產量"] / worst_m["產量"]
            power_ratio = worst_m["單位能耗"] / best_m["單位能耗"]
            saving_potential = (worst_m["單位能耗"] - best_m["單位能耗"]) / worst_m["單位能耗"]

            # --- 報告開始 ---
            st.markdown("---")
            st.title("生產效能深度診斷報告")
            st.markdown(f"**分析對象：** {factory_name} ({len(agg)}台設備) &nbsp;&nbsp; **期間：** {start_date} 至 {end_date}")

            # ==========================================
            # 圖表區 (模仿圖片樣式)
            # ==========================================
            
            # 1. 每日單位能耗趨勢圖 (折線圖) - 上方
            st.markdown("#### 每日效率趨勢 (Unit Energy Trend)")
            st.caption("數值越低代表效率越高 (越省電)")
            
            fig_trend = px.line(df, x="日期", y="單位能耗", color="機台編號", markers=True)
            fig_trend.update_layout(
                xaxis_title="", yaxis_title="單位能耗 (kWh/雙)",
                legend_title="機台", plot_bgcolor="white",
                xaxis=dict(showgrid=True, gridcolor='#eee'),
                yaxis=dict(showgrid=True, gridcolor='#eee'),
                height=350
            )
            st.plotly_chart(fig_trend, use_container_width=True)

            # 2. 總產量 vs 總耗電 (雙長條圖) - 下方
            st.markdown("#### 總產出 vs 總耗能 (Total Output vs Power)")
            
            fig_bar = go.Figure()
            # 產量 Bar
            fig_bar.add_trace(go.Bar(
                x=agg["機台編號"], y=agg["產量"], name="總產量 (雙)",
                marker_color='#95a5a6', text=agg["產量"], textposition='auto'
            ))
            # 耗電 Bar
            fig_bar.add_trace(go.Bar(
                x=agg["機台編號"], y=agg["耗電量"], name="總用電量 (kWh)",
                marker_color='#e74c3c', text=agg["耗電量"], textposition='auto',
                yaxis='y2' # 使用第二Y軸
            ))
            
            fig_bar.update_layout(
                barmode='group', # 分組並排
                yaxis=dict(title="產量 (雙)"),
                yaxis2=dict(title="用電量 (kWh)", overlaying='y', side='right'),
                legend=dict(orientation="h", y=1.1),
                plot_bgcolor="white", height=400
            )
            st.plotly_chart(fig_bar, use_container_width=True)

            # ==========================================
            # 文字分析區 (深度解讀)
            # ==========================================
            
            st.header("1. 綜合效能總結")
            st.markdown("我計算了每台設備的**單位能耗 (kWh/雙)**，數值越低代表效率越高 (越省電)。")
            
            # 製作高對比表格
            display_table = agg[["機台編號", "產量", "耗電量", "OEE", "單位能耗", "排名"]].copy()
            display_table.columns = ["設備", "總產量(雙)", "總用電量(kWh)", "平均 OEE(%)", "整體能耗效率(kWh/雙)", "排名"]
            
            st.dataframe(
                display_table.style.format({
                    "總產量(雙)": "{:,.0f}", "總用電量(kWh)": "{:,.1f}", 
                    "平均 OEE(%)": "{:.1f}", "整體能耗效率(kWh/雙)": "{:.5f}"
                }),
                use_container_width=True, hide_index=True
            )

            st.header("2. 深度分析")

            # A. 冠軍設備分析
            st.subheader(f"A. 冠軍設備：{best_m['機台編號']}")
            st.markdown(f"""
            * **壓倒性優勢**：{best_m['機台編號']} 是表現最好的設備。它的產量是 {worst_m['機台編號']} 的 <span class='good'>{output_ratio:.1f} 倍</span> ({best_m['產量']:,.0f} vs {worst_m['產量']:,.0f})，展現極高的產能優勢。
            * **高效原因**：歸功於它較高的 **OEE (平均 {best_m['OEE']:.1%})**。高稼動率意味著機器大部分時間都在有效生產，分攤了基礎能耗，使其單位能耗低至 **{best_m['單位能耗']:.5f} kWh/雙**。
            """, unsafe_allow_html=True)

            # B. 問題設備分析
            st.subheader(f"B. 問題設備：{worst_m['機台編號']}")
            st.markdown(f"""
            * **高耗能警訊**：{worst_m['機台編號']} 是效率最差的設備。它的產量最低，但用電量 ({worst_m['耗電量']:.1f} kWh) 卻與其他高產能機台相去不遠。
            * **效率低落**：每生產一雙鞋，{worst_m['機台編號']} 需要消耗 **{worst_m['單位能耗']:.5f} kWh**，這比冠軍機台多耗費了 <span class='highlight'>{power_ratio:.1f} 倍</span> 的電力。
            * **關鍵因素**：其 OEE 極低 (平均 {worst_m['OEE']:.1%})。這暗示設備可能有大量的停機、待機或故障時間，導致「光吃電不產出」的基礎負載浪費。
            """, unsafe_allow_html=True)

            # C. 中庸設備 (如果有)
            if middle_m is not None:
                st.subheader(f"C. 中庸設備：{middle_m['機台編號']}")
                st.markdown(f"""
                * **表現平平**：{middle_m['機台編號']} 的產量與 OEE 介於兩者之間。雖然不像問題設備那麼嚴重，但其單位能耗仍高於冠軍機台，仍有優化空間。
                """)

            st.header("3. 每日效率趨勢分析 (見圖表上部)")
            
            # 自動分析趨勢
            trend_desc = ""
            for m in df['機台編號'].unique():
                m_data = df[df['機台編號'] == m]
                std = m_data['單位能耗'].std()
                if std < 0.0005:
                    trend_desc += f"* **{m}**：曲線平緩，顯示生產過程相對穩定。\n"
                else:
                    trend_desc += f"* **{m}**：曲線波動較大，顯示製程不穩定，需關注特定日期的異常。\n"
            
            st.markdown(trend_desc)

            st.header("4. 建議與行動")
            
            st.markdown(f"""
            1.  **{worst_m['機台編號']} 優先檢修**：其能耗異常高且 OEE 極低，建議立即檢查是否為「待機未關機」或「頻繁故障」導致的電力浪費。
            2.  **複製 {best_m['機台編號']} 經驗**：{best_m['機台編號']} 的參數設定與操作模式顯然較優，應作為標竿 (Benchmark) 推廣至 {worst_m['機台編號']}。
            3.  **節能潛力**：若能將 {worst_m['機台編號']} 的效率提升至 {best_m['機台編號']} 的水準，其電力成本可降低約 <span class='good'>{saving_potential:.0%}</span>。
            """, unsafe_allow_html=True)
