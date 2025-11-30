import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import time
import numpy as np

# --- 1. 頁面設定 ---
st.set_page_config(page_title="生產效能智慧分析系統 Pro", layout="wide") # 改為寬螢幕模式以容納更多資訊

# CSS 優化：增強閱讀性與區塊感
st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    .stButton>button { width: 100%; border-radius: 8px; height: 3.5em; font-weight: bold; font-size: 1.1em; }
    h1 { color: #2c3e50; font-family: 'Microsoft JhengHei'; font-weight: 800; }
    h2 { color: #34495e; border-bottom: 2px solid #3498db; padding-bottom: 10px; margin-top: 40px; }
    h3 { color: #2980b9; margin-top: 20px; font-weight: 600; }
    .insight-box { background-color: #e8f6f3; padding: 15px; border-radius: 5px; border-left: 5px solid #1abc9c; margin-bottom: 20px; }
    .guide-box { background-color: #fdfefe; padding: 10px; border-radius: 5px; border: 1px solid #dcdcdc; font-size: 0.9em; color: #555; margin-bottom: 10px; }
    </style>
""", unsafe_allow_html=True)

# --- 2. 核心邏輯 (保持不變) ---

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

# --- 3. 數據輸入介面 (保持您要求的原樣) ---

st.title("🏭 生產效能智慧分析系統 Pro")
st.markdown("### 1. 數據輸入 (Data Input)")

col_input, col_param = st.columns([2, 1])

with col_input:
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
            "機台編號": st.column_config.TextColumn(label="設備/機台編號"),
            "OEE(%)": st.column_config.NumberColumn("OEE(%)", format="%.1f"),
            "產量(雙)": st.column_config.NumberColumn("產量(雙)"),
            "用電量(kWh)": st.column_config.NumberColumn("用電量(kWh)"),
        }
    )
    if st.button("🗑️ 清空表格數據"):
        st.session_state.input_data = pd.DataFrame(columns=["日期", "廠別", "機台編號", "OEE(%)", "產量(雙)", "用電量(kWh)"])
        st.rerun()

with col_param:
    st.markdown("#### 分析參數")
    elec_price = st.number_input("平均電價 (元/度)", value=3.5, step=0.1)
    target_oee = st.number_input("目標 OEE 基準 (%)", value=85.0, step=0.5)
    product_margin = st.number_input("每雙獲利估算 (元)", value=10.0, step=1.0)
    
    st.write("")
    st.write("")
    start_analysis = st.button("🚀 啟動全方位分析", type="primary")

# --- 4. 執行與分析邏輯 (重點更新區域) ---

if start_analysis:
    with st.spinner('🔄 AI 正在解讀數據趨勢、計算成本損失、撰寫診斷報告...'):
        time.sleep(1.5)
        
        # --- A. 數據前處理 ---
        df = edited_df.copy()
        
        # 1. 欄位映射
        rename_map = {
            "用電量(kWh)": "耗電量", "產量(雙)": "產量", 
            "OEE(%)": "OEE_RAW", "設備": "機台編號", "機台": "機台編號"
        }
        for user_col, sys_col in rename_map.items():
            if user_col in df.columns:
                df = df.rename(columns={user_col: sys_col})

        required = ["機台編號", "耗電量", "產量", "OEE_RAW"]
        
        if df.empty or not all(col in df.columns for col in required):
            st.error("❌ 缺少必要欄位，無法分析。請檢查輸入表格。")
        else:
            # 2. 計算基礎指標
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
            
            if "廠別" not in df.columns: df["廠別"] = "匯入廠區"
            group_col = "廠別" if df["廠別"].nunique() > 1 else "機台編號"

            # 3. 聚合數據 (機台層級總表)
            machine_agg = df.groupby("機台編號").agg({
                "OEE": "mean", "產量": "sum", "耗電量": "sum", 
                "能源損失": "sum", "產能損失機會成本": "sum", "總損失": "sum"
            }).reset_index()
            machine_agg["平均單位能耗"] = machine_agg["耗電量"] / machine_agg["產量"]
            machine_agg = machine_agg.sort_values("OEE", ascending=False) # 預設依 OEE 排名

            st.success("✅ 分析報告已生成！")
            st.markdown("---")

            # --- B. 分頁報告 ---
            tab1, tab2, tab3, tab4 = st.tabs([
                "📋 總覽與排名 (Overview)", 
                "📈 趨勢與穩定性 (Trends)", 
                "⚡ 電力與能耗深度分析 (Energy)", 
                "📝 全機台總結與診斷 (Conclusion)"
            ])

            # === Tab 1: 總覽與排名 ===
            with tab1:
                st.header("1. 生產全貌與排行榜")
                
                # KPI
                k1, k2, k3, k4 = st.columns(4)
                avg_oee_total = df["OEE"].mean()
                total_loss = df["總損失"].sum()
                k1.metric("全廠平均 OEE", f"{avg_oee_total:.1%}", delta=f"{avg_oee_total - target_oee/100:.1%}")
                k2.metric("總產量", f"{df['產量'].sum():,.0f} 雙")
                k3.metric("總耗電量", f"{df['耗電量'].sum():,.1f} kWh")
                k4.metric("總潛在損失 (NTD)", f"${total_loss:,.0f}", delta="含電費與產能損失", delta_color="inverse")

                col_t1, col_t2 = st.columns([3, 2])
                
                with col_t1:
                    st.subheader("原始數據明細 (所有紀錄)")
                    # 動態高度計算 (避免捲軸)：每行約 35px，加上表頭緩衝
                    table_height = (len(df) + 1) * 35 + 3
                    
                    display_cols = ["日期", "機台編號", "OEE", "產量", "耗電量", "單位能耗"]
                    final_table = df[display_cols].rename(columns={"OEE": "OEE(%)", "產量": "產量(雙)", "耗電量": "用電量(kWh)"})
                    
                    try:
                        st.dataframe(
                            final_table.style.format({
                                "OEE(%)": "{:.1%}", "單位能耗": "{:.5f}"
                            }).background_gradient(subset=["OEE(%)"], cmap="RdYlGn"),
                            use_container_width=True, 
                            height=table_height # 關鍵：設定高度以取消捲軸
                        )
                    except:
                        st.dataframe(final_table, use_container_width=True, height=table_height)

                with col_t2:
                    st.subheader("🏆 機台綜合實力排行榜")
                    st.markdown('<div class="guide-box">💡 這是將多天數據加總後的平均表現，用來評斷哪一台機器長期表現最好。</div>', unsafe_allow_html=True)
                    
                    # 排名圖表
                    fig_rank = px.bar(
                        machine_agg.sort_values("OEE", ascending=True), 
                        x="OEE", y="機台編號", orientation='h',
                        title="各機台平均 OEE 排名", text="OEE",
                        color="OEE", color_continuous_scale="Blues"
                    )
                    fig_rank.update_traces(texttemplate='%{text:.1%}', textposition='outside')
                    st.plotly_chart(fig_rank, use_container_width=True)
                    
                    # AI 解讀
                    top_machine = machine_agg.iloc[0]['機台編號']
                    last_machine = machine_agg.iloc[-1]['機台編號']
                    st.markdown(f"""
                    <div class="insight-box">
                    <b>🤖 AI 排名解析：</b><br>
                    在此次分析區間內，<b>{top_machine}</b> 是表現最優異的冠軍設備，平均效率最高。<br>
                    反之，<b>{last_machine}</b> 排名墊底，是拉低整體平均的主要原因。
                    </div>
                    """, unsafe_allow_html=True)

            # === Tab 2: 趨勢與穩定性 ===
            with tab2:
                st.header("2. 趨勢波動與相關性解讀")
                
                c1, c2 = st.columns(2)
                
                # --- 左圖：CV 變異係數 ---
                with c1:
                    st.subheader("A. 生產穩定度分析 (CV值)")
                    st.markdown("""
                    <div class="guide-box">
                    <b>💡 圖表怎麼看？</b><br>
                    此圖顯示設備的「不穩定程度」。<br>
                    • <b>數值越低 (長條越短)</b>：代表該機台每天表現差不多，非常穩定 (Good)。<br>
                    • <b>數值越高 (長條越長)</b>：代表該機台時好時壞 (Bad)，像神經刀一樣。
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if len(df) > 1:
                        cv_data = df.groupby("機台編號")["OEE"].agg(['mean', 'std'])
                        cv_data['CV(%)'] = (cv_data['std'] / cv_data['mean']) * 100
                        cv_data = cv_data.reset_index().sort_values('CV(%)')
                        
                        fig_cv = px.bar(cv_data, x="機台編號", y="CV(%)", text="CV(%)", 
                                      color="CV(%)", color_continuous_scale="Reds", 
                                      title="各機台 OEE 波動率 (越低越好)")
                        fig_cv.update_traces(texttemplate='%{text:.1f}%')
                        st.plotly_chart(fig_cv, use_container_width=True)
                        
                        # AI 解讀
                        most_unstable = cv_data.iloc[-1]['機台編號']
                        most_stable = cv_data.iloc[0]['機台編號']
                        st.markdown(f"""
                        <div class="insight-box">
                        <b>🤖 AI 穩定性診斷：</b><br>
                        • <b>{most_stable}</b> 是最穩定的設備，這通常代表其參數設定或操作人員手法最標準。<br>
                        • <b>{most_unstable}</b> 的波動最大，建議檢查是否受「換線頻繁」或「進料品質不一」影響。
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.info("⚠️ 數據量不足 (需至少兩天的數據才能計算波動率)。")

                # --- 右圖：相關性分析 ---
                with c2:
                    st.subheader("B. 效率 vs 能耗 關聯圖")
                    st.markdown("""
                    <div class="guide-box">
                    <b>💡 圖表怎麼看？</b><br>
                    • <b>X軸 (橫向)</b>：OEE 效率 (越右邊越好)。<br>
                    • <b>Y軸 (縱向)</b>：單位能耗 (越下面越省電)。<br>
                    • <b>理想落點</b>：圖表的<b>「右下角」</b> (高效率、低耗能)。<br>
                    • <b>異常落點</b>：圖表的<b>「左上角」</b> (低效率卻很耗電，通常是空轉)。
                    </div>
                    """, unsafe_allow_html=True)
                    
                    try:
                        fig_corr = px.scatter(
                            df, x="OEE", y="單位能耗", 
                            color="機台編號", size="產量", 
                            trendline="ols", 
                            title="OEE vs 單位能耗分佈"
                        )
                        st.plotly_chart(fig_corr, use_container_width=True)
                    except:
                        fig_corr = px.scatter(df, x="OEE", y="單位能耗", color="機台編號", size="產量")
                        st.plotly_chart(fig_corr, use_container_width=True)

                    st.markdown(f"""
                    <div class="insight-box">
                    <b>🤖 AI 關聯性解析：</b><br>
                    觀察趨勢線，若呈現<b>「左上往右下」</b>傾斜，代表工廠管理健康（效率越高越省電）。<br>
                    若發現有圓點孤零零地出現在<b>左上方</b>，該時間點該機台極可能發生了<b>「待機未關機」</b>的浪費行為。
                    </div>
                    """, unsafe_allow_html=True)

            # === Tab 3: 電力與能耗深度分析 ===
            with tab3:
                st.header("3. 電力消耗與產出效率深度分析")
                
                col_e1, col_e2 = st.columns(2)
                
                with col_e1:
                    st.subheader("A. 誰是吃電怪獸？ (總耗電量排名)")
                    st.markdown("""
                    <div class="guide-box">
                    <b>💡 圖表怎麼看？</b><br>
                    單純比較這段時間內，哪一台機器用掉最多電 (kWh)。注意：用電多不代表效率差，要配合右圖看。
                    </div>
                    """, unsafe_allow_html=True)
                    
                    fig_power_sum = px.pie(machine_agg, values="耗電量", names="機台編號", hole=0.4, title="各機台總耗電量佔比")
                    st.plotly_chart(fig_power_sum, use_container_width=True)

                with col_e2:
                    st.subheader("B. 用一度電能做多少事？ (單位能耗)")
                    st.markdown("""
                    <div class="guide-box">
                    <b>💡 圖表怎麼看？</b><br>
                    這是最公平的指標。計算生產每一雙鞋子平均要花多少電。<br>
                    • <b>柱子越低越好</b>：代表該機台省電技術最好。
                    </div>
                    """, unsafe_allow_html=True)
                    
                    fig_unit_power = px.bar(
                        machine_agg, x="機台編號", y="平均單位能耗", 
                        color="平均單位能耗", title="平均單位能耗 (kWh/雙)",
                        color_continuous_scale="Viridis_r" # 顏色反轉，數值低(省電)顯示亮色
                    )
                    st.plotly_chart(fig_unit_power, use_container_width=True)

                # 進階：電力 vs 產量 雙軸圖
                st.subheader("C. 產量與電力供需檢視 (雙軸分析)")
                st.markdown("""
                <div class="guide-box">
                <b>💡 圖表怎麼看？</b><br>
                將產量(柱狀)與用電量(折線)放在一起看。<br>
                正常情況下，柱子高(產量多)的時候，折線(用電)也要跟著高。<b>如果柱子很低，但折線卻很高，那就是異常！</b>
                </div>
                """, unsafe_allow_html=True)

                # 準備雙軸圖資料
                df_sorted = df.sort_values(["機台編號", "日期"])
                fig_dual = go.Figure()
                
                # 為了避免圖表太亂，我們以「機台+日期」為 X 軸
                x_axis_label = df_sorted["機台編號"] + " (" + df_sorted["日期"].astype(str) + ")"
                
                fig_dual.add_trace(go.Bar(
                    x=x_axis_label, y=df_sorted["產量"], name="產量 (雙)", 
                    marker_color="#3498db", opacity=0.6
                ))
                fig_dual.add_trace(go.Scatter(
                    x=x_axis_label, y=df_sorted["耗電量"], name="耗電量 (kWh)",
                    yaxis="y2", mode="lines+markers", line=dict(color="#e74c3c", width=3)
                ))
                
                fig_dual.update_layout(
                    title="產量 vs 耗電量 每日對照圖",
                    yaxis=dict(title="產量 (雙)"),
                    yaxis2=dict(title="耗電量 (kWh)", overlaying="y", side="right"),
                    xaxis=dict(title="機台 (日期)", tickangle=45),
                    legend=dict(orientation="h", y=1.1)
                )
                st.plotly_chart(fig_dual, use_container_width=True)
                
                # 電力 AI 總結
                best_power_machine = machine_agg.sort_values("平均單位能耗").iloc[0]
                worst_power_machine = machine_agg.sort_values("平均單位能耗").iloc[-1]
                
                st.markdown(f"""
                <div class="insight-box">
                <b>⚡ 電力分析 AI 結論：</b><br>
                1. <b>能源效率王</b>：<b>{best_power_machine['機台編號']}</b>。每生產一雙鞋僅需 <b>{best_power_machine['平均單位能耗']:.5f} kWh</b>。<br>
                2. <b>耗能異常機台</b>：<b>{worst_power_machine['機台編號']}</b>。其單位能耗最高，比最佳機台高出了 <b>{((worst_power_machine['平均單位能耗'] / best_power_machine['平均單位能耗']) - 1):.1%}</b>。<br>
                3. <b>建議</b>：請工程部門檢查 {worst_power_machine['機台編號']} 的馬達效率或傳動系統阻力，這可能是硬體老化或潤滑不足的徵兆。
                </div>
                """, unsafe_allow_html=True)

            # === Tab 4: 全機台總結與診斷 ===
            with tab4:
                st.header("4. 全機台 AI 診斷總結報告")
                st.markdown("以下針對每一台設備進行獨立的數據診斷與行動建議：")
                
                for index, row in machine_agg.iterrows():
                    m_name = row['機台編號']
                    m_oee = row['OEE']
                    m_loss = row['總損失']
                    m_rank = index + 1 # 目前是依 OEE 排序的
                    
                    # 邏輯判斷產生文案
                    if m_oee >= target_oee/100:
                        status = "🟢 優良 (Excellent)"
                        advice = "保持目前運作模式，可作為示範機台，將其參數複製給其他設備。"
                        box_color = "#d4edda"
                    elif m_oee >= 0.70:
                        status = "🟡 尚可 (Average)"
                        advice = "表現平穩但仍有提升空間。建議分析短暫停機原因，目標提升 5-10% 稼動率。"
                        box_color = "#fff3cd"
                    else:
                        status = "🔴 嚴重異常 (Critical)"
                        advice = f"該機台為主要虧損來源 (損失 NT$ {m_loss:,.0f})。請立即安排停機檢修，確認是設備故障還是排程問題。"
                        box_color = "#f8d7da"

                    # 顯示卡片
                    st.markdown(f"""
                    <div style="background-color: {box_color}; padding: 20px; border-radius: 10px; margin-bottom: 15px; border: 1px solid #ddd;">
                        <h3 style="margin-top:0;">🔧 設備：{m_name}</h3>
                        <p><b>• 綜合排名：</b> 第 {m_rank} 名<br>
                        <b>• 平均 OEE：</b> {m_oee:.1%} <br>
                        <b>• 狀態評估：</b> <strong>{status}</strong><br>
                        <b>• 潛在財務損失：</b> NT$ {m_loss:,.0f}<br>
                        <b>• AI 行動建議：</b> {advice}</p>
                    </div>
                    """, unsafe_allow_html=True)
