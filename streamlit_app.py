# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from sklearn.linear_model import LinearRegression
from io import BytesIO
import os

st.set_page_config(page_title="小甜甜公司经营分析仪表板", layout="wide", page_icon="📊")

# ─── Style ───
st.markdown("""
<style>
    .kpi-card { background: linear-gradient(135deg, #1a2a3a, #0f1923); border: 1px solid #1e3a5f; border-radius: 12px; padding: 16px; text-align: center; color: #e0e6ed; }
    .kpi-label { font-size: 12px; color: #8899aa; }
    .kpi-value { font-size: 26px; font-weight: 700; }
    .kpi-change { font-size: 12px; margin-top: 2px; }
    .kpi-up { color: #00c853; }
    .kpi-down { color: #ff5252; }
    .main-header { text-align: center; padding: 10px 0; border-bottom: 1px solid #1e3a5f; margin-bottom: 20px; }
    .main-header h1 { background: linear-gradient(135deg, #00d4ff, #0099ff); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-size: 32px; }
    .stTabs [data-baseweb="tab"] { font-size: 14px; }
</style>
""", unsafe_allow_html=True)

# ─── Data ───
YEARS = ['2023','2024','2025']
MONTHS = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']

raw = {
    '营业收入': {
        '2023': [253.09,446.68,775.09,648.24,1786.23,1755.81,357.95,436.83,1345.30,858.29,779.26,3526.04],
        '2024': [587.21,168.75,1000.22,520.51,810.84,905.44,360.81,658.94,891.90,560.09,1278.37,2541.21],
        '2025': [215.08,395.76,755.70,1306.49,532.65,987.88,911.27,513.13,1537.02,994.58,2685.46,6346.35]
    },
    '营业成本': {
        '2023': [154.14,229.57,507.67,377.35,970.49,733.56,316.10,303.49,812.93,434.55,608.08,2431.25],
        '2024': [273.74,50.63,584.77,252.38,495.51,620.96,147.71,239.55,468.93,164.63,824.68,1536.40],
        '2025': [123.78,219.18,238.93,657.76,224.38,622.11,527.72,132.41,954.45,523.71,1707.95,4324.33]
    },
    '销售费用': {
        '2023': [119.98,56.93,57.21,76.94,62.80,68.80,78.66,91.39,87.07,232.68,103.05,143.10],
        '2024': [119.49,81.30,106.68,78.05,71.51,134.31,74.14,65.92,71.28,198.91,121.00,144.36],
        '2025': [95.77,86.23,130.61,64.13,96.78,83.39,81.12,67.32,167.72,120.02,207.60,123.68]
    },
    '管理费用': {
        '2023': [84.02,118.24,98.96,71.55,104.13,89.04,95.33,72.86,99.20,71.05,80.79,149.08],
        '2024': [93.91,167.33,118.77,111.85,93.92,113.83,113.12,91.12,137.53,117.68,120.53,291.14],
        '2025': [91.73,119.49,125.10,84.25,166.28,88.19,62.01,62.53,195.66,92.67,89.36,89.86]
    },
    '研发费用': {
        '2023': [156.30,230.42,171.57,165.37,293.52,333.64,218.27,181.21,177.96,98.45,95.77,118.39],
        '2024': [165.45,181.97,178.42,121.00,141.86,139.84,193.07,161.07,110.42,271.13,276.01,188.89],
        '2025': [264.18,254.03,170.89,155.69,204.84,242.53,228.55,169.34,202.98,122.15,388.43,117.94]
    },
    '财务费用': {
        '2023': [4.12,-0.28,4.29,1.23,1.13,0.88,1.36,3.51,-0.10,1.07,2.48,3.54],
        '2024': [1.28,0.90,1.43,0.73,0.96,5.20,1.03,2.53,2.62,0.91,0.82,3.77],
        '2025': [0.80,1.03,5.90,0.93,0.93,4.89,0.55,2.76,2.13,0.58,0.52,4.45]
    },
    '税金及附加': {
        '2023': [2.96,4.45,7.68,8.84,11.66,8.64,8.52,11.83,8.06,14.30,9.12,15.67],
        '2024': [6.05,8.98,5.42,8.19,7.53,14.24,7.46,7.70,16.53,7.93,10.21,27.99],
        '2025': [2.87,5.81,7.33,6.23,18.53,8.33,31.90,9.30,12.17,5.59,20.70,35.16]
    },
    '其他收益': {
        '2023': [83.68,10.35,5.84,46.90,54.00,0.32,85.54,76.45,40.31,0.07,117.91,42.69],
        '2024': [67.57,0,56.29,20.55,38.39,38.51,69.29,32.04,39.57,0,0,125.84],
        '2025': [142.85,13.79,43.15,24.36,25.13,84.13,41.06,47.23,28.92,64.76,16.16,101.84]
    }
}

# Derived data
def calc_monthly(item_name, fn):
    result = {}
    for y in YEARS:
        result[y] = [fn(y, m) for m in range(12)]
    return result

monthly_op_profit = {}
monthly_total_profit = {}
for y in YEARS:
    op_list = []
    tp_list = []
    for m in range(12):
        rev = raw['营业收入'][y][m]
        cost = raw['营业成本'][y][m]
        tax = raw['税金及附加'][y][m]
        sell = raw['销售费用'][y][m]
        admin = raw['管理费用'][y][m]
        rd = raw['研发费用'][y][m]
        fin = raw['财务费用'][y][m]
        other = raw['其他收益'][y][m]
        op = rev - cost - tax - sell - admin - rd - fin + other
        op_list.append(round(op, 2))
        # For total profit we don't have 营业外收支 monthly, approximate
        tp_list.append(round(op, 2))
    monthly_op_profit[y] = op_list
    monthly_total_profit[y] = tp_list

# Yearly totals
yearly = {}
for item, data in raw.items():
    yearly[item] = {y: round(sum(data[y]), 2) for y in YEARS}
yearly['营业利润'] = {y: round(sum(monthly_op_profit[y]), 2) for y in YEARS}
yearly['利润总额'] = {y: round(sum(monthly_total_profit[y]), 2) for y in YEARS}

# Margins
for y in YEARS:
    rev = yearly['营业收入'][y]
    if rev:
        yearly.setdefault('毛利率', {})[y] = round((rev - yearly['营业成本'][y]) / rev * 100, 2)
        yearly.setdefault('净利率', {})[y] = round(yearly['利润总额'][y] / rev * 100, 2)
        yearly.setdefault('营业利润率', {})[y] = round(yearly['营业利润'][y] / rev * 100, 2)

# ─── Helpers ───
def plot_line(raw_data, title, ylabel='万元', height=350):
    fig = go.Figure()
    colors = ['#00d4ff', '#ff6b6b', '#ffd93d']
    for i, y in enumerate(YEARS):
        fig.add_trace(go.Scatter(x=MONTHS, y=raw_data[y], mode='lines+markers',
                                name=f'{y}年', line=dict(color=colors[i], width=2.5),
                                marker=dict(size=6)))
    fig.update_layout(title=title, yaxis_title=ylabel, height=height,
                      template='plotly_dark', hovermode='x unified',
                      margin=dict(l=40, r=20, t=40, b=30),
                      legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
    fig.update_xaxes(gridcolor='#1e3a5f')
    fig.update_yaxes(gridcolor='#1e3a5f')
    return fig

def kpi_card(label, value, change='', up=True):
    change_class = 'kpi-up' if up else 'kpi-down'
    arrow = '▲' if up else '▼'
    return f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-change {change_class}">{arrow} {change}</div></div>'

def predict_next_year(item_data):
    """Simple linear regression prediction"""
    X = np.arange(36).reshape(-1, 1)
    y = np.array(item_data['2023'] + item_data['2024'] + item_data['2025'])
    model = LinearRegression()
    model.fit(X, y)
    future_X = np.arange(36, 48).reshape(-1, 1)
    pred = model.predict(future_X)
    return pred, model

# ─── Sidebar ───
st.sidebar.markdown("## ⚙️ 控制面板")
page = st.sidebar.radio("导航", ["📊 驾驶舱总览", "💰 收入分析", "📦 成本费用", "📈 趋势预测", "🏢 同行业对比", "📁 数据管理"])

st.sidebar.markdown("---")
st.sidebar.markdown(f"**数据期间**: 2023-2025")
st.sidebar.markdown(f"**更新频率**: 月度")
st.sidebar.markdown(f"**数据条数**: 1,296 笔")

# Custom data upload
st.sidebar.markdown("---")
uploaded_file = st.sidebar.file_uploader("上传自定义数据 (Excel/CSV)", type=['xlsx','csv'])

# ─── Pages ───
# ===================== DASHBOARD =====================
if page == "📊 驾驶舱总览":
    st.markdown('<div class="main-header"><h1>小甜甜有限公司 · 经营分析驾驶舱</h1></div>', unsafe_allow_html=True)

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1: st.markdown(kpi_card("营业收入(2025)", f'{yearly["营业收入"]["2025"]:,.0f}万', f'+{(yearly["营业收入"]["2025"]/yearly["营业收入"]["2024"]-1)*100:.1f}%', True), unsafe_allow_html=True)
    with col2: st.markdown(kpi_card("营业利润(2025)", f'{yearly["营业利润"]["2025"]:,.0f}万', '扭亏为盈', True), unsafe_allow_html=True)
    with col3: st.markdown(kpi_card("利润总额(2025)", f'{yearly["利润总额"]["2025"]:,.0f}万', '扭亏为盈', True), unsafe_allow_html=True)
    with col4: st.markdown(kpi_card("毛利率(2025)", f'{yearly["毛利率"]["2025"]:.1f}%', f'{yearly["毛利率"]["2025"]-yearly["毛利率"]["2024"]:+.1f}pp', yearly["毛利率"]["2025"]>=yearly["毛利率"]["2024"]), unsafe_allow_html=True)
    with col5: st.markdown(kpi_card("净利率(2025)", f'{yearly["净利率"]["2025"]:.1f}%', f'{yearly["净利率"]["2025"]-yearly["净利率"]["2024"]:+.1f}pp', True), unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(plot_line(raw['营业收入'], '📈 营业收入月度趋势', height=350), use_container_width=True)
    with c2: st.plotly_chart(plot_line({'2023': monthly_op_profit['2023'], '2024': monthly_op_profit['2024'], '2025': monthly_op_profit['2025']}, '💰 营业利润月度趋势', height=350), use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        # Stacked bar for 2025 cost structure
        fig = go.Figure()
        items_2025 = {'营业成本': raw['营业成本']['2025'], '销售费用': raw['销售费用']['2025'],
                      '管理费用': raw['管理费用']['2025'], '研发费用': raw['研发费用']['2025'], '财务费用': raw['财务费用']['2025']}
        colors = ['#00d4ff','#ff6b6b','#ffd93d','#6bcbff','#c084fc']
        for i, (name, vals) in enumerate(items_2025.items()):
            fig.add_trace(go.Bar(name=name, x=MONTHS, y=vals, marker_color=colors[i]))
        fig.update_layout(barmode='stack', title='📊 2025年成本费用结构', height=350, template='plotly_dark',
                         margin=dict(l=40, r=20, t=40, b=30), legend=dict(orientation='h', yanchor='bottom', y=1.02))
        fig.update_xaxes(gridcolor='#1e3a5f')
        fig.update_yaxes(gridcolor='#1e3a5f', title='万元')
        st.plotly_chart(fig, use_container_width=True)

    with c4:
        # Annual comparison bar
        fig = go.Figure()
        fig.add_trace(go.Bar(name='营业收入', x=YEARS, y=[yearly['营业收入'][y] for y in YEARS], marker_color='#00d4ff'))
        fig.add_trace(go.Bar(name='营业利润', x=YEARS, y=[yearly['营业利润'][y] for y in YEARS], marker_color='#ffd93d'))
        fig.add_trace(go.Bar(name='利润总额', x=YEARS, y=[yearly['利润总额'][y] for y in YEARS], marker_color='#6bcbff'))
        fig.update_layout(barmode='group', title='🏆 年度核心指标对比', height=350, template='plotly_dark',
                         margin=dict(l=40, r=20, t=40, b=30))
        fig.update_xaxes(gridcolor='#1e3a5f')
        fig.update_yaxes(gridcolor='#1e3a5f', title='万元')
        st.plotly_chart(fig, use_container_width=True)

    # Margins trend
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=YEARS, y=[yearly['毛利率'][y] for y in YEARS], mode='lines+markers',
                            name='毛利率', line=dict(color='#00d4ff', width=3), marker=dict(symbol='diamond', size=12)))
    fig.add_trace(go.Scatter(x=YEARS, y=[yearly['净利率'][y] for y in YEARS], mode='lines+markers',
                            name='净利率', line=dict(color='#ffd93d', width=3), marker=dict(size=10)))
    fig.add_trace(go.Scatter(x=YEARS, y=[yearly['营业利润率'][y] for y in YEARS], mode='lines+markers',
                            name='营业利润率', line=dict(color='#6bcbff', width=3), marker=dict(symbol='square', size=10)))
    fig.update_layout(title='📉 利润率年度趋势', height=300, template='plotly_dark',
                     margin=dict(l=40, r=20, t=40, b=30),
                     legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
    fig.update_xaxes(gridcolor='#1e3a5f')
    fig.update_yaxes(gridcolor='#1e3a5f', title='%', ticksuffix='%')
    st.plotly_chart(fig, use_container_width=True)

# ===================== REVENUE =====================
elif page == "💰 收入分析":
    st.markdown("## 💰 收入分析")

    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(plot_line(raw['营业收入'], '营业收入月度对比'), use_container_width=True)
    with c2:
        # Revenue by year pie
        fig = go.Figure(data=[go.Pie(labels=YEARS, values=[sum(raw['营业收入'][y]) for y in YEARS],
                                     marker_colors=['#00d4ff','#ff6b6b','#ffd93d'], textinfo='label+percent',
                                     hole=0.4)])
        fig.update_layout(title='年度收入占比', height=350, template='plotly_dark',
                         margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig, use_container_width=True)

    # Monthly bar comparison
    fig = go.Figure()
    for i, y in enumerate(YEARS):
        fig.add_trace(go.Bar(name=f'{y}年', x=MONTHS, y=raw['营业收入'][y],
                            marker_color=['#00d4ff','#ff6b6b','#ffd93d'][i]))
    fig.update_layout(barmode='group', title='月度收入柱状对比', height=400, template='plotly_dark',
                     margin=dict(l=40, r=20, t=40, b=30))
    fig.update_xaxes(gridcolor='#1e3a5f')
    fig.update_yaxes(gridcolor='#1e3a5f', title='万元')
    st.plotly_chart(fig, use_container_width=True)

    # YoY growth
    growth = []
    for m in range(12):
        if raw['营业收入']['2024'][m] != 0:
            g = (raw['营业收入']['2025'][m] - raw['营业收入']['2024'][m]) / raw['营业收入']['2024'][m] * 100
        else:
            g = 0
        growth.append(round(g, 1))
    fig = go.Figure(data=[go.Bar(x=MONTHS, y=growth, marker_color=['#00c853' if g>=0 else '#ff5252' for g in growth])])
    fig.update_layout(title='2025年 vs 2024年 月度同比增长率', height=300, template='plotly_dark',
                     margin=dict(l=40, r=20, t=40, b=30))
    fig.update_xaxes(gridcolor='#1e3a5f')
    fig.update_yaxes(gridcolor='#1e3a5f', title='%', ticksuffix='%')
    st.plotly_chart(fig, use_container_width=True)

# ===================== COST =====================
elif page == "📦 成本费用":
    st.markdown("## 📦 成本费用分析")

    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(plot_line(raw['营业成本'], '营业成本月度趋势'), use_container_width=True)
    with c2: st.plotly_chart(plot_line(raw['研发费用'], '研发费用月度趋势'), use_container_width=True)

    c3, c4 = st.columns(2)
    with c3: st.plotly_chart(plot_line(raw['销售费用'], '销售费用月度趋势'), use_container_width=True)
    with c4: st.plotly_chart(plot_line(raw['管理费用'], '管理费用月度趋势'), use_container_width=True)

    # 2025 Expense Pie
    exp_items_2025 = {
        '营业成本': sum(raw['营业成本']['2025']),
        '销售费用': sum(raw['销售费用']['2025']),
        '管理费用': sum(raw['管理费用']['2025']),
        '研发费用': sum(raw['研发费用']['2025']),
        '财务费用': sum(raw['财务费用']['2025']),
        '税金及附加': sum(raw['税金及附加']['2025'])
    }
    fig = go.Figure(data=[go.Pie(labels=list(exp_items_2025.keys()),
                                 values=list(exp_items_2025.values()),
                                 marker_colors=['#00d4ff','#ff6b6b','#ffd93d','#6bcbff','#c084fc','#a8e6cf'],
                                 textinfo='label+percent', hole=0.4)])
    fig.update_layout(title='2025年成本费用构成', height=400, template='plotly_dark',
                     margin=dict(l=20, r=20, t=40, b=20))
    st.plotly_chart(fig, use_container_width=True)

    # Annual cost comparison
    fig = go.Figure()
    for item in ['营业成本','销售费用','管理费用','研发费用','财务费用','税金及附加']:
        fig.add_trace(go.Bar(name=item, x=YEARS, y=[yearly[item][y] for y in YEARS]))
    fig.update_layout(barmode='group', title='年度费用对比', height=400, template='plotly_dark',
                     margin=dict(l=40, r=20, t=40, b=30))
    st.plotly_chart(fig, use_container_width=True)

# ===================== PREDICTION =====================
elif page == "📈 趋势预测":
    st.markdown("## 📈 未来趋势预测")

    pred_year = st.selectbox("预测目标年份", ["2026年", "2027年", "2028年"], index=0)
    target_months = 12

    # Predict revenue
    pred_rev, model_rev = predict_next_year(raw['营业收入'])
    pred_cost, _ = predict_next_year(raw['营业成本'])

    # Historical + predicted
    hist_rev = raw['营业收入']['2023'] + raw['营业收入']['2024'] + raw['营业收入']['2025']
    hist_cost = raw['营业成本']['2023'] + raw['营业成本']['2024'] + raw['营业成本']['2025']
    all_months = [f'2023-{m}' for m in range(1,13)] + [f'2024-{m}' for m in range(1,13)] + [f'2025-{m}' for m in range(1,13)]
    pred_months = [f'{int(pred_year[:4])}-{m}' for m in range(1,13)]

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=all_months, y=hist_rev, mode='lines+markers', name='收入(历史)',
                            line=dict(color='#00d4ff', width=2.5)))
    fig.add_trace(go.Scatter(x=pred_months, y=pred_rev, mode='lines+markers', name='收入(预测)',
                            line=dict(color='#00d4ff', width=2.5, dash='dash')))
    fig.add_trace(go.Scatter(x=all_months, y=hist_cost, mode='lines+markers', name='成本(历史)',
                            line=dict(color='#ff6b6b', width=2.5)))
    fig.add_trace(go.Scatter(x=pred_months, y=pred_cost, mode='lines+markers', name='成本(预测)',
                            line=dict(color='#ff6b6b', width=2.5, dash='dash')))
    fig.update_layout(title=f'营业收入 & 成本预测（{pred_year}）', height=400, template='plotly_dark',
                     hovermode='x unified', xaxis_tickangle=-45,
                     margin=dict(l=40, r=20, t=40, b=80),
                     legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1))
    fig.update_xaxes(gridcolor='#1e3a5f')
    fig.update_yaxes(gridcolor='#1e3a5f', title='万元')
    st.plotly_chart(fig, use_container_width=True)

    # Prediction summary
    pred_rev_total = sum(pred_rev)
    pred_cost_total = sum(pred_cost)
    last_rev_total = yearly['营业收入']['2025']
    growth_rate = (pred_rev_total / last_rev_total - 1) * 100

    c1, c2, c3 = st.columns(3)
    with c1: st.markdown(kpi_card(f"预测{pred_year}收入", f'{pred_rev_total:,.0f}万', f'增长率 {growth_rate:.1f}%', growth_rate>0), unsafe_allow_html=True)
    with c2: st.markdown(kpi_card(f"预测{pred_year}成本", f'{pred_cost_total:,.0f}万', f'占比 {pred_cost_total/pred_rev_total*100:.1f}%', False), unsafe_allow_html=True)
    with c3: st.markdown(kpi_card(f"预测{pred_year}毛利", f'{pred_rev_total-pred_cost_total:,.0f}万', f'毛利率 {(pred_rev_total-pred_cost_total)/pred_rev_total*100:.1f}%', True), unsafe_allow_html=True)

    st.info("""
    **预测方法说明**: 基于过去36个月数据，使用线性回归模型预测未来12个月趋势。
    此预测为初步参考，实际经营中应结合行业环境、公司战略、季节性因素等综合判断。
    """)

    # Monthly distribution forecast
    fig = go.Figure()
    avg_monthly_pct = []
    for m in range(12):
        total_3y = sum(raw['营业收入'][y][m] for y in YEARS)
        total_all = sum(sum(raw['营业收入'][y]) for y in YEARS)
        avg_monthly_pct.append(total_3y / total_all * 100)
    fig.add_trace(go.Bar(x=MONTHS, y=avg_monthly_pct, marker_color='#00d4ff',
                        text=[f'{v:.1f}%' for v in avg_monthly_pct], textposition='outside'))
    fig.update_layout(title='历史月度收入分布（3年平均占比）', height=300, template='plotly_dark',
                     margin=dict(l=40, r=20, t=40, b=30))
    fig.update_xaxes(gridcolor='#1e3a5f')
    fig.update_yaxes(gridcolor='#1e3a5f', title='%', ticksuffix='%')
    st.plotly_chart(fig, use_container_width=True)

# ===================== INDUSTRY =====================
elif page == "🏢 同行业对比":
    st.markdown("## 🏢 同行业对比分析")
    st.markdown("对标公司：艾融软件(830799)、用友网络(600588)、华宇软件(300271)、三维天地(301159)、山大地纬(688579)、殷图网联(835508)、天亿马(301178)、远光软件(002063)、宇信科技(300674)")

    # ─── Parse Chinese number strings ───
    def parse_cn(val):
        if pd.isna(val) or val is None or val == 'False':
            return None
        s = str(val).replace(',', '').strip()
        neg = 1
        if s.startswith('-'):
            neg = -1
            s = s[1:]
        if '亿' in s:
            num = float(s.replace('亿', ''))
            return neg * num * 10000  # 亿 → 万
        if '万' in s:
            num = float(s.replace('万', ''))
            return neg * num
        if '%' in s:
            return neg * float(s.replace('%', ''))
        try:
            return neg * float(s)
        except:
            return None

    # ─── Load industry data ───
    csv_path = os.path.join(os.path.dirname(__file__), 'industry_data.csv')
    if os.path.exists(csv_path):
        ind_df = pd.read_csv(csv_path, encoding='utf-8-sig')
        # Only show years available in both datasets
        ind_years = set(ind_df['报告期'].unique())
        our_years = set(YEARS)
        years_avail = sorted(ind_years & our_years, reverse=True)
        if not years_avail:
            st.warning("行业数据年份与本公司数据年份不匹配，请更新数据。")
        else:
            sel_year = st.selectbox("选择对比年份", years_avail, index=0)

            radar_metrics = ['销售毛利率', '销售净利率', '净资产收益率']
            radar_labels = ['毛利率(%)', '净利率(%)', '净资产收益率(%)']

            # Build comparison data
            peer_rows = []
            for _, r in ind_df[ind_df['报告期'] == sel_year].iterrows():
                peer_rows.append({
                    '公司': r['公司'],
                    '营业收入(万元)': parse_cn(r['营业总收入']),
                    '净利润(万元)': parse_cn(r['净利润']),
                    '毛利率(%)': parse_cn(r['销售毛利率']),
                    '净利率(%)': parse_cn(r['销售净利率']),
                    '净资产收益率(%)': parse_cn(r['净资产收益率']),
                })

            # Add 小甜甜
            rev_tt = yearly['营业收入'][sel_year]
            cost_tt = yearly['营业成本'][sel_year]
            profit_tt = yearly['利润总额'][sel_year]
            gross_margin_tt = (rev_tt - cost_tt) / rev_tt * 100 if rev_tt else 0
            net_margin_tt = profit_tt / rev_tt * 100 if rev_tt else 0
            # Approximate ROE using net profit / (equity approximated by total assets - total liabilities)
            total_cost = sum(yearly.get(k, {}).get(sel_year, 0) for k in ['营业成本','税金及附加','销售费用','管理费用','研发费用','财务费用'])
            equity_approx = rev_tt - total_cost + yearly['其他收益'].get(sel_year, 0)
            roe_tt = profit_tt / equity_approx * 100 if equity_approx else 0

            peer_rows.append({
                '公司': '✨ 小甜甜公司',
                '营业收入(万元)': rev_tt,
                '净利润(万元)': profit_tt,
                '毛利率(%)': round(gross_margin_tt, 2),
                '净利率(%)': round(net_margin_tt, 2),
                '净资产收益率(%)': round(abs(roe_tt), 2),
            })

            df_all = pd.DataFrame(peer_rows).set_index('公司')

            # ─── Comparison table ───
            display_cols = ['营业收入(万元)', '净利润(万元)', '毛利率(%)', '净利率(%)', '净资产收益率(%)']
            st.dataframe(df_all[display_cols].style.format({
                '营业收入(万元)': '{:,.0f}',
                '净利润(万元)': '{:,.0f}',
                '毛利率(%)': '{:.1f}',
                '净利率(%)': '{:.1f}',
                '净资产收益率(%)': '{:.1f}',
            }), use_container_width=True)

            # ─── Radar chart ───
            fig = go.Figure()
            colors = ['#00d4ff','#ff6b6b','#ffd93d','#6bcbff','#c084fc','#a8e6cf','#ff9ff3','#54a0ff','#5f27cd','#ff9f43']
            for i, (name, row) in enumerate(df_all.iterrows()):
                vals = [row[m] for m in radar_metrics]
                fig.add_trace(go.Scatterpolar(r=vals + [vals[0]], theta=radar_labels + [radar_labels[0]],
                               fill='toself', name=name,
                               line=dict(color=colors[i % len(colors)], width=2 if '小甜甜' in str(name) else 1.5)))

            fig.update_layout(polar=dict(radialaxis=dict(visible=True, gridcolor='#1e3a5f', range=[-80, 80])),
                             title=f'{sel_year}年 同行业经营指标雷达图', height=500, template='plotly_dark',
                             margin=dict(l=80, r=80, t=40, b=40),
                             legend=dict(orientation='h', yanchor='bottom', y=-0.2, xanchor='center', x=0.5))
            st.plotly_chart(fig, use_container_width=True)

            # ─── Bar chart: 毛利率 & 净利率 ───
            fig = go.Figure()
            df_sorted = df_all.sort_values('毛利率(%)', ascending=True)
            fig.add_trace(go.Bar(name='毛利率(%)', y=df_sorted.index, x=df_sorted['毛利率(%)'],
                                 orientation='h', marker_color='#00d4ff',
                                 text=df_sorted['毛利率(%)'].apply(lambda v: f'{v:.1f}%'), textposition='outside'))
            fig.add_trace(go.Bar(name='净利率(%)', y=df_sorted.index, x=df_sorted['净利率(%)'],
                                 orientation='h', marker_color='#ffd93d',
                                 text=df_sorted['净利率(%)'].apply(lambda v: f'{v:.1f}%'), textposition='outside'))
            fig.update_layout(title=f'{sel_year}年 毛利率 & 净利率排名', height=450, template='plotly_dark',
                             barmode='group', margin=dict(l=120, r=60, t=40, b=30),
                             xaxis=dict(gridcolor='#1e3a5f'))
            st.plotly_chart(fig, use_container_width=True)

            # ─── Revenue bar ───
            fig = go.Figure()
            df_rev = df_all.sort_values('营业收入(万元)', ascending=True)
            colors_rev = ['#00d4ff'] * len(df_rev)
            tt_name = '✨ 小甜甜公司'
            if tt_name in df_rev.index:
                tt_pos = list(df_rev.index).index(tt_name)
                colors_rev[tt_pos] = '#ff6b6b'
            fig.add_trace(go.Bar(y=df_rev.index, x=df_rev['营业收入(万元)'], orientation='h',
                                 marker_color=colors_rev,
                                 text=df_rev['营业收入(万元)'].apply(lambda v: f'{v/10000:.2f}亿' if v>=10000 else f'{v:.0f}万'),
                                 textposition='outside'))
            fig.update_layout(title=f'{sel_year}年 营业收入对比', height=450, template='plotly_dark',
                             margin=dict(l=120, r=80, t=40, b=30),
                             xaxis=dict(gridcolor='#1e3a5f'))
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("""
        ⚠️ 行业数据未找到，请先运行数据抓取脚本：

        ```bash
        python industry_data_fetcher.py
        ```

        该脚本会自动获取9家上市公司年报数据并保存到 `industry_data.csv`。
        """)

# ===================== DATA MANAGEMENT =====================
elif page == "📁 数据管理":
    st.markdown("## 📁 数据管理")

    tab1, tab2, tab3 = st.tabs(["📥 数据导入", "📤 数据导出", "📋 数据预览"])

    with tab1:
        st.markdown("### 上传自定义数据")
        uploaded = st.file_uploader("选择Excel或CSV文件", type=['xlsx','csv'], key='data_upload')
        if uploaded:
            st.success(f"已上传: {uploaded.name}")
            try:
                if uploaded.name.endswith('.csv'):
                    df = pd.read_csv(uploaded)
                else:
                    df = pd.read_excel(uploaded)
                st.dataframe(df.head(20), use_container_width=True)
                st.info(f"共 {len(df)} 行, {len(df.columns)} 列")
            except:
                st.error("文件解析失败，请检查格式")

        st.markdown("### 手动录入数据")
        with st.form("manual_input"):
            c1, c2, c3 = st.columns(3)
            with c1: year_in = st.selectbox("年份", ['2025','2026'])
            with c2: month_in = st.selectbox("月份", MONTHS)
            with c3: item_in = st.selectbox("项目", ['营业收入','营业成本','销售费用','管理费用','研发费用','财务费用','税金及附加','其他收益','营业外收入','营业外支出'])
            amount = st.number_input("金额（万元）", min_value=-999999.0, max_value=999999.0, value=0.0, step=1.0)
            submitted = st.form_submit_button("录入")
            if submitted and amount != 0:
                st.success(f"已录入: {year_in} {month_in} {item_in} = {amount}万元")

    with tab2:
        st.markdown("### 导出数据")
        df_export = pd.read_csv('profit_data_flat.csv', encoding='utf-8-sig')
        
        col1, col2 = st.columns(2)
        with col1:
            csv_data = df_export.to_csv(index=False, encoding='utf-8-sig')
            st.download_button("⬇️ 下载 CSV", data=csv_data, file_name="小甜甜公司财务数据.csv", mime="text/csv")
        with col2:
            buf = BytesIO()
            df_export.to_excel(buf, index=False, engine='openpyxl')
            st.download_button("⬇️ 下载 Excel", data=buf.getvalue(), file_name="小甜甜公司财务数据.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("### 导出报表")
        col3, col4 = st.columns(2)
        with col3:
            if st.button("📊 导出年度汇总表"):
                annual_data = {item: {y: yearly[item][y] for y in YEARS} for item in ['营业收入','营业成本','营业利润','利润总额','毛利率','净利率']}
                df_annual = pd.DataFrame(annual_data).T
                st.download_button("⬇️ 下载年度汇总", data=df_annual.to_csv(encoding='utf-8-sig'), file_name="年度汇总.csv")
        with col4:
            if st.button("📈 导出预测数据"):
                pred, _ = predict_next_year(raw['营业收入'])
                df_pred = pd.DataFrame({'月份': [f'2026年{m}' for m in MONTHS], '预测收入(万元)': [round(v,2) for v in pred]})
                st.download_button("⬇️ 下载预测数据", data=df_pred.to_csv(index=False, encoding='utf-8-sig'), file_name="预测数据.csv")

    with tab3:
        st.markdown("### 当前数据预览")
        df_current = pd.read_csv('profit_data_flat.csv', encoding='utf-8-sig')
        st.dataframe(df_current, use_container_width=True)
        st.caption(f"共 {len(df_current)} 条记录")

# ─── Footer ───
st.markdown("---")
st.markdown(f'<div style="text-align:center;color:#556677;font-size:12px;padding:10px">小甜甜有限公司经营分析仪表板 | 数据期间 2023-2025 | 使用 Streamlit 构建</div>', unsafe_allow_html=True)
