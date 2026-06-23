# -*- coding: utf-8 -*-
"""生成 Excel 版财务经营分析仪表板"""
import sys; sys.stdout.reconfigure(encoding='utf-8')
import xlsxwriter
import os

OUTPUT = os.path.join(os.path.dirname(__file__), '财务仪表板.xlsx')
wb = xlsxwriter.Workbook(OUTPUT)

# ─── Raw data ───
YEARS = ['2023','2024','2025']
MONTHS = [f'{m}月' for m in range(1,13)]

raw = {
    '营业收入': {'2023': [253.09,446.68,775.09,648.24,1786.23,1755.81,357.95,436.83,1345.30,858.29,779.26,3526.04],
                '2024': [587.21,168.75,1000.22,520.51,810.84,905.44,360.81,658.94,891.90,560.09,1278.37,2541.21],
                '2025': [215.08,395.76,755.70,1306.49,532.65,987.88,911.27,513.13,1537.02,994.58,2685.46,6346.35]},
    '营业成本': {'2023': [154.14,229.57,507.67,377.35,970.49,733.56,316.10,303.49,812.93,434.55,608.08,2431.25],
                '2024': [273.74,50.63,584.77,252.38,495.51,620.96,147.71,239.55,468.93,164.63,824.68,1536.40],
                '2025': [123.78,219.18,238.93,657.76,224.38,622.11,527.72,132.41,954.45,523.71,1707.95,4324.33]},
    '销售费用': {'2023': [119.98,56.93,57.21,76.94,62.80,68.80,78.66,91.39,87.07,232.68,103.05,143.10],
                '2024': [119.49,81.30,106.68,78.05,71.51,134.31,74.14,65.92,71.28,198.91,121.00,144.36],
                '2025': [95.77,86.23,130.61,64.13,96.78,83.39,81.12,67.32,167.72,120.02,207.60,123.68]},
    '管理费用': {'2023': [84.02,118.24,98.96,71.55,104.13,89.04,95.33,72.86,99.20,71.05,80.79,149.08],
                '2024': [93.91,167.33,118.77,111.85,93.92,113.83,113.12,91.12,137.53,117.68,120.53,291.14],
                '2025': [91.73,119.49,125.10,84.25,166.28,88.19,62.01,62.53,195.66,92.67,89.36,89.86]},
    '研发费用': {'2023': [156.30,230.42,171.57,165.37,293.52,333.64,218.27,181.21,177.96,98.45,95.77,118.39],
                '2024': [165.45,181.97,178.42,121.00,141.86,139.84,193.07,161.07,110.42,271.13,276.01,188.89],
                '2025': [264.18,254.03,170.89,155.69,204.84,242.53,228.55,169.34,202.98,122.15,388.43,117.94]},
    '财务费用': {'2023': [4.12,-0.28,4.29,1.23,1.13,0.88,1.36,3.51,-0.10,1.07,2.48,3.54],
                '2024': [1.28,0.90,1.43,0.73,0.96,5.20,1.03,2.53,2.62,0.91,0.82,3.77],
                '2025': [0.80,1.03,5.90,0.93,0.93,4.89,0.55,2.76,2.13,0.58,0.52,4.45]},
    '税金及附加': {'2023': [2.96,4.45,7.68,8.84,11.66,8.64,8.52,11.83,8.06,14.30,9.12,15.67],
                   '2024': [6.05,8.98,5.42,8.19,7.53,14.24,7.46,7.70,16.53,7.93,10.21,27.99],
                   '2025': [2.87,5.81,7.33,6.23,18.53,8.33,31.90,9.30,12.17,5.59,20.70,35.16]},
    '其他收益': {'2023': [83.68,10.35,5.84,46.90,54.00,0.32,85.54,76.45,40.31,0.07,117.91,42.69],
                '2024': [67.57,0,56.29,20.55,38.39,38.51,69.29,32.04,39.57,0,0,125.84],
                '2025': [142.85,13.79,43.15,24.36,25.13,84.13,41.06,47.23,28.92,64.76,16.16,101.84]},
}

# ─── Styles ───
title_fmt = wb.add_format({'bold':True,'font_size':16,'font_color':'#FFFFFF','bg_color':'#0f1923','align':'center','valign':'vcenter','border':0})
kpi_label_fmt = wb.add_format({'font_size':9,'font_color':'#8899aa','align':'center','bg_color':'#1a2a3a','border':1,'border_color':'#1e3a5f','top':1,'left':1,'right':1})
kpi_value_fmt = wb.add_format({'font_size':22,'font_color':'#e0e6ed','align':'center','bg_color':'#1a2a3a','border':1,'border_color':'#1e3a5f','num_format':'#,##0'})
kpi_pct_fmt = wb.add_format({'font_size':22,'font_color':'#e0e6ed','align':'center','bg_color':'#1a2a3a','border':1,'border_color':'#1e3a5f','num_format':'0.0%'})
kpi_change_up = wb.add_format({'font_size':10,'font_color':'#00c853','align':'center','bg_color':'#1a2a3a','border':1,'border_color':'#1e3a5f'})
kpi_change_dn = wb.add_format({'font_size':10,'font_color':'#ff5252','align':'center','bg_color':'#1a2a3a','border':1,'border_color':'#1e3a5f'})
section_fmt = wb.add_format({'bold':True,'font_size':13,'font_color':'#00d4ff','bottom':1,'bottom_color':'#1e3a5f','bg_color':'#0f1923'})
data_val = wb.add_format({'font_size':10,'font_color':'#e0e6ed','bg_color':'#0f1923','num_format':'#,##0.00','border':1,'border_color':'#1e3a5f'})
data_int = wb.add_format({'font_size':10,'font_color':'#e0e6ed','bg_color':'#0f1923','num_format':'#,##0','border':1,'border_color':'#1e3a5f'})
data_pct = wb.add_format({'font_size':10,'font_color':'#e0e6ed','bg_color':'#0f1923','num_format':'0.0%','border':1,'border_color':'#1e3a5f'})
header_fmt = wb.add_format({'bold':True,'font_size':10,'font_color':'#00d4ff','bg_color':'#1a2a3a','border':1,'border_color':'#1e3a5f','align':'center','text_wrap':True})
label_fmt = wb.add_format({'font_size':10,'font_color':'#8899aa','bg_color':'#0f1923','border':1,'border_color':'#1e3a5f'})
total_fmt = wb.add_format({'bold':True,'font_size':10,'font_color':'#ffd93d','bg_color':'#0f1923','num_format':'#,##0.00','border':1,'border_color':'#1e3a5f'})
note_fmt = wb.add_format({'font_size':8,'font_color':'#556677','italic':True,'bg_color':'#0f1923'})

# ─── Helper: yearly totals & margins ───
def sum_year(item, y): return sum(raw[item][y])
def yearly_data():
    yd = {}
    for item in raw:
        yd[item] = {y: sum_year(item, y) for y in YEARS}
    yd['营业利润'] = {}
    yd['利润总额'] = {}
    for y in YEARS:
        rev = yd['营业收入'][y]
        cost = yd['营业成本'][y]
        tax = yd['税金及附加'][y]
        sell = yd['销售费用'][y]
        admin = yd['管理费用'][y]
        rd = yd['研发费用'][y]
        fin = yd['财务费用'][y]
        other = yd['其他收益'][y]
        op = rev - cost - tax - sell - admin - rd - fin + other
        yd['营业利润'][y] = round(op, 2)
        yd['利润总额'][y] = round(op, 2)  # no 营业外收支
    for y in YEARS:
        r = yd['营业收入'][y]
        yd.setdefault('毛利率',{})[y] = round((r-yd['营业成本'][y])/r*100 if r else 0, 2)
        yd.setdefault('净利率',{})[y] = round(yd['利润总额'][y]/r*100 if r else 0, 2)
        yd.setdefault('营业利润率',{})[y] = round(yd['营业利润'][y]/r*100 if r else 0, 2)
    return yd

yd = yearly_data()

# ═══════════════════════════════════════
# SHEET 1: 数据源
# ═══════════════════════════════════════
ws_data = wb.add_worksheet('数据源')
ws_data.hide()
ws_data.set_tab_color('#1e3a5f')
ws_data.write('A1', '项目', header_fmt)
ws_data.write('B1', '年份', header_fmt)
for mi, m in enumerate(MONTHS, 3):
    col_letter = xlsxwriter.utility.xl_col_to_name(mi-1)
    ws_data.write(f'{col_letter}1', m, header_fmt)
ws_data.set_column('A:A', 14)
ws_data.set_column('B:B', 8)
ws_data.set_column('C:N', 11)

row = 1
for item in ['营业收入','营业成本','销售费用','管理费用','研发费用','财务费用','税金及附加','其他收益']:
    for y in YEARS:
        ws_data.write(row, 0, item, label_fmt)
        ws_data.write(row, 1, y, data_int)
        for mi in range(12):
            ws_data.write(row, mi+2, raw[item][y][mi], data_val)
        row += 1

# ═══════════════════════════════════════
# SHEET 2: 驾驶舱总览
# ═══════════════════════════════════════
ws = wb.add_worksheet('驾驶舱总览')
ws.set_tab_color('#00d4ff')
ws.hide_gridlines(2)
ws.set_landscape()
ws.set_paper(9)  # A4
ws.set_margins(0.3, 0.3, 0.3, 0.3)
ws.set_column('A:E', 22)

# Background
bg = wb.add_format({'bg_color':'#0f1923'})
ws.merge_range('A1:E1', '', bg)
ws.merge_range('A2:E2', '', bg)

# Title
title_row = 2
ws.merge_range(title_row, 0, title_row, 4, '小甜甜有限公司 · 经营分析驾驶舱', title_fmt)
ws.set_row(title_row, 40)

# KPI Row
kpi_row = 4
kpi_items = [
    ('营业收入(2025)', f'{yd["营业收入"]["2025"]:,.0f}万', f'+{(yd["营业收入"]["2025"]/yd["营业收入"]["2024"]-1)*100:.1f}%', True),
    ('营业利润(2025)', f'{yd["营业利润"]["2025"]:,.0f}万', '扭亏为盈', True),
    ('利润总额(2025)', f'{yd["利润总额"]["2025"]:,.0f}万', '扭亏为盈', True),
    ('毛利率(2025)', f'{yd["毛利率"]["2025"]:.1f}%', f'{yd["毛利率"]["2025"]-yd["毛利率"]["2024"]:+.1f}pp', yd["毛利率"]["2025"]>=yd["毛利率"]["2024"]),
    ('净利率(2025)', f'{yd["净利率"]["2025"]:.1f}%', f'{yd["净利率"]["2025"]-yd["净利率"]["2024"]:+.1f}pp', True),
]

kpi_bg = wb.add_format({'bg_color':'#1a2a3a','border':1,'border_color':'#1e3a5f','text_wrap':True,'valign':'vcenter'})
kpi_val2 = wb.add_format({'bg_color':'#1a2a3a','border':1,'border_color':'#1e3a5f','font_size':22,'font_color':'#e0e6ed','align':'center','valign':'vcenter'})
kpi_up2 = wb.add_format({'bg_color':'#1a2a3a','border':1,'border_color':'#1e3a5f','font_size':10,'font_color':'#00c853','align':'center'})
kpi_dn2 = wb.add_format({'bg_color':'#1a2a3a','border':1,'border_color':'#1e3a5f','font_size':10,'font_color':'#ff5252','align':'center'})
kpi_ttl2 = wb.add_format({'bg_color':'#1a2a3a','border':1,'border_color':'#1e3a5f','font_size':9,'font_color':'#8899aa','align':'center','valign':'vcenter'})

ws.set_row(kpi_row, 18)
ws.set_row(kpi_row+1, 30)
ws.set_row(kpi_row+2, 18)
for ci, (label, val, chg, up) in enumerate(kpi_items):
    ws.merge_range(kpi_row, ci, kpi_row+2, ci, '', kpi_bg)
    ws.write(kpi_row, ci, label, kpi_ttl2)
    ws.write(kpi_row+1, ci, val, kpi_val2)
    ws.write(kpi_row+2, ci, chg, kpi_up2 if up else kpi_dn2)

# Revenue line chart
chart_row = 9
ws.write(chart_row, 0, '营业收入月度趋势', section_fmt)
ws.merge_range(chart_row, 0, chart_row, 4, '营业收入月度趋势', section_fmt)

rev_chart = wb.add_chart({'type':'line'})
for yi, y in enumerate(YEARS):
    rev_chart.add_series({
        'name': f'{y}年',
        'categories': ['数据源', 1, 2, 1, 13],
        'values': ['数据源', yi+1, 2, yi+1, 13],
        'line': {'color': ['#00d4ff','#ff6b6b','#ffd93d'][yi], 'width': 2.5},
        'marker': {'type':'circle','size':5},
    })
rev_chart.set_title({'name':'营业收入月度趋势','name_font':{'color':'#00d4ff','size':12}})
rev_chart.set_y_axis({'name':'万元','name_font':{'color':'#8899aa'},'gridlines':{'visible':False}})
rev_chart.set_legend({'font':{'color':'#8899aa'}})
rev_chart.set_size({'width':500,'height':280})
ws.insert_chart(chart_row+1, 0, rev_chart)

# Annual comparison bar
ann_chart = wb.add_chart({'type':'column'})
for yi, y in enumerate(YEARS):
    ann_chart.add_series({
        'name': f'{y}年',
        'categories': ['数据源', 8, 0, 8, 0],
        'values': ['数据源', 8, 0, 8, 0],
    })
# Redo properly
ann_chart = wb.add_chart({'type':'column'})
cats = ['营业收入','营业利润','利润总额']
colors_ann = ['#00d4ff','#ffd93d','#6bcbff']
for ci, cat in enumerate(cats):
    vals = [yd[cat][y] for y in YEARS]
    ann_chart.add_series({
        'name': cat,
        'categories': ['驾驶舱总览', 100, 100, 100, 100],  # dummy, will fix below
        'values': ['驾驶舱总览', 100, 100, 100, 100],
    })

# Actually let me write yearly data to a small table and chart from it
yr_row = chart_row + 18
ws.merge_range(yr_row, 0, yr_row, 4, '年度核心指标', section_fmt)
yr_row += 1
ws.write(yr_row, 0, '指标', header_fmt)
for yi, y in enumerate(YEARS):
    ws.write(yr_row, yi+1, y, header_fmt)
yr_row += 1
for cat in ['营业收入','营业利润','利润总额']:
    ws.write(yr_row, 0, cat, label_fmt)
    for yi, y in enumerate(YEARS):
        ws.write(yr_row, yi+1, yd[cat][y], data_int)
    yr_row += 1

ann_chart = wb.add_chart({'type':'column'})
for yi, y in enumerate(YEARS):
    ann_chart.add_series({
        'name': y,
        'categories': ['驾驶舱总览', yr_row-3, 0, yr_row-1, 0],
        'values': ['驾驶舱总览', yr_row-3, yi+1, yr_row-1, yi+1],
        'fill': {'color': ['#00d4ff','#ff6b6b','#ffd93d'][yi]},
    })
ann_chart.set_title({'name':'年度核心指标对比','name_font':{'color':'#ffd93d','size':12}})
ann_chart.set_y_axis({'name':'万元','name_font':{'color':'#8899aa'}})
ann_chart.set_legend({'font':{'color':'#8899aa'}})
ann_chart.set_size({'width':500,'height':280})
ws.insert_chart(chart_row+1, 3, ann_chart)

# Margin trend
mar_row = yr_row + 2
ws.merge_range(mar_row, 0, mar_row, 4, '利润率年度趋势', section_fmt)
mar_row += 1
ws.write(mar_row, 0, '指标', header_fmt)
for yi, y in enumerate(YEARS):
    ws.write(mar_row, yi+1, y, header_fmt)
mar_row += 1
for cat in ['毛利率','净利率','营业利润率']:
    ws.write(mar_row, 0, cat, label_fmt)
    for yi, y in enumerate(YEARS):
        ws.write(mar_row, yi+1, yd[cat][y]/100, data_pct)
    mar_row += 1

mar_chart = wb.add_chart({'type':'line'})
for yi, y in enumerate(YEARS):
    mar_chart.add_series({
        'name': y,
        'categories': ['驾驶舱总览', mar_row-3, 0, mar_row-1, 0],
        'values': ['驾驶舱总览', mar_row-3, yi+1, mar_row-1, yi+1],
        'line': {'color': ['#00d4ff','#ffd93d','#6bcbff'][yi], 'width': 3},
        'marker': {'type':'diamond','size':8},
    })
mar_chart.set_title({'name':'利润率年度趋势','name_font':{'color':'#6bcbff','size':12}})
mar_chart.set_y_axis({'name':'%','name_font':{'color':'#8899aa'},'num_format':'0%'})
mar_chart.set_legend({'font':{'color':'#8899aa'}})
mar_chart.set_size({'width':500,'height':280})
ws.insert_chart(mar_row+1, 0, mar_chart)

# ═══════════════════════════════════════
# SHEET 3: 收入分析
# ═══════════════════════════════════════
ws_rev = wb.add_worksheet('收入分析')
ws_rev.set_tab_color('#ff6b6b')
ws_rev.hide_gridlines(2)
ws_rev.set_column('A:E', 22)
ws_rev.merge_range('A1:E1', '收入分析', title_fmt)
ws_rev.set_row(0, 40)

# Revenue data table
rev_table_row = 3
ws_rev.write(rev_table_row, 0, '月份', header_fmt)
for yi, y in enumerate(YEARS):
    ws_rev.write(rev_table_row, yi+1, y, header_fmt)
rev_table_row += 1
for mi, m in enumerate(MONTHS):
    ws_rev.write(rev_table_row+mi, 0, m, label_fmt)
    for yi, y in enumerate(YEARS):
        ws_rev.write(rev_table_row+mi, yi+1, raw['营业收入'][y][mi], data_val)
rev_table_end = rev_table_row + 11
ws_rev.write(rev_table_end, 0, '合计', total_fmt)
for yi, y in enumerate(YEARS):
    ws_rev.write(rev_table_end, yi+1, yd['营业收入'][y], total_fmt)

# Revenue line chart
rev_line = wb.add_chart({'type':'line'})
for yi, y in enumerate(YEARS):
    rev_line.add_series({
        'name': y,
        'categories': ['收入分析', rev_table_row, 0, rev_table_end-1, 0],
        'values': ['收入分析', rev_table_row, yi+1, rev_table_end-1, yi+1],
        'line': {'color': ['#00d4ff','#ff6b6b','#ffd93d'][yi], 'width': 2.5},
    })
rev_line.set_title({'name':'营业收入月度对比','name_font':{'color':'#ff6b6b','size':12}})
rev_line.set_y_axis({'name':'万元','name_font':{'color':'#8899aa'}})
rev_line.set_size({'width':550,'height':300})
ws_rev.insert_chart(rev_table_end+1, 0, rev_line)

# Annual pie chart
pie = wb.add_chart({'type':'pie'})
pie.add_series({
    'name':'年度收入占比',
    'categories': ['收入分析', rev_table_end, 0, rev_table_end, 2],
    'values': ['收入分析', rev_table_end, 1, rev_table_end, 3],
})
pie.set_title({'name':'年度收入占比','name_font':{'color':'#ffd93d','size':12}})
pie.set_size({'width':450,'height':300})
ws_rev.insert_chart(rev_table_end+1, 3, pie)

# YoY growth
growth_row = rev_table_end + 18
ws_rev.write(growth_row, 0, '2025 vs 2024 同比增长率', section_fmt)
growth_row += 1
ws_rev.write(growth_row, 0, '月份', header_fmt)
ws_rev.write(growth_row, 1, '增长率', header_fmt)
growth_row += 1
for mi in range(12):
    ws_rev.write(growth_row+mi, 0, MONTHS[mi], label_fmt)
    base = raw['营业收入']['2024'][mi]
    g = (raw['营业收入']['2025'][mi] - base) / base if base else 0
    ws_rev.write(growth_row+mi, 1, g, data_pct)

growth_bar = wb.add_chart({'type':'column'})
growth_bar.add_series({
    'name':'同比增长率',
    'categories': ['收入分析', growth_row, 0, growth_row+11, 0],
    'values': ['收入分析', growth_row, 1, growth_row+11, 1],
    'fill': {'color':'#00d4ff'},
})
growth_bar.set_title({'name':'2025年 vs 2024年 月度同比增长率','name_font':{'color':'#ff6b6b','size':12}})
growth_bar.set_y_axis({'name':'%','num_format':'0%'})
growth_bar.set_size({'width':550,'height':280})
ws_rev.insert_chart(growth_row+13, 0, growth_bar)

# ═══════════════════════════════════════
# SHEET 4: 成本费用
# ═══════════════════════════════════════
ws_cost = wb.add_worksheet('成本费用')
ws_cost.set_tab_color('#ffd93d')
ws_cost.hide_gridlines(2)
ws_cost.set_column('A:E', 22)
ws_cost.merge_range('A1:E1', '成本费用分析', title_fmt)
ws_cost.set_row(0, 40)

# Cost data table
cost_table_row = 3
ws_cost.write(cost_table_row, 0, '月份', header_fmt)
for yi, y in enumerate(YEARS):
    ws_cost.write(cost_table_row, yi+1, y, header_fmt)
cost_table_row += 1
for mi, m in enumerate(MONTHS):
    ws_cost.write(cost_table_row+mi, 0, m, label_fmt)
    for yi, y in enumerate(YEARS):
        ws_cost.write(cost_table_row+mi, yi+1, raw['营业成本'][y][mi], data_val)
ct_end = cost_table_row + 11
ws_cost.write(ct_end, 0, '合计', total_fmt)
for yi, y in enumerate(YEARS):
    ws_cost.write(ct_end, yi+1, yd['营业成本'][y], total_fmt)

cost_line = wb.add_chart({'type':'line'})
for yi, y in enumerate(YEARS):
    cost_line.add_series({
        'name': y,
        'categories': ['成本费用', cost_table_row, 0, ct_end-1, 0],
        'values': ['成本费用', cost_table_row, yi+1, ct_end-1, yi+1],
        'line': {'color': ['#00d4ff','#ff6b6b','#ffd93d'][yi], 'width': 2.5},
    })
cost_line.set_title({'name':'营业成本月度趋势','name_font':{'color':'#ffd93d','size':12}})
cost_line.set_y_axis({'name':'万元'})
cost_line.set_size({'width':520,'height':280})
ws_cost.insert_chart(ct_end+1, 0, cost_line)

# 2025 expense pie
exp_items = ['营业成本','销售费用','管理费用','研发费用','财务费用','税金及附加']
exp_vals = [yd[item]['2025'] for item in exp_items]
pie_row = ct_end + 16
ws_cost.write(pie_row, 0, '2025年成本费用构成', section_fmt)
pie_row += 1
for i, item in enumerate(exp_items):
    ws_cost.write(pie_row+i, 0, item, label_fmt)
    ws_cost.write(pie_row+i, 1, exp_vals[i], data_int)

exp_pie = wb.add_chart({'type':'pie'})
exp_pie.add_series({
    'name':'2025年成本费用构成',
    'categories': ['成本费用', pie_row, 0, pie_row+5, 0],
    'values': ['成本费用', pie_row, 1, pie_row+5, 1],
})
exp_pie.set_title({'name':'2025年成本费用构成','name_font':{'color':'#ffd93d','size':12}})
exp_pie.set_size({'width':450,'height':300})
ws_cost.insert_chart(ct_end+1, 3, exp_pie)

# Annual cost comparison
ac_row = pie_row + 7
ws_cost.write(ac_row, 0, '年度费用对比', section_fmt)
ac_row += 1
ws_cost.write(ac_row, 0, '项目', header_fmt)
for yi, y in enumerate(YEARS):
    ws_cost.write(ac_row, yi+1, y, header_fmt)
ac_row += 1
for item in exp_items:
    ws_cost.write(ac_row, 0, item, label_fmt)
    for yi, y in enumerate(YEARS):
        ws_cost.write(ac_row, yi+1, yd[item][y], data_int)
    ac_row += 1

ac_chart = wb.add_chart({'type':'column'})
for yi, y in enumerate(YEARS):
    ac_chart.add_series({
        'name': y,
        'categories': ['成本费用', ac_row-len(exp_items), 0, ac_row-1, 0],
        'values': ['成本费用', ac_row-len(exp_items), yi+1, ac_row-1, yi+1],
        'fill': {'color': ['#00d4ff','#ff6b6b','#ffd93d'][yi]},
    })
ac_chart.set_title({'name':'年度费用对比','name_font':{'color':'#ffd93d','size':12}})
ac_chart.set_y_axis({'name':'万元'})
ac_chart.set_size({'width':520,'height':280})
ws_cost.insert_chart(ac_row+1, 0, ac_chart)

# ═══════════════════════════════════════
# SHEET 5: 趋势预测
# ═══════════════════════════════════════
ws_pred = wb.add_worksheet('趋势预测')
ws_pred.set_tab_color('#c084fc')
ws_pred.hide_gridlines(2)
ws_pred.set_column('A:E', 22)
ws_pred.merge_range('A1:E1', '未来趋势预测', title_fmt)
ws_pred.set_row(0, 40)

from sklearn.linear_model import LinearRegression
import numpy as np

def predict(item):
    X = np.arange(36).reshape(-1, 1)
    y = np.array(raw[item]['2023'] + raw[item]['2024'] + raw[item]['2025'])
    model = LinearRegression().fit(X, y)
    pred = model.predict(np.arange(36, 48).reshape(-1, 1))
    return pred

pred_rev = predict('营业收入')
pred_cost = predict('营业成本')

hist_rev = raw['营业收入']['2023'] + raw['营业收入']['2024'] + raw['营业收入']['2025']
hist_cost = raw['营业成本']['2023'] + raw['营业成本']['2024'] + raw['营业成本']['2025']
all_months = [f'2023-{m}' for m in range(1,13)] + [f'2024-{m}' for m in range(1,13)] + [f'2025-{m}' for m in range(1,13)]
pred_mons = [f'2026-{m}' for m in range(1,13)]

# Write prediction data
pd_row = 3
ws_pred.write(pd_row, 0, '月份', header_fmt)
ws_pred.write(pd_row, 1, '收入(实际)', header_fmt)
ws_pred.write(pd_row, 2, '成本(实际)', header_fmt)
pd_row += 1
for i in range(36):
    ws_pred.write(pd_row+i, 0, all_months[i], label_fmt)
    ws_pred.write(pd_row+i, 1, hist_rev[i], data_val)
    ws_pred.write(pd_row+i, 2, hist_cost[i], data_val)
hist_end = pd_row + 36

# Prediction section
ws_pred.write(hist_end+1, 0, '预测数据', section_fmt)
ws_pred.write(hist_end+2, 0, '月份', header_fmt)
ws_pred.write(hist_end+2, 1, '收入(预测)', header_fmt)
ws_pred.write(hist_end+2, 2, '成本(预测)', header_fmt)
for i in range(12):
    ws_pred.write(hist_end+3+i, 0, pred_mons[i], label_fmt)
    ws_pred.write(hist_end+3+i, 1, round(pred_rev[i],2), data_val)
    ws_pred.write(hist_end+3+i, 2, round(pred_cost[i],2), data_val)
pred_end = hist_end + 3 + 11

# Prediction chart
pred_chart = wb.add_chart({'type':'line'})
pred_chart.add_series({
    'name':'收入(历史)',
    'categories': ['趋势预测', pd_row, 0, hist_end-1, 0],
    'values': ['趋势预测', pd_row, 1, hist_end-1, 1],
    'line': {'color':'#00d4ff','width':2.5},
})
pred_chart.add_series({
    'name':'收入(预测)',
    'categories': ['趋势预测', hist_end+3, 0, pred_end, 0],
    'values': ['趋势预测', hist_end+3, 1, pred_end, 1],
    'line': {'color':'#00d4ff','width':2.5,'dash_type':'dash'},
})
pred_chart.add_series({
    'name':'成本(历史)',
    'categories': ['趋势预测', pd_row, 0, hist_end-1, 0],
    'values': ['趋势预测', pd_row, 2, hist_end-1, 2],
    'line': {'color':'#ff6b6b','width':2.5},
})
pred_chart.add_series({
    'name':'成本(预测)',
    'categories': ['趋势预测', hist_end+3, 0, pred_end, 0],
    'values': ['趋势预测', hist_end+3, 2, pred_end, 2],
    'line': {'color':'#ff6b6b','width':2.5,'dash_type':'dash'},
})
pred_chart.set_title({'name':'营业收入 & 成本预测（2026年）','name_font':{'color':'#c084fc','size':12}})
pred_chart.set_y_axis({'name':'万元'})
pred_chart.set_size({'width':800,'height':380})
ws_pred.insert_chart(pd_row, 3, pred_chart)

# Prediction KPIs
pk_row = pred_end + 3
ws_pred.write(pk_row, 0, '预测摘要', section_fmt)
pk_row += 1
pred_rev_total = sum(pred_rev)
pred_cost_total = sum(pred_cost)
growth = (pred_rev_total / yd['营业收入']['2025'] - 1) * 100
ws_pred.write(pk_row, 0, '预测2026年收入', label_fmt)
ws_pred.write(pk_row, 1, round(pred_rev_total), data_int)
ws_pred.write(pk_row, 2, f'增长率 {growth:.1f}%', kpi_change_up if growth>0 else kpi_change_dn)
pk_row += 1
ws_pred.write(pk_row, 0, '预测2026年成本', label_fmt)
ws_pred.write(pk_row, 1, round(pred_cost_total), data_int)
pk_row += 1
ws_pred.write(pk_row, 0, '预测毛利', label_fmt)
ws_pred.write(pk_row, 1, round(pred_rev_total - pred_cost_total), data_int)

# ═══════════════════════════════════════
# SHEET 6: 同行业对比
# ═══════════════════════════════════════
ws_ind = wb.add_worksheet('同行业对比')
ws_ind.set_tab_color('#54a0ff')
ws_ind.hide_gridlines(2)
ws_ind.set_column('A:E', 22)
ws_ind.merge_range('A1:E1', '同行业对比分析', title_fmt)
ws_ind.set_row(0, 40)

ws_ind.write(2, 0, '对标公司：艾融软件(830799)、用友网络(600588)、华宇软件(300271)、三维天地(301159)、山大地纬(688579)、殷图网联(835508)、天亿马(301178)、远光软件(002063)、宇信科技(300674)', note_fmt)
ws_ind.merge_range(2, 0, 2, 4, '', note_fmt)

# Try to load industry data
import pandas as pd
csv_path = os.path.join(os.path.dirname(__file__), 'industry_data.csv')
if os.path.exists(csv_path):
    ind_df = pd.read_csv(csv_path, encoding='utf-8-sig')
    ind_row = 4
    ws_ind.write(ind_row, 0, '2024年行业对比数据', section_fmt)
    ws_ind.merge_range(ind_row, 0, ind_row, 4, '2024年行业对比数据', section_fmt)
    ind_row += 1

    def parse_cn(val):
        if pd.isna(val) or val is None or val == 'False': return None
        s = str(val).replace(',','').strip()
        neg = 1
        if s.startswith('-'): neg = -1; s = s[1:]
        if '亿' in s: return neg * float(s.replace('亿','')) * 10000
        if '万' in s: return neg * float(s.replace('万',''))
        if '%' in s: return neg * float(s.replace('%',''))
        try: return neg * float(s)
        except: return None

    yr_sel = '2024'
    peer_data = ind_df[ind_df['报告期'].astype(str) == yr_sel]

    # Headers
    headers = ['公司','营业收入(万元)','净利润(万元)','毛利率','净利率','净资产收益率']
    for ci, h in enumerate(headers):
        ws_ind.write(ind_row, ci, h, header_fmt)
    ind_row += 1

    # Our company
    rev_tt = yd['营业收入'][yr_sel]
    cost_tt = yd['营业成本'][yr_sel]
    profit_tt = yd['利润总额'][yr_sel]
    gross_m = (rev_tt - cost_tt) / rev_tt * 100
    net_m = profit_tt / rev_tt * 100
    ws_ind.write(ind_row, 0, '✨ 小甜甜公司', label_fmt)
    ws_ind.write(ind_row, 1, round(rev_tt), data_int)
    ws_ind.write(ind_row, 2, round(profit_tt), data_int)
    ws_ind.write(ind_row, 3, f'{gross_m:.1f}%', data_val)
    ws_ind.write(ind_row, 4, f'{net_m:.1f}%', data_val)
    ws_ind.write(ind_row, 5, '-', data_val)
    ind_row += 1

    # Peer companies
    for _, r in peer_data.iterrows():
        ws_ind.write(ind_row, 0, r['公司'], label_fmt)
        ws_ind.write(ind_row, 1, round(parse_cn(r['营业总收入']) or 0), data_int)
        ws_ind.write(ind_row, 2, round(parse_cn(r['净利润']) or 0), data_int)
        ws_ind.write(ind_row, 3, r['销售毛利率'] if '%' in str(r['销售毛利率']) else f'{r["销售毛利率"]:.1f}%', data_val)
        ws_ind.write(ind_row, 4, r['销售净利率'] if '%' in str(r['销售净利率']) else f'{r["销售净利率"]:.1f}%', data_val)
        ws_ind.write(ind_row, 5, r['净资产收益率'] if '%' in str(r['净资产收益率']) else f'{r["净资产收益率"]:.1f}%', data_val)
        ind_row += 1
else:
    ws_ind.write(3, 0, '行业数据未找到，请先运行 industry_data_fetcher.py', note_fmt)

# ─── Close ───
wb.close()
print(f'Excel dashboard saved to: {OUTPUT}')
