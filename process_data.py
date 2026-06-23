# -*- coding: utf-8 -*-
import sys, openpyxl, csv, os
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('附件0-小甜甜公司数据.xlsx', data_only=True)
ws = wb['利润表']

rows_data = {}
for i, row in enumerate(ws.iter_rows(values_only=True), 1):
    vals = [v for v in row]
    if i >= 4 and i <= 27:
        label = str(vals[0]).strip().replace('\n','')
        rows_data[label] = vals

income_labels = {
    '营业收入': '一、营业收入',
    '营业成本': '减：营业成本',
    '税金及附加': '税金及附加',
    '销售费用': '销售费用',
    '管理费用': '管理费用',
    '研发费用': '研发费用',
    '财务费用': '财务费用',
    '其他收益': '加：其他收益',
    '信用减值损失': '信用减值损失（损失以"-"号填列）',
    '资产处置收益': '资产处置收益（损失以"-"号填列）',
    '营业外收入': '加：营业外收入',
    '营业外支出': '减：营业外支出'
}

with open('profit_data_flat.csv', 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['年份','月份','项目','金额（元）','金额（万元）'])
    
    for year_col_start, year_offset in [(28,'2023'),(15,'2024'),(2,'2025')]:
        months_list = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']
        for mi in range(12):
            for short_name, long_label in income_labels.items():
                if long_label in rows_data:
                    val = rows_data[long_label][year_col_start + mi]
                    if val and isinstance(val, (int, float)) and val != 0:
                        writer.writerow([year_offset, months_list[mi], short_name, round(val, 2), round(val/10000, 4)])

    # Derived items: 营业利润, 利润总额, 净利润
    for year_col_start, year_offset in [(28,'2023'),(15,'2024'),(2,'2025')]:
        months_list = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']
        for mi in range(12):
            def get_val(label, idx, default=0):
                v = rows_data.get(label, [0]*50)[idx]
                return v if v and isinstance(v, (int, float)) else default

            rev = get_val('一、营业收入', year_col_start + mi)
            cost = get_val('减：营业成本', year_col_start + mi)
            tax = get_val('税金及附加', year_col_start + mi)
            sell = get_val('销售费用', year_col_start + mi)
            admin = get_val('管理费用', year_col_start + mi)
            rd = get_val('研发费用', year_col_start + mi)
            fin = get_val('财务费用', year_col_start + mi)
            other_inc = get_val('加：其他收益', year_col_start + mi)
            credit_loss = get_val('信用减值损失（损失以"-"号填列）', year_col_start + mi)
            asset_gain = get_val('资产处置收益（损失以"-"号填列）', year_col_start + mi)
            
            op_profit = rev - cost - tax - sell - admin - rd - fin + other_inc + credit_loss + asset_gain
            writer.writerow([year_offset, months_list[mi], '营业利润', round(op_profit, 2), round(op_profit/10000, 4)])
            
            oi = get_val('加：营业外收入', year_col_start + mi)
            oe = get_val('减：营业外支出', year_col_start + mi)
            total_profit = op_profit + oi - oe
            writer.writerow([year_offset, months_list[mi], '利润总额', round(total_profit, 2), round(total_profit/10000, 4)])

print('CSV created: profit_data_flat.csv')
print(f'File size: {os.path.getsize("profit_data_flat.csv")} bytes')
