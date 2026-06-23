# -*- coding: utf-8 -*-
import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

# 1. 仪表板需求
wb = openpyxl.load_workbook('仪表板需求.xlsx', data_only=True)
print('===== 仪表板需求 =====')
print('Sheets:', wb.sheetnames)
for name in wb.sheetnames:
    ws = wb[name]
    print(f'\n--- Sheet: {name} ---')
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if any(v is not None for v in row):
            vals = []
            for v in row:
                if isinstance(v, float) and v == int(v):
                    vals.append(str(int(v)))
                elif v is not None:
                    vals.append(str(v)[:120])
                else:
                    vals.append('')
            print(f'  R{i}: {vals}')

# 2. 附件1-收入图表 summary sheets
print('\n\n===== 附件1-收入图表 =====')
wb2 = openpyxl.load_workbook('附件1-收入图表.xlsx', data_only=True)
for name in ['营业收入', '年度对比', '饼图-月度分布', '柱状图-同期对比', '收入占比环图', '年度对比', '同比增长率']:
    if name in wb2.sheetnames:
        ws = wb2[name]
        print(f'\n--- Sheet: {name} ---')
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            if any(v is not None for v in row):
                vals = []
                for v in row:
                    if isinstance(v, float) and v == int(v):
                        vals.append(str(int(v)))
                    elif v is not None:
                        vals.append(str(v)[:60])
                    else:
                        vals.append('')
                print(f'  {vals}')

# 3. 附件2-成本费用图表 summary
print('\n\n===== 附件2-成本费用图表 =====')
wb3 = openpyxl.load_workbook('附件2-成本费用图表.xlsx', data_only=True)
for name in wb3.sheetnames[:8]:
    ws = wb3[name]
    print(f'\n--- Sheet: {name} ---')
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if any(v is not None for v in row):
            vals = []
            for v in row:
                if isinstance(v, float) and v == int(v):
                    vals.append(str(int(v)))
                elif v is not None:
                    vals.append(str(v)[:60])
                else:
                    vals.append('')
            print(f'  {vals}')

# 4. 利润表 full
print('\n\n===== 利润表 Full =====')
ws4 = wb['利润表']
for i, row in enumerate(ws4.iter_rows(values_only=True), 1):
    if any(v is not None for v in row):
        vals = []
        for v in row:
            if v is not None:
                vals.append(str(v)[:30])
            else:
                vals.append('')
        print(f'  R{i}: {vals}')
