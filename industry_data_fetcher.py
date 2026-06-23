# -*- coding: utf-8 -*-
"""行业对比数据抓取脚本 - 使用 AkShare 获取9家上市公司年报数据"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import akshare as ak
import pandas as pd
import os

OUTPUT = os.path.join(os.path.dirname(__file__), 'industry_data.csv')

PEERS = {
    '艾融软件': '830799',
    '用友网络': '600588',
    '华宇软件': '300271',
    '三维天地': '301159',
    '山大地纬': '688579',
    '殷图网联': '835508',
    '天亿马': '301178',
    '远光软件': '002063',
    '宇信科技': '300674',
}

rows = []
for name, code in PEERS.items():
    print(f'Fetching {name}({code})...', end=' ')
    try:
        df = ak.stock_financial_abstract_ths(symbol=code)
        annual = df[df['报告期'].str.endswith('12-31')].tail(3)
        for _, r in annual.iterrows():
            rows.append({
                '公司': name,
                '股票代码': code,
                '报告期': r['报告期'][:4],
                '营业总收入': r.get('营业总收入'),
                '净利润': r.get('净利润'),
                '销售毛利率': r.get('销售毛利率'),
                '销售净利率': r.get('销售净利率'),
                '净资产收益率': r.get('净资产收益率'),
                '基本每股收益': r.get('基本每股收益'),
                '每股净资产': r.get('每股净资产'),
            })
        print(f'{len(annual)} rows')
    except Exception as e:
        print(f'Error: {e}')

result = pd.DataFrame(rows)
result.to_csv(OUTPUT, index=False, encoding='utf-8-sig')
print(f'\nSaved {len(result)} rows to {OUTPUT}')
print(result.head(20).to_string(index=False))
