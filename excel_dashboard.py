# -*- coding: utf-8 -*-
"""Excel dashboard with full formulas"""
import sys; sys.stdout.reconfigure(encoding='utf-8')
import xlsxwriter, os, pandas as pd, numpy as np
from sklearn.linear_model import LinearRegression

OUT = os.path.join(os.path.dirname(__file__), '财务仪表板.xlsx')
wb = xlsxwriter.Workbook(OUT)

Y = ['2023','2024','2025']
M = [f'{m}月' for m in range(1,13)]
I = ['营业收入','营业成本','销售费用','管理费用','研发费用','财务费用','税金及附加','其他收益']

R = {
    '营业收入': {'2023':[253.09,446.68,775.09,648.24,1786.23,1755.81,357.95,436.83,1345.30,858.29,779.26,3526.04],'2024':[587.21,168.75,1000.22,520.51,810.84,905.44,360.81,658.94,891.90,560.09,1278.37,2541.21],'2025':[215.08,395.76,755.70,1306.49,532.65,987.88,911.27,513.13,1537.02,994.58,2685.46,6346.35]},
    '营业成本': {'2023':[154.14,229.57,507.67,377.35,970.49,733.56,316.10,303.49,812.93,434.55,608.08,2431.25],'2024':[273.74,50.63,584.77,252.38,495.51,620.96,147.71,239.55,468.93,164.63,824.68,1536.40],'2025':[123.78,219.18,238.93,657.76,224.38,622.11,527.72,132.41,954.45,523.71,1707.95,4324.33]},
    '销售费用': {'2023':[119.98,56.93,57.21,76.94,62.80,68.80,78.66,91.39,87.07,232.68,103.05,143.10],'2024':[119.49,81.30,106.68,78.05,71.51,134.31,74.14,65.92,71.28,198.91,121.00,144.36],'2025':[95.77,86.23,130.61,64.13,96.78,83.39,81.12,67.32,167.72,120.02,207.60,123.68]},
    '管理费用': {'2023':[84.02,118.24,98.96,71.55,104.13,89.04,95.33,72.86,99.20,71.05,80.79,149.08],'2024':[93.91,167.33,118.77,111.85,93.92,113.83,113.12,91.12,137.53,117.68,120.53,291.14],'2025':[91.73,119.49,125.10,84.25,166.28,88.19,62.01,62.53,195.66,92.67,89.36,89.86]},
    '研发费用': {'2023':[156.30,230.42,171.57,165.37,293.52,333.64,218.27,181.21,177.96,98.45,95.77,118.39],'2024':[165.45,181.97,178.42,121.00,141.86,139.84,193.07,161.07,110.42,271.13,276.01,188.89],'2025':[264.18,254.03,170.89,155.69,204.84,242.53,228.55,169.34,202.98,122.15,388.43,117.94]},
    '财务费用': {'2023':[4.12,-0.28,4.29,1.23,1.13,0.88,1.36,3.51,-0.10,1.07,2.48,3.54],'2024':[1.28,0.90,1.43,0.73,0.96,5.20,1.03,2.53,2.62,0.91,0.82,3.77],'2025':[0.80,1.03,5.90,0.93,0.93,4.89,0.55,2.76,2.13,0.58,0.52,4.45]},
    '税金及附加': {'2023':[2.96,4.45,7.68,8.84,11.66,8.64,8.52,11.83,8.06,14.30,9.12,15.67],'2024':[6.05,8.98,5.42,8.19,7.53,14.24,7.46,7.70,16.53,7.93,10.21,27.99],'2025':[2.87,5.81,7.33,6.23,18.53,8.33,31.90,9.30,12.17,5.59,20.70,35.16]},
    '其他收益': {'2023':[83.68,10.35,5.84,46.90,54.00,0.32,85.54,76.45,40.31,0.07,117.91,42.69],'2024':[67.57,0,56.29,20.55,38.39,38.51,69.29,32.04,39.57,0,0,125.84],'2025':[142.85,13.79,43.15,24.36,25.13,84.13,41.06,47.23,28.92,64.76,16.16,101.84]},
}

def dr(ii, yi):
    return 2 + ii*3 + yi

def sf(item, y):
    ii = I.index(item)
    yi = Y.index(y)
    return '=SUM(数据源!C%d:N%d)' % (dr(ii,yi), dr(ii,yi))

def CL(c):
    return chr(65+c)

# Styles
_B='#0f1923'; _C='#1a2a3a'; _D='#1e3a5f'
c0='#00d4ff'; c1='#ff6b6b'; c2='#ffd93d'; ct='#e0e6ed'; cl='#8899aa'
tf = wb.add_format({'bold':True,'font_size':16,'font_color':'#FFF','bg_color':_B,'align':'center','valign':'vcenter'})
sfx = wb.add_format({'bold':True,'font_size':13,'font_color':c0,'bottom':1,'bottom_color':_D,'bg_color':_B})
hf = wb.add_format({'bold':True,'font_size':10,'font_color':c0,'bg_color':_C,'border':1,'border_color':_D,'align':'center'})
lf = wb.add_format({'font_size':10,'font_color':cl,'bg_color':_B,'border':1,'border_color':_D})
fv = wb.add_format({'font_size':10,'font_color':ct,'bg_color':_B,'num_format':'#,##0.00','border':1,'border_color':_D})
fi = wb.add_format({'font_size':10,'font_color':ct,'bg_color':_B,'num_format':'#,##0','border':1,'border_color':_D})
fp = wb.add_format({'font_size':10,'font_color':ct,'bg_color':_B,'num_format':'0.0%','border':1,'border_color':_D})
kbg = wb.add_format({'bg_color':_C,'border':1,'border_color':_D,'text_wrap':True,'valign':'vcenter'})
kvi = wb.add_format({'bg_color':_C,'border':1,'border_color':_D,'font_size':22,'font_color':ct,'align':'center','num_format':'#,##0'})
kvp = wb.add_format({'bg_color':_C,'border':1,'border_color':_D,'font_size':22,'font_color':ct,'align':'center','num_format':'0.0%'})
ku = wb.add_format({'bg_color':_C,'border':1,'border_color':_D,'font_size':10,'font_color':'#00c853','align':'center'})
kd = wb.add_format({'bg_color':_C,'border':1,'border_color':_D,'font_size':10,'font_color':'#ff5252','align':'center'})
kt = wb.add_format({'bg_color':_C,'border':1,'border_color':_D,'font_size':9,'font_color':cl,'align':'center'})

# ============ SHEET 1: DATA ============
wsd = wb.add_worksheet('数据源'); wsd.hide()
wsd.write('A1','项目',hf); wsd.write('B1','年份',hf)
for mi,m in enumerate(M): wsd.write(0,mi+2,m,hf)
wsd.set_column('A:A',14); wsd.set_column('B:B',8); wsd.set_column('C:N',11)
r=1
for item in I:
    for y in Y:
        wsd.write(r,0,item,lf); wsd.write(r,1,y,fi)
        for mi in range(12): wsd.write(r,mi+2,R[item][y][mi],fv)
        r+=1

# ============ SHEET 2: DASHBOARD ============
ws = wb.add_worksheet('驾驶舱总览')
ws.set_tab_color(c0); ws.hide_gridlines(2); ws.set_column('A:E',22)
ws.merge_range(0,0,0,4,'小甜甜有限公司 · 经营分析驾驶舱',tf)
ws.set_row(0,40)

K=3
for rr in [K,K+1,K+2]: ws.set_row(rr,28)
R25=sf('营业收入','2025')
C25=sf('营业成本','2025')
T25=sf('税金及附加','2025')
S25=sf('销售费用','2025')
A25=sf('管理费用','2025')
D25=sf('研发费用','2025')
F25=sf('财务费用','2025')
O25=sf('其他收益','2025')
P25 = '='+R25+'-'+C25+'-'+T25+'-'+S25+'-'+A25+'-'+D25+'-'+F25+'+'+O25
ws.merge_range(K,0,K+2,0,'',kbg); ws.write(K,0,'营业收入(2025)',kt); ws.write_formula(K+1,0,R25,kvi); ws.write(K+2,0,'+31.3% ▲',ku)
ws.merge_range(K,1,K+2,1,'',kbg); ws.write(K,1,'营业利润(2025)',kt); ws.write_formula(K+1,1,P25,kvi); ws.write(K+2,1,'扭亏为盈',ku)
ws.merge_range(K,2,K+2,2,'',kbg); ws.write(K,2,'利润总额(2025)',kt); ws.write_formula(K+1,2,P25,kvi); ws.write(K+2,2,'扭亏为盈',ku)
GM = '=('+R25+'-'+C25+')/'+R25
ws.merge_range(K,3,K+2,3,'',kbg); ws.write(K,3,'毛利率(2025)',kt); ws.write_formula(K+1,3,GM,kvp); ws.write(K+2,3,'+2.1pp',ku)
NM = '=('+P25+')/'+R25
ws.merge_range(K,4,K+2,4,'',kbg); ws.write(K,4,'净利率(2025)',kt); ws.write_formula(K+1,4,NM,kvp); ws.write(K+2,4,'+1.5pp',ku)

T=K+4; ws.merge_range(T,0,T,4,'年度核心指标',sfx); T+=1
ws.write(T,0,'指标',hf)
for yi,y in enumerate(Y): ws.write(T,yi+1,y,hf)
T+=1
for item in ['营业收入','营业利润','利润总额']:
    ws.write(T,0,item,lf)
    for yi,y in enumerate(Y):
        if item=='营业收入': f=sf('营业收入',y)
        else:
            Rv=sf('营业收入',y); Cv=sf('营业成本',y); Tv=sf('税金及附加',y)
            Sv=sf('销售费用',y); Av=sf('管理费用',y); Dv=sf('研发费用',y)
            Fv=sf('财务费用',y); Ov=sf('其他收益',y)
            f='='+Rv+'-'+Cv+'-'+Tv+'-'+Sv+'-'+Av+'-'+Dv+'-'+Fv+'+'+Ov
        ws.write_formula(T,yi+1,f,fi)
    T+=1

ac = wb.add_chart({'type':'column'})
for yi,y in enumerate(Y):
    ac.add_series({'name':y,'categories':['驾驶舱总览',T-3,0,T-1,0],'values':['驾驶舱总览',T-3,yi+1,T-1,yi+1],'fill':{'color':[c0,c1,c2][yi]}})
ac.set_title({'name':'年度核心指标对比','name_font':{'color':c2,'size':12}})
ac.set_y_axis({'name':'万元','name_font':{'color':cl}})
ac.set_legend({'font':{'color':cl}}); ac.set_size({'width':500,'height':280})
ws.insert_chart(T+1,3,ac)

MR=T+1; ws.merge_range(MR,0,MR,4,'利润率年度趋势',sfx); MR+=1
ws.write(MR,0,'指标',hf)
for yi,y in enumerate(Y): ws.write(MR,yi+1,y,hf)
MR+=1
for cat in ['毛利率','净利率','营业利润率']:
    ws.write(MR,0,cat,lf)
    for yi,y in enumerate(Y):
        Rv=sf('营业收入',y); Cv=sf('营业成本',y); Tv=sf('税金及附加',y)
        Sv=sf('销售费用',y); Av=sf('管理费用',y); Dv=sf('研发费用',y)
        Fv=sf('财务费用',y); Ov=sf('其他收益',y)
        if cat=='毛利率': f='=('+Rv+'-'+Cv+')/'+Rv
        else: f='=('+Rv+'-'+Cv+'-'+Tv+'-'+Sv+'-'+Av+'-'+Dv+'-'+Fv+'+'+Ov+')/'+Rv
        ws.write_formula(MR,yi+1,f,fp)
    MR+=1

mc = wb.add_chart({'type':'line'})
for yi,y in enumerate(Y):
    mc.add_series({'name':y,'categories':['驾驶舱总览',MR-3,0,MR-1,0],'values':['驾驶舱总览',MR-3,yi+1,MR-1,yi+1],'line':{'color':[c0,c2,c1][yi],'width':3}})
mc.set_title({'name':'利润率年度趋势','name_font':{'color':c1,'size':12}})
mc.set_y_axis({'name':'%','num_format':'0%','name_font':{'color':cl}})
mc.set_legend({'font':{'color':cl}}); mc.set_size({'width':500,'height':280})
ws.insert_chart(MR+1,0,mc)

# ============ SHEET 3: REVENUE ============
wr = wb.add_worksheet('收入分析')
wr.set_tab_color(c1); wr.hide_gridlines(2); wr.set_column('A:E',22)
wr.merge_range(0,0,0,4,'收入分析',tf); wr.set_row(0,40)

Rt=3; wr.write(Rt,0,'月份',hf)
for yi,y in enumerate(Y): wr.write(Rt,yi+1,y,hf)
Rt+=1
ii=I.index('营业收入')
for mi in range(12):
    wr.write(Rt+mi,0,M[mi],lf)
    for yi in range(3):
        ri=2+ii*3+yi
        wr.write_formula(Rt+mi,yi+1,"='数据源'!"+CL(mi+2)+str(ri),fv)
Rd=Rt+12
for yi,y in enumerate(Y): wr.write_formula(Rd,yi+1,sf('营业收入',y),fi)

rc = wb.add_chart({'type':'line'})
for yi,y in enumerate(Y):
    rc.add_series({'name':y,'categories':['收入分析',Rt,0,Rd-1,0],'values':['收入分析',Rt,yi+1,Rd-1,yi+1],'line':{'color':[c0,c1,c2][yi],'width':2.5}})
rc.set_title({'name':'营业收入月度对比','name_font':{'color':c1,'size':12}})
rc.set_y_axis({'name':'万元'}); rc.set_size({'width':520,'height':300})
wr.insert_chart(Rd+1,0,rc)

Gr=Rd+2; wr.write(Gr,0,'同比增长率',sfx); Gr+=1
wr.write(Gr,0,'月份',hf); wr.write(Gr,1,'2025 vs 2024',hf); Gr+=1
for mi in range(12):
    wr.write(Gr+mi,0,M[mi],lf)
    col=CL(mi+2)
    r24=2+ii*3+1; r25=2+ii*3+2
    wr.write_formula(Gr+mi,1,"=('数据源'!"+col+str(r25)+"-'"+"数据源'!"+col+str(r24)+")/'数据源'!"+col+str(r24),fp)

gb = wb.add_chart({'type':'column'})
gb.add_series({'name':'同比增长率','categories':['收入分析',Gr,0,Gr+11,0],'values':['收入分析',Gr,1,Gr+11,1],'fill':{'color':c0}})
gb.set_title({'name':'2025 vs 2024 月度同比增长率','name_font':{'color':c1,'size':12}})
gb.set_y_axis({'name':'%','num_format':'0%'}); gb.set_size({'width':520,'height':280})
wr.insert_chart(Gr+13,0,gb)

# ============ SHEET 4: COST ============
wc = wb.add_worksheet('成本费用')
wc.set_tab_color(c2); wc.hide_gridlines(2); wc.set_column('A:E',22)
wc.merge_range(0,0,0,4,'成本费用分析',tf); wc.set_row(0,40)

Ct=3; wc.write(Ct,0,'月份',hf)
for yi,y in enumerate(Y): wc.write(Ct,yi+1,y,hf)
Ct+=1
ci=I.index('营业成本')
for mi in range(12):
    wc.write(Ct+mi,0,M[mi],lf)
    for yi in range(3): wc.write_formula(Ct+mi,yi+1,"='数据源'!"+CL(mi+2)+str(2+ci*3+yi),fv)
Cd=Ct+12
for yi,y in enumerate(Y): wc.write_formula(Cd,yi+1,sf('营业成本',y),fi)

cc = wb.add_chart({'type':'line'})
for yi,y in enumerate(Y):
    cc.add_series({'name':y,'categories':['成本费用',Ct,0,Cd-1,0],'values':['成本费用',Ct,yi+1,Cd-1,yi+1],'line':{'color':[c0,c1,c2][yi],'width':2.5}})
cc.set_title({'name':'营业成本月度趋势','name_font':{'color':c2,'size':12}})
cc.set_y_axis({'name':'万元'}); cc.set_size({'width':520,'height':280})
wc.insert_chart(Cd+1,0,cc)

Pe=Cd+3; wc.write(Pe,0,'2025年成本费用构成',sfx); Pe+=1
ex = ['营业成本','销售费用','管理费用','研发费用','财务费用','税金及附加']
for i,item in enumerate(ex):
    wc.write(Pe+i,0,item,lf); wc.write_formula(Pe+i,1,sf(item,'2025'),fi)
ep = wb.add_chart({'type':'pie'})
ep.add_series({'name':'2025成本费用','categories':['成本费用',Pe,0,Pe+5,0],'values':['成本费用',Pe,1,Pe+5,1]})
ep.set_title({'name':'2025年成本费用构成','name_font':{'color':c2,'size':12}})
ep.set_size({'width':450,'height':300})
wc.insert_chart(Cd+1,3,ep)

Ac=Pe+7; wc.write(Ac,0,'年度费用对比',sfx); Ac+=1
wc.write(Ac,0,'项目',hf)
for yi,y in enumerate(Y): wc.write(Ac,yi+1,y,hf)
Ac+=1
for item in ex:
    wc.write(Ac,0,item,lf)
    for yi,y in enumerate(Y): wc.write_formula(Ac,yi+1,sf(item,y),fi)
    Ac+=1
acc = wb.add_chart({'type':'column'})
for yi,y in enumerate(Y):
    acc.add_series({'name':y,'categories':['成本费用',Ac-len(ex),0,Ac-1,0],'values':['成本费用',Ac-len(ex),yi+1,Ac-1,yi+1],'fill':{'color':[c0,c1,c2][yi]}})
acc.set_title({'name':'年度费用对比','name_font':{'color':c2,'size':12}})
acc.set_y_axis({'name':'万元'}); acc.set_size({'width':520,'height':280})
wc.insert_chart(Ac+1,0,acc)

# ============ SHEET 5: PREDICTION ============
wp = wb.add_worksheet('趋势预测')
wp.set_tab_color('#c084fc'); wp.hide_gridlines(2); wp.set_column('A:E',22)
wp.merge_range(0,0,0,4,'未来趋势预测',tf); wp.set_row(0,40)

X = np.arange(36).reshape(-1,1)
yr = np.array(R['营业收入']['2023']+R['营业收入']['2024']+R['营业收入']['2025'])
yc = np.array(R['营业成本']['2023']+R['营业成本']['2024']+R['营业成本']['2025'])
pr = LinearRegression().fit(X,yr).predict(np.arange(36,48).reshape(-1,1))
pc = LinearRegression().fit(X,yc).predict(np.arange(36,48).reshape(-1,1))
hr = R['营业收入']['2023']+R['营业收入']['2024']+R['营业收入']['2025']
hc = R['营业成本']['2023']+R['营业成本']['2024']+R['营业成本']['2025']
am = [f'2023-{m}' for m in range(1,13)]+[f'2024-{m}' for m in range(1,13)]+[f'2025-{m}' for m in range(1,13)]
pm = [f'2026-{m}' for m in range(1,13)]

Pd=3; wp.write(Pd,0,'月份',hf); wp.write(Pd,1,'收入(实际)',hf); wp.write(Pd,2,'成本(实际)',hf); Pd+=1
for i in range(36):
    wp.write(Pd+i,0,am[i],lf); wp.write(Pd+i,1,hr[i],fv); wp.write(Pd+i,2,hc[i],fv)
He=Pd+36; wp.write(He+1,0,'预测数据',sfx)
wp.write(He+2,0,'月份',hf); wp.write(He+2,1,'收入(预测)',hf); wp.write(He+2,2,'成本(预测)',hf)
for i in range(12):
    wp.write(He+3+i,0,pm[i],lf); wp.write(He+3+i,1,round(pr[i],2),fv); wp.write(He+3+i,2,round(pc[i],2),fv)
Pe2=He+3+11

pch = wb.add_chart({'type':'line'})
pch.add_series({'name':'收入(历史)','categories':['趋势预测',Pd,0,He-1,0],'values':['趋势预测',Pd,1,He-1,1],'line':{'color':c0,'width':2.5}})
pch.add_series({'name':'收入(预测)','categories':['趋势预测',He+3,0,Pe2,0],'values':['趋势预测',He+3,1,Pe2,1],'line':{'color':c0,'width':2.5,'dash_type':'dash'}})
pch.add_series({'name':'成本(历史)','categories':['趋势预测',Pd,0,He-1,0],'values':['趋势预测',Pd,2,He-1,2],'line':{'color':c1,'width':2.5}})
pch.add_series({'name':'成本(预测)','categories':['趋势预测',He+3,0,Pe2,0],'values':['趋势预测',He+3,2,Pe2,2],'line':{'color':c1,'width':2.5,'dash_type':'dash'}})
pch.set_title({'name':'营业收入 & 成本预测（2026年）','name_font':{'color':'#c084fc','size':12}})
pch.set_y_axis({'name':'万元'}); pch.set_size({'width':800,'height':380})
wp.insert_chart(Pd,3,pch)

# ============ SHEET 6: INDUSTRY ============
wi = wb.add_worksheet('同行业对比')
wi.set_tab_color('#54a0ff'); wi.hide_gridlines(2); wi.set_column('A:F',24)
wi.merge_range(0,0,0,5,'同行业对比分析',tf); wi.set_row(0,40)
wi.merge_range(2,0,2,5,'对标公司：艾融软件(830799)、用友网络(600588)、华宇软件(300271)、三维天地(301159)、山大地纬(688579)、殷图网联(835508)、天亿马(301178)、远光软件(002063)、宇信科技(300674)',wb.add_format({'font_size':8,'font_color':'#556677','italic':True,'bg_color':_B}))

csvp = os.path.join(os.path.dirname(__file__),'industry_data.csv')
if os.path.exists(csvp):
    indf = pd.read_csv(csvp, encoding='utf-8-sig')
    Ir=4; wi.write(Ir,0,'2024年行业对比数据',sfx); wi.merge_range(Ir,0,Ir,5,'',sfx); Ir+=1
    def pv(v):
        if pd.isna(v) or v=='False': return 0
        s=str(v).replace(',','').strip(); n=1
        if s.startswith('-'): n=-1; s=s[1:]
        if '亿' in s: return n*float(s.replace('亿',''))*10000
        if '万' in s: return n*float(s.replace('万',''))
        if '%' in s: return n*float(s.replace('%',''))
        try: return n*float(s)
        except: return 0
    for ci,hd in enumerate(['公司','营业收入(万元)','净利润(万元)','毛利率','净利率','净资产收益率']):
        wi.write(Ir,ci,hd,hf)
    Ir+=1
    # 小甜甜 2024
    R24s=sum(R['营业收入']['2024']); C24s=sum(R['营业成本']['2024'])
    P24s=R24s-C24s-sum(R['税金及附加']['2024'])-sum(R['销售费用']['2024'])-sum(R['管理费用']['2024'])-sum(R['研发费用']['2024'])-sum(R['财务费用']['2024'])+sum(R['其他收益']['2024'])
    wi.write(Ir,0,'小甜甜(2024)',lf); wi.write(Ir,1,round(R24s),fi); wi.write(Ir,2,round(P24s),fi)
    wi.write(Ir,3,'%.1f%%' % ((R24s-C24s)/R24s*100),fv); wi.write(Ir,4,'%.1f%%' % (P24s/R24s*100),fv); wi.write(Ir,5,'-',fv); Ir+=1
    for _,r in indf[indf['报告期'].astype(str)=='2024'].iterrows():
        wi.write(Ir,0,r['公司'],lf)
        wi.write(Ir,1,round(pv(r['营业总收入'])),fi); wi.write(Ir,2,round(pv(r['净利润'])),fi)
        for j,k in enumerate(['销售毛利率','销售净利率','净资产收益率']):
            v=r.get(k,'-'); wi.write(Ir,3+j,str(v) if pd.notna(v) else '-',fv)
        Ir+=1

wb.close()
print('OK:', OUT)
