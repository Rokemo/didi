import win32com.client as w
OUT = r"D:\倍优跟单进度表（模板）_v10.xlsx"
app=w.Dispatch("Ket.Application"); app.Visible=False; app.DisplayAlerts=False
wb=app.Workbooks.Open(OUT)
ws=wb.Worksheets("00使用说明")
# find an empty row after used range
r=80
while ws.Range("A%d"%r).Value not in (None,""): r+=1
ws.Range("A%d"%r).Value="【v10 性能优化说明】"
ws.Range("A%d"%(r+1)).Value=("1) 全部整列引用(如 A:A)已收口为 A2:A100000，每个 SUMIFS 扫描行数从 104 万降到 10 万（约 10 倍）；"
    "2) 4.2 万处 TODAY() 已统一引用本表 C2（今日日期），全表仅 1 处易失函数；"
    "3) 经 WPS 全量重算复核，16 张表公式错误合计=0，KPI 不变。"
    "注意：若原始数据超过 10 万行，请把各表求和上限 100000 同步调大。")
ws.Range("A%d"%r).Font.Bold=True
wb.Save()
wb.Close(False); app.Quit(); print("note added at row", r, "+1; saved v10")
