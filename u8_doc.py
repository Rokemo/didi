# -*- coding: utf-8 -*-
"""u8: refresh stale hardcoded numbers in 00使用说明 to current 13058-row baseline."""
import sys, io, time
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
SRC = r"D:\倍优跟单进度表（模板）_v9.xlsx"
t0 = time.time()
app = win32.DispatchEx("Ket.Application"); app.Visible=False; app.DisplayAlerts=False
try: app.AskToUpdateLinks=False
except Exception: pass
wb = app.Workbooks.Open(SRC)
if wb.ReadOnly:
    print("!! 只读中止"); wb.Close(False); app.Quit(); sys.exit(1)
ws = wb.Worksheets("00使用说明")
updates = {
    "A1": "跨境采购跟单进度表 v9 · 使用说明（数据基准 2026-08-06 · 全量 13,058 行）",
    "A50": "🔴 高 | 产品资料表「工艺品类」尚未填写 → 类目全待确认 | 本次导入的产品资料表 1,678 行的「工艺品类」列为空。类目不再由系统推断，需你在「产品资料表」按产品编号补全工艺品类（可下拉选标准类目），02/03 才会出排期，否则整列显示「待确认」。",
    "A51": "🔴 高 | 大量行逾期，其中多行逾期超过 180 天 | 最久的已挂账 1,314 天。这类尾数订单挂在表上没有意义，建议单独开清尾评审（核销 / 退单 / 转内销）。",
    "A56": "🟠 中 | 大量行交货日期在未来 → 交付周期分档以「未到期」为主 | 原始数据 13,058 行中大量行交货日期晚于今天，交付周期=基准日−交货日期 为负，判为「未到期」（在途正常单）。但已有相当部分行超过 90 天阈值被判为「冗余」（冗余 31,292 件 / 13.86 万元，占比 30.7%），需优先催办；待更多订单实际逾期后，①分档与⑤接单判断会逐步体现占压。",
}
for addr, txt in updates.items():
    ws.Range(addr).Value = txt
    print("[%5.1fs] set %s" % (time.time() - t0, addr))
wb.Save()
print("[%5.1fs] saved" % (time.time() - t0))
wb.Close(False); app.Quit()
print("[%5.1fs] done" % (time.time() - t0))
