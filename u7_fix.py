# -*- coding: utf-8 -*-
"""u7: per-cell absolute rewrite of 03 M/N rows 503..3000 to break the
WPS shared-formula block (anchor $502) that is immune to array/range writes.
Per-cell unique absolute formulas -> WPS cannot re-share on save."""
import sys, io, time
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
SRC = r"D:\倍优跟单进度表（模板）_v9.xlsx"
t0 = time.time()
app = win32.DispatchEx("Ket.Application")
app.Visible = False; app.DisplayAlerts = False
try: app.AskToUpdateLinks = False
except Exception: pass
wb = app.Workbooks.Open(SRC)
if wb.ReadOnly:
    print("!! 只读中止"); wb.Close(False); app.Quit(); sys.exit(1)
wb.Application.Calculation = -4135
wb.Application.ScreenUpdating = False
wb.Application.EnableEvents = False
ws = wb.Worksheets("03工序进度")
R0, R1 = 503, 3000
m_tpl = lambda r: ('=IF($J%d=0,"已交付",IFERROR(INDEX(INDEX(\'08参数-字典\'!$B$15:$J$20,'
            'MATCH($E%d,\'08参数-字典\'!$A$15:$A$20,0),0),MATCH(0,$BQ%d:$BY%d,0)),"全部完成"))'
            % (r, r, r, r))
n_tpl = lambda r: ('=IF($J%d=0,"全部完成",IF($K%d>=$J%d,"全部完成",'
            'IF(INDEX($AY%d:$BG%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<TODAY(),"工序延期",'
            'IF(INDEX($AP%d:$AX%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<=TODAY(),"进行中","未开工"))))'
            % (r, r, r, r, r, r, r, r, r, r, r))
done = 0
for r in range(R0, R1 + 1):
    ws.Range("M%d" % r).Formula = m_tpl(r)
    ws.Range("N%d" % r).Formula = n_tpl(r)
    done += 1
    if (r - R0 + 1) % 200 == 0:
        print("[%5.1fs] %d/%d  r=%d" % (time.time() - t0, done, R1 - R0 + 1, r))
print("[%5.1fs] per-cell done %d..%d" % (time.time() - t0, R0, R1))
wb.Application.Calculation = -4105
wb.Save()
print("[%5.1fs] saved" % (time.time() - t0))
wb.Close(False); app.Quit()
print("[%5.1fs] done" % (time.time() - t0))
