# -*- coding: utf-8 -*-
import sys, io, time, shutil
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
SRC = r"D:\倍优跟单进度表（模板）_v8.xlsx"
DST = r"D:\倍优跟单进度表（模板）_v9.xlsx"
t0=time.time()
def log(m): print("[%6.1fs] %s"%(time.time()-t0,m)); sys.stdout.flush()
shutil.copy2(SRC, DST)
app = win32.DispatchEx("Ket.Application"); app.Visible=False; app.DisplayAlerts=False
app.ScreenUpdating=False; app.EnableEvents=False
wb = app.Workbooks.Open(DST)
if wb.ReadOnly: log("只读"); wb.Close(False); app.Quit(); sys.exit(1)
app.Calculation=-4135
ws3 = wb.Worksheets("03工序进度")
# 先清空损坏块公式
ws3.Range("M503:N13060").Value = ""
log("已清空 M503:N13060")
# 100行相对填充测试(仅前100行)
m3rel = lambda r: ('=IF($J%d=0,"已交付",IFERROR(INDEX(INDEX(\'08参数-字典\'!$B$15:$J$20,'
         'MATCH($E%d,\'08参数-字典\'!$A$15:$A$20,0),0),MATCH(0,$BQ%d:$BY%d,0)),"全部完成"))')%(r,r,r,r)
n3rel = lambda r: ('=IF($J%d=0,"全部完成",IF($K%d>=$J%d,"全部完成",'
         'IF(INDEX($AY%d:$BG%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<TODAY(),"工序延期",'
         'IF(INDEX($AP%d:$AX%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<=TODAY(),"进行中","未开工"))))')%(r,r,r,r,r,r,r,r,r,r,r)
ws3.Range("M503:M602").Formula = m3rel(503)
ws3.Range("N503:N602").Formula = n3rel(503)
log("100行填充完成")
def f(a):
    try: return ws3.Range(a).Formula
    except Exception as e: return "ERR:%s"%e
for a in ("N520","N550","N575","N602","M575"):
    log("%s = %s"%(a, f(a)))
wb.Close(False); app.Quit(); log("done")
