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
app.Calculation=-4135; wb.Application.Calculation=-4135
ws3 = wb.Worksheets("03工序进度")
ws3.EnableCalculation = False   # 关键：禁止本表重算
log("EnableCalculation=False 已设")
n3rel = ('=IF($J%d=0,"全部完成",IF($K%d>=$J%d,"全部完成",'
         'IF(INDEX($AY%d:$BG%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<TODAY(),"工序延期",'
         'IF(INDEX($AP%d:$AX%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<=TODAY(),"进行中","未开工"))))')%(
         503,503,503,503,503,503,503,503,503,503,503)
t1=time.time()
for r in range(503, 603):
    ws3.Range("N%d"%r).Formula = n3rel
dt = time.time()-t1
log("逐格写 100 格耗时 %.1fs (%.1f ms/格)" % (dt, dt*1000/100))
def f(a):
    try: return ws3.Range(a).Formula
    except Exception as e: return "ERR:%s"%e
for a in ("N520","N575","N602"):
    log("%s = %s"%(a, f(a)[:55]))
wb.Close(False); app.Quit(); log("done")
