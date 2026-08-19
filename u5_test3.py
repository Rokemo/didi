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
m3rel = ('=IF($J3=0,"已交付",IFERROR(INDEX(INDEX(\'08参数-字典\'!$B$15:$J$20,'
         'MATCH($E3,\'08参数-字典\'!$A$15:$A$20,0),0),MATCH(0,$BQ3:$BY3,0)),"全部完成"))')
n3rel = ('=IF($J3=0,"全部完成",IF($K3>=$J3,"全部完成",'
         'IF(INDEX($AY3:$BG3,IFERROR(MATCH(0,$BQ3:$BY3,0),1))<TODAY(),"工序延期",'
         'IF(INDEX($AP3:$AX3,IFERROR(MATCH(0,$BQ3:$BY3,0),1))<=TODAY(),"进行中","未开工"))))')
# 仅逐格写首格(503)破除共享锚点，再整列相对填充
ws3.Range("M503").Formula = m3rel  # 相对公式(行号3，填充时按503基准调整)
ws3.Range("N503").Formula = n3rel
ws3.Range("M503:M13060").Formula = m3rel
ws3.Range("N503:N13060").Formula = n3rel
log("首格破除+整列填充完成")
def f(a):
    try: return ws3.Range(a).Formula
    except Exception as e: return "ERR:%s"%e
for a in ("N503","N520","N575","N5030","N13060","M575"):
    log("%s = %s"%(a, f(a)[:56]))
wb.Close(False); app.Quit(); log("done")
