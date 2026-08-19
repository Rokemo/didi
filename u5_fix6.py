# -*- coding: utf-8 -*-
"""u5_fix6: 破解 03 M/N 共享公式块(503..13060) 的完整快速修复。
方法：按 400 行分块；每块先逐格写首格(相对公式，破除共享锚点)，再整块相对填充。
每块 <500 行，规避 WPS 大范围填充的共享公式 stride 漂移(500行后错位)。
"""
import sys, io, time, shutil, os
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SRC = r"D:\倍优跟单进度表（模板）_v8.xlsx"
DST = r"D:\倍优跟单进度表（模板）_v9.xlsx"
R0, R1, CH = 503, 13060, 400
t0 = time.time()
def log(m):
    s = "[%6.1fs] %s" % (time.time()-t0, m); print(s); sys.stdout.flush()

shutil.copy2(SRC, DST); log("复制可写目标 -> %s" % DST)
app = win32.DispatchEx("Ket.Application")
app.Visible = False; app.DisplayAlerts = False
app.ScreenUpdating = False; app.EnableEvents = False
try: app.AskToUpdateLinks = False
except Exception: pass
wb = app.Workbooks.Open(DST)
if wb.ReadOnly:
    log("!! 只读 中止"); wb.Close(False); app.Quit(); sys.exit(1)
app.Calculation = -4135; wb.Application.Calculation = -4135
log("已打开 v9")

ws3 = wb.Worksheets("03工序进度")
def m_rel(r):
    return ('=IF($J%d=0,"已交付",IFERROR(INDEX(INDEX(\'08参数-字典\'!$B$15:$J$20,'
            'MATCH($E%d,\'08参数-字典\'!$A$15:$A$20,0),0),MATCH(0,$BQ%d:$BY%d,0)),"全部完成"))') % (r, r, r, r)
def n_rel(r):
    return ('=IF($J%d=0,"全部完成",IF($K%d>=$J%d,"全部完成",'
            'IF(INDEX($AY%d:$BG%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<TODAY(),"工序延期",'
            'IF(INDEX($AP%d:$AX%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<=TODAY(),"进行中","未开工"))))') % (
            r, r, r, r, r, r, r, r, r, r, r)

nchunks = 0
for s in range(R0, R1+1, CH):
    e = min(s+CH-1, R1)
    # 逐格写首格(相对公式，行号=s)破除共享锚点
    ws3.Range("M%d" % s).Formula = m_rel(s)
    ws3.Range("N%d" % s).Formula = n_rel(s)
    # 整块相对填充(块<500行，无 stride 漂移)
    ws3.Range("M%d:M%d" % (s, e)).Formula = m_rel(s)
    ws3.Range("N%d:N%d" % (s, e)).Formula = n_rel(s)
    nchunks += 1
log("分块完成: %d 块" % nchunks)

def fget(ws, a):
    try: return ws.Range(a).Formula
    except Exception as e: return "ERR:%s" % e
for a in ("N503","N575","N902","N903","N5030","N13060","M575","M5030"):
    log("%s = %s" % (a, fget(ws3, a)[:52]))

mt_before = os.path.getmtime(DST)
wb.Save()
mt_after = os.path.getmtime(DST)
log("保存 %s -> %s %s" % (time.strftime("%H:%M:%S", time.localtime(mt_before)),
                          time.strftime("%H:%M:%S", time.localtime(mt_after)),
                          "✔" if mt_after > mt_before else "✘未写入"))
wb.Close(False); app.Quit()
log("完成")
