# -*- coding: utf-8 -*-
"""u5_fix4: 逐格写入 03 M/N 的损坏共享公式块(503..13060)。
WPS ET 格式下，共享公式块无法用整列/分块填充覆盖，只能逐格写破除。
优化：ScreenUpdating/EnableEvents 关闭 + 手动计算。
"""
import sys, io, time, shutil, os
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SRC = r"D:\倍优跟单进度表（模板）_v8.xlsx"
DST = r"D:\倍优跟单进度表（模板）_v9.xlsx"
R0, R1 = 503, 13060   # 损坏块
t0 = time.time()
def log(m):
    s = "[%6.1fs] %s" % (time.time()-t0, m); print(s); sys.stdout.flush()

shutil.copy2(SRC, DST); log("复制可写目标 -> %s (覆盖)" % DST)
app = win32.DispatchEx("Ket.Application")
app.Visible = False; app.DisplayAlerts = False
app.ScreenUpdating = False; app.EnableEvents = False
try: app.AskToUpdateLinks = False
except Exception: pass
wb = app.Workbooks.Open(DST)
if wb.ReadOnly:
    log("!! 目标只读 中止"); wb.Close(False); app.Quit(); sys.exit(1)
app.Calculation = -4135
wb.Application.Calculation = -4135
log("已打开 v9，计算模式=手动，屏显/事件关闭")

ws3 = wb.Worksheets("03工序进度")
def m_f(r):
    return ('=IF($J%d=0,"已交付",IFERROR(INDEX(INDEX(\'08参数-字典\'!$B$15:$J$20,'
            'MATCH($E%d,\'08参数-字典\'!$A$15:$A$20,0),0),MATCH(0,$BQ%d:$BY%d,0)),"全部完成"))') % (r, r, r, r)
def n_f(r):
    return ('=IF($J%d=0,"全部完成",IF($K%d>=$J%d,"全部完成",'
            'IF(INDEX($AY%d:$BG%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<TODAY(),"工序延期",'
            'IF(INDEX($AP%d:$AX%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<=TODAY(),"进行中","未开工"))))') % (
            r, r, r, r, r, r, r, r, r, r, r)

total = (R1-R0+1)*2
done = 0
for r in range(R0, R1+1):
    ws3.Range("M%d" % r).Formula = m_f(r)
    ws3.Range("N%d" % r).Formula = n_f(r)
    done += 2
    if done % 2000 == 0:
        log("  已写 %d/%d" % (done, total))
log("03 M/N 损坏块逐格写入完成")

def fget(ws, a):
    try: return ws.Range(a).Formula
    except Exception as e: return "ERR:%s" % e
log("抽样 03 N503  =%s" % fget(ws3, "N503"))
log("抽样 03 N575  =%s" % fget(ws3, "N575"))
log("抽样 03 N5030 =%s" % fget(ws3, "N5030"))
log("抽样 03 N13060=%s" % fget(ws3, "N13060"))

mt_before = os.path.getmtime(DST)
wb.Save()
mt_after = os.path.getmtime(DST)
log("保存 %s -> %s %s" % (time.strftime("%H:%M:%S", time.localtime(mt_before)),
                          time.strftime("%H:%M:%S", time.localtime(mt_after)),
                          "✔" if mt_after > mt_before else "✘未写入"))
wb.Close(False); app.Quit()
log("完成")
