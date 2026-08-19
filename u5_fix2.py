# -*- coding: utf-8 -*-
"""u5_fix2: 针对 03 共享公式块(sheet 内 502+ 行原为同一 $BQ$502 共享公式)，
先 ClearContents 打破共享，再标量填充。11 AZ / 06 L 已正确，仍重做以保证一致。
"""
import sys, io, time, shutil, os
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SRC = r"D:\倍优跟单进度表（模板）_v8.xlsx"
DST = r"D:\倍优跟单进度表（模板）_v9.xlsx"
t0 = time.time()
def log(m):
    s = "[%6.1fs] %s" % (time.time()-t0, m); print(s); sys.stdout.flush()

shutil.copy2(SRC, DST); log("复制可写目标 -> %s (覆盖)" % DST)
app = win32.DispatchEx("Ket.Application")
app.Visible = False; app.DisplayAlerts = False
try: app.AskToUpdateLinks = False
except Exception: pass
wb = app.Workbooks.Open(DST)
if wb.ReadOnly:
    log("!! 目标只读 中止"); wb.Close(False); app.Quit(); sys.exit(1)
app.Calculation = -4135
wb.Application.Calculation = -4135
log("已打开 v9，计算模式=手动")

# 03 M/N: 先清后填
ws3 = wb.Worksheets("03工序进度")
ws3.Range("M3:N13060").ClearContents()
log("03 M/N ClearContents 完成")
m3rel = ('=IF($J3=0,"已交付",IFERROR(INDEX(INDEX(\'08参数-字典\'!$B$15:$J$20,'
         'MATCH($E3,\'08参数-字典\'!$A$15:$A$20,0),0),MATCH(0,$BQ3:$BY3,0)),"全部完成"))')
n3rel = ('=IF($J3=0,"全部完成",IF($K3>=$J3,"全部完成",'
         'IF(INDEX($AY3:$BG3,IFERROR(MATCH(0,$BQ3:$BY3,0),1))<TODAY(),"工序延期",'
         'IF(INDEX($AP3:$AX3,IFERROR(MATCH(0,$BQ3:$BY3,0),1))<=TODAY(),"进行中","未开工"))))')
ws3.Range("M3:M13060").Formula = m3rel
ws3.Range("N3:N13060").Formula = n3rel
log("03 M/N 填充完成")

# 11 AZ
ws11 = wb.Worksheets("11合同冗余")
azrel = ("=IF($AV11=\"\",\"\",IF(OR('原始数据表-畅享'!$G3=\"\",'原始数据表-畅享'!$G3=\"1899-12-30\"),\"\","
         "IF(ISNUMBER('原始数据表-畅享'!$G3),'原始数据表-畅享'!$G3,DATEVALUE('原始数据表-畅享'!$G3))))")
ws11.Range("AZ11:AZ13068").Formula = azrel

# 06 L
ws6 = wb.Worksheets("06合同级汇总")
ws6.Range("L5:L2958").Formula = "=IF($K5=\"\",\"\",$K5-TODAY())"

def fget(ws, a):
    try: return ws.Range(a).Formula
    except Exception as e: return "ERR:%s" % e
log("抽样 03 M575  =%s" % fget(ws3, "M575"))
log("抽样 03 N3    =%s" % fget(ws3, "N3"))
log("抽样 03 N500  =%s" % fget(ws3, "N500"))
log("抽样 03 N501  =%s" % fget(ws3, "N501"))
log("抽样 03 N502  =%s" % fget(ws3, "N502"))
log("抽样 03 N575  =%s" % fget(ws3, "N575"))
log("抽样 03 N13060=%s" % fget(ws3, "N13060"))
log("抽样 11 AZ12387=%s" % fget(ws11, "AZ12387"))

mt_before = os.path.getmtime(DST)
wb.Save()
mt_after = os.path.getmtime(DST)
log("保存 %s -> %s %s" % (time.strftime("%H:%M:%S", time.localtime(mt_before)),
                          time.strftime("%H:%M:%S", time.localtime(mt_after)),
                          "✔" if mt_after > mt_before else "✘未写入"))
wb.Close(False); app.Quit()
log("完成")
