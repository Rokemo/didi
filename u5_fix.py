# -*- coding: utf-8 -*-
"""u5_fix (标量相对公式整列填充版): 修复 v8 复核发现的 4 类公式错误 -> v9
WPS COM 的二维数组赋值会错位/转置，故改用：对整列范围赋【单个相对公式字符串】，
由 WPS 原生向下填充(相对引用自动递增)。计算模式 Open 后设手动。
"""
import sys, io, time, shutil, os
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SRC = r"D:\倍优跟单进度表（模板）_v8.xlsx"
DST = r"D:\倍优跟单进度表（模板）_v9.xlsx"
BAK = r"D:\倍优跟单进度表（模板）_v8_bak_before_u5.xlsx"
t0 = time.time()
def log(m):
    s = "[%6.1fs] %s" % (time.time()-t0, m); print(s); sys.stdout.flush()

shutil.copy2(SRC, BAK); log("备份源 -> %s" % BAK)
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

# 1) 03 M (相对公式，整列标量赋值)
ws3 = wb.Worksheets("03工序进度")
m3rel = ('=IF($J3=0,"已交付",IFERROR(INDEX(INDEX(\'08参数-字典\'!$B$15:$J$20,'
         'MATCH($E3,\'08参数-字典\'!$A$15:$A$20,0),0),MATCH(0,$BQ3:$BY3,0)),"全部完成"))')
n3rel = ('=IF($J3=0,"全部完成",IF($K3>=$J3,"全部完成",'
         'IF(INDEX($AY3:$BG3,IFERROR(MATCH(0,$BQ3:$BY3,0),1))<TODAY(),"工序延期",'
         'IF(INDEX($AP3:$AX3,IFERROR(MATCH(0,$BQ3:$BY3,0),1))<=TODAY(),"进行中","未开工"))))')
ws3.Range("M3:M13060").Formula = m3rel
ws3.Range("N3:N13060").Formula = n3rel
log("03 M/N 整列标量填充完成")

# 2) 11 AZ
ws11 = wb.Worksheets("11合同冗余")
azrel = ("=IF($AV11=\"\",\"\",IF(OR('原始数据表-畅享'!$G3=\"\",'原始数据表-畅享'!$G3=\"1899-12-30\"),\"\","
         "IF(ISNUMBER('原始数据表-畅享'!$G3),'原始数据表-畅享'!$G3,DATEVALUE('原始数据表-畅享'!$G3))))")
ws11.Range("AZ11:AZ13068").Formula = azrel
log("11 AZ 整列标量填充完成")

# 3) 06 L
ws6 = wb.Worksheets("06合同级汇总")
l5rel = "=IF($K5=\"\",\"\",$K5-TODAY())"
ws6.Range("L5:L2958").Formula = l5rel
log("06 L 整列标量填充完成")

def fget(ws, a):
    try: return ws.Range(a).Formula
    except Exception as e: return "ERR:%s" % e
log("抽样 03 N3    =%s" % fget(ws3, "N3"))
log("抽样 03 N575  =%s" % fget(ws3, "N575"))
log("抽样 03 N13060=%s" % fget(ws3, "N13060"))
log("抽样 11 AZ11    =%s" % fget(ws11, "AZ11"))
log("抽样 11 AZ12387 =%s" % fget(ws11, "AZ12387"))
log("抽样 11 AZ13068 =%s" % fget(ws11, "AZ13068"))
log("抽样 06 L5    =%s" % fget(ws6, "L5"))
log("抽样 06 L2958 =%s" % fget(ws6, "L2958"))

mt_before = os.path.getmtime(DST)
wb.Save()
mt_after = os.path.getmtime(DST)
log("保存 %s -> %s %s" % (time.strftime("%H:%M:%S", time.localtime(mt_before)),
                          time.strftime("%H:%M:%S", time.localtime(mt_after)),
                          "✔" if mt_after > mt_before else "✘未写入"))
wb.Close(False); app.Quit()
log("完成")
