# -*- coding: utf-8 -*-
"""u5_fix3: 分块(<=400行)标量填充，规避 WPS 对 >~500 行相对填充生成共享公式(锚定502)的缺陷。
修复 03 M/N、11 AZ、06 L。
"""
import sys, io, time, shutil, os
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SRC = r"D:\倍优跟单进度表（模板）_v8.xlsx"
DST = r"D:\倍优跟单进度表（模板）_v9.xlsx"
CHUNK = 400
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

def fill_col(ws, col, r0, r1, rel_formula):
    n = r1 - r0 + 1; done = 0
    for s in range(r0, r1+1, CHUNK):
        e = min(s+CHUNK-1, r1)
        # 为块首行构造相对公式(行号=s)，整块标量赋值(块<阈值，不触发共享)
        f = rel_formula(s)
        ws.Range("%s%d:%s%d" % (col, s, col, e)).Formula = f
        done += (e-s+1)
    log("  %s 已写 %d/%d" % (col, done, n))

# 03 M/N
ws3 = wb.Worksheets("03工序进度")
m3rel = lambda r: ('=IF($J%d=0,"已交付",IFERROR(INDEX(INDEX(\'08参数-字典\'!$B$15:$J$20,'
         'MATCH($E%d,\'08参数-字典\'!$A$15:$A$20,0),0),MATCH(0,$BQ%d:$BY%d,0)),"全部完成"))') % (r, r, r, r)
n3rel = lambda r: ('=IF($J%d=0,"全部完成",IF($K%d>=$J%d,"全部完成",'
         'IF(INDEX($AY%d:$BG%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<TODAY(),"工序延期",'
         'IF(INDEX($AP%d:$AX%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<=TODAY(),"进行中","未开工"))))') % (
         r, r, r, r, r, r, r, r, r, r, r)
log("03 M ..."); fill_col(ws3, "M", 3, 13060, m3rel)
log("03 N ..."); fill_col(ws3, "N", 3, 13060, n3rel)

# 11 AZ
ws11 = wb.Worksheets("11合同冗余")
azrel = lambda r: ("=IF($AV%d=\"\",\"\",IF(OR('原始数据表-畅享'!$G%d=\"\",'原始数据表-畅享'!$G%d=\"1899-12-30\"),\"\","
         "IF(ISNUMBER('原始数据表-畅享'!$G%d),'原始数据表-畅享'!$G%d,DATEVALUE('原始数据表-畅享'!$G%d))))") % (r, r-8, r-8, r-8, r-8, r-8)
log("11 AZ ..."); fill_col(ws11, "AZ", 11, 13068, azrel)

# 06 L
ws6 = wb.Worksheets("06合同级汇总")
lrel = lambda r: "=IF($K%d=\"\",\"\",$K%d-TODAY())" % (r, r)
log("06 L ..."); fill_col(ws6, "L", 5, 2958, lrel)

def fget(ws, a):
    try: return ws.Range(a).Formula
    except Exception as e: return "ERR:%s" % e
log("抽样 03 M575  =%s" % fget(ws3, "M575"))
log("抽样 03 N500  =%s" % fget(ws3, "N500"))
log("抽样 03 N502  =%s" % fget(ws3, "N502"))
log("抽样 03 N575  =%s" % fget(ws3, "N575"))
log("抽样 03 N5030 =%s" % fget(ws3, "N5030"))
log("抽样 03 N13060=%s" % fget(ws3, "N13060"))
log("抽样 11 AZ12387=%s" % fget(ws11, "AZ12387"))
log("抽样 06 L2958 =%s" % fget(ws6, "L2958"))

mt_before = os.path.getmtime(DST)
wb.Save()
mt_after = os.path.getmtime(DST)
log("保存 %s -> %s %s" % (time.strftime("%H:%M:%S", time.localtime(mt_before)),
                          time.strftime("%H:%M:%S", time.localtime(mt_after)),
                          "✔" if mt_after > mt_before else "✘未写入"))
wb.Close(False); app.Quit()
log("完成")
