# -*- coding: utf-8 -*-
"""FINAL fix (u6): rewrite broken formula columns with ABSOLUTE formulas
(row number baked into every reference) so WPS cannot re-share them.

Targets (from v9_verify: 1377 errors):
  1) 03工序进度 M/N  rows 503..13060  (shared-formula block anchored at 502)
  2) 11合同冗余 AZ   rows 11..13068   (text zero-date "1899-12-30" -> #VALUE! cascade)
  3) 06合同级汇总 L   rows 5..2958     (blank total-row K -> #VALUE!)

01驾驶舱 B154/B155 are =SUM('11合同冗余'!E5/E6) -> resolve once 11 cascade fixed.
"""
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
    print("!! 文件只读 中止（可能被 WPS 打开）"); wb.Close(False); app.Quit(); sys.exit(1)
wb.Application.Calculation = -4135        # manual calc -> fast writes
wb.Application.ScreenUpdating = False
wb.Application.EnableEvents = False
print("[%5.1fs] opened v9 (calc=manual)" % (time.time() - t0))

# ---- formula builders (absolute, row baked in) ----
def m_f(r):
    return ('=IF($J%d=0,"已交付",IFERROR(INDEX(INDEX(\'08参数-字典\'!$B$15:$J$20,'
            'MATCH($E%d,\'08参数-字典\'!$A$15:$A$20,0),0),MATCH(0,$BQ%d:$BY%d,0)),"全部完成"))'
            % (r, r, r, r))
def n_f(r):
    return ('=IF($J%d=0,"全部完成",IF($K%d>=$J%d,"全部完成",'
            'IF(INDEX($AY%d:$BG%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<TODAY(),"工序延期",'
            'IF(INDEX($AP%d:$AX%d,IFERROR(MATCH(0,$BQ%d:$BY%d,0),1))<=TODAY(),"进行中","未开工"))))'
            % (r, r, r, r, r, r, r, r, r, r, r))
def az_f(r):
    g = r - 8
    return ('=IF($AV%d="","",IF(OR(\'原始数据表-畅享\'!$G%d="",\'原始数据表-畅享\'!$G%d="1899-12-30"),"",'
            'IF(ISNUMBER(\'原始数据表-畅享\'!$G%d),\'原始数据表-畅享\'!$G%d,DATEVALUE(\'原始数据表-畅享\'!$G%d))))'
            % (r, g, g, g, g, g))
def l_f(r):
    return '=IF($K%d="","",$K%d-TODAY())' % (r, r)

def chunk_write_2col(ws, colA, colB, r0, r1, fa, fb, label, CH=2000):
    n = r1 - r0 + 1
    done = 0
    for s in range(r0, r1 + 1, CH):
        e = min(s + CH - 1, r1)
        arr = [[fa(r), fb(r)] for r in range(s, e + 1)]
        ws.Range("%s%d:%s%d" % (colA, s, colB, e)).Formula = arr
        done += (e - s + 1)
        print("[%5.1fs] %s %d..%d (%d/%d)" % (time.time() - t0, label, s, e, done, n))
    print("[%5.1fs] %s DONE %d..%d" % (time.time() - t0, label, r0, r1))

def chunk_write_1col(ws, col, r0, r1, f, label, CH=2000):
    n = r1 - r0 + 1
    done = 0
    for s in range(r0, r1 + 1, CH):
        e = min(s + CH - 1, r1)
        arr = [[f(r)] for r in range(s, e + 1)]
        ws.Range("%s%d:%s%d" % (col, s, col, e)).Formula = arr
        done += (e - s + 1)
        print("[%5.1fs] %s %d..%d (%d/%d)" % (time.time() - t0, label, s, e, done, n))
    print("[%5.1fs] %s DONE %d..%d" % (time.time() - t0, label, r0, r1))

ws3 = wb.Worksheets("03工序进度")
chunk_write_2col(ws3, "M", "N", 503, 13060, m_f, n_f, "03 M/N")

ws11 = wb.Worksheets("11合同冗余")
chunk_write_1col(ws11, "AZ", 11, 13068, az_f, "11 AZ")

ws6 = wb.Worksheets("06合同级汇总")
chunk_write_1col(ws6, "L", 5, 2958, l_f, "06 L")

print("[%5.1fs] saving..." % (time.time() - t0))
wb.Application.Calculation = -4105   # automatic so next open recalcs
wb.Save()
print("[%5.1fs] saved v9" % (time.time() - t0))
wb.Close(False); app.Quit()
print("[%5.1fs] done" % (time.time() - t0))
