# -*- coding: utf-8 -*-
"""补齐 02订单明细 上没生效的写入，并逐步回读定位原因"""
import sys, io, os, time
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

DST = r"D:\倍优跟单进度表（模板）_v8.xlsx"
RAW_LAST = 13060
t0 = time.time()
def log(m):
    print("[%6.1fs] %s" % (time.time() - t0, m)); sys.stdout.flush()

M0 = os.path.getmtime(DST)
app = win32.DispatchEx("Ket.Application")
app.Visible = False; app.DisplayAlerts = False
try: app.AskToUpdateLinks = False
except Exception: pass
app.ScreenUpdating = False
wb = app.Workbooks.Open(DST)
app.Calculation = -4135
log("打开 v8 只读=%s 计算=%s" % (wb.ReadOnly, app.Calculation))

ws = wb.Worksheets("02订单明细")
log("--- 现状 ---")
log("  保护=%s  行数=%d 列数=%d" % (ws.ProtectContents, ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count))
for a in ("A3", "A2629", "A13060", "AA2", "AA3", "P3", "P13060", "Q3", "E3", "H3", "J3"):
    log("  %-8s = %s" % (a, str(ws.Range(a).Formula)[:80]))

log("--- 试写单格 AA3 ---")
try:
    ws.Range("AA3").Formula = "=1+1"
    log("  写后回读 AA3 = %s" % ws.Range("AA3").Formula)
except Exception as e:
    log("  单格写异常: %s" % e)

log("--- 试写小区间 P3:P5 ---")
try:
    ws.Range("P3:P5").Formula = '=IFERROR($O3-$N3,"")'
    log("  回读 P3=%s P5=%s" % (ws.Range("P3").Formula, ws.Range("P5").Formula))
except Exception as e:
    log("  小区间写异常: %s" % e)


def write_chunk(col, r0, r1, fml, step=2000):
    """分块写，逐块回读确认"""
    bad = 0
    for s in range(r0, r1 + 1, step):
        e = min(s + step - 1, r1)
        rng = "%s%d:%s%d" % (col, s, col, e)
        try:
            ws.Range(rng).Formula = fml.replace("@R", str(s))
        except Exception as ex:
            log("  %s 写异常 %s" % (rng, ex)); bad += 1; continue
        back = ws.Range("%s%d" % (col, s)).Formula
        if not (isinstance(back, str) and back.startswith("=")):
            log("  %s 回读失败: %r" % (rng, back)); bad += 1
    return bad


log("--- 分块重写 P / Q ---")
b1 = write_chunk("P", 3, RAW_LAST, '=IFERROR($O@R-$N@R,"")')
b2 = write_chunk("Q", 3, RAW_LAST, '=IFERROR($O@R-TODAY(),"")')
log("  P 失败块 %d，Q 失败块 %d；回读 P3=%s P13060=%s" % (
    b1, b2, ws.Range("P3").Formula, ws.Range("P13060").Formula))

log("--- 分块重写 AA（SKU首次出现辅助列）---")
ws.Cells(2, 27).Value2 = "SKU首次出现(辅助)"
log("  AA2 回读 = %r" % ws.Cells(2, 27).Value2)
b3 = write_chunk("AA", 3, RAW_LAST,
                 '=IF($C@R="","",IF(COUNTIFS($C$3:$C@R,$C@R,$D$3:$D@R,$D@R)=1,1,0))')
log("  AA 失败块 %d；回读 AA3=%s AA13060=%s" % (b3, ws.Range("AA3").Formula, ws.Range("AA13060").Formula))

log("--- 分块重写 A 列序号 ---")
STEP = 2000
for s in range(3, RAW_LAST + 1, STEP):
    e = min(s + STEP - 1, RAW_LAST)
    ws.Range("A%d:A%d" % (s, e)).Value2 = tuple((i - 2,) for i in range(s, e + 1))
log("  回读 A3=%s A2629=%s A13060=%s" % (
    ws.Range("A3").Value2, ws.Range("A2629").Value2, ws.Range("A13060").Value2))

log("--- 隐藏 AA 列 ---")
try:
    ws.Columns("AA").Hidden = True
    log("  AA 列隐藏=%s" % ws.Columns("AA").Hidden)
except Exception as e:
    log("  隐藏失败 %s" % e)

log("--- 最终回读 ---")
for a in ("A3", "A2629", "A13060", "AA2", "AA3", "AA13060", "P3", "P13060", "Q3", "E3", "H3", "J3"):
    log("  %-8s = %s" % (a, str(ws.Range(a).Formula)[:88]))

log("保存 ...")
wb.Save(); wb.Close(False)
app.Calculation = -4105
app.Quit()
m1 = os.path.getmtime(DST)
log("落盘 %s -> %s  %s" % (time.strftime("%H:%M:%S", time.localtime(M0)),
                          time.strftime("%H:%M:%S", time.localtime(m1)),
                          "已写入 ✔" if m1 > M0 else "！！未写入"))
