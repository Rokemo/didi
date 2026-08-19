# -*- coding: utf-8 -*-
"""第1步：v6 -> v7，把「行级一一对应」的表扩到全量 13058 行。
02订单明细 : 行3..13060  (行r <-> 原始行r)
03工序进度 : 行3..13060  (行r <-> 02行r)
04甘特视图 : 行4..13061  (行r <-> 02行r-1)
11⑥辅助区 : 行11..13068 (行r <-> 原始行r-8)
做法：复制模板行 -> 粘贴到新增区间（单次 COM 调用，公式相对引用自动位移，格式一并继承）。
"""
import os, shutil, sys, io, time
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

SRC = r"D:\倍优跟单进度表（模板）_v6.xlsx"
DST = r"D:\倍优跟单进度表（模板）_v7.xlsx"

RAW_LAST = 13060          # 原始数据最后一行
N = RAW_LAST - 3 + 1      # 13058 行业务数据

xlPasteAll = -4104
xlCalcManual = -4135
xlCalcAuto = -4105

if os.path.exists(DST):
    os.remove(DST)
shutil.copyfile(SRC, DST)
print("已复制 -> %s" % DST)

app = win32.DispatchEx("Ket.Application")
app.Visible = False
app.DisplayAlerts = False
try:
    app.AskToUpdateLinks = False
except Exception:
    pass
app.Calculation = xlCalcManual
app.ScreenUpdating = False

wb = app.Workbooks.Open(DST)
t0 = time.time()


def extend(sheet, tmpl_row, first_new, last_new, last_col_letter, clear_cols=()):
    ws = wb.Worksheets(sheet)
    if last_new < first_new:
        print("  [%s] 无需扩展" % sheet)
        return
    src = ws.Range("A%d:%s%d" % (tmpl_row, last_col_letter, tmpl_row))
    dst = ws.Range("A%d:%s%d" % (first_new, last_col_letter, last_new))
    src.Copy()
    dst.PasteSpecial(xlPasteAll)
    app.CutCopyMode = False
    for c in clear_cols:
        ws.Range(ws.Cells(first_new, c), ws.Cells(last_new, c)).ClearContents()
    print("  [%s] 模板行%d -> 填充 %d..%d  (+%d 行)  用时%.1fs"
          % (sheet, tmpl_row, first_new, last_new, last_new - first_new + 1, time.time() - t0))


def used_last_row(sheet):
    ws = wb.Worksheets(sheet)
    ur = ws.UsedRange
    return ur.Row + ur.Rows.Count - 1


print("\n=== 扩展前行数 ===")
for s in ("02订单明细", "03工序进度", "04甘特视图", "11合同冗余"):
    print("  %-12s 末行 %d" % (s, used_last_row(s)))

print("\n=== 开始扩展 ===")
# 02订单明细 A..Z, 现有 3..2628
extend("02订单明细", 3, 2629, RAW_LAST, "Z")

# 03工序进度 A..BY(77列), 现有 3..3000；手工列 Q T W Z AC AF AI AL AO = 17,20,...,41
MANUAL_03 = [17 + i * 3 for i in range(9)]
extend("03工序进度", 3, 3001, RAW_LAST, "BY", clear_cols=MANUAL_03)

# 04甘特视图 A..AD(30列), 现有 4..503, 目标 4..13061
extend("04甘特视图", 4, 504, RAW_LAST + 1, "AD")

# 11⑥辅助区 AV..BI, 现有 11..3611, 目标 11..13068
ws11 = wb.Worksheets("11合同冗余")
src = ws11.Range("AV11:BI11")
dst = ws11.Range("AV3612:BI%d" % (RAW_LAST + 8))
src.Copy()
dst.PasteSpecial(xlPasteAll)
app.CutCopyMode = False
print("  [11合同冗余⑥辅助区] AV11:BI11 -> AV3612:BI%d  (+%d 行)  用时%.1fs"
      % (RAW_LAST + 8, RAW_LAST + 8 - 3612 + 1, time.time() - t0))

print("\n=== 扩展后行数 ===")
for s in ("02订单明细", "03工序进度", "04甘特视图", "11合同冗余"):
    print("  %-12s 末行 %d" % (s, used_last_row(s)))

wb.Save()
wb.Close(False)
app.Calculation = xlCalcAuto
app.ScreenUpdating = True
app.Quit()
print("\n完成，用时 %.1fs" % (time.time() - t0))
