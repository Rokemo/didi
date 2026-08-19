# -*- coding: utf-8 -*-
"""用 WPS/Excel COM 打开工作簿强制全量重算，扫描公式错误"""
import os, sys, time
import win32com.client as win32
from win32com.client import constants

SRC = r"C:\Users\beiyou201\WorkBuddy\2026-07-31-18-27-53\理单-跟单进度跟踪表_v6.xlsx"
DST = r"C:\Users\beiyou201\WorkBuddy\2026-07-31-18-27-53\_verify_out.xlsx"

ERRS = ["#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NULL!", "#NUM!"]

app = None
for pid in ["Ket.Application", "Excel.Application"]:
    try:
        app = win32.DispatchEx(pid)
        print("COM:", pid)
        break
    except Exception as e:
        print("fail", pid, e)
if app is None:
    sys.exit("no COM app")

app.Visible = False
app.DisplayAlerts = False
try:
    app.AskToUpdateLinks = False
except Exception:
    pass

wb = app.Workbooks.Open(SRC)
app.Calculation = -4105  # xlCalculationAutomatic
app.CalculateFullRebuild()
time.sleep(2)
try:
    app.CalculateUntilAsyncQueriesDone()
except Exception:
    pass

total_err = 0
report = []
for ws in wb.Worksheets:
    name = ws.Name
    used = ws.UsedRange
    nr, nc = used.Rows.Count, used.Columns.Count
    vals = used.Value2 if nr * nc > 1 else [[used.Value2]]
    cnt = {}
    samples = []
    r0, c0 = used.Row, used.Column
    for i in range(nr):
        row = vals[i] if isinstance(vals, tuple) else vals
        if not isinstance(row, (tuple, list)):
            row = [row]
        for j in range(len(row)):
            v = row[j]
            if isinstance(v, str) and v in ERRS:
                cnt[v] = cnt.get(v, 0) + 1
                if len(samples) < 8:
                    samples.append("%s%d=%s" % (chr(64 + c0 + j) if c0 + j <= 26 else "C%d" % (c0+j), r0 + i, v))
            elif isinstance(v, int) and v in (-2146826281, -2146826246, -2146826252, -2146826259, -2146826273, -2146826288):
                m = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826252: "#NAME?",
                     -2146826259: "#NULL!", -2146826273: "#NUM!", -2146826288: "#REF!"}
                e = m.get(v, "#ERR")
                cnt[e] = cnt.get(e, 0) + 1
                if len(samples) < 8:
                    samples.append("R%dC%d=%s" % (r0 + i, c0 + j, e))
    n = sum(cnt.values())
    total_err += n
    report.append((name, nr, nc, n, cnt, samples))

print("\n%-18s %6s %5s %8s  %s" % ("SHEET", "rows", "cols", "errors", "detail"))
for name, nr, nc, n, cnt, s in report:
    print("%-18s %6d %5d %8d  %s %s" % (name, nr, nc, n, cnt if cnt else "", "; ".join(s)))
print("\n=== TOTAL FORMULA ERRORS: %d ===" % total_err)

# 抽样关键指标
def gv(sheet, addr):
    try:
        return wb.Worksheets(sheet).Range(addr).Text
    except Exception as e:
        return "ERR:%s" % e

print("\n--- 驾驶舱 KPI 抽样 ---")
for addr, lb in [("A5", "订单总量"), ("C5", "已交付"), ("E5", "待交付"), ("G5", "完成率"),
                 ("I5", "逾期行数"), ("K5", "逾期件数"), ("A8", "呆滞行数"), ("C8", "紧急行数"),
                 ("E8", "在制母合同"), ("G8", "类目待确认"), ("I8", "在制订单行"), ("K8", "平均周期")]:
    print("  %-10s %s = %s" % (lb, addr, gv("01驾驶舱", addr)))

print("\n--- 02订单明细 抽样(第3/13/40/100行) ---")
for r in (3, 13, 40, 100, 183, 500):
    print("  r%-4d %s | %s | 待交%s 完成率%s | 状态%s 逾期%s 风险%s | 工序%s %s %s" % (
        r, gv("02订单明细", "C%d" % r), gv("02订单明细", "F%d" % r),
        gv("02订单明细", "H%d" % r), gv("02订单明细", "I%d" % r),
        gv("02订单明细", "R%d" % r), gv("02订单明细", "S%d" % r), gv("02订单明细", "T%d" % r),
        gv("02订单明细", "U%d" % r), gv("02订单明细", "V%d" % r), gv("02订单明细", "W%d" % r)))

print("\n--- 03工序进度 抽样 ---")
for r in (3, 13, 40, 100):
    print("  r%-4d 类目%s 档%s 适用%s 完成%s 进度%s 在制[%s] 卡点[%s] P1计划%s~%s P9计划%s~%s" % (
        r, gv("03工序进度", "E%d" % r), gv("03工序进度", "F%d" % r),
        gv("03工序进度", "J%d" % r), gv("03工序进度", "K%d" % r), gv("03工序进度", "L%d" % r),
        gv("03工序进度", "M%d" % r), gv("03工序进度", "N%d" % r),
        gv("03工序进度", "O%d" % r), gv("03工序进度", "P%d" % r),
        gv("03工序进度", "AM%d" % r), gv("03工序进度", "AN%d" % r)))

print("\n--- 06合同级汇总 抽样 ---")
for r in (5, 6, 20, 52):
    print("  r%-3d %s %s 单据%s SKU%s 总量%s 待交%s 完成%s 最紧%s 状态%s 可出货[%s]" % (
        r, gv("06合同级汇总", "B%d" % r), gv("06合同级汇总", "C%d" % r),
        gv("06合同级汇总", "D%d" % r), gv("06合同级汇总", "E%d" % r), gv("06合同级汇总", "F%d" % r),
        gv("06合同级汇总", "H%d" % r), gv("06合同级汇总", "I%d" % r),
        gv("06合同级汇总", "K%d" % r), gv("06合同级汇总", "M%d" % r), gv("06合同级汇总", "O%d" % r)))

print("\n--- 04甘特 抽样 ---")
for r in (4, 20, 60):
    row = [gv("04甘特视图", "%s%d" % (c, r)) for c in ["I", "K", "M", "O", "Q", "S", "U", "W", "Y", "AA"]]
    print("  r%-3d %s %s -> %s" % (r, gv("04甘特视图", "D%d" % r), gv("04甘特视图", "H%d" % r), row))
print("  周表头:", [gv("04甘特视图", "%s3" % c) for c in ["I", "J", "K", "L", "M", "N"]])

print("\n--- 05逾期专项 抽样 ---")
for r in (4, 5, 50, 133):
    print("  r%-4d %s %s 逾期%s 账龄%s 风险%s %s" % (
        r, gv("05逾期呆滞专项", "B%d" % r), gv("05逾期呆滞专项", "G%d" % r),
        gv("05逾期呆滞专项", "C%d" % r), gv("05逾期呆滞专项", "D%d" % r),
        gv("05逾期呆滞专项", "E%d" % r), gv("05逾期呆滞专项", "N%d" % r)))

print("\n--- 销售预测表 抽样(前3组合 + 计数) ---")
print("  组合行 A4/A5/A6:", gv("销售预测表", "A4"), "/", gv("销售预测表", "A5"), "/", gv("销售预测表", "A6"))
print("  D4 是否空白(待填):", gv("销售预测表", "D4"))

print("\n--- 11合同冗余 KPI 抽样 ---")
for addr, lb in [("E3", "在手待交付总量"), ("E4", "在手待交付金额"), ("E5", "冗余数量"),
                 ("E6", "冗余金额"), ("E7", "冗余金额占比"), ("E8", "冗余行数"), ("E9", "不可接单组合数")]:
    print("  %-12s %s = %s" % (lb, addr, gv("11合同冗余", addr)))

wb.SaveAs(DST, 51)
wb.Close(False)
app.Quit()
print("\nsaved verify copy:", DST)
