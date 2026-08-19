# -*- coding: utf-8 -*-
"""v7 全量重算 + 0 错误扫描 + 关键指标抽样（只读，不保存）"""
import sys, io, time
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SRC = r"D:\倍优跟单进度表（模板）_v7.xlsx"
xlCellTypeFormulas, xlErrors = -4123, 16

t0 = time.time()
OUT = open("v7_verify_log.txt", "w", encoding="utf-8")
def log(m):
    s = "[%6.1fs] %s" % (time.time() - t0, m)
    print(s); sys.stdout.flush(); OUT.write(s + "\n"); OUT.flush()

app = win32.DispatchEx("Ket.Application")
app.Visible = False; app.DisplayAlerts = False
try: app.AskToUpdateLinks = False
except Exception: pass
wb = app.Workbooks.Open(SRC)
log("已打开 v7")

app.Calculation = -4105
t1 = time.time()
app.CalculateFullRebuild()
try: app.CalculateUntilAsyncQueriesDone()
except Exception: pass
recalc = time.time() - t1
log("★ 全量重算耗时 %.1f 秒" % recalc)

total = 0
log("%-18s %8s %6s  %s" % ("SHEET", "错误数", "行数", "样例"))
for ws in wb.Worksheets:
    try:
        used = ws.UsedRange
        nr = used.Rows.Count
    except Exception:
        nr = 0
    n, samples = 0, []
    try:
        rng = ws.Cells.SpecialCells(xlCellTypeFormulas, xlErrors)
        n = rng.Count
        k = 0
        for area in rng.Areas:
            for c in area.Cells:
                samples.append("%s=%s" % (c.Address(0, 0), c.Text))
                k += 1
                if k >= 6: break
            if k >= 6: break
    except Exception:
        n = 0
    total += n
    log("%-18s %8d %6d  %s" % (ws.Name, n, nr, "; ".join(samples)))
log("=== 公式错误合计: %d ===" % total)


def gv(sheet, addr):
    try:
        return wb.Worksheets(sheet).Range(addr).Text
    except Exception as e:
        return "ERR:%s" % e

log("--- 01驾驶舱 KPI ---")
log("  A2 = %s" % gv("01驾驶舱", "A2"))
for addr, lb in [("A5", "订单总量"), ("C5", "已交付"), ("E5", "待交付"), ("G5", "完成率"),
                 ("I5", "逾期行数"), ("K5", "逾期件数"), ("A8", "呆滞行数"), ("C8", "紧急行数"),
                 ("E8", "在制母合同"), ("G8", "类目待确认"), ("I8", "在制订单行"), ("K8", "平均周期")]:
    log("  %-10s %s = %s" % (lb, addr, gv("01驾驶舱", addr)))

log("--- 02订单明细 抽样 ---")
for r in (3, 100, 2628, 2629, 8000, 13060):
    log("  r%-6d 序%s 合同%s SKU%s 类目%s 总%s 已交%s 待交%s 完成%s 合同总%s 档%s 状态%s 逾期%s 风险%s" % (
        r, gv("02订单明细", "A%d" % r), gv("02订单明细", "C%d" % r), gv("02订单明细", "D%d" % r),
        gv("02订单明细", "E%d" % r), gv("02订单明细", "F%d" % r), gv("02订单明细", "G%d" % r),
        gv("02订单明细", "H%d" % r), gv("02订单明细", "I%d" % r), gv("02订单明细", "J%d" % r),
        gv("02订单明细", "K%d" % r), gv("02订单明细", "R%d" % r), gv("02订单明细", "S%d" % r),
        gv("02订单明细", "T%d" % r)))

log("--- 03工序进度 抽样 ---")
for r in (3, 3000, 3001, 13060):
    log("  r%-6d SKU%s 类目%s 适用%s 完成%s 进度%s 在制[%s]" % (
        r, gv("03工序进度", "D%d" % r), gv("03工序进度", "E%d" % r), gv("03工序进度", "J%d" % r),
        gv("03工序进度", "K%d" % r), gv("03工序进度", "L%d" % r), gv("03工序进度", "M%d" % r)))

log("--- 04甘特视图 抽样 ---")
for r in (4, 504, 13061):
    log("  r%-6d %s | %s" % (r, gv("04甘特视图", "D%d" % r), gv("04甘特视图", "H%d" % r)))

log("--- 05逾期呆滞专项 抽样 ---")
for r in (4, 5, 1500, 2752):
    log("  r%-6d 状态%s 逾期%s 账龄%s SKU%s" % (
        r, gv("05逾期呆滞专项", "B%d" % r), gv("05逾期呆滞专项", "C%d" % r),
        gv("05逾期呆滞专项", "D%d" % r), gv("05逾期呆滞专项", "G%d" % r)))

log("--- 06合同级汇总 抽样 ---")
for r in (5, 6, 1500, 2957, 2958):
    log("  r%-6d 合同%s 单据%s SKU种类%s 总量%s 待交%s 完成%s 最早%s 最紧%s 最大逾期%s 可出货%s" % (
        r, gv("06合同级汇总", "B%d" % r), gv("06合同级汇总", "D%d" % r), gv("06合同级汇总", "E%d" % r),
        gv("06合同级汇总", "F%d" % r), gv("06合同级汇总", "H%d" % r), gv("06合同级汇总", "I%d" % r),
        gv("06合同级汇总", "J%d" % r), gv("06合同级汇总", "K%d" % r), gv("06合同级汇总", "N%d" % r),
        gv("06合同级汇总", "O%d" % r)))

log("--- 10发货需求判断 ---")
for addr, lb in [("E3", "需求合计"), ("E4", "未入库"), ("E5", "缺口"), ("E6", "缺口行"), ("E7", "满足率")]:
    log("  %-8s %s = %s" % (lb, addr, gv("10发货需求判断", addr)))
for r in (9, 10, 11):
    log("  r%-3d %s|%s 需求%s 未入库%s 差额%s 判定%s 主行%s" % (
        r, gv("10发货需求判断", "B%d" % r), gv("10发货需求判断", "C%d" % r),
        gv("10发货需求判断", "I%d" % r), gv("10发货需求判断", "J%d" % r),
        gv("10发货需求判断", "L%d" % r), gv("10发货需求判断", "M%d" % r),
        gv("10发货需求判断", "N%d" % r)))

log("--- 11合同冗余 KPI ---")
for addr, lb in [("B6", "在手行数"), ("E3", "在手总量"), ("E4", "在手金额"), ("E5", "冗余量"),
                 ("E6", "冗余金额"), ("E7", "冗余占比"), ("E8", "冗余行数"), ("E9", "不可接单")]:
    log("  %-10s %s = %s" % (lb, addr, gv("11合同冗余", addr)))
log("  ①分档 B11..B16 = %s" % [gv("11合同冗余", "B%d" % r) for r in range(11, 17)])
log("  ②运营组别 r11: %s 量%s 金额%s" % (gv("11合同冗余", "H11"), gv("11合同冗余", "I11"), gv("11合同冗余", "J11")))
log("  ③厂商    r11: %s 量%s 金额%s" % (gv("11合同冗余", "S11"), gv("11合同冗余", "T11"), gv("11合同冗余", "U11")))
log("  ④产品编号 r11: %s 量%s 金额%s" % (gv("11合同冗余", "AA11"), gv("11合同冗余", "AB11"), gv("11合同冗余", "AC11")))
log("  ⑤接单判断 r11: %s|%s 预测%s 等级%s 判断%s" % (
    gv("11合同冗余", "AI11"), gv("11合同冗余", "AJ11"), gv("11合同冗余", "AO11"),
    gv("11合同冗余", "AQ11"), gv("11合同冗余", "AU11")))

log("--- 销售预测表 ---")
for r in (4, 5, 5274):
    log("  r%-6d %s|%s key%s 预测%s 品类%s 在手%s" % (
        r, gv("销售预测表", "A%d" % r), gv("销售预测表", "B%d" % r), gv("销售预测表", "C%d" % r),
        gv("销售预测表", "D%d" % r), gv("销售预测表", "E%d" % r), gv("销售预测表", "F%d" % r)))

wb.Close(False)
app.Quit()
log("完成（未保存，源文件未改动）")
OUT.close()
