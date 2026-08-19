# -*- coding: utf-8 -*-
"""v8 全量重算 + 逐格取值扫描公式错误 + 关键指标抽样（只读，不保存）

注意：WPS 的 SpecialCells(xlCellTypeFormulas, xlErrors) 不可靠（有错也返回空），
必须用 Value2 逐块取值判断错误常量。
"""
import sys, io, time
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SRC = r"D:\倍优跟单进度表（模板）_v9.xlsx"
ERRMAP = {-2146826288: "#NULL!", -2146826281: "#DIV/0!", -2146826273: "#VALUE!",
          -2146826265: "#REF!", -2146826259: "#NAME?", -2146826252: "#NUM!",
          -2146826246: "#N/A"}
ERRTXT = set(ERRMAP.values())

t0 = time.time()
OUT = open("v9_verify_log.txt", "w", encoding="utf-8")
def log(m):
    s = "[%6.1fs] %s" % (time.time() - t0, m)
    print(s); sys.stdout.flush(); OUT.write(s + "\n"); OUT.flush()

def CL(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

app = win32.DispatchEx("Ket.Application")
app.Visible = False; app.DisplayAlerts = False
try: app.AskToUpdateLinks = False
except Exception: pass
wb = app.Workbooks.Open(SRC)
log("已打开 v9（只读=%s）" % wb.ReadOnly)

app.Calculation = -4105
t1 = time.time()
app.CalculateFullRebuild()
try: app.CalculateUntilAsyncQueriesDone()
except Exception: pass
log("★ 全量重算耗时 %.1f 秒" % (time.time() - t1))

total = 0
detail = {}
for ws in wb.Worksheets:
    used = ws.UsedRange
    r0, c0 = used.Row, used.Column
    nr, nc = used.Rows.Count, used.Columns.Count
    cnt, samples, colcnt = {}, [], {}
    STEP = 2000
    for s in range(0, nr, STEP):
        h = min(STEP, nr - s)
        rng = ws.Range(ws.Cells(r0 + s, c0), ws.Cells(r0 + s + h - 1, c0 + nc - 1))
        v = rng.Value2
        if not isinstance(v, tuple):
            v = ((v,),)
        for i, row in enumerate(v):
            if not isinstance(row, tuple):
                row = (row,)
            for j, x in enumerate(row):
                e = None
                if isinstance(x, int) and x in ERRMAP:
                    e = ERRMAP[x]
                elif isinstance(x, str) and x in ERRTXT:
                    e = x
                if e:
                    cnt[e] = cnt.get(e, 0) + 1
                    col = CL(c0 + j)
                    colcnt[col] = colcnt.get(col, 0) + 1
                    if len(samples) < 5:
                        samples.append("%s%d=%s" % (col, r0 + s + i, e))
    n = sum(cnt.values())
    total += n
    detail[ws.Name] = (n, cnt, colcnt, samples)
    log("%-16s 行%6d 列%3d  错误 %5d  %s %s" % (
        ws.Name, nr, nc, n, cnt if cnt else "", "; ".join(samples)))
log("================ 公式错误合计: %d ================" % total)
for name, (n, cnt, colcnt, s) in detail.items():
    if n:
        log("  %s 按列分布: %s" % (name, sorted(colcnt.items(), key=lambda x: -x[1])[:15]))


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
log("  ①工厂首行 A14=%s 行数%s 总量%s 待交%s 逾期%s 最久逾期%s" % (
    gv("01驾驶舱", "A14"), gv("01驾驶舱", "B14"), gv("01驾驶舱", "C14"),
    gv("01驾驶舱", "E14"), gv("01驾驶舱", "G14"), gv("01驾驶舱", "I14")))
log("  ②类目首行 A107=%s 行数%s 最久逾期%s" % (gv("01驾驶舱", "A107"), gv("01驾驶舱", "B107"), gv("01驾驶舱", "I107")))
log("  ③跟单员首行 A119=%s 行数%s 最久逾期%s" % (gv("01驾驶舱", "A119"), gv("01驾驶舱", "B119"), gv("01驾驶舱", "I119")))

log("--- 02订单明细 抽样 ---")
for r in (3, 2628, 2629, 8000, 13060):
    log("  r%-6d 序%s 合同%s SKU%s 类目%s 总%s 已交%s 待交%s 完成%s 合同总%s 周期%s 状态%s 逾期%s" % (
        r, gv("02订单明细", "A%d" % r), gv("02订单明细", "C%d" % r), gv("02订单明细", "D%d" % r),
        gv("02订单明细", "E%d" % r), gv("02订单明细", "F%d" % r), gv("02订单明细", "G%d" % r),
        gv("02订单明细", "H%d" % r), gv("02订单明细", "I%d" % r), gv("02订单明细", "J%d" % r),
        gv("02订单明细", "P%d" % r), gv("02订单明细", "R%d" % r), gv("02订单明细", "S%d" % r)))

log("--- 06合同级汇总 抽样 ---")
for r in (5, 6, 1500, 2957, 2958):
    log("  r%-6d 合同%s 单据%s SKU种类%s 总量%s 待交%s 完成%s 最早%s 最紧%s 最大逾期%s" % (
        r, gv("06合同级汇总", "B%d" % r), gv("06合同级汇总", "D%d" % r), gv("06合同级汇总", "E%d" % r),
        gv("06合同级汇总", "F%d" % r), gv("06合同级汇总", "H%d" % r), gv("06合同级汇总", "I%d" % r),
        gv("06合同级汇总", "J%d" % r), gv("06合同级汇总", "K%d" % r), gv("06合同级汇总", "N%d" % r)))

log("--- 10发货需求判断 ---")
for addr, lb in [("E3", "需求合计"), ("E4", "未入库"), ("E5", "缺口"), ("E6", "缺口行"), ("E7", "满足率")]:
    log("  %-8s %s = %s" % (lb, addr, gv("10发货需求判断", addr)))
for r in (9, 10, 11):
    log("  r%-3d %s|%s 需求%s 未入库%s 判定%s 主行[%s]" % (
        r, gv("10发货需求判断", "B%d" % r), gv("10发货需求判断", "C%d" % r),
        gv("10发货需求判断", "I%d" % r), gv("10发货需求判断", "J%d" % r),
        gv("10发货需求判断", "M%d" % r), gv("10发货需求判断", "N%d" % r)))

log("--- 11合同冗余 KPI ---")
for addr, lb in [("B6", "在手行数"), ("E3", "在手总量"), ("E4", "在手金额"), ("E5", "冗余量"),
                 ("E6", "冗余金额"), ("E7", "冗余占比"), ("E8", "冗余行数"), ("E9", "不可接单")]:
    log("  %-10s %s = %s" % (lb, addr, gv("11合同冗余", addr)))
log("  ①分档 = %s" % [gv("11合同冗余", "B%d" % r) for r in range(11, 17)])
log("  ②组别 r11 %s 量%s 金额%s | ③厂商 r11 %s 量%s | ④产品 r11 %s 量%s | ⑤组合 r11 %s 在手%s 预测%s 判断%s" % (
    gv("11合同冗余", "H11"), gv("11合同冗余", "I11"), gv("11合同冗余", "J11"),
    gv("11合同冗余", "S11"), gv("11合同冗余", "T11"),
    gv("11合同冗余", "AA11"), gv("11合同冗余", "AB11"),
    gv("11合同冗余", "AI11"), gv("11合同冗余", "AL11"), gv("11合同冗余", "AO11"), gv("11合同冗余", "AU11")))

log("--- 销售预测表 ---")
for r in (4, 5274):
    log("  r%-6d %s|%s 在手%s" % (r, gv("销售预测表", "A%d" % r), gv("销售预测表", "B%d" % r),
                                gv("销售预测表", "F%d" % r)))

log("--- 05逾期呆滞专项 ---")
for r in (4, 2752):
    log("  r%-6d 状态%s 逾期%s 账龄%s" % (r, gv("05逾期呆滞专项", "B%d" % r),
                                     gv("05逾期呆滞专项", "C%d" % r), gv("05逾期呆滞专项", "D%d" % r)))

wb.Close(False)
app.Quit()
log("完成（未保存，源文件未改动）")
OUT.close()
