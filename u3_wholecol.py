# -*- coding: utf-8 -*-
"""第3步：整列改造 + 逻辑修正。

整列规则（避免循环引用 / 性能塌方）：
  ✔ 转整列：SUMIFS/SUMIF/COUNTIFS/COUNTIF/AVERAGEIFS/AVERAGEIF/VLOOKUP/MINIFS/MAXIFS
            的区域参数，且两端行号都是绝对($)。
  ✔ 转整列：带表名限定的 SUM（跨表汇总）。
  ✘ 不转  ：不带表名的裸 SUM —— 那是区块小计（如 =SUM(B11:B16)），转了会循环/串区。
  ✘ 不转  ：SUMPRODUCT 等数组公式 —— 整列会真的算满 104 万行，直接卡死；改用 MINIFS/MAXIFS。
  ✘ 不转  ：参数表 07/08/09 —— 固定尺寸查找表，转整列只会拖慢。
  ✘ 不转  ：扩展区域（$B$9:$B9）与横向区域（$AP3:$AX3）。
"""
import re, sys, io, time, os, shutil, traceback
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SRC = r"D:\倍优跟单进度表（模板）_v7.xlsx"
DST = r"D:\倍优跟单进度表（模板）_v8.xlsx"
RAW_LAST, AUX_LAST = 13060, 13068
xlCalcManual, xlCalcAuto = -4135, -4105

t0 = time.time()
LOGF = open("u3_log.txt", "w", encoding="utf-8")
def log(m):
    s = "[%6.1fs] %s" % (time.time() - t0, m)
    print(s); sys.stdout.flush()
    LOGF.write(s + "\n"); LOGF.flush()

EXCLUDE_SHEETS = ("07参数-工序节拍", "08参数-字典", "09检验标准")
FUNCS = ("SUMIFS", "SUMIF", "COUNTIFS", "COUNTIF", "AVERAGEIFS", "AVERAGEIF",
         "VLOOKUP", "MINIFS", "MAXIFS", "SUM")
FUNC_RE = re.compile(r"\b(" + "|".join(FUNCS) + r")\s*\(", re.I)
RANGE_RE = re.compile(
    r"(?P<sheet>'[^']+'!|[A-Za-z\u4e00-\u9fa5_][\w\u4e00-\u9fa5\-]*!)?"
    r"\$(?P<c1>[A-Z]{1,3})\$(?P<r1>\d+):\$(?P<c2>[A-Z]{1,3})\$(?P<r2>\d+)")


def span(s, i):
    d, q = 0, False
    while i < len(s):
        ch = s[i]
        if ch == '"':
            q = not q
        elif not q:
            if ch == "(":
                d += 1
            elif ch == ")":
                d -= 1
                if d == 0:
                    return i
        i += 1
    return len(s) - 1


def conv(fml):
    """返回 (新公式, 改动次数)"""
    if not isinstance(fml, str) or not fml.startswith("="):
        return fml, 0
    out, cnt = fml, 0
    pos, guard = 0, 0
    while guard < 200:
        guard += 1
        m = FUNC_RE.search(out, pos)
        if not m:
            break
        fname = m.group(1).upper()
        op = m.end() - 1
        cp = span(out, op)
        body = out[op + 1:cp]
        if fname == "SUM" and "!" not in body:
            pos = m.end()
            continue
        if re.search(r"\bSUMPRODUCT\s*\(", body, re.I):
            pos = m.end()
            continue

        def rep(mm):
            sh = mm.group("sheet") or ""
            if any(("'%s'!" % x) == sh or ("%s!" % x) == sh for x in EXCLUDE_SHEETS):
                return mm.group(0)
            r1, r2 = int(mm.group("r1")), int(mm.group("r2"))
            if r2 >= 1048576 or r2 - r1 < 3:
                return mm.group(0)
            return "%s$%s:$%s" % (sh, mm.group("c1"), mm.group("c2"))

        nb, n = RANGE_RE.subn(rep, body)
        if n:
            out = out[:op + 1] + nb + out[cp:]
            cnt += n
            cp = op + 1 + len(nb)
        pos = cp
    return out, cnt


def CL(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s


FAILS = []
def guard(tag, fn):
    """执行 fn，失败只记录不中断"""
    try:
        return fn()
    except Exception as e:
        FAILS.append("%s -> %s" % (tag, e))
        log("  !! %s 失败: %s" % (tag, e))
        return None


# v7 被 WPS 遗留的 ~$ 锁文件占用 -> 只读打开 -> Save() 静默失效。改为复制到 v8 再改。
shutil.copy2(SRC, DST)
log("已复制 v7 -> v8（%d 字节）" % os.path.getsize(DST))
MTIME0 = os.path.getmtime(DST)

app = win32.DispatchEx("Ket.Application")
app.Visible = False; app.DisplayAlerts = False
try: app.AskToUpdateLinks = False
except Exception: pass
app.ScreenUpdating = False
wb = app.Workbooks.Open(DST)
# 关键：WPS 打开工作簿后会用文件自带的计算模式覆盖 app 设置，必须在 Open 之后再设手动
app.Calculation = xlCalcManual
try: wb.Application.Calculation = xlCalcManual
except Exception: pass
ro = bool(wb.ReadOnly)
log("已打开 v8（只读=%s，计算模式=%s，-4135 即手动）" % (ro, app.Calculation))
if ro:
    log("!!! 工作簿是只读的，Save 会被静默丢弃，终止")
    wb.Close(False); app.Quit(); sys.exit(1)

total = 0
ok = True
try:
    # ======================================================== A. 先修逻辑 bug
    log("=== A. 逻辑修正 ===")
    Q02 = "'02订单明细'!"

    # A1. 02 序号列被复制粘贴污染 -> 重写 1..13058
    ws = wb.Worksheets("02订单明细")
    ws.Range("A3:A%d" % RAW_LAST).Value2 = tuple((i,) for i in range(1, RAW_LAST - 1))
    log("A1  02序号列 A3:A%d 重写为 1..%d" % (RAW_LAST, RAW_LAST - 2))

    # A2. 02 新增 AA 列「SKU首次出现」——供 06「含SKU种类」高效去重计数
    ws.Cells(2, 27).Value2 = "SKU首次出现(辅助)"
    ws.Range("AA3:AA%d" % RAW_LAST).Formula = \
        '=IF($C3="","",IF(COUNTIFS($C$3:$C3,$C3,$D$3:$D3,$D3)=1,1,0))'
    ws.Range(ws.Cells(2, 27), ws.Cells(RAW_LAST, 27)).EntireColumn.Hidden = True
    log("A2  02 新增辅助列 AA「SKU首次出现」（已隐藏）")

    # A8. 02 计划周期 P 列做容错（原 =O-N，遇到文本型日期会 #VALUE!，并污染驾驶舱平均周期）
    ws.Range("P3:P%d" % RAW_LAST).Formula = '=IFERROR($O3-$N3,"")'
    ws.Range("Q3:Q%d" % RAW_LAST).Formula = '=IFERROR($O3-TODAY(),"")'
    log("A8  02 计划周期P/剩余天数Q 加 IFERROR 容错")

    # A3. 06 用 MINIFS/MAXIFS 替掉全表 SUMPRODUCT 数组公式
    ws = wb.Worksheets("06合同级汇总")
    L6 = 2957
    fix06 = [
        ("E", '=SUMIF(%s$C:$C,$B5,%s$AA:$AA)' % (Q02, Q02)),
        ("J", '=IFERROR(MINIFS(%s$N:$N,%s$C:$C,$B5,%s$N:$N,">0"),"")' % (Q02, Q02, Q02)),
        ("K", '=IFERROR(MINIFS(%s$O:$O,%s$C:$C,$B5,%s$O:$O,">0"),"")' % (Q02, Q02, Q02)),
        ("N", '=IFERROR(MAXIFS(%s$S:$S,%s$C:$C,$B5),0)' % (Q02, Q02)),
    ]
    for col, f in fix06:
        ws.Range("%s5:%s%d" % (col, col, L6)).Formula = f
    log("A3  06 含SKU种类/最早合同日/最紧出货日/最大逾期天数 -> SUMIF+MINIFS/MAXIFS")

    # A4. 10发货需求判断「★主行」判定 bug
    ws = wb.Worksheets("10发货需求判断")
    ws.Range("N9:N608").Formula = \
        '=IF($B9="","",IF(COUNTIFS($B$9:$B9,$B9,$C$9:$C9,$C9)=1,"★主行",""))'
    log("A4  10「★主行」改为 COUNTIFS 双键判定")

    # A5. 11 顶部 KPI 覆盖全部辅助区
    ws = wb.Worksheets("11合同冗余")
    ws.Range("B6").Formula = '=COUNTIF($AV$11:$AV$1048576,"?*")'
    for cell, col in (("E3", "BC"), ("E4", "BE"), ("E5", "BG"), ("E6", "BH")):
        ws.Range(cell).Formula = "=SUM($%s$11:$%s$1048576)" % (col, col)
    ws.Range("E8").Formula = '=COUNTIF($BF$11:$BF$1048576,"冗余")'
    log("A5  11 顶部 KPI 改为开放区间（覆盖 %d 行辅助区）" % (AUX_LAST - 10))

    # A6. 驾驶舱 A2 / E8
    ws = wb.Worksheets("01驾驶舱")
    guard("A6-A2", lambda: setattr(ws.Range("A2"), "Formula", (
        '=" 数据基准日： "&TEXT(TODAY(),"yyyy年m月d日")&"    |    在制订单行 "'
        '&COUNTIF(%s$H:$H,">0")&" 行 / 共 "&COUNTA(%s$D$3:$D$1048576)'
        '&" 行    |    口径：一行 = 一个「合同号 + SKU」"' % (Q02, Q02))))
    guard("A6-E8", lambda: setattr(ws.Range("E8"), "Formula",
        '=COUNTIFS(\'06合同级汇总\'!$B:$B,"?*",\'06合同级汇总\'!$H:$H,">0")'))
    # A9. 平均计划周期 K8：原 AVERAGE 被 P 列的 #VALUE! 污染
    guard("A9-K8", lambda: setattr(ws.Range("K8"), "Formula",
        '=IFERROR(ROUND(AVERAGE(%s$P:$P),0),0)' % Q02))
    log("A6/A9  驾驶舱 A2/E8/K8 修正")

    # A7. 驾驶舱三个维度块「最久逾期天数」-> MAXIFS
    blocks = []
    r = 10
    while r < 400:
        v = ws.Cells(r, 1).Value2
        if isinstance(v, str) and v.strip() in ("工厂", "标准工艺类目", "跟单员"):
            name = v.strip()
            r0 = r + 1
            r1 = r0
            while r1 < 20000:
                nx = ws.Cells(r1 + 1, 1).Value2
                if nx is None or (isinstance(nx, str) and nx.strip() in ("", "合计")):
                    break
                r1 += 1
            blocks.append((name, r0, r1))
            r = r1 + 1
        else:
            r += 1
    KEYCOL = {"工厂": "B", "标准工艺类目": "E", "跟单员": "M"}
    for name, r0, r1 in blocks:
        kc = KEYCOL[name]
        f = '=IFERROR(MAXIFS(%s$S:$S,%s$%s:$%s,$A%d),0)' % (Q02, Q02, kc, kc, r0)
        guard("A7-%s" % name,
              lambda rng="I%d:I%d" % (r0, r1), ff=f: setattr(ws.Range(rng), "Formula", ff))
        log("A7  驾驶舱「%s」块 I%d:I%d -> MAXIFS" % (name, r0, r1))

    # ======================================================== B. 整列改造
    log("=== B. 整列改造 ===")
    UNIFORM = [
        ("02订单明细",      3,  RAW_LAST,   range(2, 28),    ()),
        ("03工序进度",      3,  RAW_LAST,   range(1, 78),    tuple(17 + i * 3 for i in range(9))),
        ("04甘特视图",      4,  RAW_LAST+1, range(1, 31),    ()),
        ("06合同级汇总",    5,  2957,       range(3, 17),    ()),
        ("10发货需求判断",  9,  608,        range(2, 15),    ()),
        ("销售预测表",      4,  5274,       range(3, 7),     ()),
        ("11合同冗余",     11,  AUX_LAST,   range(48, 62),   ()),
        ("11合同冗余",     11,  84,         range(9, 16),    ()),
        ("11合同冗余",     11,  98,         range(20, 27),   ()),
        ("11合同冗余",     11,  1860,       range(28, 35),   ()),
        ("11合同冗余",     11,  5281,       range(37, 48),   ()),
        ("11合同冗余",     11,  16,         range(2, 7),     ()),
    ]
    # 驾驶舱三个维度块也按列区间批量改（逐格写会触发依赖树重建，慢 300 倍）
    for _n, _r0, _r1 in blocks:
        UNIFORM.append(("01驾驶舱", _r0, _r1, range(2, 13), ()))

    for sheet, r0, r1, cols, skip in UNIFORM:
        ws = wb.Worksheets(sheet)
        hit = []
        for c in cols:
            if c in skip:
                continue
            try:
                f = ws.Cells(r0, c).Formula
            except Exception as e:
                FAILS.append("%s R%dC%d 读失败: %s" % (sheet, r0, c, e)); continue
            nf, n = conv(f)
            if n and nf != f:
                rng = "%s%d:%s%d" % (CL(c), r0, CL(c), r1)
                try:
                    ws.Range(rng).Formula = nf
                    hit.append(CL(c)); total += n
                except Exception as e:
                    FAILS.append("%s %s 写失败: %s" % (sheet, rng, e))
                    log("  !! %s %s 写失败: %s" % (sheet, rng, e))
        log("  %-14s 行%d..%-6d 改造列: %s" % (sheet, r0, r1, " ".join(hit) or "（无）"))

    # 零散单元格（KPI 区、合计行）逐格处理
    log("=== B2. 零散单元格 ===")
    SPOT = [("01驾驶舱", 1, 200, 1, 12), ("11合同冗余", 1, 10, 1, 61),
            ("10发货需求判断", 1, 8, 1, 14), ("06合同级汇总", 2958, 2958, 1, 16)]
    for sheet, ra, rb, ca, cb in SPOT:
        ws = wb.Worksheets(sheet)
        chg, err = 0, 0
        # 批量读一次（快），只对需要改的少数格逐个写
        try:
            g = ws.Range(ws.Cells(ra, ca), ws.Cells(rb, cb)).Formula
        except Exception as e:
            FAILS.append("%s 批量读失败: %s" % (sheet, e)); g = None
        todo = []
        if g is not None:
            if not isinstance(g, tuple):
                g = ((g,),)
            for i, row in enumerate(g):
                if not isinstance(row, tuple):
                    row = (row,)
                for j, v in enumerate(row):
                    if not isinstance(v, str) or not v.startswith("="):
                        continue
                    nf, n = conv(v)
                    if n and nf != v:
                        todo.append((ra + i, ca + j, nf, n))
        log("  %-14s 零散区 r%d..%d 待改 %d 格" % (sheet, ra, rb, len(todo)))
        for rr, cc, nf, n in todo:
            try:
                ws.Cells(rr, cc).Formula = nf
                chg += 1; total += n
            except Exception as e:
                err += 1
                FAILS.append("%s %s%d 写失败: %s" % (sheet, CL(cc), rr, e))
        log("  %-14s 零散区 已改 %d 格（失败 %d）" % (sheet, chg, err))

    log("整列改造合计 %d 处" % total)
except Exception:
    ok = False
    tb = traceback.format_exc()
    log("!!! 异常中断:\n" + tb)

if ok:
    log("=== C. 写入回读校验 ===")
    for sh, ad in (("01驾驶舱", "A2"), ("01驾驶舱", "K8"), ("02订单明细", "A3"),
                   ("02订单明细", "AA3"), ("02订单明细", "P3"), ("06合同级汇总", "D5"),
                   ("10发货需求判断", "N9"), ("11合同冗余", "B6"), ("销售预测表", "F4")):
        try:
            log("  %s!%s = %s" % (sh, ad, str(wb.Worksheets(sh).Range(ad).Formula)[:90]))
        except Exception as e:
            log("  %s!%s 读失败 %s" % (sh, ad, e))
    log("保存 ...")
    wb.Save(); wb.Close(False)
else:
    log("发生异常，放弃保存")
    wb.Close(False)
app.Calculation = xlCalcAuto
app.ScreenUpdating = True
app.Quit()
if FAILS:
    log("局部失败 %d 条：" % len(FAILS))
    for x in FAILS[:40]:
        log("   - " + x)
if ok:
    m1 = os.path.getmtime(DST)
    log("落盘校验：mtime %s -> %s（%s），大小 %d" % (
        time.strftime("%H:%M:%S", time.localtime(MTIME0)),
        time.strftime("%H:%M:%S", time.localtime(m1)),
        "已写入 ✔" if m1 > MTIME0 else "！！未写入", os.path.getsize(DST)))
log("完成 ok=%s" % ok)
LOGF.close()
