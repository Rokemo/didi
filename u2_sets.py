# -*- coding: utf-8 -*-
"""第2步：重建「集合枚举型」表 —— 05逾期 / 06合同汇总 / 11维度块 / 销售预测表 / 01驾驶舱维度块。
原则：先复制模板行拿到格式，再整块 bulk 写入公式/取值；手工列按业务键回填。
"""
import sys, io, time, datetime
import win32com.client as win32
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

DST = r"D:\倍优跟单进度表（模板）_v7.xlsx"
RAW_LAST = 13060
xlPasteAll, xlPasteFormats = -4104, -4122
xlCalcManual, xlCalcAuto = -4135, -4105
xlUp, xlDown = -4162, -4121

t0 = time.time()
def log(m):
    print("[%6.1fs] %s" % (time.time() - t0, m)); sys.stdout.flush()

def CL(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

app = win32.DispatchEx("Ket.Application")
app.Visible = False; app.DisplayAlerts = False
try: app.AskToUpdateLinks = False
except Exception: pass
app.Calculation = xlCalcManual
app.ScreenUpdating = False
wb = app.Workbooks.Open(DST)

# ---------------------------------------------------------------- 读原始数据
log("读取原始数据表-畅享 ...")
raw = wb.Worksheets("原始数据表-畅享").UsedRange.Value2
rows = raw[2:]                     # 行3..13060
C = lambda n: n - 1
I_SKU, I_TEAM, I_OP, I_CDATE, I_DDATE = C(3), C(4), C(5), C(6), C(7)
I_BUYER, I_FAC, I_NAME = C(9), C(14), C(20)
I_PRICE, I_BOX, I_QTY, I_CON, I_LEFT = C(28), C(29), C(30), C(31), C(33)
I_GRP = C(67)

def s(v): return "" if v is None else str(v).strip()
def f(v):
    try: return float(v)
    except Exception: return 0.0

today = datetime.date.today()
EPOCH = datetime.date(1899, 12, 30)
today_ser = (today - EPOCH).days

# 按 (合同,SKU) 聚合未到货 —— 02!H 的口径
grp_left = {}
for r in rows:
    grp_left[(s(r[I_CON]), s(r[I_SKU]))] = grp_left.get((s(r[I_CON]), s(r[I_SKU])), 0.0) + f(r[I_LEFT])

# ---- 05 逾期集合：02 行号 (= 原始行号)，条件 待交付>0 且 交货日期<今天
late = []
for i, r in enumerate(rows):
    excel_row = 3 + i
    key = (s(r[I_CON]), s(r[I_SKU]))
    if grp_left.get(key, 0.0) <= 0:
        continue
    d = r[I_DDATE]
    try: ds = float(d)
    except Exception: continue
    if ds >= today_ser:
        continue
    late.append((today_ser - int(ds), excel_row, key))
late.sort(key=lambda x: (-x[0], x[1]))
log("05 逾期行 = %d" % len(late))

# ---- 06 合同集合
con_left = {}
for r in rows:
    con_left[s(r[I_CON])] = con_left.get(s(r[I_CON]), 0.0) + f(r[I_LEFT])
contracts = sorted(con_left.keys(), key=lambda k: (-con_left[k], k))
log("06 合同数 = %d" % len(contracts))

# ---- 11 维度：一律按「未到货金额」降序
def dim(idx):
    amt = {}
    for r in rows:
        k = s(r[idx])
        amt[k] = amt.get(k, 0.0) + f(r[I_LEFT]) * f(r[I_PRICE])
    return sorted(amt.keys(), key=lambda k: (-amt[k], k))

groups, facs, skus = dim(I_GRP), dim(I_FAC), dim(I_SKU)
camt = {}
for r in rows:
    k = (s(r[I_SKU]), s(r[I_OP]))
    camt[k] = camt.get(k, 0.0) + f(r[I_LEFT]) * f(r[I_PRICE])
combos = sorted(camt.keys(), key=lambda k: (-camt[k], k[0], k[1]))
log("11 维度：组别%d 厂商%d SKU%d 组合%d" % (len(groups), len(facs), len(skus), len(combos)))

# ---- 驾驶舱维度
fac_box = {}
buyer_box = {}
for r in rows:
    fac_box[s(r[I_FAC])] = fac_box.get(s(r[I_FAC]), 0.0) + f(r[I_BOX])
    buyer_box[s(r[I_BUYER])] = buyer_box.get(s(r[I_BUYER]), 0.0) + f(r[I_BOX])
dash_facs = sorted(fac_box.keys(), key=lambda k: (-fac_box[k], k))
dash_buyers = sorted(buyer_box.keys(), key=lambda k: (-buyer_box[k], k))
log("驾驶舱：工厂%d 跟单员%d" % (len(dash_facs), len(dash_buyers)))

# ---------------------------------------------------------------- 工具
def fill_down(ws, r_tmpl, c1, c2, r_from, r_to):
    """复制模板行的某列区间，向下铺到 r_from..r_to（格式+公式一次到位）"""
    if r_to < r_from: return
    ws.Range(ws.Cells(r_tmpl, c1), ws.Cells(r_tmpl, c2)).Copy()
    ws.Range(ws.Cells(r_from, c1), ws.Cells(r_to, c2)).PasteSpecial(xlPasteAll)
    app.CutCopyMode = False

def put_col(ws, col, r_from, values):
    """整列 bulk 写值"""
    if not values: return
    ws.Range(ws.Cells(r_from, col), ws.Cells(r_from + len(values) - 1, col)).Value2 = \
        tuple((v,) for v in values)

def put_block(ws, r_from, c_from, grid):
    if not grid: return
    ws.Range(ws.Cells(r_from, c_from),
             ws.Cells(r_from + len(grid) - 1, c_from + len(grid[0]) - 1)).Formula = \
        tuple(tuple(row) for row in grid)

# ================================================================ 05 逾期呆滞专项
log("--- 05逾期呆滞专项 ---")
ws = wb.Worksheets("05逾期呆滞专项")
OLD_LAST = 93
n = len(late)
NEW_LAST = 3 + n
# 备份手工列 O..R（按 合同号+SKU 回填）
bak = {}
old = ws.Range("A4:R%d" % OLD_LAST).Value2
for row in old:
    k = (s(row[6]), s(row[7]))
    man = [row[14], row[15], row[16], row[17]]
    if any(x is not None and str(x).strip() != "" for x in man):
        bak[k] = man
log("  备份手工催办记录 %d 条" % len(bak))

if NEW_LAST > OLD_LAST:
    fill_down(ws, 4, 1, 18, OLD_LAST + 1, NEW_LAST)
elif NEW_LAST < OLD_LAST:
    ws.Range("%d:%d" % (NEW_LAST + 1, OLD_LAST)).Delete()

Q = "'02订单明细'!"
grid = []
for i, (days, er, key) in enumerate(late):
    r = 4 + i
    grid.append([
        i + 1,
        '=IF(%sH%d<=0,"已解决",IF(%sQ%d<0,"仍逾期","已改期"))' % (Q, er, Q, er),
        "=%sS%d" % (Q, er),
        '=IF($C%d=0,"—",IF($C%d<=30,"1-30天",IF($C%d<=90,"31-90天",IF($C%d<=180,"91-180天",">180天呆滞"))))' % (r, r, r, r),
        "=%sT%d" % (Q, er), "=%sB%d" % (Q, er), "=%sC%d" % (Q, er), "=%sD%d" % (Q, er),
        "=%sF%d" % (Q, er), "=%sH%d" % (Q, er), "=%sI%d" % (Q, er), "=%sO%d" % (Q, er),
        "=%sM%d" % (Q, er), "=%sV%d" % (Q, er),
    ])
put_block(ws, 4, 1, grid)
ws.Range("O4:R%d" % NEW_LAST).ClearContents()
# 回填手工
hit = 0
if bak:
    for i, (days, er, key) in enumerate(late):
        if key in bak:
            ws.Range(ws.Cells(4 + i, 15), ws.Cells(4 + i, 18)).Value2 = tuple(bak[key])
            hit += 1
log("  写入 %d 行（4..%d），回填手工 %d 条" % (n, NEW_LAST, hit))

# ================================================================ 06 合同级汇总
log("--- 06合同级汇总 ---")
ws = wb.Worksheets("06合同级汇总")
OLD_DATA_LAST, OLD_SUM = 82, 83
n = len(contracts)
NEW_DATA_LAST = 4 + n
if n > (OLD_DATA_LAST - 4):
    # 在合计行上方插入行 -> SUM(D5:D82) 自动扩展为 SUM(D5:D<新末行>)
    ws.Range("%d:%d" % (OLD_DATA_LAST, OLD_DATA_LAST + (n - (OLD_DATA_LAST - 4)) - 1)).Insert()
    fill_down(ws, 5, 1, 16, OLD_DATA_LAST, NEW_DATA_LAST)
put_col(ws, 1, 5, list(range(1, n + 1)))
put_col(ws, 2, 5, contracts)
log("  写入 %d 份合同（5..%d），合计行 %d" % (n, NEW_DATA_LAST, NEW_DATA_LAST + 1))

# ================================================================ 11 合同冗余 维度块
log("--- 11合同冗余 维度块 ---")
ws = wb.Worksheets("11合同冗余")
AUX_LAST = RAW_LAST + 8   # 13068

def rebuild_block(tag, c0, ncol, keys, old_last, old_sum, ratio_col_off, sum_cols, key_is_pair=False,
                  manual_offs=(), bak_key_col=None):
    """c0=块首列(1-based)  ncol=块列数  old_last=原末行  old_sum=原合计行(0=无)
       ratio_col_off=占比列相对偏移  sum_cols=需要合计的列偏移列表"""
    n = len(keys)
    new_last = 10 + n
    new_sum = new_last + 1 if old_sum else 0
    c1, c2 = c0, c0 + ncol - 1
    # 备份手工列
    bak = {}
    if manual_offs and old_last >= 11:
        vals = ws.Range(ws.Cells(11, c1), ws.Cells(old_last, c2)).Value2
        for row in vals:
            k = s(row[bak_key_col])
            man = [row[o] for o in manual_offs]
            if any(x is not None and str(x).strip() != "" for x in man):
                bak[k] = man
    # 清掉旧合计行
    if old_sum:
        ws.Range(ws.Cells(old_sum, c1), ws.Cells(old_sum, c2)).Clear()
    # 占比分母改指向新合计行
    if ratio_col_off is not None and new_sum:
        cell = ws.Cells(11, c0 + ratio_col_off)
        cell.Formula = cell.Formula.replace("$%d," % old_sum, "$%d," % new_sum)
    # 铺模板行
    if new_last > 11:
        fill_down(ws, 11, c1, c2, 12, new_last)
    elif new_last < old_last:
        ws.Range(ws.Cells(new_last + 1, c1), ws.Cells(old_last, c2)).Clear()
    # 写维度值
    if key_is_pair:
        put_block(ws, 11, c0, [[a, b] for a, b in keys])
    else:
        put_col(ws, c0, 11, keys)
    # 新合计行
    if new_sum:
        ws.Cells(new_sum, c0).Value2 = "合计"
        for off in sum_cols:
            cc = c0 + off
            ws.Cells(new_sum, cc).Formula = "=SUM(%s11:%s%d)" % (CL(cc), CL(cc), new_last)
        rc = c0 + ratio_col_off
        ws.Cells(new_sum, rc).Formula = "=IFERROR($%s%d/$%s%d,0)" % (CL(rc - 1), new_sum, CL(rc - 1), new_sum)
    # 回填手工
    hit = 0
    for i, k in enumerate(keys):
        kk = k if not key_is_pair else k[0]
        if kk in bak:
            for j, o in enumerate(manual_offs):
                ws.Cells(11 + i, c0 + o).Value2 = bak[kk][j]
            hit += 1
    log("  %s: %d 行 (11..%d) 合计行=%s 回填手工%d" % (tag, n, new_last, new_sum or "无", hit))
    return new_last, new_sum

# ② 按运营组别 H..Q(8..17) 旧 11..26 合计27  占比=L(off4) 合计列 off 1,2,3,5,6  手工 P,Q = off 8,9
rebuild_block("②按运营组别", 8, 10, groups, 26, 27, 4, [1, 2, 3, 5, 6],
              manual_offs=(8, 9), bak_key_col=0)
# ③ 按厂商 S..Z(19..26) 旧 11..25 合计26  占比=W(off4)
rebuild_block("③按厂商", 19, 8, facs, 25, 26, 4, [1, 2, 3, 5, 6])
# ④ 按产品编号 AA..AH(27..34) 旧 11..197 合计198  占比=AE(off4)
rebuild_block("④按产品编号", 27, 8, skus, 197, 198, 4, [1, 2, 3, 5, 6])
# ⑤ 接单判断 AI..AU(35..47) 旧 11..309 无合计
rebuild_block("⑤接单判断", 35, 13, combos, 309, 0, None, [], key_is_pair=True)

# ================================================================ 销售预测表
log("--- 销售预测表 ---")
ws = wb.Worksheets("销售预测表")
OLD_LAST = 503
n = len(combos)
NEW_LAST = 3 + n
oldv = ws.Range("A4:F%d" % OLD_LAST).Value2
fbak = {}
for row in oldv:
    if row[3] is not None and str(row[3]).strip() != "":
        fbak[(s(row[0]), s(row[1]))] = row[3]
log("  备份已填预测 %d 条" % len(fbak))
if NEW_LAST > OLD_LAST:
    fill_down(ws, 4, 1, 6, OLD_LAST + 1, NEW_LAST)
elif NEW_LAST < OLD_LAST:
    ws.Range("A%d:F%d" % (NEW_LAST + 1, OLD_LAST)).Clear()
put_block(ws, 4, 1, [[a, b] for a, b in combos])
ws.Range("D4:D%d" % NEW_LAST).ClearContents()
hit = 0
for i, k in enumerate(combos):
    if k in fbak:
        ws.Cells(4 + i, 4).Value2 = fbak[k]; hit += 1
log("  写入 %d 组合（4..%d），回填预测 %d 条" % (n, NEW_LAST, hit))

# ================================================================ 01 驾驶舱
log("--- 01驾驶舱 ---")
ws = wb.Worksheets("01驾驶舱")
def rebuild_dash(tag, r0, old_last, keys):
    """自下而上调用；插入/删除行让 WPS 自动修正 合计 SUM 与下方区块位置"""
    n, cur = len(keys), old_last - r0 + 1
    if n > cur:
        ws.Range("%d:%d" % (r0 + 1, r0 + (n - cur))).Insert()
        fill_down(ws, r0, 1, 9, r0 + 1, r0 + n - 1)
    elif n < cur:
        ws.Range("%d:%d" % (r0 + n, old_last)).Delete()
    put_col(ws, 1, r0, keys)
    log("  %s: %d -> %d 行 (%d..%d)" % (tag, cur, n, r0, r0 + n - 1))

rebuild_dash("③按跟单员", 46, 49, dash_buyers)   # 先下后上
rebuild_dash("①按工厂", 14, 28, dash_facs)

log("保存 ...")
wb.Save()
wb.Close(False)
app.Calculation = xlCalcAuto
app.ScreenUpdating = True
app.Quit()
log("完成")
