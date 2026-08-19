# -*- coding: utf-8 -*-
"""跟单进度表 v3 —— 数据提取(WPS读原始导出) + 参数层 + 常量定义"""
import re, os, datetime
import win32com.client as win32

NEW_FILE = r"D:\原始数据.xlsx"                       # 用户维护的主数据源（WPS私有格式，需WPS读）
OLD_FILE = r"D:\理单-品质协同非核心工厂订单汇总跟进表.xlsx"  # 旧文件，仅取检验标准(sheet8)
PROD_FILE = r"D:\产品资料模板.xlsx"                   # 产品资料表（用户维护：产品编号 -> 工艺品类）
SHIP_FILE = r"D:\发货需求模板.xlsx"                   # 发货需求表（用户维护：SKU/团队/发货数量/发货时间）

# ---------- 颜色 / 样式常量 ----------
C_INPUT   = "0070C0"   # 蓝字 = 手工输入
C_FORMULA = "000000"   # 黑字 = 本表公式
C_XREF    = "00843D"   # 绿字 = 跨表引用
C_WARN    = "C00000"   # 红字 = 告警
F_HEADER  = "1F4E79"   # 表头深蓝
F_SUBHEAD = "2E75B6"   # 次级表头
F_ASSUME  = "FFF2CC"   # 黄底 = 假设/待确认
F_INPUT   = "EAF3FB"   # 浅蓝底 = 输入区
F_GREY    = "F2F2F2"
F_KPI     = "DDEBF7"

SH_README = "00使用说明"
SH_DASH   = "01驾驶舱"
SH_RAW    = "原始数据表"        # 新增：用户手动维护的主数据源
SH_ORDER  = "02订单明细"
SH_PROC   = "03工序进度"
SH_GANTT  = "04甘特视图"
SH_LATE   = "05逾期呆滞专项"
SH_CONT   = "06合同级汇总"
SH_TACT   = "07参数-工序节拍"
SH_DICT   = "08参数-字典"
SH_QC     = "09检验标准"
SH_PROD   = "产品资料表"
SH_SHIP   = "发货需求表"
SH_SHIPCHK = "10发货需求判断"
SH_FCST   = "销售预测表"        # 用户维护：产品编号+运营专员 -> 未来3月预计出货数量
SH_REDUN  = "11合同冗余"        # 超期未交付冗余分析 + 接单判断

def q(name):
    return "'%s'" % name


# ================= WPS 读取（兼容 WPS 私有格式/加密） =================
def wps_read(path, sheet_index=1, header_row=1):
    """用 WPS COM 打开任意格式表格，返回 list[dict(header->value)]。

    header_row: 表头所在行（1-based）。原始数据表-畅享 的表头在第 2 行，数据从第 3 行起，
    故调用时传 header_row=2。其余表（产品资料表/发货需求表/09检验标准）表头在第 3 行。
    """
    app = win32.DispatchEx("Ket.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        app.AskToUpdateLinks = False
    except Exception:
        pass
    wb = app.Workbooks.Open(path)
    ws = wb.Worksheets(sheet_index)
    ur = ws.UsedRange
    nr, nc = ur.Rows.Count, ur.Columns.Count
    vals = ur.Value2
    app.Quit()
    if nr == 0 or nc == 0:
        return []
    hidx = header_row - 1
    hdrs = [vals[hidx][c] if c < len(vals[hidx]) else None for c in range(nc)]
    out = []
    for r in range(header_row, nr):
        row = {}
        for c in range(nc):
            row[hdrs[c]] = vals[r][c]
        out.append(row)
    return out


def _serial_to_date(v):
    if isinstance(v, (int, float)):
        try:
            return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(round(v)))
        except Exception:
            return None
    return None


def _to_int(v):
    try:
        return int(float(v)) if v not in (None, "") else 0
    except Exception:
        return 0


# 原始导出 -> 跟单订单结构
RAW_DATE_COLS = {"合同日期", "交货日期", "出货日期", "原出货日期", "原定出货日期",
                 "交货日期 ", "入库日期", "验货通过时间"}
# 02订单明细 对应原始表列（1-based 列号，A=1）
RAWCOL = {
    "工厂":       14,   # 厂商简称
    "合同号":     31,   # 采购合同
    "SKU":         3,   # 产品编号
    "原始品类":   69,   # 报关材质（材质口径）
    "订单总数量": 30,   # 合同数量
    "已交付数量": 41,   # 已交货数量
    "跟单员":      9,   # 跟单专员
    "合同日期":    6,
    "计划出货日":  7,   # 交货日期（用户确认用此列）
    "团队":        4,   # 团队（发货需求汇总条件之一）
    "未入库数量": 43,   # 未入库数量（发货需求模块供应侧汇总口径）
    "合同数量":    30,  # 合同数量（发货需求模块参考）
    # ↓ 11合同冗余 模块用（口径对齐 D:\倍优待交付情况6.2.xlsx）
    "运营专员":     5,   # E  运营专员（冗余接单判断的汇总维度之一）
    "运营组别":    67,   # BO 运营组别（冗余的组织维度，用户指定：只用运营组别）
    "采购单价":    28,   # AB 采购单价（冗余金额 = 未到货总数量 × 采购单价）
    "未到货总数量": 33,  # AG 未到货总数量（冗余数量口径，用户指定，与参考表一致）
    "中文品名":    20,   # T  中文品名
}
# 原始表手工维护区：列号在读取导出后由 set_manual_cols() 动态计算（= 导出列数 + 1..4），
# 这样无论导出是 85 列还是 88 列，手工列都紧贴在导出列之后，绝不与导出列冲突。
MANUAL_ORDER = ["最新跟进日期", "跟进结论", "跟进备注", "是否新单"]
MANUAL_COLS = {}   # 由 build_main 在读取后填充：{名称: 1-based 列号}
def set_manual_cols(ncols):
    """根据导出实际列数，确定 4 个手工维护列的 1-based 位置，并同步更新 RAWCOL。

    关键：必须原地修改 MANUAL_COLS（.clear()+.update()），不能整体重新赋值。
    否则 build_sheets_a/b 通过 `from build_part1 import *` 持有的引用会停留在
    旧的空字典对象上，看不到本次更新，导致列号错位问题复发。
    """
    global MANUAL_COLS, RAWCOL
    MANUAL_COLS.clear()
    MANUAL_COLS.update({nm: ncols + i for i, nm in enumerate(MANUAL_ORDER, start=1)})
    RAWCOL["是否新单"] = MANUAL_COLS["是否新单"]


def norm_raw(raw_rows):
    """把原始导出行规整为跟单订单结构，并返回去重维度集合"""
    orders, mats, facs, buyers = [], set(), set(), set()
    for row in raw_rows:
        fac = (row.get("厂商简称") or "").strip()
        con = (row.get("采购合同") or "").strip()
        sku = (row.get("产品编号") or "").strip()
        qty = _to_int(row.get("合同数量"))
        done = _to_int(row.get("已交货数量"))
        cat = (row.get("报关材质") or "").strip()
        buyer = (row.get("跟单专员") or "").strip()
        cdate = _serial_to_date(row.get("合同日期"))
        sdate = _serial_to_date(row.get("交货日期"))   # 计划出货日 = 交货日期
        orders.append(dict(
            fac=fac, con=con, sku=sku, qty=qty, done=done,
            rem=max(qty - done, 0), cat=cat, buyer=buyer,
            cdate=cdate, sdate=sdate, new="否", raw=row))
        mats.add(cat); facs.add(fac); buyers.add(buyer)
    return orders, mats, facs, buyers


def load_qc():
    """从旧文件 sheet8 读产中检验标准（无表头，前2行跳过，列向前补齐）"""
    app = win32.DispatchEx("Ket.Application")
    app.Visible = False; app.DisplayAlerts = False
    try:
        app.AskToUpdateLinks = False
    except Exception:
        pass
    wb = app.Workbooks.Open(OLD_FILE)
    ws = wb.Worksheets(8)
    ur = ws.UsedRange
    nr, nc = ur.Rows.Count, ur.Columns.Count
    vals = ur.Value2
    app.Quit()
    out, last = [], [""] * 7
    for r in range(2, nr):   # 跳过前2行标题
        vv = [(vals[r][c] if c < len(vals[r]) else None) for c in range(7)]
        if not any(vv):
            continue
        if not vv[0]:
            vv[0] = last[0]
        last = vv
        out.append(vv)
    return out


def _to_num(v):
    """文本/数值 -> float，失败返回 0.0（原始导出里数值常为 '100.0' 文本形态）"""
    if v is None or v is False:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.strip())
        except Exception:
            return 0.0
    return 0.0


def redun_dims(raw_rows):
    """11合同冗余 模块的去重维度。

    返回 dict：运营组别 / 厂商 / SKU / SKU×运营专员组合，
    均按「未到货金额（= 未到货总数量 × 采购单价）」降序排列，
    这样表里默认第一行就是冗余最重的对象，不用用户再手动排序。
    """
    grp, fac, sku, combo = {}, {}, {}, {}
    for row in raw_rows:
        amt = _to_num(row.get("未到货总数量")) * _to_num(row.get("采购单价"))
        g = (row.get("运营组别") or "").strip()
        f = (row.get("厂商简称") or "").strip()
        s = (row.get("产品编号") or "").strip()
        o = (row.get("运营专员") or "").strip()
        if g:
            grp[g] = grp.get(g, 0.0) + amt
        if f:
            fac[f] = fac.get(f, 0.0) + amt
        if s:
            sku[s] = sku.get(s, 0.0) + amt
            combo[(s, o)] = combo.get((s, o), 0.0) + amt
    srt = lambda d: [k for k, _ in sorted(d.items(), key=lambda x: (-x[1], str(x[0])))]
    return dict(groups=srt(grp), facs=srt(fac), skus=srt(sku), combos=srt(combo))


def load_prod():
    """读取用户维护的产品资料表（WPS私有格式），返回 list[dict]"""
    return wps_read(PROD_FILE, 1)


def load_ship():
    """读取用户维护的发货需求表（WPS私有格式），返回 list[dict]"""
    return wps_read(SHIP_FILE, 1)


# ================= 材质 -> 标准工艺类目 映射 =================
def map_one_cat(m):
    """报关材质(材料口径) -> 标准工艺类目。关键字优先级已调好，避免 PE藤 误判树脂。"""
    if not m:
        return ("待确认", "材质缺失，请在08字典①区补映射")
    s = str(m)
    if "特斯林" in s:
        return ("特斯林", "报关材质关键字推断")
    if "藤" in s:                      # 编藤/PE藤/铁管PE藤 等
        return ("编藤", "报关材质关键字推断")
    if "杉木" in s or "木" in s:
        return ("待确认", "实木类无工序模板，需新增(原6模板无实木)")
    if any(k in s for k in ("铁", "铝", "铁制", "铁管", "铝合金", "不锈钢")):
        return ("铁艺铝艺-普通", "关键字推断(刷漆/转印子类待你确认)")
    if s in ("PP", "PE", "HDPE") or "树脂" in s or "PE" in s:
        return ("树脂", "报关材质关键字推断")
    return ("待确认", "无法自动推断，请手填标准类目")


def build_cat_map(distinct_materials):
    """生成 ①区映射行：(原始品类, 标准工艺类目, 依据, is_assume)"""
    out = []
    for m in sorted(distinct_materials, key=lambda x: (x is None, str(x))):
        mapped, note = map_one_cat(m)
        out.append((m if m else "(空)", mapped, note, mapped == "待确认"))
    return out


# ================= 工序统一框架 =================
STAGES = ["下料", "焊接-大件", "焊接-小件", "喷涂·底漆",
          "表面处理-大件", "表面处理-小件", "坐垫·软装", "装配", "包装"]
NA = "NA"

# 标准工艺类目 -> 9道工序的本地叫法（"—" = 该类目不适用）
CAT_STAGE_NAME = {
    "特斯林":     ["下料", "焊接-外露", "焊接-拉布", "—", "喷涂线-外露", "喷涂线-拉布", "—", "拉布机装配", "流水线包装"],
    "编藤":       ["下料", "焊接-大件", "焊接-小件", "—", "编藤-大件", "编藤-小件", "坐垫", "—", "流水线包装"],
    "铁艺铝艺-刷漆": ["下料", "焊接-大件", "焊接-小件", "—", "刷漆-大件", "刷漆-小件", "坐垫", "—", "流水线包装"],
    "铁艺铝艺-转印": ["下料", "焊接-大件", "焊接-小件", "喷涂", "转印-大件", "转印-小件", "坐垫", "—", "流水线包装"],
    "铁艺铝艺-普通": ["下料", "焊接-大件", "焊接-小件", "喷涂", "—", "—", "坐垫", "—", "流水线包装"],
    "树脂":       ["下料", "—", "—", "—", "—", "—", "—", "—", "流水线包装"],
}
CAT_ORDER = ["特斯林", "编藤", "铁艺铝艺-刷漆", "铁艺铝艺-转印", "铁艺铝艺-普通", "树脂"]

# 节拍参数：(开工偏移 = 合同日 + a, 完工偏移 = 出货日 + b[负数])
TACT = {
 ("特斯林", "1000以内", "否"):   [(5,-17),(10,-11),(8,-14),NA,(15,-8),(12,-11),NA,(15,-8),(20,-3)],
 ("特斯林", "1000-3000", "否"):  [(5,-25),(15,-16),(10,-19),NA,(23,-11),(18,-14),NA,(20,-11),(28,-6)],
 ("特斯林", "1000以内", "是"):   [(5,-37),(10,-23),(8,-32),NA,(18,-18),(14,-25),NA,(17,-10),(27,-8)],
 ("编藤", "1000以内", "否"):     [(5,-40),(10,-33),(12,-23),NA,(20,-8),(27,-8),(25,-5),NA,(32,-3)],
 ("编藤", "1000-3000", "否"):    [(7,-46),(22,-36),(17,-31),NA,(47,-11),(29,-11),(47,-11),NA,(57,-1)],
 ("编藤", "1000以内", "是"):     [(5,-43),(12,-38),(10,-28),NA,(20,-13),(16,-8),(10,-3),NA,(25,-3)],
 ("铁艺铝艺-刷漆", "1000以内", "否"):  [(3,-30),(6,-23),(8,-16),NA,(16,-8),(15,-6),(20,-4),NA,(30,-3)],
 ("铁艺铝艺-刷漆", "1000-3000", "否"): [(3,-47),(8,-37),(11,-28),NA,(23,-17),(21,-13),(30,-8),NA,(45,-7)],
 ("铁艺铝艺-刷漆", "1000以内", "是"):  [(3,-33),(6,-18),(7,-16),NA,(16,-8),(15,-8),(20,-3),NA,(35,-3)],
 ("铁艺铝艺-转印", "1000以内", "否"):  [(5,-33),(10,-26),(12,-21),(14,-16),(22,-6),(24,-6),(20,-3),NA,(29,-1)],
 ("铁艺铝艺-转印", "1000-3000", "否"): [(8,-39),(18,-31),(18,-28),(23,-21),(38,-6),(33,-6),(30,-7),NA,(46,-1)],
 ("铁艺铝艺-转印", "1000以内", "是"):  [(3,-33),(8,-26),(10,-21),(12,-16),(20,-6),(22,-6),(20,-3),NA,(27,-1)],
 ("铁艺铝艺-普通", "1000以内", "否"):  [(5,-19),(10,-15),(12,-13),(11,-6),NA,NA,(20,-6),NA,(27,-5)],
 ("铁艺铝艺-普通", "1000-3000", "否"): [(8,-27),(15,-17),(18,-12),(18,-2),NA,NA,(30,-2),NA,(45,-1)],
 ("铁艺铝艺-普通", "1000以内", "是"):  [(3,-21),(8,-17),(10,-15),(9,-8),NA,NA,(20,-8),NA,(27,-7)],
 ("树脂", "1000以内", "否"):     [(10,-25),NA,NA,NA,NA,NA,NA,NA,(25,-10)],
 ("树脂", "1000-3000", "否"):    [(20,-35),NA,NA,NA,NA,NA,NA,NA,(40,-20)],
 ("树脂", "1000以内", "是"):     [(10,-30),NA,NA,NA,NA,NA,NA,NA,(30,-20)],
}
TIERS = ["1000以内", "1000-3000", "3000以上"]
NEWS = ["否", "是"]


def tact_rows():
    """生成节拍参数行，缺失组合用回落值并标注"""
    out = []
    for cat in CAT_ORDER:
        for tier in TIERS:
            for nw in NEWS:
                key = (cat, tier, nw)
                if key in TACT:
                    out.append((cat, tier, nw, TACT[key], "实测节拍（源自原模板公式反解）", False))
                else:
                    src_tier = "1000-3000" if tier == "3000以上" else tier
                    for cand in [(cat, src_tier, nw), (cat, src_tier, "否"),
                                 (cat, "1000-3000", "否"), (cat, "1000以内", "否")]:
                        if cand in TACT:
                            out.append((cat, tier, nw, TACT[cand],
                                        "【回落值·待工厂确认】参照 %s/%s" % (cand[1], cand[2]), True))
                            break
    return out


# 下拉列表默认值
DEFAULT_QC_LIST = ["已解决", "已改期", "催办中", "待确认", "已出货", "无需跟进"]


if __name__ == "__main__":
    rows = wps_read(NEW_FILE, 1)
    orders, mats, facs, buyers = norm_raw(rows)
    print("原始数据行:", len(rows), "-> 订单行:", len(orders))
    print("去重材质:", len(mats), "工厂:", len(facs), "跟单员:", len(buyers))
    print("节拍行:", len(tact_rows()), "类目映射行:", len(build_cat_map(mats)))
    print("样例:", orders[0]["fac"], orders[0]["con"], orders[0]["sku"], orders[0]["cat"])
