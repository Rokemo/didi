# -*- coding: utf-8 -*-
"""跟单网页原型 - 工序/甘特/订单明细/合同汇总/发货判断 计算逻辑层

常量与算法移植自 WPS 跟单表 build_part1.py / build_sheets_b/c：
  - 9 道统一工序 STAGES
  - 各标准工艺类目本地工序名 CAT_STAGE_NAME
  - 工序节拍 TACT（开工偏移=合同日+a，完工偏移=出货日+b，b 为负）
  - 交期预警阈值 THRESHOLDS
纯函数，不碰数据库，便于单测。
"""
import datetime

# ---------------- 工序统一框架 ----------------
STAGES = ["下料", "焊接-大件", "焊接-小件", "喷涂·底漆",
          "表面处理-大件", "表面处理-小件", "坐垫·软装", "装配", "包装"]

# 标准工艺类目 -> 9 道工序本地名（"—" = 该类目不走此工序）
CAT_STAGE_NAME = {
    "特斯林":     ["下料", "焊接-外露", "焊接-拉布", "—", "喷涂线-外露", "喷涂线-拉布", "—", "拉布机装配", "流水线包装"],
    "编藤":       ["下料", "焊接-大件", "焊接-小件", "—", "编藤-大件", "编藤-小件", "坐垫", "—", "流水线包装"],
    "铁艺铝艺-刷漆": ["下料", "焊接-大件", "焊接-小件", "—", "刷漆-大件", "刷漆-小件", "坐垫", "—", "流水线包装"],
    "铁艺铝艺-转印": ["下料", "焊接-大件", "焊接-小件", "喷涂", "转印-大件", "转印-小件", "坐垫", "—", "流水线包装"],
    "铁艺铝艺-普通": ["下料", "焊接-大件", "焊接-小件", "喷涂", "—", "—", "坐垫", "—", "流水线包装"],
    "树脂":       ["下料", "—", "—", "—", "—", "—", "—", "—", "流水线包装"],
}
CAT_ORDER = ["特斯林", "编藤", "铁艺铝艺-刷漆", "铁艺铝艺-转印", "铁艺铝艺-普通", "树脂"]

# 节拍：(开工偏移=合同日+a, 完工偏移=出货日+b[负])
TACT = {
 ("特斯林", "1000以内", "否"):   [(5,-17),(10,-11),(8,-14),"NA",(15,-8),(12,-11),"NA",(15,-8),(20,-3)],
 ("特斯林", "1000-3000", "否"):  [(5,-25),(15,-16),(10,-19),"NA",(23,-11),(18,-14),"NA",(20,-11),(28,-6)],
 ("特斯林", "1000以内", "是"):   [(5,-37),(10,-23),(8,-32),"NA",(18,-18),(14,-25),"NA",(17,-10),(27,-8)],
 ("编藤", "1000以内", "否"):     [(5,-40),(10,-33),(12,-23),"NA",(20,-8),(27,-8),(25,-5),"NA",(32,-3)],
 ("编藤", "1000-3000", "否"):    [(7,-46),(22,-36),(17,-31),"NA",(47,-11),(29,-11),(47,-11),"NA",(57,-1)],
 ("编藤", "1000以内", "是"):     [(5,-43),(12,-38),(10,-28),"NA",(20,-13),(16,-8),(10,-3),"NA",(25,-3)],
 ("铁艺铝艺-刷漆", "1000以内", "否"):  [(3,-30),(6,-23),(8,-16),"NA",(16,-8),(15,-6),(20,-4),"NA",(30,-3)],
 ("铁艺铝艺-刷漆", "1000-3000", "否"): [(3,-47),(8,-37),(11,-28),"NA",(23,-17),(21,-13),(30,-8),"NA",(45,-7)],
 ("铁艺铝艺-刷漆", "1000以内", "是"):  [(3,-33),(6,-18),(7,-16),"NA",(16,-8),(15,-8),(20,-3),"NA",(35,-3)],
 ("铁艺铝艺-转印", "1000以内", "否"):  [(5,-33),(10,-26),(12,-21),(14,-16),(22,-6),(24,-6),(20,-3),"NA",(29,-1)],
 ("铁艺铝艺-转印", "1000-3000", "否"): [(8,-39),(18,-31),(18,-28),(23,-21),(38,-6),(33,-6),(30,-7),"NA",(46,-1)],
 ("铁艺铝艺-转印", "1000以内", "是"):  [(3,-33),(8,-26),(10,-21),(12,-16),(20,-6),(22,-6),(20,-3),"NA",(27,-1)],
 ("铁艺铝艺-普通", "1000以内", "否"):  [(5,-19),(10,-15),(12,-13),(11,-6),"NA","NA",(20,-6),"NA",(27,-5)],
 ("铁艺铝艺-普通", "1000-3000", "否"): [(8,-27),(15,-17),(18,-12),(18,-2),"NA","NA",(30,-2),"NA",(45,-1)],
 ("铁艺铝艺-普通", "1000以内", "是"):  [(3,-21),(8,-17),(10,-15),(9,-8),"NA","NA",(20,-8),"NA",(27,-7)],
 ("树脂", "1000以内", "否"):     [(10,-25),"NA","NA","NA","NA","NA","NA","NA",(25,-10)],
 ("树脂", "1000-3000", "否"):    [(20,-35),"NA","NA","NA","NA","NA","NA","NA",(40,-20)],
 ("树脂", "1000以内", "是"):     [(10,-30),"NA","NA","NA","NA","NA","NA","NA",(30,-20)],
}
TIERS = ["1000以内", "1000-3000", "3000以上"]
NEWS = ["否", "是"]

# 交期预警阈值（对应 08字典③区）
THRESHOLDS = {"urgent": 7, "warn": 15, "dead": 180, "lr1": 0.8, "lr2": 0.5}

# 网页未采集「合同日期」时的回退前置期（天）——仅用于推算计划开工，可后续在网页补录合同日期后更精确
DEFAULT_LEAD_DAYS = 30


def today_str():
    return datetime.date.today().isoformat()


def parse_date(v):
    if not v:
        return None
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(s[:10], fmt).date()
        except Exception:
            pass
    return None


def add_days(d, n):
    if d is None:
        return None
    return (d + datetime.timedelta(days=n)).isoformat()


def tact_lookup(cat, tier, is_new):
    """返回 9 道工序的 (a,b) 偏移列表，缺失组合按规则回落。"""
    key = (cat, tier, is_new)
    if key in TACT:
        return TACT[key]
    src_tier = "1000-3000" if tier == "3000以上" else tier
    for cand in [(cat, src_tier, is_new), (cat, src_tier, "否"),
                (cat, "1000-3000", "否"), (cat, "1000以内", "否")]:
        if cand in TACT:
            return TACT[cand]
    # 类目不在模板里：全部 NA
    return ["NA"] * 9


def stage_names(cat):
    """返回该类目 9 道工序本地名（含 '—' 不适用）。"""
    return CAT_STAGE_NAME.get(cat, ["—"] * 9)


def resolve_tier(contract_total):
    """按合同总数判定数量档位。"""
    try:
        t = float(contract_total or 0)
    except Exception:
        t = 0
    if t <= 1000:
        return "1000以内"
    if t <= 3000:
        return "1000-3000"
    return "3000以上"


def _offset_pair(ob):
    """把 TACT 元素规整成 (a, b) 或 None(NA)。"""
    if ob == "NA" or ob is None:
        return None
    a, b = ob[0], ob[1]
    if a is None and b is None:
        return None
    return (a if a is not None else 0, b if b is not None else 0)


def compute_proc(category, tier, is_new, contract_date, delivery_date, actual_dates, base=None):
    """计算单个订单行（合同号+SKU）的 9 道工序进度。

    入参：
      category      标准工艺类目（来自产品资料表；空/未知 -> 无适用工序）
      tier          数量档位（1000以内/1000-3000/3000以上）
      is_new        是否新单（是/否）
      contract_date 合同日期（YYYY-MM-DD，可空 -> 用 DEFAULT_LEAD_DAYS 回退）
      delivery_date 计划出货日（YYYY-MM-DD）
      actual_dates  {stage_idx(0-8): "YYYY-MM-DD" 或 ""}  ★实际完成
      base          基准日（默认今天）
    返回：{stages:[...], applicable, done, progress, current, block}
      stages[i] = {idx, name, applicable(bool), plan_start, due, actual, done(bool), status}
      status: done / overdue / doing / notstart / na
    """
    if base is None:
        base = today_str()
    base_d = parse_date(base)
    cat = category if category in CAT_STAGE_NAME else None
    names = stage_names(cat) if cat else ["—"] * 9
    offsets = tact_lookup(cat, tier, is_new) if cat else ["NA"] * 9

    cdate = parse_date(contract_date)
    if cdate is None and delivery_date:
        cdate = parse_date(delivery_date)
        lead = -DEFAULT_LEAD_DAYS  # 合同日 = 出货日 - 30 天（回退）
    else:
        lead = 0
    ddate = parse_date(delivery_date)

    stages = []
    applicable = 0
    done = 0
    current = None
    block = "无适用工序"
    for i in range(9):
        nm = names[i] if i < len(names) else "—"
        ob = offsets[i] if i < len(offsets) else "NA"
        pair = _offset_pair(ob)
        is_app = pair is not None and nm != "—"
        plan_start = due = None
        actual = (actual_dates or {}).get(i) or ""
        actual_d = parse_date(actual)
        st = "na"
        if is_app:
            applicable += 1
            a, b = pair
            # 计划开工 = 合同日 + a；完工 = 出货日 + b
            ps = cdate + datetime.timedelta(days=(a + lead)) if cdate else None
            du = ddate + datetime.timedelta(days=b) if ddate else None
            plan_start = ps.isoformat() if ps else None
            due = du.isoformat() if du else None
            if actual_d:
                done += 1
                st = "done"
            else:
                if base_d and du:
                    if du < base_d:
                        st = "overdue"
                    elif du == base_d:
                        st = "doing"
                    else:
                        st = "notstart"
                else:
                    st = "notstart"
                if current is None:
                    current = nm  # 第一个未完成的工序 = 当前在制
        stages.append({
            "idx": i, "name": nm, "applicable": is_app,
            "plan_start": plan_start, "due": due,
            "actual": actual, "done": bool(actual_d), "status": st,
        })

    progress = (done / applicable) if applicable else 0
    if applicable == 0:
        block = "无适用工序"
    elif done >= applicable:
        block = "全部完成"
    else:
        # 取第一个未完成工序的 due 判定卡点
        first_undone = next((s for s in stages if s["applicable"] and not s["done"]), None)
        if first_undone and first_undone["due"]:
            du = parse_date(first_undone["due"])
            if base_d and du < base_d:
                block = "工序延期"
            elif base_d and du == base_d:
                block = "进行中"
            else:
                block = "未开工"
        else:
            block = "未开工"

    return {
        "stages": stages, "applicable": applicable, "done": done,
        "progress": progress, "current": current or "", "block": block,
    }


def gantt_weeks(base=None, n=22):
    """生成 n 周的时间轴（每周末日期，含本周往前 4 周、往后 n-5 周）。"""
    if base is None:
        base = today_str()
    base_d = parse_date(base) or datetime.date.today()
    # 本周一
    monday = base_d - datetime.timedelta(days=base_d.weekday())
    start = monday - datetime.timedelta(days=28)  # 前 4 周
    weeks = []
    for i in range(n):
        wk_end = start + datetime.timedelta(days=7 * i + 6)
        weeks.append({
            "idx": i,
            "week_start": (start + datetime.timedelta(days=7 * i)).isoformat(),
            "week_end": wk_end.isoformat(),
            "is_current": (start + datetime.timedelta(days=7 * i) <= base_d <= wk_end),
        })
    return weeks


def gantt_cell(plan_starts, week_end):
    """该合同 9 道工序中，计划开工 <= week_end 的累计道数（用于甘特条带）。"""
    we = parse_date(week_end)
    if we is None:
        return 0
    cnt = 0
    for ps in plan_starts:
        pd = parse_date(ps)
        if pd and pd <= we:
            cnt += 1
    return cnt


# ---------------- 订单明细(02) 派生字段 ----------------
def order_detail_fields(order_qty, delivered_qty, outstanding_qty, delivery_date,
                        complete_rate, contract_total, qty_tier, is_new,
                        follow_days, base=None):
    """返回 02 表的派生字段：待交付/完成率/数量档位/交期状态/逾期天数/风险等级。"""
    if base is None:
        base = today_str()
    base_d = parse_date(base)
    order_qty = float(order_qty or 0)
    delivered = float(delivered_qty or 0)
    outstanding = float(outstanding_qty or 0)
    remain = outstanding  # 待交付口径 = 未到货总数量
    cr = complete_rate if complete_rate is not None else (delivered / order_qty if order_qty else 0)
    tier = qty_tier or resolve_tier(contract_total)
    ddate = parse_date(delivery_date)
    days_to_ship = (ddate - base_d).days if (ddate and base_d) else None

    # 交期状态
    if remain <= 0:
        status = "已交付"
    elif days_to_ship is None:
        status = "未知"
    elif days_to_ship < 0:
        status = "已逾期"
    elif days_to_ship <= THRESHOLDS["urgent"]:
        status = "紧急"
    elif days_to_ship <= THRESHOLDS["warn"]:
        status = "预警"
    else:
        status = "正常"
    # 逾期天数
    overdue_days = max(0, -days_to_ship) if (days_to_ship is not None and remain > 0) else 0
    # 风险等级
    if remain <= 0:
        risk = "无"
    elif overdue_days > THRESHOLDS["dead"]:
        risk = "呆滞"
    elif overdue_days > 0:
        risk = "高"
    elif days_to_ship is not None and days_to_ship <= THRESHOLDS["urgent"] and cr < THRESHOLDS["lr1"]:
        risk = "高"
    elif days_to_ship is not None and days_to_ship <= THRESHOLDS["warn"] and cr < THRESHOLDS["lr2"]:
        risk = "中"
    else:
        risk = "低"
    return {
        "remain": remain, "complete_rate": cr, "tier": tier,
        "days_to_ship": days_to_ship, "status": status,
        "overdue_days": overdue_days, "risk": risk,
    }
