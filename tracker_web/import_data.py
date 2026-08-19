# -*- coding: utf-8 -*-
"""
导入真实数据（优先）或生成种子数据（兜底）。
真实数据：用 WPS COM 读取 D:\倍优跟单进度表（模板）_v10.xlsx 的四张源表。
表头行不固定（首行多为说明），脚本自动探测表头行。
"""
import os
import sys
import random
import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
from server import init_db, DB, conn_db, PROD_COLS, SHIP_COLS, FC_COLS, RAW_COLS

V10 = r"D:\倍优跟单进度表（模板）_v10.xlsx"

RAW_MAP = {
    "contract_no": ["采购合同"], "product_code": ["产品编号"], "team": ["团队"],
    "operator": ["运营专员", "运营"], "operator_group": ["运营组别", "跟单组别"],
    "factory": ["厂商简称", "厂商"], "order_qty": ["合同数量", "订单总数量"],
    "delivered_qty": ["已交付数量"], "undelivered_qty": ["未入库数量"],
    "outstanding_qty": ["未到货总数量"], "unit_price": ["采购单价"],
    "delivery_date": ["交货日期"], "customs_material": ["报关材质"],
    "followup_date": ["最新跟进日期"], "followup_conclusion": ["跟进结论"],
    "followup_note": ["跟进备注"], "is_new": ["是否新单"],
}
PROD_MAP = {"product_code": ["产品编号"], "product_name": ["中文品名", "品名"], "craft_category": ["工艺品类"]}
SHIP_MAP = {"sku": ["SKU", "sku"], "team": ["团队"], "operator": ["运营"], "factory": ["工厂"],
            "follower": ["跟单"], "ship_qty": ["发货数量"], "ship_date": ["发货时间", "发货日期"]}
FC_MAP = {"product_code": ["产品编号"], "operator": ["运营专员", "运营"], "forecast_qty": ["未来3月预计出货数量", "预计出货"]}

RAW_KW = ["产品编号", "团队", "运营专员", "交货日期", "采购单价", "合同数量", "厂商简称"]
PROD_KW = ["产品编号", "中文品名", "工艺品类"]
SHIP_KW = ["SKU", "团队", "运营", "工厂", "跟单", "发货数量", "发货时间"]
FC_KW = ["产品编号", "运营专员", "未来3月预计出货数量"]

RAW_NUM = {"order_qty", "delivered_qty", "undelivered_qty", "outstanding_qty", "unit_price"}
SHIP_NUM = {"ship_qty"}
FC_NUM = {"forecast_qty"}


def find_idx(headers, subs):
    for i, h in enumerate(headers):
        if h is None:
            continue
        hs = str(h).strip()
        for s in subs:
            if s in hs:
                return i
    return -1


def map_cols(headers, themap):
    return {k: find_idx(headers, subs) for k, subs in themap.items()}


def detect_header_row(vals, keywords, max_scan=6):
    best, best_i = -1, 0
    for r in range(min(max_scan, len(vals))):
        row = vals[r]
        score = sum(1 for v in row if v is not None and any(k in str(v) for k in keywords))
        if score > best:
            best, best_i = score, r
    return best_i


def to_num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except Exception:
        return None


def to_str(v):
    if v is None:
        return ""
    return str(v).strip()


def parse_date(v):
    if v is None:
        return ""
    s = str(v).strip()
    if len(s) >= 10 and s[4] == "-":
        return s[:10]
    return s


def extract_rows(vals, keywords, themap, numeric, datekeys, skip_pred):
    hi = detect_header_row(vals, keywords)
    headers = [vals[hi][j] if j < len(vals[hi]) else None for j in range(len(vals[hi]))]
    m = map_cols(headers, themap)
    rows = []
    for r in range(hi + 1, len(vals)):
        row = vals[r]
        d = {}
        for k, i in m.items():
            if i < 0 or i >= len(row):
                d[k] = None
                continue
            v = row[i]
            if k in numeric:
                d[k] = to_num(v)
            elif k in datekeys:
                d[k] = parse_date(v)
            else:
                d[k] = to_str(v)
        if skip_pred(d):
            continue
        rows.append(d)
    return rows


def locate(wb, name):
    for ws in wb.Worksheets:
        if ws.Name == name:
            return ws
    for ws in wb.Worksheets:
        if name in ws.Name:
            return ws
    return None


def import_from_wps():
    import win32com.client
    app = win32com.client.Dispatch("Ket.Application")
    app.Visible = False
    app.DisplayAlerts = False
    wb = app.Workbooks.Open(V10, False, True)
    try:
        c = conn_db()
        c.execute("DELETE FROM raw_orders"); c.execute("DELETE FROM products")
        c.execute("DELETE FROM shipping_demands"); c.execute("DELETE FROM sales_forecast")

        raw_ws = locate(wb, "原始数据表-畅享")
        prod_ws = locate(wb, "产品资料表")
        ship_ws = locate(wb, "发货需求表")
        fc_ws = locate(wb, "销售预测表")

        if raw_ws:
            vals = raw_ws.UsedRange.Value
            rows = extract_rows(vals, RAW_KW, RAW_MAP, RAW_NUM, {"delivery_date"},
                                lambda d: not d.get("product_code") and not d.get("contract_no"))
            for d in rows:
                if d["delivered_qty"] is None and d["order_qty"] is not None and d["outstanding_qty"] is not None:
                    d["delivered_qty"] = max(0, d["order_qty"] - d["outstanding_qty"])
            _bulk(c, "raw_orders", RAW_COLS, rows)
            print("原始数据：导入 %d 行" % len(rows))

        if prod_ws:
            vals = prod_ws.UsedRange.Value
            rows = extract_rows(vals, PROD_KW, PROD_MAP, set(), set(),
                                lambda d: not d.get("product_code"))
            _bulk(c, "products", PROD_COLS, rows)
            print("产品资料：导入 %d 行" % len(rows))

        if ship_ws:
            vals = ship_ws.UsedRange.Value
            rows = extract_rows(vals, SHIP_KW, SHIP_MAP, SHIP_NUM, {"ship_date"},
                                lambda d: not d.get("sku"))
            _bulk(c, "shipping_demands", SHIP_COLS, rows)
            print("发货需求：导入 %d 行" % len(rows))

        if fc_ws:
            vals = fc_ws.UsedRange.Value
            rows = extract_rows(vals, FC_KW, FC_MAP, FC_NUM, set(),
                                lambda d: not d.get("product_code"))
            _bulk(c, "sales_forecast", FC_COLS, rows)
            print("销售预测：导入 %d 行" % len(rows))

        c.commit(); c.close()
        return 1
    finally:
        wb.Close(False)
        app.Quit()


def _bulk(c, table, cols, rows):
    ph = ",".join(["?"] * len(cols))
    sql = "INSERT INTO %s (%s) VALUES (%s)" % (table, ",".join(cols), ph)
    c.executemany(sql, [tuple(d.get(col) for col in cols) for d in rows])


def seed():
    print("WPS 不可用，生成结构化种子数据用于演示...")
    random.seed(20260817)
    c = conn_db()
    c.execute("DELETE FROM raw_orders"); c.execute("DELETE FROM products")
    c.execute("DELETE FROM shipping_demands"); c.execute("DELETE FROM sales_forecast")

    factories = ["倍优一厂", "倍优二厂", "义乌先锋", "深圳宏达", "宁波海纳", "广州顺成",
                 "福建锐泽", "无锡精工", "青岛万洋", "东莞启航", "苏州锦程", "佛山卓越",
                 "温州联众", "杭州云栖", "成都蜀源"]
    teams = ["团队A", "团队B", "团队C", "团队D"]
    groups = ["组别1", "组别2", "组别3"]
    operators = ["张敏", "李强", "王芳", "刘洋", "陈静", "赵磊", "孙颖", "周涛"]
    followers = ["跟单甲", "跟单乙", "跟单丙", "跟单丁"]
    materials = ["塑料", "金属", "纺织", "硅胶", "木制", "陶瓷"]
    cats = ["家居", "数码配件", "户外", "服饰", "美妆", "宠物"]
    today = datetime.date.today()

    prods = []
    for i in range(60):
        code = "NB%03d" % (i + 1)
        prods.append((code, "产品%s" % code, random.choice(cats)))
    c.executemany("INSERT INTO products (product_code,product_name,craft_category) VALUES (?,?,?)", prods)

    raw = []
    for i in range(2500):
        pc = random.choice(prods)[0]
        days = random.randint(-200, 120)
        dd = today + datetime.timedelta(days=days)
        order = random.choice([500, 800, 1000, 1500, 2000, 3000])
        delivered = int(order * random.uniform(0.3, 1.0))
        if delivered > order:
            delivered = order
        outstanding = order - delivered
        price = round(random.uniform(8, 120), 2)
        raw.append((
            "HB%04d" % random.randint(1, 78), pc, random.choice(teams),
            random.choice(operators), random.choice(groups), random.choice(factories),
            order, delivered, random.randint(0, 200), outstanding, price,
            dd.isoformat(), random.choice(materials),
            (today - datetime.timedelta(days=random.randint(0, 30))).isoformat() if random.random() < 0.5 else "",
            random.choice(["生产中", "已发货", "待确认", "清尾中"]),
            random.choice(["", "正常", "需跟进"]), random.choice(["是", "否"]),
        ))
    c.executemany(
        "INSERT INTO raw_orders (contract_no,product_code,team,operator,operator_group,factory,"
        "order_qty,delivered_qty,undelivered_qty,outstanding_qty,unit_price,delivery_date,"
        "customs_material,followup_date,followup_conclusion,followup_note,is_new) VALUES ("
        + ",".join(["?"] * 17) + ")", raw)

    ship = []
    for i in range(150):
        ship.append(("SKU%04d" % (i + 1), random.choice(teams), random.choice(operators),
                     random.choice(factories), random.choice(followers), random.randint(100, 2000),
                     (today + datetime.timedelta(days=random.randint(0, 60))).isoformat()))
    c.executemany("INSERT INTO shipping_demands (sku,team,operator,factory,follower,ship_qty,ship_date) VALUES (?,?,?,?,?,?,?)", ship)

    fc = []
    for p in random.sample(prods, 30):
        fc.append((p[0], random.choice(operators), random.randint(50, 600)))
    c.executemany("INSERT INTO sales_forecast (product_code,operator,forecast_qty) VALUES (?,?,?)", fc)

    c.commit(); c.close()
    print("种子数据：原始数据 %d 行 / 产品 %d / 发货 %d / 预测 %d" % (len(raw), len(prods), len(ship), len(fc)))


if __name__ == "__main__":
    init_db()
    ok = 0
    try:
        ok = import_from_wps()
    except Exception as e:
        print("WPS 导入失败：%s" % e)
    if not ok:
        seed()
    print("导入完成。数据库：%s" % DB)
