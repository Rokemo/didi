// 跟单网页原型 v0.2 - 前端逻辑（原生 JS，无构建）
const $ = (s) => document.querySelector(s);
const $$ = (s) => document.querySelectorAll(s);
let charts = {};
let rawState = { page: 1, size: 100, search: "", cols: [], labels: {}, operator: "", group: "", factory: "", selAll: false };
let ovState = { page: 1, size: 200, base: "", group: "", operator: "", factory: "", bucket: "" };
let rdState = { base: "", sub: "by_group" };
let odState = { page: 1, size: 200, search: "", operator: "", group: "", factory: "" };
let pcState = { page: 1, size: 100, search: "", filters: {}, sel: new Set() };
let gtState = { page: 1, size: 200, search: "", operator: "", group: "", factory: "", base: "" };
let scState = { start: "", end: "" };
let lastProcRows = [];
let currentRole = "";
let lastProcAllKeys = [];
let lastProcAllTotal = 0;
let procInspOptions = [];
const INSPECTION_FOCUS_FALLBACK = ["金加工阶段","外发返厂","半成品成型","半成品备件","预组装阶段","入箱前","原材料到厂阶段","产前样","喷塑环节","来料检查","打磨/喷塑检验","预包装阶段","样品开发阶段","下料后","预加工","产前样组装测试"];
let procStages = [];
// 模块 tab 名 -> HTML DOM id 前缀（index.html 用短前缀 prod/ship/fc）
const TAB_ID = { products: "prod", shipping: "ship", forecast: "fc" };

// ---------- 通用 fetch ----------
async function api(path, opts) {
  const r = await fetch(path, opts);
  if (r.status === 401) { throw new Error("未登录"); }
  return r;
}
async function apiJson(path, opts) {
  const r = await api(path, opts);
  return r.json();
}

// ---------- 登录（免登录模式下登录遮罩已移除，函数做空安全保护） ----------
function showLogin() {
  const el = $("#loginOverlay");
  if (!el) return;
  el.classList.remove("hidden");
  const ub = $("#userBox"); if (ub) ub.classList.add("hidden");
}
function hideLogin() {
  const el = $("#loginOverlay");
  if (!el) return;
  el.classList.add("hidden");
}
async function refreshMe() {
  try {
    const me = await apiJson("/api/me");
    currentRole = me.role;
    $("#userBox").classList.remove("hidden");
    $("#userName").textContent = me.display_name || me.username;
    const rb = { admin: "管理员", "跟单员": "跟单员", viewer: "只读" }[me.role] || me.role;
    $("#userRole").textContent = rb;
    applyRole();
    return true;
  } catch (e) {
    return false;
  }
}
// 免登录模式：登录/退出按钮相关 DOM 已移除，以下逻辑仅在元素存在时绑定
const loginBtn = $("#loginBtn");
if (loginBtn) {
  loginBtn.onclick = async () => {
    const u = $("#loginUser").value.trim(), p = $("#loginPw").value;
    const r = await fetch("/api/login", { method: "POST", headers: jh(), body: JSON.stringify({ username: u, password: p }) });
    if (r.ok) {
      $("#loginErr").textContent = "";
      hideLogin();
      await refreshMe();
      loadDashSilent();
    } else {
      $("#loginErr").textContent = "用户名或密码错误";
    }
  };
  $("#loginPw").onkeydown = (e) => { if (e.key === "Enter") loginBtn.click(); };
}
const logoutBtn = $("#logoutBtn");
if (logoutBtn) {
  logoutBtn.onclick = async () => {
    await fetch("/api/logout", { method: "POST" });
    currentRole = "";
    showLogin();
  };
}

// ---------- 角色权限 ----------
function applyRole() {
  const writer = currentRole === "admin" || currentRole === "跟单员";
  const admin = currentRole === "admin";
  $$("[data-need]").forEach((el) => {
    const need = el.dataset.need;
    let show = false;
    if (need === "admin") show = admin;
    else if (need === "writer") show = writer;
    el.style.display = show ? "" : "none";
  });
}

// ---------- Tabs ----------
$$(".tab").forEach((t) => {
  t.onclick = () => {
    $$(".tab").forEach((x) => x.classList.remove("active"));
    $$(".panel").forEach((x) => x.classList.remove("active"));
    t.classList.add("active");
    const id = t.dataset.tab;
    $("#" + id).classList.add("active");
    if (id === "dash") loadDash();
    if (id === "raw") { ensureFilters(); loadRaw(); }
    if (id === "overdue") { ensureFilters(); loadOverdue(); }
    if (id === "redundancy") loadRedundancy();
    if (id === "products") loadTable("products", "/api/products", prodFields(), "products");
    if (id === "shipping") loadTable("shipping", "/api/shipping", shipFields(), "shipping_demands");
    if (id === "forecast") loadTable("forecast", "/api/forecast", fcFields(), "sales_forecast");
    if (id === "order") { ensureFilters(); loadOrderDetail(); }
    if (id === "proc") { ensureFilters(); loadProc(); }
    if (id === "gantt") { ensureFilters(); loadGantt(); }
    if (id === "contract") loadContract();
    if (id === "shipchk") loadShipchk();
    if (id === "params") loadParams();
  };
});

// ---------- 驾驶舱 ----------
async function loadDash() {
  const k = await apiJson("/api/kpi");
  const cards = [
    { l: "订单总数量", v: fmt(k.order_qty), c: "" },
    { l: "已交付数量", v: fmt(k.delivered_qty), c: "" },
    { l: "完成率", v: k.completion_rate + "%", c: "" },
    { l: "在手待交付", v: fmt(k.outstanding_qty), c: "" },
    { l: "在手金额(元)", v: fmt(k.outstanding_amt), c: "" },
    { l: "逾期行数", v: k.overdue, c: k.overdue > 0 ? "warn" : "" },
    { l: "合同数", v: k.contracts, c: "" },
    { l: "厂商数 / 产品数", v: k.factories + " / " + k.products, c: "" },
  ];
  $("#kpiCards").innerHTML = cards.map((c) =>
    `<div class="kpi ${c.c}"><div class="v">${c.v}</div><div class="l">${c.l}</div></div>`).join("");
  drawChart("chartRate", "doughnut",
    ["已完成", "未完成"],
    [k.delivered_qty, Math.max(0, k.order_qty - k.delivered_qty)],
    ["#0c447c", "#cbd5e1"]);
  drawChart("chartFactory", "bar",
    k.by_factory.map((x) => x.name || "未填"),
    k.by_factory.map((x) => Math.round(x.amt)),
    ["#185fa5"]);
  drawChart("chartGroup", "bar",
    k.by_group.map((x) => x.name || "未填"),
    k.by_group.map((x) => x.n),
    ["#ba7517"]);
}
async function loadDashSilent() { try { await loadDash(); } catch (e) {} }

function drawChart(id, type, labels, data, colors) {
  if (charts[id]) charts[id].destroy();
  const ctx = document.getElementById(id);
  if (!ctx) return;
  charts[id] = new Chart(ctx, {
    type,
    data: { labels, datasets: [{ data, backgroundColor: colors, borderWidth: 0 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: type === "doughnut" } },
      scales: type === "bar" ? { y: { beginAtZero: true } } : {},
    },
  });
}

// ---------- 维度筛选选项 ----------
let filtersLoaded = false;
async function ensureFilters() {
  if (filtersLoaded) return;
  try {
    const f = await apiJson("/api/filters");
    fillSelect($("#rawOperator"), f.operator);
    fillSelect($("#rawGroup"), f.operator_group);
    fillSelect($("#rawFactory"), f.factory);
    fillSelect($("#ovOperator"), f.operator);
    fillSelect($("#ovGroup"), f.operator_group);
    fillSelect($("#ovFactory"), f.factory);
    ["od", "pc", "gt"].forEach((p) => {
      fillSelect($("#" + p + "Operator"), f.operator);
      fillSelect($("#" + p + "Group"), f.operator_group);
      fillSelect($("#" + p + "Factory"), f.factory);
    });
    filtersLoaded = true;
  } catch (e) {}
}
function fillSelect(sel, arr) {
  const first = sel.options[0];
  sel.innerHTML = "";
  sel.appendChild(first);
  (arr || []).forEach((v) => {
    const o = document.createElement("option");
    o.value = v; o.textContent = v; sel.appendChild(o);
  });
}

// ---------- 原始数据 CRUD + 筛选 ----------
async function loadRaw() {
  if (!rawState.cols.length) {
    const meta = await apiJson("/api/columns");
    rawState.cols = meta.cols; rawState.labels = meta.labels;
  }
  const params = new URLSearchParams({
    page: rawState.page, size: rawState.size, search: rawState.search,
    operator: rawState.operator, operator_group: rawState.group, factory: rawState.factory,
  });
  const d = await apiJson("/api/raw?" + params.toString());
  const cols = rawState.cols, labels = rawState.labels;
  let html = "<table><thead><tr><th class=\"chk\"><input type=\"checkbox\" id=\"rawCheckAll\"></th><th>ID</th>";
  cols.forEach((c) => (html += `<th>${labels[c]}</th>`));
  html += "<th>操作</th></tr></thead><tbody>";
  d.rows.forEach((r) => {
    html += `<tr><td class="chk"><input type="checkbox" class="rawChk" value="${r.id}"></td><td>${r.id}</td>`;
    cols.forEach((c) => (html += `<td>${esc(r[c])}</td>`));
    html += `<td class="row-actions"><button class="btn" data-edit="${r.id}" data-need="writer">编辑</button><button class="btn danger" data-del="${r.id}" data-need="admin">删</button></td></tr>`;
  });
  html += "</tbody></table>";
  $("#rawTableWrap").innerHTML = html;
  renderPager(d.total, d.page, d.size, "#rawPager", (p) => { rawState.page = p; loadRaw(); });
  $("#rawTotal").textContent = d.total;
  $("#rawSelAll").checked = rawState.selAll;
  $$("#rawTableWrap [data-edit]").forEach((b) => b.onclick = () => editRaw(+b.dataset.edit, cols, labels));
  $$("#rawTableWrap [data-del]").forEach((b) => b.onclick = async () => {
    if (!confirm("确认删除该行？")) return;
    await fetch("/api/raw/" + b.dataset.del, { method: "DELETE" });
    loadRaw(); loadDashSilent();
  });
  $("#rawCheckAll").onchange = (e) => {
    $$("#rawTableWrap .rawChk").forEach((b) => (b.checked = e.target.checked));
  };
  applyRole();
}

function editRaw(id, cols, labels) {
  const row = [...$$("#rawTableWrap tr")].find((tr) => tr.querySelector(`[data-edit="${id}"]`));
  const cells = row ? row.children : [];
  const vals = {};
  // 第0列=勾选框，第1列=ID，其后为数据列
  cols.forEach((c, i) => (vals[c] = cells[i + 2] ? cells[i + 2].textContent : ""));
  openModal("编辑原始数据 #" + id, cols.map((c) => field(c, labels[c], vals[c])), async (v) => {
    const r = await fetch("/api/raw/" + id, { method: "PUT", headers: jh(), body: JSON.stringify(v) });
    if (!r.ok) { showModalErr(await r.json()); return; }
    closeModal(); loadRaw(); loadDashSilent();
  });
}

$("#rawAddBtn").onclick = () => {
  const cols = rawState.cols, labels = rawState.labels;
  openModal("新增原始数据", cols.map((c) => field(c, labels[c], "")), async (v) => {
    const r = await fetch("/api/raw", { method: "POST", headers: jh(), body: JSON.stringify(v) });
    if (!r.ok) { showModalErr(await r.json()); return; }
    closeModal(); rawState.page = 1; loadRaw(); loadDashSilent();
  });
};
$("#rawBulkBtn").onclick = () => { $("#bulkText").value = ""; $("#bulkMsg").textContent = ""; $("#bulkModal").classList.remove("hidden"); };
$("#bulkClose").onclick = () => $("#bulkModal").classList.add("hidden");
$("#bulkCancel").onclick = () => $("#bulkModal").classList.add("hidden");

let bulkRows = [];
$("#bulkParse").onclick = async () => {
  const text = $("#bulkText").value.trim();
  if (!text) { $("#bulkMsg").textContent = "请先粘贴内容"; return; }
  const meta = await apiJson("/api/columns");
  const labels = meta.labels; // key->中文
  const inv = {}; Object.keys(labels).forEach((k) => (inv[labels[k]] = k));
  const lines = text.split(/\r?\n/).filter((l) => l.trim() !== "");
  if (lines.length < 2) { $("#bulkMsg").textContent = "至少需表头+1 行数据"; return; }
  const head = lines[0].split(/\t/).map((s) => s.trim());
  const keymap = head.map((h) => inv[h] || null);
  bulkRows = [];
  for (let i = 1; i < lines.length; i++) {
    const cells = lines[i].split(/\t/);
    const obj = {};
    keymap.forEach((k, idx) => { if (k) obj[k] = (cells[idx] || "").trim(); });
    bulkRows.push(obj);
  }
  $("#bulkMsg").textContent = `已解析 ${bulkRows.length} 行，匹配列：${keymap.filter(Boolean).map((k) => labels[k]).join("、") || "无"}。点"确认导入"写入（后端会逐行校验）。`;
};
$("#bulkImport").onclick = async () => {
  if (!bulkRows.length) { $("#bulkMsg").textContent = "请先解析"; return; }
  const r = await fetch("/api/raw/bulk", { method: "POST", headers: jh(), body: JSON.stringify({ rows: bulkRows }) });
  const res = await r.json();
  if (res.error_count > 0) {
    const sample = res.errors.slice(0, 8).map((e) => `第${e.row}行: ${e.errors.join("; ")}`).join("\n");
    $("#bulkMsg").textContent = `导入 ${res.inserted} 行，${res.error_count} 行校验未通过：\n${sample}`;
  } else {
    $("#bulkMsg").textContent = `成功导入 ${res.inserted} 行！`;
    $("#bulkModal").classList.add("hidden");
    rawState.page = 1; loadRaw(); loadDashSilent();
  }
};

$("#rawSearchBtn").onclick = () => { rawState.search = $("#rawSearch").value; rawState.page = 1; resetSelAll(); loadRaw(); };
$("#rawSearch").onkeydown = (e) => { if (e.key === "Enter") $("#rawSearchBtn").click(); };
$("#rawOperator").onchange = () => { rawState.operator = $("#rawOperator").value; rawState.page = 1; resetSelAll(); loadRaw(); };
$("#rawGroup").onchange = () => { rawState.group = $("#rawGroup").value; rawState.page = 1; resetSelAll(); loadRaw(); };
$("#rawFactory").onchange = () => { rawState.factory = $("#rawFactory").value; rawState.page = 1; resetSelAll(); loadRaw(); };
function resetSelAll() {
  rawState.selAll = false;
  $("#rawSelAll").checked = false;
  $("#rawDelBtn").textContent = "批量删除";
}

// ---------- 批量删除 + Excel 导入 ----------
$("#rawDelBtn").onclick = async () => {
  // 跨页全选：按当前搜索/筛选条件整批删除所有匹配行
  if (rawState.selAll) {
    const n = parseInt($("#rawTotal").textContent || "0", 10);
    if (!n) { alert("当前筛选条件下没有可删除的记录"); return; }
    const cond = "搜索=" + (rawState.search || "（空）") +
      " 跟单员=" + (rawState.operator || "（全部）") +
      " 组别=" + (rawState.group || "（全部）") +
      " 厂商=" + (rawState.factory || "（全部）");
    if (!confirm("⚠️ 危险操作：确认删除全部 " + n + " 条匹配记录？\n这些记录跨越所有分页，删除后不可恢复！\n\n筛选条件：" + cond)) return;
    const r = await fetch("/api/raw/batch", {
      method: "DELETE", headers: jh(),
      body: JSON.stringify({
        scope: "all", search: rawState.search, operator: rawState.operator,
        operator_group: rawState.group, factory: rawState.factory,
      }),
    });
    if (!r.ok) { alert("删除失败：" + ((await r.json()).error || "")); return; }
    const res = await r.json();
    rawState.selAll = false; $("#rawSelAll").checked = false; $("#rawDelBtn").textContent = "批量删除";
    rawState.page = 1; loadRaw(); loadDashSilent();
    alert("已删除 " + (res.deleted || 0) + " 条记录");
    return;
  }
  const ids = [...$$("#rawTableWrap .rawChk:checked")].map((b) => +b.value);
  if (!ids.length) { alert("请先勾选要删除的记录"); return; }
  if (!confirm("确认删除选中的 " + ids.length + " 条记录？此操作不可恢复！")) return;
  const r = await fetch("/api/raw/batch", { method: "DELETE", headers: jh(), body: JSON.stringify({ ids }) });
  if (!r.ok) { alert("删除失败：" + ((await r.json()).error || "")); return; }
  loadRaw(); loadDashSilent();
};
// 全选所有匹配项（跨页）开关
$("#rawSelAll").onchange = (e) => {
  rawState.selAll = e.target.checked;
  if (rawState.selAll) {
    // 取消本页逐行勾选，避免语义混淆
    $$("#rawTableWrap .rawChk").forEach((b) => (b.checked = false));
  }
  $("#rawDelBtn").textContent = rawState.selAll ? ("删除全部 " + ($("#rawTotal").textContent || "0") + " 条") : "批量删除";
};
$("#rawExcelBtn").onclick = () => $("#rawExcelFile").click();
$("#rawExcelFile").onchange = async () => {
  const file = $("#rawExcelFile").files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append("file", file);
  $("#rawExcelBtn").textContent = "导入中…";
  try {
    const r = await fetch("/api/raw/import", { method: "POST", body: fd });
    const res = await r.json();
    if (res.error_count > 0) {
      const sample = res.errors.slice(0, 8).map((e) => "第" + e.row + "行: " + e.errors.join("; ")).join("\n");
      alert("导入 " + res.inserted + " 行，" + res.error_count + " 行校验未通过：\n" + sample);
      if (res.inserted > 0) { rawState.page = 1; loadRaw(); loadDashSilent(); }
    } else {
      alert("成功从 Excel 导入 " + res.inserted + " 行！");
      rawState.page = 1; loadRaw(); loadDashSilent();
    }
  } catch (e) {
    alert("导入失败：" + e);
  } finally {
    $("#rawExcelBtn").textContent = "从Excel导入";
    $("#rawExcelFile").value = "";
  }
};

// ---------- 02/03/04 筛选与搜索 ----------
$("#odSearchBtn").onclick = () => { odState.search = $("#odSearch").value; odState.page = 1; loadOrderDetail(); };
$("#odSearch").onkeydown = (e) => { if (e.key === "Enter") $("#odSearchBtn").click(); };
$("#odOperator").onchange = () => { odState.page = 1; loadOrderDetail(); };
$("#odGroup").onchange = () => { odState.page = 1; loadOrderDetail(); };
$("#odFactory").onchange = () => { odState.page = 1; loadOrderDetail(); };

// ---- 工序进度（03）绑定 ----
$("#pcSearchBtn").onclick = () => { pcState.search = $("#pcSearch").value.trim(); pcState.page = 1; loadProc(); };
$("#pcSearch").onkeydown = (e) => { if (e.key === "Enter") $("#pcSearchBtn").click(); };
$("#pcBatchBtn").onclick = () => openBatchStageModal();
$("#pcColEditBtn").onclick = () => openColumnEditor();
$("#ceClose").onclick = () => $("#ceModal").classList.add("hidden");
$("#ceCancel").onclick = () => $("#ceModal").classList.add("hidden");
// 点击弹层外部关闭
document.addEventListener("click", (e) => {
  if (!e.target.closest(".col-filter-pop") && !e.target.closest(".col-filter")) closeColFilterPop();
});
$("#pcBulkClose").onclick = () => $("#pcBulkModal").classList.add("hidden");
$("#pcBulkCancel").onclick = () => $("#pcBulkModal").classList.add("hidden");
$("#pcTplBtn").onclick = () => { window.location.href = "/api/proc/template"; };
$("#pcImportBtn").onclick = () => $("#pcExcelFile").click();
$("#pcExcelFile").onchange = async (e) => {
  const file = e.target.files[0];
  e.target.value = "";
  if (!file) return;
  const fd = new FormData();
  fd.append("file", file);
  try {
    const r = await fetch("/api/proc/import", { method: "POST", body: fd });
    const res = await r.json().catch(() => ({ error: "服务器未返回 JSON" }));
    if (!r.ok || res.error) {
      if (res.error_count) {
        const detail = (res.errors || []).slice(0, 10).map((x) => "第" + x.row + "行: " + x.errors.join("、")).join("\n");
        alert("导入部分失败（成功 " + (res.inserted || 0) + " 行）：\n" + detail + (res.errors.length > 10 ? "\n…" : ""));
      } else {
        alert("导入失败：" + (res.error || "未知错误"));
      }
    } else {
      toast("导入成功，新增 " + res.inserted + " 行", true);
      loadProc();
    }
  } catch (err) {
    alert("导入出错：" + err.message);
  }
};
$("#pcDetailClose").onclick = () => $("#pcDetailModal").classList.add("hidden");
$("#pcDetailClose2").onclick = () => $("#pcDetailModal").classList.add("hidden");

$("#gtSearchBtn").onclick = () => { gtState.search = $("#gtSearch").value; gtState.page = 1; loadGantt(); };
$("#gtSearch").onkeydown = (e) => { if (e.key === "Enter") $("#gtSearchBtn").click(); };
$("#gtOperator").onchange = () => { gtState.page = 1; loadGantt(); };
$("#gtGroup").onchange = () => { gtState.page = 1; loadGantt(); };
$("#gtFactory").onchange = () => { gtState.page = 1; loadGantt(); };
$("#gtBase").onchange = () => { gtState.page = 1; loadGantt(); };

$("#scBtn").onclick = () => loadShipchk();
$("#ctBase").onchange = () => loadContract();

// ---------- 逾期呆滞 ----------
function todayISO() { return new Date().toISOString().slice(0, 10); }
async function loadOverdue() {
  ovState.base = $("#ovBase").value || todayISO();
  ovState.group = $("#ovGroup").value;
  ovState.operator = $("#ovOperator").value;
  ovState.factory = $("#ovFactory").value;
  ovState.bucket = $("#ovBucket").value;
  const params = new URLSearchParams({
    base: ovState.base, group: ovState.group, operator: ovState.operator,
    factory: ovState.factory, bucket: ovState.bucket, page: ovState.page, size: ovState.size,
  });
  const d = await apiJson("/api/overdue?" + params.toString());
  const s = d.summary;
  const bc = s.buckets;
  const cards = [
    { l: "逾期总行数", v: s.overdue_count, c: s.overdue_count > 0 ? "warn" : "" },
    { l: "逾期总金额(元)", v: fmt(s.overdue_amt), c: "" },
    { l: "呆滞行数(>90天无跟进)", v: s.stale_count, c: s.stale_count > 0 ? "warn" : "" },
    { l: "0-30天", v: bc["0-30"].count, c: "" },
    { l: "31-90天", v: bc["31-90"].count, c: "" },
    { l: "91-180天", v: bc["91-180"].count, c: "" },
    { l: "181-365天", v: bc["181-365"].count, c: "" },
    { l: "365天以上", v: bc["365+"].count, c: "warn" },
  ];
  $("#ovSummary").innerHTML = cards.map((c) =>
    `<div class="kpi ${c.c}"><div class="v">${c.v}</div><div class="l">${c.l}</div></div>`).join("");
  drawChart("chartOvBucket", "bar", Object.keys(bc), Object.values(bc).map((x) => x.count), ["#ba7517"]);
  drawChart("chartOvAmt", "bar", Object.keys(bc), Object.values(bc).map((x) => Math.round(x.amt)), ["#185fa5"]);
  const cols = [
    { k: "id", l: "ID" }, { k: "contract_no", l: "合同号" }, { k: "product_code", l: "产品编号" },
    { k: "operator", l: "运营专员" }, { k: "operator_group", l: "运营组别" }, { k: "factory", l: "厂商" },
    { k: "outstanding_qty", l: "未到货量" }, { k: "amt", l: "金额" }, { k: "days", l: "逾期天数" },
    { k: "bucket", l: "档" }, { k: "stale", l: "呆滞", fmt: (v) => (v ? "是" : "否") }, { k: "followup_date", l: "最近跟进" },
  ];
  renderTable("#ovTableWrap", cols, d.rows);
  renderPager(d.total, d.page, d.size, "#ovPager", (p) => { ovState.page = p; loadOverdue(); });
}
$("#ovBase").onchange = () => { ovState.page = 1; loadOverdue(); };
$("#ovGroup").onchange = () => { ovState.page = 1; loadOverdue(); };
$("#ovOperator").onchange = () => { ovState.page = 1; loadOverdue(); };
$("#ovFactory").onchange = () => { ovState.page = 1; loadOverdue(); };
$("#ovBucket").onchange = () => { ovState.page = 1; loadOverdue(); };

// ---------- 合同冗余 ----------
async function loadRedundancy() {
  rdState.base = $("#rdBase").value || todayISO();
  const d = await apiJson("/api/redundancy?base=" + rdState.base);
  const s = d.summary;
  const cards = [
    { l: "在手未交总量", v: fmt(s.total_qty), c: "" },
    { l: "在手未交金额(元)", v: fmt(s.total_amt), c: "" },
    { l: "涉及产品数", v: s.product_count, c: "" },
    { l: "不可接单组合数", v: s.no_take_count, c: s.no_take_count > 0 ? "warn" : "" },
  ];
  $("#rdSummary").innerHTML = cards.map((c) =>
    `<div class="kpi ${c.c}"><div class="v">${c.v}</div><div class="l">${c.l}</div></div>`).join("");
  const cyc = s.cycle;
  drawChart("chartCyc", "bar", Object.keys(cyc), Object.values(cyc).map((x) => Math.round(x.amt)), ["#0c447c", "#185fa5", "#2e8bc0", "#ba7517", "#c0392b", "#7d3c98"]);
  renderRdTable(d);
}
$$(".sub-tab").forEach((t) => {
  t.onclick = () => {
    $$(".sub-tab").forEach((x) => x.classList.remove("active"));
    t.classList.add("active");
    rdState.sub = t.dataset.sub;
    loadRedundancy();
  };
});
$("#rdBase").onchange = () => loadRedundancy();

function renderRdTable(d) {
  const sub = rdState.sub;
  let cols, rows;
  if (sub === "by_group") {
    cols = [{ k: "name", l: "运营组别" }, { k: "qty", l: "未交量" }, { k: "amt", l: "未交金额" }, { k: "count", l: "行数" }];
    rows = d.by_group;
  } else if (sub === "by_factory") {
    cols = [{ k: "name", l: "厂商" }, { k: "qty", l: "未交量" }, { k: "amt", l: "未交金额" }, { k: "count", l: "行数" }];
    rows = d.by_factory;
  } else if (sub === "by_product") {
    cols = [{ k: "product_code", l: "产品编号" }, { k: "qty", l: "公司级未交量" }, { k: "amt", l: "未交金额" },
      { k: "forecast", l: "未来3月预测" }, { k: "grade", l: "等级" }, { k: "decision", l: "接单判断", badge: true }];
    rows = d.by_product;
  } else {
    cols = [{ k: "product_code", l: "产品编号" }, { k: "operator", l: "运营专员" }, { k: "qty", l: "未交量" },
      { k: "amt", l: "未交金额" }, { k: "forecast", l: "预测" }, { k: "grade", l: "等级" }, { k: "decision", l: "接单判断", badge: true }];
    rows = d.by_sku_operator;
  }
  renderTable("#rdTableWrap", cols, rows);
}

// ---------- 通用表格 ----------
function renderTable(sel, cols, rows) {
  let html = "<table><thead><tr>";
  cols.forEach((c) => (html += `<th>${c.l}</th>`));
  html += "</tr></thead><tbody>";
  rows.forEach((r) => {
    html += "<tr>";
    cols.forEach((c) => {
      let v = r[c.k];
      if (c.fmt) v = c.fmt(v);
      if (c.badge) {
        html += `<td><span class="badge ${badgeClass(v)}">${esc(v)}</span></td>`;
      } else {
        html += `<td>${esc(v)}</td>`;
      }
    });
    html += "</tr>";
  });
  html += "</tbody></table>";
  $(sel).innerHTML = html;
}

function prodFields() { return [f("product_code", "产品编号"), f("product_name", "中文品名"), f("craft_category", "工艺品类")]; }
function shipFields() { return [f("sku", "SKU"), f("team", "团队"), f("operator", "运营"), f("factory", "工厂"), f("follower", "跟单"), f("ship_qty", "发货数量", "number"), f("ship_date", "发货时间", "date")]; }
function fcFields() { return [f("product_code", "产品编号"), f("operator", "运营专员"), f("forecast_qty", "未来3月预计出货数量", "number")]; }

async function loadTable(tab, apiPath, fields, table) {
  const idp = TAB_ID[tab] || tab;  // DOM id 前缀：products->prod / shipping->ship / forecast->fc
  const d = await apiJson(apiPath);
  let html = "<table><thead><tr>";
  fields.forEach((fdef) => (html += `<th>${fdef.label}</th>`));
  html += "<th>操作</th></tr></thead><tbody>";
  d.rows.forEach((r) => {
    html += "<tr>";
    fields.forEach((fdef) => (html += `<td>${esc(r[fdef.name])}</td>`));
    html += `<td><button class="btn danger" data-del="${r.id}" data-need="admin">删</button></td></tr>`;
  });
  html += "</tbody></table>";
  $("#" + idp + "TableWrap").innerHTML = html;
  const addBtn = $("#" + idp + "AddBtn");
  if (addBtn) {
    addBtn.dataset.need = (tab === "shipping") ? "writer" : "admin";
    addBtn.onclick = () => openModal("新增" + tab, fields, async (v) => {
      const r = await fetch(apiPath, { method: "POST", headers: jh(), body: JSON.stringify(v) });
      if (!r.ok) { showModalErr(await r.json()); return; }
      closeModal(); loadTable(tab, apiPath, fields, table);
    });
  }
  $$("#" + idp + "TableWrap [data-del]").forEach((b) => b.onclick = async () => {
    if (!confirm("确认删除？")) return;
    await fetch(apiPath + "/" + b.dataset.del, { method: "DELETE" });
    loadTable(tab, apiPath, fields, table);
  });
  applyRole();
}

// ---------- 02 订单明细 ----------
async function loadOrderDetail() {
  odState.operator = $("#odOperator").value;
  odState.group = $("#odGroup").value;
  odState.factory = $("#odFactory").value;
  const params = new URLSearchParams({
    page: odState.page, size: odState.size, search: odState.search,
    operator: odState.operator, operator_group: odState.group, factory: odState.factory,
  });
  const d = await apiJson("/api/order-detail?" + params.toString());
  const cols = [
    { k: "contract_no", l: "合同号" }, { k: "product_code", l: "产品编号" },
    { k: "team", l: "团队" }, { k: "operator", l: "运营专员" }, { k: "factory", l: "厂商" },
    { k: "craft_category", l: "工艺品类", fmt: (v) => v || "待确认" },
    { k: "order_qty", l: "订单总量" }, { k: "delivered_qty", l: "已交付" },
    { k: "outstanding_qty", l: "未到货总量" }, { k: "contract_total", l: "合同总数" },
    { k: "tier", l: "数量档位" }, { k: "remain", l: "待交付" },
    { k: "complete_rate", l: "完成率", fmt: (v) => (v * 100).toFixed(1) + "%" },
    { k: "status", l: "交期状态", badge: true },
    { k: "overdue_days", l: "逾期天数" }, { k: "risk", l: "风险等级", badge: true },
  ];
  renderTable("#odTableWrap", cols, d.rows);
  renderPager(d.total, d.page, d.size, "#odPager", (p) => { odState.page = p; loadOrderDetail(); });
  // 本页汇总
  let remSum = 0, odSum = 0, overdue = 0;
  d.rows.forEach((r) => {
    remSum += (+r.remain || 0);
    odSum += (+r.outstanding_qty || 0) * (+r.unit_price || 0);
    if (r.status === "已逾期") overdue++;
  });
  const cards = [
    { l: "本页合同行数", v: d.rows.length + " / 共 " + d.total, c: "" },
    { l: "本页待交付合计", v: fmt(remSum), c: "" },
    { l: "本页在手金额(元)", v: fmt(Math.round(odSum)), c: "" },
    { l: "本页逾期行数", v: overdue, c: overdue > 0 ? "warn" : "" },
  ];
  $("#odSummary").innerHTML = cards.map((c) => `<div class="kpi ${c.c}"><div class="v">${c.v}</div><div class="l">${c.l}</div></div>`).join("");
}

// ---------- 03 工序进度（独立数据源，按 产品编号+合同号 汇总） ----------
function fmtNum(v) {
  if (v === null || v === undefined || v === "") return "-";
  const n = Number(v);
  if (isNaN(n)) return v;
  return n.toLocaleString("zh-CN", { maximumFractionDigits: 2 });
}

/* ====================================================
   工序进度：列定义 + 列状态(可见/顺序/排序/筛选) + Excel 式列筛选 + 列名编辑
   ==================================================== */
const PC_COLUMNS = [
  { k: "checkbox", label: "", type: "checkbox", hideable: false },
  { k: "contract_no", label: "合同号", type: "text", filterable: true },
  { k: "product_code", label: "产品编号", type: "text", filterable: true },
  { k: "spu", label: "SPU", type: "text", filterable: true },
  { k: "craft_category", label: "工艺品类", type: "text", filterable: true },
  { k: "order_qty", label: "订单数量", type: "number", filterable: true },
  { k: "contract_date", label: "合同日期", type: "date", filterable: true },
  { k: "ship_date", label: "出货日期", type: "date", filterable: true },
  { k: "customer_complaint", label: "产品主要客诉点", type: "edit_text", meta: "customer_complaint", filterable: true },
  { k: "inspection_focus", label: "检验重点", type: "edit_select", meta: "inspection_focus", filterable: true },
  { k: "inspection_time", label: "验货时间", type: "edit_date", meta: "inspection_time", filterable: true },
  { k: "inspection_result", label: "验货结果", type: "edit_text", meta: "inspection_result", filterable: true },
  { k: "stages", label: "工序进度（状态机）", type: "stages", filterable: false },
  { k: "current", label: "当前工序", type: "text", filterable: true },
  { k: "block", label: "卡点", type: "text", filterable: true },
  { k: "progress", label: "完成度", type: "progress", filterable: true },
  { k: "detail", label: "明细", type: "detail", filterable: false, hideable: false },
  { k: "actions", label: "操作", type: "actions", filterable: false, hideable: false },
];
const PC_DEFAULT_ORDER = PC_COLUMNS.map((c) => c.k);
const pcColState = { visible: {}, order: [], sort: { key: null, dir: null }, filters: {} };
function loadColState() {
  try {
    const o = JSON.parse(localStorage.getItem("pc_col_state") || "{}");
    pcColState.order = (Array.isArray(o.order) && o.order.length) ? o.order.slice() : PC_DEFAULT_ORDER.slice();
    pcColState.visible = o.visible || {};
    pcColState.sort = o.sort || { key: null, dir: null };
    pcColState.filters = o.filters || {};
  } catch (e) { pcColState.order = PC_DEFAULT_ORDER.slice(); }
  PC_COLUMNS.forEach((c) => { if (!(c.k in pcColState.visible)) pcColState.visible[c.k] = true; });
  PC_COLUMNS.forEach((c) => { if (!pcColState.order.includes(c.k)) pcColState.order.push(c.k); });
}
function saveColState() {
  localStorage.setItem("pc_col_state", JSON.stringify({
    visible: pcColState.visible, order: pcColState.order,
    sort: pcColState.sort, filters: pcColState.filters
  }));
}
function visCols() {
  return pcColState.order.filter((k) => pcColState.visible[k] !== false)
    .map((k) => PC_COLUMNS.find((c) => c.k === k)).filter(Boolean);
}

/* 渲染工序进度表：应用排序 + 列筛选 + 列可见性 */
function renderProcTable() {
  const rows = lastProcRows || [];
  const cs = visCols();
  let rs = rows.slice();
  if (pcColState.sort.key && pcColState.sort.dir) {
    const sk = pcColState.sort.key, dir = pcColState.sort.dir;
    rs.sort((a, b) => {
      if (sk === "progress") { const va = a.progress || 0, vb = b.progress || 0; return dir === "asc" ? va - vb : vb - va; }
      const va = (a[sk] || "").toString(), vb = (b[sk] || "").toString();
      return dir === "asc" ? va.localeCompare(vb, "zh-CN") : vb.localeCompare(va, "zh-CN");
    });
  }
  Object.keys(pcColState.filters || {}).forEach((k) => {
    const f = pcColState.filters[k]; if (!f) return;
    if (f.values && f.values.size) rs = rs.filter((r) => f.values.has(String(r[k] || "")));
    if (f.text) rs = rs.filter((r) => String(r[k] || "").toLowerCase().includes(f.text.toLowerCase()));
  });
  const writer = currentRole !== "viewer";
  let head = "<tr>";
  cs.forEach((c) => {
    if (c.type === "checkbox") head += '<th class="chk"><input type="checkbox" id="pcSelAll"></th>';
    else if (c.hideable === false) head += `<th>${esc(c.label)}</th>`;
    else head += `<th>${esc(c.label)}<button class="col-filter" data-col-filter="${c.k}" aria-label="列筛选">▼</button></th>`;
  });
  head += "</tr>";
  let html = `<table><thead>${head}</thead><tbody>`;
  rs.forEach((r) => {
    html += "<tr>";
    cs.forEach((c) => {
      const k = c.k;
      if (c.type === "checkbox") html += `<td class="chk"><input type='checkbox' class='pcSel' data-pkey='${esc(r.pkey)}' ${pcState.sel.has(r.pkey) ? "checked" : ""}></td>`;
      else if (c.type === "number") html += `<td class="num">${fmtNum(r[k])}</td>`;
      else if (c.type === "edit_text") html += `<td class="edit-cell"><input data-meta="${c.meta}" data-pc="${esc(r.product_code)}" data-cn="${esc(r.contract_no)}" value="${esc(r[k])}" placeholder="点击填写"></td>`;
      else if (c.type === "edit_select") {
        const opts = ["<option value=''>—</option>"].concat(procInspOptions.map((o) => `<option ${o === r[k] ? "selected" : ""}>${esc(o)}</option>`)).join("");
        html += `<td class="edit-cell"><select data-meta="${c.meta}" data-pc="${esc(r.product_code)}" data-cn="${esc(r.contract_no)}">${opts}</select></td>`;
      }
      else if (c.type === "edit_date") html += `<td class="edit-cell"><input type="date" data-meta="${c.meta}" data-pc="${esc(r.product_code)}" data-cn="${esc(r.contract_no)}" value="${esc(r[k])}"></td>`;
      else if (c.type === "stages") {
        const chips = r.stages.map((s) => {
          const cls = s.applicable ? ("stage-" + s.status) : "stage-na";
          const t = `${s.name}｜计划开工:${s.plan_start || "-"}｜计划完工:${s.due || "-"}｜实际:${s.actual || "-"}`;
          return `<span class="stg ${cls}" title="${esc(t)}" data-pkey="${esc(r.pkey)}" data-idx="${s.idx}">${esc(s.name)}</span>`;
        }).join("");
        html += `<td class="stg-row">${chips}</td>`;
      }
      else if (c.type === "progress") html += `<td class="num"><div class="pbar"><i style="width:${(r.progress * 100).toFixed(0)}%"></i></div><span class="pct">${(r.progress * 100).toFixed(0)}%</span></td>`;
      else if (c.type === "detail") html += `<td><button class="btn small" data-detail data-pc="${esc(r.product_code)}" data-cn="${esc(r.contract_no)}">明细(${r.detail_count})</button></td>`;
      else if (c.type === "actions") html += `<td><button class="btn ${writer ? "primary" : ""}" data-proc="${esc(r.pkey)}" ${writer ? "" : "disabled title='只读'"}>${writer ? "录入实际完成" : "查看"}</button></td>`;
      else html += `<td>${esc(r[k])}</td>`;
    });
    html += "</tr>";
  });
  html += "</tbody></table>";
  $("#pcTableWrap").innerHTML = html;
  const sa = $("#pcSelAll");
  if (sa) {
    sa.checked = rs.length > 0 && rs.every((r) => pcState.sel.has(r.pkey));
    sa.onchange = () => {
      $$("#pcTableWrap .pcSel").forEach((cb) => {
        const pk = cb.dataset.pkey;
        if (sa.checked) pcState.sel.add(pk); else pcState.sel.delete(pk);
        cb.checked = sa.checked;
      });
      $("#pcSelCount").textContent = `已选 ${pcState.sel.size} 行`;
    };
    $$("#pcTableWrap .pcSel").forEach((cb) => cb.onchange = () => {
      const pk = cb.dataset.pkey;
      if (cb.checked) pcState.sel.add(pk); else pcState.sel.delete(pk);
      $("#pcSelCount").textContent = `已选 ${pcState.sel.size} 行`;
    });
  }
}

/* 列筛选弹层（Excel 式：排序 + 按值筛选 + 隐藏） */
function closeColFilterPop() {
  document.querySelectorAll(".col-filter-pop").forEach((p) => p.remove());
}
function openColFilterPop(btn, key) {
  closeColFilterPop();
  const col = PC_COLUMNS.find((c) => c.k === key);
  const rect = btn.getBoundingClientRect();
  const pop = document.createElement("div");
  pop.className = "col-filter-pop";
  pop.dataset.colKey = key;
  const cur = pcColState.sort;
  const f = pcColState.filters[key] || {};
  pop.innerHTML = `
    <div class="cf-section">
      <div class="cf-label">排序</div>
      <button data-cf-sort="asc" class="${cur.key === key && cur.dir === "asc" ? "cur" : ""}">↑ 升序</button>
      <button data-cf-sort="desc" class="${cur.key === key && cur.dir === "desc" ? "cur" : ""}">↓ 降序</button>
      <button data-cf-sort="" class="${!(cur.key === key && cur.dir) ? "cur" : ""}">无</button>
    </div>
    ${col && col.filterable ? `
    <div class="cf-section">
      <div class="cf-label">按值筛选</div>
      <input class="cf-text" placeholder="包含..." value="${esc(f.text || "")}">
      <div class="cf-values" data-key="${key}"></div>
    </div>` : ""}
    <div class="cf-section cf-foot">
      <button class="cf-hide">隐藏此列</button>
    </div>
  `;
  document.body.appendChild(pop);
  pop.style.left = rect.left + "px";
  pop.style.top = (rect.bottom + 6) + "px";
  if (pop.querySelector(".cf-values")) {
    const vals = Array.from(new Set(lastProcRows.map((r) => String(r[key] || ""))))
      .filter((v) => v !== "").sort((a, b) => a.localeCompare(b, "zh-CN"));
    const f2 = pcColState.filters[key] || {};
    const checkedAll = !f2.values || f2.values.size === 0 || f2.values.size === vals.length;
    const wrap = pop.querySelector(".cf-values");
    wrap.innerHTML = `<label class="cf-all"><input type="checkbox" data-cf-all ${checkedAll ? "checked" : ""}> 全选 (${vals.length})</label>` +
      vals.map((v) => {
        const on = f2.values && f2.values.size > 0 ? f2.values.has(v) : checkedAll;
        return `<label><input type="checkbox" data-cf-val="${esc(v)}" ${on ? "checked" : ""}> ${esc(v)}</label>`;
      }).join("");
  }
  pop.addEventListener("click", (e) => {
    const s = e.target.closest("[data-cf-sort]");
    if (s) {
      pcColState.sort = s.dataset.cfSort ? { key, dir: s.dataset.cfSort } : { key: null, dir: null };
      saveColState(); closeColFilterPop(); renderProcTable(); return;
    }
    const all = e.target.closest("[data-cf-all]");
    if (all) {
      const on = all.checked;
      pop.querySelectorAll("[data-cf-val]").forEach((c) => c.checked = on);
      applyCfVals(key, pop); return;
    }
    if (e.target.closest("[data-cf-val]")) { applyCfVals(key, pop); return; }
    if (e.target.closest(".cf-hide")) {
      pcColState.visible[key] = false;
      saveColState(); closeColFilterPop(); renderProcTable(); return;
    }
  });
  const ti = pop.querySelector(".cf-text");
  if (ti) ti.addEventListener("input", (e) => {
    if (!pcColState.filters[key]) pcColState.filters[key] = { values: null, text: "" };
    pcColState.filters[key].text = e.target.value;
    saveColState(); renderProcTable();
  });
}
function applyCfVals(key, pop) {
  const checks = pop.querySelectorAll("[data-cf-val]:checked");
  const set = new Set(Array.from(checks).map((c) => c.dataset.cfVal));
  if (!pcColState.filters[key]) pcColState.filters[key] = { values: null, text: "" };
  pcColState.filters[key].values = set;
  saveColState(); renderProcTable();
}

/* 列名编辑弹窗（统一管理：显示 + 顺序） */
function openColumnEditor() {
  $("#ceModalBody").innerHTML = `
    <p class="hint small">勾选控制是否显示；点击 ↑↓ 调整顺序。设置存于浏览器本地（localStorage），不影响他人。</p>
    <div class="ce-list" id="ceList"></div>
    <div class="ce-actions"><button class="btn small" id="ceReset">重置默认</button></div>
  `;
  renderColEditor();
  $("#ceReset").onclick = () => {
    pcColState.order = PC_DEFAULT_ORDER.slice();
    pcColState.visible = {}; PC_COLUMNS.forEach((c) => pcColState.visible[c.k] = true);
    pcColState.sort = { key: null, dir: null }; pcColState.filters = {};
    saveColState(); renderColEditor(); renderProcTable();
  };
  $("#ceModal").classList.remove("hidden");
}
function renderColEditor() {
  const html = pcColState.order.map((k, i) => {
    const c = PC_COLUMNS.find((x) => x.k === k);
    if (!c || c.hideable === false) return "";
    const vis = pcColState.visible[k] !== false;
    const upDis = i === 0 ? "disabled" : "";
    const dnDis = i === pcColState.order.length - 1 ? "disabled" : "";
    return `<div class="ce-row" data-key="${k}">
      <label class="ce-vis"><input type="checkbox" class="ce-chk" ${vis ? "checked" : ""}> 显示</label>
      <span class="ce-label">${esc(c.label)}</span>
      <button class="ce-up" ${upDis}>↑</button>
      <button class="ce-down" ${dnDis}>↓</button>
    </div>`;
  }).join("");
  $("#ceList").innerHTML = html;
  $("#ceList").querySelectorAll(".ce-chk").forEach((c) => c.onchange = (e) => {
    const key = e.target.closest(".ce-row").dataset.key;
    pcColState.visible[key] = e.target.checked;
    saveColState(); renderProcTable();
  });
  $("#ceList").querySelectorAll(".ce-up").forEach((b) => b.onclick = () => {
    const key = b.closest(".ce-row").dataset.key;
    const i = pcColState.order.indexOf(key);
    if (i > 0) { pcColState.order.splice(i - 1, 0, pcColState.order.splice(i, 1)[0]); saveColState(); renderColEditor(); renderProcTable(); }
  });
  $("#ceList").querySelectorAll(".ce-down").forEach((b) => b.onclick = () => {
    const key = b.closest(".ce-row").dataset.key;
    const i = pcColState.order.indexOf(key);
    if (i < pcColState.order.length - 1) { pcColState.order.splice(i + 1, 0, pcColState.order.splice(i, 1)[0]); saveColState(); renderColEditor(); renderProcTable(); }
  });
}

async function loadProc() {
  pcState.search = $("#pcSearch").value.trim();
  const params = new URLSearchParams({ page: pcState.page, size: pcState.size, search: pcState.search });
  const d = await apiJson("/api/proc?" + params.toString());
  lastProcRows = d.rows;
  lastProcAllKeys = d.all_keys || d.rows.map((r) => r.pkey);
  procInspOptions = d.inspection_focus_options || INSPECTION_FOCUS_FALLBACK;
  procStages = d.stages || [];
  $("#pcSelCount").textContent = `已选 ${pcState.sel.size} 行`;
  renderProcTable();
  renderPager(d.total, d.page, d.size, "#pcPager", (p) => { pcState.page = p; loadProc(); });
}

// 工序进度交互：一次性事件委托（色块/按钮/明细/内联编辑），避免翻页后绑定丢失
(function bindProcDelegated() {
  document.body.addEventListener("click", (e) => {
    const wrap = $("#pcTableWrap");
    if (!wrap || !wrap.contains(e.target)) return;
    const cfBtn = e.target.closest && e.target.closest(".col-filter");
    if (cfBtn) { openColFilterPop(cfBtn, cfBtn.dataset.colFilter); return; }
    const stg = e.target.closest && e.target.closest(".stg");
    if (stg) {
      if (currentRole === "viewer") return;
      openProcStageModal(stg.dataset.pkey, +stg.dataset.idx);
      return;
    }
    const btn = e.target.closest && e.target.closest("[data-proc]");
    if (btn) {
      if (btn.disabled) return;
      openProcModal(btn.dataset.proc, lastProcRows);
      return;
    }
    const det = e.target.closest && e.target.closest("[data-detail]");
    if (det) {
      openProcDetail(det.dataset.pc, det.dataset.cn);
      return;
    }
  });
  document.body.addEventListener("change", (e) => {
    const el = e.target.closest && e.target.closest("[data-meta]");
    if (!el) return;
    const wrap = $("#pcTableWrap");
    if (!wrap || !wrap.contains(el)) return;
    saveProcMeta(el.dataset.pc, el.dataset.cn, el.dataset.meta, el.value, el);
  });
})();

function openProcModal(pkey, rows) {
  const r = rows.find((x) => x.pkey === pkey);
  if (!r) return;
  $("#modalTitle").textContent = "录入实际完成 · " + r.contract_no + " / " + r.product_code;
  $("#modalErr").textContent = "";
  let body = "";
  r.stages.forEach((s) => {
    const dis = s.applicable ? "" : "disabled";
    body += `<div class="field proc-stage"><label>${esc(s.name)}${s.applicable ? "" : "（不适用）"}</label><input data-sidx="${s.idx}" type="date" value="${esc(s.actual)}" ${dis}></div>`;
  });
  $("#modalBody").innerHTML = body;
  $("#modalSave").onclick = async () => {
    let ok = 0, fail = 0;
    for (const inp of $$("#modalBody [data-sidx]")) {
      const idx = +inp.dataset.sidx;
      const val = inp.value.trim();
      const orig = r.stages[idx].actual || "";
      if (val === orig) continue;
      const res = await fetch("/api/proc/stage", {
        method: "POST", headers: jh(),
        body: JSON.stringify({ pkey: pkey, stage_idx: idx, actual_date: val }),
      });
      if (res.ok) ok++; else fail++;
    }
    if (fail === 0) { closeModal(); loadProc(); loadDashSilent(); }
    else $("#modalErr").textContent = "部分保存失败：" + fail + " 项";
  };
  $("#modal").classList.remove("hidden");
}
function openProcStageModal(pkey, idx) {
  const r = lastProcRows.find((x) => x.pkey === pkey);
  const cur = r ? (r.stages[idx].actual || "") : "";
  const name = r ? r.stages[idx].name : "";
  $("#modalTitle").textContent = "录入「" + name + "」实际完成";
  $("#modalErr").textContent = "";
  $("#modalBody").innerHTML = `<div class="field"><label>实际完成日期（留空=清除）</label><input id="procOneDate" type="date" value="${esc(cur)}"></div>`;
  $("#modalSave").onclick = async () => {
    const val = $("#procOneDate").value.trim();
    const res = await fetch("/api/proc/stage", {
      method: "POST", headers: jh(),
      body: JSON.stringify({ pkey: pkey, stage_idx: idx, actual_date: val }),
    });
    if (res.ok) { closeModal(); loadProc(); loadDashSilent(); }
    else $("#modalErr").textContent = "保存失败";
  };
  $("#modal").classList.remove("hidden");
}
async function openProcDetail(pc, cn) {
  $("#pcDetailTitle").textContent = "导入明细 · " + cn + " / " + pc;
  $("#pcDetailBody").innerHTML = "<p class='hint'>加载中…</p>";
  $("#pcDetailModal").classList.remove("hidden");
  try {
    const d = await apiJson("/api/proc/detail?product_code=" + encodeURIComponent(pc) + "&contract_no=" + encodeURIComponent(cn));
    if (!d.rows.length) { $("#pcDetailBody").innerHTML = "<p class='hint'>暂无明细数据。</p>"; return; }
    let h = "<table><thead><tr><th>ID</th><th>合同号</th><th>产品编号</th><th>SPU</th><th class='num'>订单数量</th><th>合同日期</th><th>出货日期</th><th>工艺品类</th><th>导入批次</th></tr></thead><tbody>";
    d.rows.forEach((x) => {
      h += `<tr><td>${x.id}</td><td>${esc(x.contract_no)}</td><td>${esc(x.product_code)}</td><td>${esc(x.spu)}</td><td class="num">${fmtNum(x.order_qty)}</td><td>${esc(x.contract_date)}</td><td>${esc(x.ship_date)}</td><td>${esc(x.craft_category)}</td><td>${esc(x.import_batch || "")}</td></tr>`;
    });
    h += "</tbody></table>";
    $("#pcDetailBody").innerHTML = h;
  } catch (e) {
    $("#pcDetailBody").innerHTML = "<p class='err'>加载失败：" + esc(e.message) + "</p>";
  }
}
function saveProcMeta(pc, cn, field, val, el) {
  const optText = (field === "inspection_focus" && el && el.options && el.selectedIndex >= 0)
    ? el.options[el.selectedIndex].text : val;
  fetch("/api/proc/meta", {
    method: "POST", headers: jh(),
    body: JSON.stringify({ product_code: pc, contract_no: cn, [field]: val }),
  }).then((r) => {
    if (r.ok) toast("已保存：" + field, true);
    else return r.json().then((j) => { throw new Error(j.error || ("HTTP " + r.status)); });
  }).catch((e) => toast("保存失败：" + e.message, false));
}
async function openBatchStageModal() {
  const writer = currentRole !== "viewer";
  if (!writer) { alert("当前为只读角色，无法批量更新"); return; }
  const stages = procStages.map((n, i) => `<option value="${i}">${i + 1}. ${esc(n)}</option>`).join("");
  const selCount = pcState.sel.size;
  const allCount = lastProcAllKeys.length;
  const body = `
    <div class="field" style="grid-column:1 / -1">
      <label>应用范围</label>
      <select id="pcbScope">
        <option value="selected">仅当前勾选 (${selCount} 行)</option>
        <option value="all">当前搜索/筛选下的全部 (${allCount} 行)</option>
      </select>
    </div>
    <div class="field"><label>工序</label><select id="pcbStage">${stages}</select></div>
    <div class="field"><label>实际完成日期（留空=清除该工序）</label><input id="pcbDate" type="date"></div>
    <div id="pcbSummary" class="hint small" style="grid-column:1 / -1"></div>
  `;
  $("#pcBulkModalBody").innerHTML = body;
  $("#pcBulkModalErr").textContent = "";
  $("#pcBulkModalSave").onclick = async () => {
    const scope = $("#pcbScope").value;
    const stage_idx = +$("#pcbStage").value;
    const actual_date = $("#pcbDate").value.trim();
    if (scope === "selected" && selCount === 0) {
      $("#pcBulkModalErr").textContent = "当前未勾选任何行，请先勾选，或切换到「当前搜索/筛选下的全部」";
      return;
    }
    let body = { stage_idx, actual_date };
    if (scope === "all") {
      body.scope = "all";
      body.search = pcState.search;
    } else {
      body.pkeys = Array.from(pcState.sel);
    }
    $("#pcBulkModalSave").disabled = true;
    $("#pcBulkModalErr").textContent = "提交中…";
    try {
      const r = await fetch("/api/proc/stage-batch", {
        method: "POST", headers: jh(), body: JSON.stringify(body),
      });
      const j = await r.json();
      if (r.ok) {
        $("#pcBulkModalErr").textContent = `✅ 已更新 ${j.updated || 0} 行 / 目标 ${j.target_count} 行`;
        $("#pcBulkModalSave").disabled = false;
        pcState.sel.clear();
        loadProc(); loadDashSilent();
      } else {
        $("#pcBulkModalErr").textContent = "❌ " + (j.error || ("HTTP " + r.status));
        $("#pcBulkModalSave").disabled = false;
      }
    } catch (e) {
      $("#pcBulkModalErr").textContent = "❌ 网络错误：" + e.message;
      $("#pcBulkModalSave").disabled = false;
    }
  };
  $("#pcBulkModal").classList.remove("hidden");
}

// ---------- 04 甘特 ----------
async function loadGantt() {
  gtState.operator = $("#gtOperator").value;
  gtState.group = $("#gtGroup").value;
  gtState.factory = $("#gtFactory").value;
  gtState.base = $("#gtBase").value || todayISO();
  const params = new URLSearchParams({
    base: gtState.base, page: gtState.page, size: gtState.size, search: gtState.search,
    operator: gtState.operator, operator_group: gtState.group, factory: gtState.factory,
  });
  const d = await apiJson("/api/gantt?" + params.toString());
  const weeks = d.weeks || [];
  const cur = weeks.findIndex((w) => w.is_current);
  let head = "<th>合同号</th><th>产品</th><th>厂商</th><th>工艺品类</th><th>交货日期</th><th class=\"gantt-head\">甘特（22 周）</th>";
  let html = `<table class="gantt"><thead><tr>${head}</tr><tr class="wk-head"><th></th><th></th><th></th><th></th><th></th><th><div class="wk-cells">` +
    weeks.map((w, i) => `<span class="wk ${i === cur ? "cur" : ""}" title="${w.week_start}~${w.week_end}">${w.week_start.slice(5)}</span>`).join("") +
    `</div></th></tr></thead><tbody>`;
  let overdueRows = 0;
  d.rows.forEach((r) => {
    const applicable = Math.max(1, r.plan_starts.length);
    const dd = parseISO(r.delivery_date);
    const baseD = parseISO(gtState.base);
    const isOverdue = dd && baseD && dd < baseD && (+r.outstanding_qty > 0);
    if (isOverdue) overdueRows++;
    const cells = r.cells.map((c, i) => {
      const ratio = c / applicable;
      const intensity = c > 0 ? (0.25 + ratio * 0.75) : 0;
      const wEnd = weeks[i] ? weeks[i].week_end : null;
      const red = isOverdue && wEnd && dd && parseISO(wEnd) >= dd;
      const bg = red ? "background:#f06565;" : (c > 0 ? `background:rgba(12,68,124,${intensity.toFixed(2)});` : "");
      const curMark = i === cur ? "box-shadow:inset 0 0 0 2px #ba7517;" : "";
      return `<span class="gc" style="${bg}${curMark}"></span>`;
    }).join("");
    html += `<tr><td>${esc(r.contract_no)}</td><td>${esc(r.product_code)}</td><td>${esc(r.factory)}</td><td>${esc(r.craft_category) || "待确认"}</td><td>${esc(r.delivery_date)}</td><td><div class="gbar">${cells}</div></td></tr>`;
  });
  html += "</tbody></table>";
  $("#gtTableWrap").innerHTML = html;
  renderPager(d.rows.length ? d.rows.length : 1, gtState.page, gtState.size, "#gtPager", (p) => { gtState.page = p; loadGantt(); });
  const cards = [
    { l: "本页合同行数", v: d.rows.length, c: "" },
    { l: "已逾期(红)行数", v: overdueRows, c: overdueRows > 0 ? "warn" : "" },
    { l: "当前周", v: cur >= 0 ? weeks[cur].week_start : "-", c: "" },
  ];
  $("#gtSummary").innerHTML = cards.map((c) => `<div class="kpi ${c.c}"><div class="v">${c.v}</div><div class="l">${c.l}</div></div>`).join("");
}
function parseISO(s) { if (!s) return null; const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; }

// ---------- 06 合同汇总 ----------
async function loadContract() {
  const base = $("#ctBase").value || todayISO();
  const rows = await apiJson("/api/contract?base=" + base);
  let qtySum = 0, amtSum = 0, done = 0, overdue = 0;
  rows.forEach((r) => {
    qtySum += (+r.outstanding_qty || 0);
    amtSum += (+r.outstanding_qty || 0) * (+r.unit_price || 0);
    if (r.status === "已交清") done++;
    if (r.status === "已逾期") overdue++;
  });
  const cards = [
    { l: "合同数", v: rows.length, c: "" },
    { l: "在手未交总量", v: fmt(qtySum), c: "" },
    { l: "在手未交金额(元)", v: fmt(Math.round(amtSum)), c: "" },
    { l: "已交清合同", v: done, c: "" },
    { l: "已逾期合同", v: overdue, c: overdue > 0 ? "warn" : "" },
  ];
  $("#ctSummary").innerHTML = cards.map((c) => `<div class="kpi ${c.c}"><div class="v">${c.v}</div><div class="l">${c.l}</div></div>`).join("");
  const cols = [
    { k: "contract_no", l: "合同号" }, { k: "factory", l: "厂商" },
    { k: "lines", l: "行数" }, { k: "skus", l: "SKU数" },
    { k: "order_qty", l: "订单总量" }, { k: "delivered_qty", l: "已交付" },
    { k: "outstanding_qty", l: "未到货量" },
    { k: "complete_rate", l: "完成率", fmt: (v) => (v * 100).toFixed(1) + "%" },
    { k: "min_dd", l: "最早交期" }, { k: "max_dd", l: "最晚交期" },
    { k: "days_to_ship", l: "距交期(天)", fmt: (v) => (v == null ? "-" : v) },
    { k: "status", l: "状态", badge: true },
    { k: "whole", l: "整单出货", badge: true },
  ];
  renderTable("#ctTableWrap", cols, rows);
}

// ---------- 10 发货需求判断 ----------
async function loadShipchk() {
  scState.start = $("#scStart").value;
  scState.end = $("#scEnd").value;
  const params = new URLSearchParams({ start: scState.start, end: scState.end });
  const d = await apiJson("/api/shipchk?" + params.toString());
  const k = d.kpi || {};
  const cards = [
    { l: "窗口内需求合计", v: fmt(k.demand_total), c: "" },
    { l: "供应(未入库)合计", v: fmt(k.supply_total), c: "" },
    { l: "差额(供应-需求)", v: fmt(k.gap), c: k.gap < 0 ? "warn" : "" },
    { l: "缺口行数", v: k.gap_rows, c: k.gap_rows > 0 ? "warn" : "" },
    { l: "满足率", v: ((k.satisfy_rate || 0) * 100).toFixed(1) + "%", c: "" },
  ];
  $("#scSummary").innerHTML = cards.map((c) => `<div class="kpi ${c.c}"><div class="v">${c.v}</div><div class="l">${c.l}</div></div>`).join("");
  const cols = [
    { k: "product_code", l: "产品编号" }, { k: "team", l: "团队" },
    { k: "demand", l: "需求" }, { k: "supply", l: "供应(未入库)" },
    { k: "contract", l: "合同数量" }, { k: "outstanding", l: "未到货量" },
    { k: "diff", l: "差额", fmt: (v) => fmt(v) },
    { k: "status", l: "状态", badge: true },
  ];
  renderTable("#scTableWrap", cols, d.rows || []);
}

// ---------- 07/08/09 参数·字典·标准 ----------
async function loadParams() {
  const d = await apiJson("/api/params");
  // ① 工序节拍
  const stages = d.stages || [];
  let html = "<table><thead><tr><th>工艺类目</th><th>数量档位</th><th>是否新单</th>" +
    stages.map((s) => `<th>${esc(s)}</th>`).join("") + "</tr></thead><tbody>";
  (d.tact || []).forEach((t) => {
    const offs = t.offsets.map((o) => {
      if (!o || o === "NA" || (o[0] == null && o[1] == null)) return "—";
      return o[0] + "," + o[1];
    }).join("</td><td>");
    html += `<tr><td>${esc(t.cat)}</td><td>${esc(t.tier)}</td><td>${esc(t.is_new)}</td><td>${offs}</td></tr>`;
  });
  html += "</tbody></table>";
  $("#pmTactWrap").innerHTML = html;
  // ② 阈值与字典
  const th = d.thresholds || {};
  const dict = [
    { l: "交期-紧急阈值(天)", v: th.urgent },
    { l: "交期-预警阈值(天)", v: th.warn },
    { l: "呆滞阈值(天)", v: th.dead },
    { l: "完成率<该值且紧急=高风险", v: th.lr1 },
    { l: "完成率<该值且预警=中风险", v: th.lr2 },
    { l: "数量档位", v: (d.tiers || []).join(" / ") },
    { l: "是否新单取值", v: (d.news || []).join(" / ") },
    { l: "标准工艺类目顺序", v: (d.cat_order || []).join(" / ") },
    { l: "默认前置期(无合同日时,天)", v: d.default_lead_days },
  ];
  $("#pmDictWrap").innerHTML = dict.map((x) => `<div class="kpi"><div class="v">${esc(x.v)}</div><div class="l">${x.l}</div></div>`).join("");
  // ⑨ 检验标准（产中检验参考，源自源模板 09 检验标准）
  $("#pmQcWrap").innerHTML = `<p class="hint">本表为产中检验标准参考框架，结构对应源文件「09检验标准」（首件检验 / 过程巡检 / 成品终检 / 抽样方案）。具体抽样数与合格判据以工厂质量协议为准；如需把该页内容完全搬上网页，请告知，我可将其建模为可维护数据表。</p>`;
}

// ---------- 备份 ----------
$("#backupBtn").onclick = async () => {
  const r = await fetch("/api/backup", { method: "POST" });
  const res = await r.json();
  $("#backupMsg").textContent = res.ok ? "已备份：" + res.path : ("备份失败：" + (res.error || ""));
};

// ---------- 工具 ----------
function field(name, label, value, type) {
  return { name, label, value: value || "", type: type || (isNum(name) ? "number" : isDate(name) ? "date" : "text") };
}
function f(name, label, type) { return field(name, label, "", type); }
function isNum(n) { return /qty|price|amt/.test(n); }
function isDate(n) { return /date/.test(n); }
function esc(v) { return v == null ? "" : String(v).replace(/[&<>"]/g, (s) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[s])); }
function badgeClass(v) {
  const s = String(v || "");
  if (/不可接单|已逾期|缺口|缺|紧急|高风险|呆滞/.test(s)) return "bad";
  if (/可接单|已交付|正常|满足|✔|整单出货/.test(s)) return "ok";
  return "wait";
}
function fmt(n) { return (n == null ? 0 : n).toLocaleString("zh-CN"); }
function jh() { return { "Content-Type": "application/json" }; }
function toast(msg, ok) {
  let box = $("#toastBox");
  if (!box) {
    box = document.createElement("div");
    box.id = "toastBox";
    box.className = "toast-box";
    document.body.appendChild(box);
  }
  const el = document.createElement("div");
  el.className = "toast " + (ok ? "ok" : "bad");
  el.textContent = msg;
  box.appendChild(el);
  requestAnimationFrame(() => el.classList.add("show"));
  setTimeout(() => { el.classList.remove("show"); setTimeout(() => el.remove(), 300); }, 2400);
}
function renderPager(total, page, size, sel, cb) {
  const pages = Math.ceil(total / size) || 1;
  let h = `共 ${total} 行 · 第 ${page}/${pages} 页 `;
  h += `<button ${page <= 1 ? "disabled" : ""} data-p="${page - 1}">上一页</button>`;
  h += `<button ${page >= pages ? "disabled" : ""} data-p="${page + 1}">下一页</button>`;
  $(sel).innerHTML = h;
  $$(sel + " [data-p]").forEach((b) => (b.onclick = () => cb(+b.dataset.p)));
}
async function loadDashSilent2() { try { await loadDash(); } catch (e) {} }

// ---------- 弹窗 ----------
function openModal(title, fields, onSave) {
  $("#modalTitle").textContent = title;
  $("#modalErr").textContent = "";
  $("#modalBody").innerHTML = fields.map((fdef) =>
    `<div class="field"><label>${fdef.label}</label><input data-name="${fdef.name}" type="${fdef.type}" value="${esc(fdef.value)}"></div>`).join("");
  $("#modalSave").onclick = () => {
    const v = {};
    $$("#modalBody [data-name]").forEach((inp) => {
      const val = inp.value.trim();
      v[inp.dataset.name] = isNum(inp.dataset.name) ? (val === "" ? null : +val) : val;
    });
    onSave(v);
  };
  $("#modal").classList.remove("hidden");
}
function showModalErr(res) {
  const det = res.details ? res.details.join("；") : (res.error || "保存失败");
  $("#modalErr").textContent = det;
}
function closeModal() { $("#modal").classList.add("hidden"); }
$("#modalClose").onclick = closeModal;
$("#modalCancel").onclick = closeModal;

// ---------- 产品资料/发货需求/销售预测：新增 + Excel导入 + 模板下载 ----------
const ADD_TITLE = { products: "产品资料", shipping: "发货需求", forecast: "销售预测" };
const TBL_NAME = { products: "products", shipping: "shipping_demands", forecast: "sales_forecast" };
const FIELD_FN = { products: prodFields, shipping: shipFields, forecast: fcFields };

function bindAdd(tab) {
  const apiPath = "/" + (tab === "products" ? "api/products" : "api/" + tab);
  const el = $("#" + (TAB_ID[tab] || tab) + "AddBtn");
  if (!el) return;
  el.onclick = () => {
    openModal("新增" + ADD_TITLE[tab], FIELD_FN[tab]().map((fd) => field(fd.name, fd.label, "", fd.type)),
      async (v) => {
        const r = await fetch(apiPath, { method: "POST", headers: jh(), body: JSON.stringify(v) });
        const res = await r.json();
        if (!r.ok) { showModalErr(res); return; }
        closeModal();
        loadTable(tab, apiPath, FIELD_FN[tab](), TBL_NAME[tab]);
      });
  };
}

function bindImport(tab) {
  const apiPath = "/" + (tab === "products" ? "api/products" : "api/" + tab);
  const idp = TAB_ID[tab] || tab;
  const fileId = idp + "ExcelFile";
  const impBtn = $("#" + idp + "ImportBtn");
  const fileEl = $("#" + fileId);
  if (!impBtn || !fileEl) return;
  impBtn.onclick = () => fileEl.click();
  fileEl.onchange = async (e) => {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    const fd = new FormData();
    fd.append("file", file);
    // 注意：不要手动设置 Content-Type，浏览器会自动带上 multipart 边界；cookie 随请求发送
    const r = await fetch(apiPath + "/import", { method: "POST", body: fd });
    let res;
    try { res = await r.json(); } catch (err) { res = { error: "服务器未返回 JSON" }; }
    if (!r.ok || res.error) {
      if (res.error_count) {
        const detail = (res.errors || []).slice(0, 8).map((x) => "第" + x.row + "行: " + x.errors.join("、")).join("\n");
        alert("导入部分失败（成功 " + (res.inserted || 0) + " 行）：\n" + detail + (res.errors.length > 8 ? "\n…" : ""));
      } else {
        alert("导入失败：" + (res.error || "未知错误"));
      }
    } else {
      alert("导入成功，新增 " + res.inserted + " 行" + (res.error_count ? "（" + res.error_count + " 行被跳过，见下方详情）" : ""));
    }
    loadTable(tab, apiPath, FIELD_FN[tab](), TBL_NAME[tab]);
  };
}

function bindTemplate(tab) {
  const apiPath = "/" + (tab === "products" ? "api/products" : "api/" + tab);
  const el = $("#" + (TAB_ID[tab] || tab) + "TplBtn");
  if (!el) return;
  el.onclick = () => { window.location.href = apiPath + "/template"; };
}

["products", "shipping", "forecast"].forEach((tab) => { bindAdd(tab); bindImport(tab); bindTemplate(tab); });

// ---------- 启动 ----------
(async function init() {
  // 免登录模式：默认以管理员权限进入，无需账号密码
  currentRole = "admin";
  hideLogin();
  loadColState();
  try { await refreshMe(); } catch (e) {}
  if (!currentRole) currentRole = "admin";
  applyRole();
  if ($("#logoutBtn")) $("#logoutBtn").style.display = "none";  // 免登录时无退出意义
  $("#ovBase").value = todayISO();
  $("#rdBase").value = todayISO();
  $("#gtBase").value = todayISO();
  $("#ctBase").value = todayISO();
  loadDash();
})();
