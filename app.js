(() => {
  "use strict";
  const el = id => document.getElementById(id);
  const storeKey = "packshift.plan.v1";
  const format = new Intl.NumberFormat("zh-CN", {maximumFractionDigits:2});
  const money = new Intl.NumberFormat("zh-CN", {minimumFractionDigits:2, maximumFractionDigits:2});
  const captions = {example:"示例数据", edited:"本地方案", imported:"已导入"};
  let plan = Planner.sample(), source = "example", filter = "all", pending = null, readVersion = 0, noticeTimer;

  function notify(message) {
    clearTimeout(noticeTimer);
    el("status").textContent = message;
    noticeTimer = setTimeout(() => { el("status").textContent = ""; }, 7000);
  }
  try {
    const saved = localStorage.getItem(storeKey);
    if (saved) {
      const data = JSON.parse(saved);
      plan = Planner.validate(data.plan);
      source = Object.hasOwn(captions,data.source) ? data.source : "edited";
    }
  } catch {
    notify("本地方案暂时无法读取，已显示示例。");
  }

  function save(next, label, message) {
    plan = Planner.validate(next);
    source = label;
    try {
      localStorage.setItem(storeKey, JSON.stringify({plan,source}));
    } catch {
      message += "。浏览器无法保存，请导出 Excel 备份";
    }
    render();
    notify(message);
  }

  function make(tag, className, value) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (value !== undefined) node.textContent = value;
    return node;
  }
  function edit(id) {
    Editor.open(plan, id, next => save(next, "edited", "已更新产品排期"));
  }
  function renderRow(item) {
    const tr = make("tr");
    const product = make("td");
    product.append(make("span","product-name",item.name), make("span","product-id",item.id));
    const stock = make("td");
    const bar = make("div",item.stock ? "bar" : "bar no-stock");
    const used = make("i");
    used.style.width = (item.stock ? 100 * item.used/item.stock : 0) + "%";
    bar.setAttribute("aria-hidden","true");
    bar.append(used);
    stock.append(make("span","stock-count",format.format(item.used) + " / " + format.format(item.stock)), bar);
    const action = make("td");
    const button = make("button","text-button edit-btn","编辑");
    button.setAttribute("aria-label", "编辑 " + item.name);
    button.addEventListener("click",() => edit(item.id));
    action.append(button);
    tr.append(product,stock,make("td","numeric",format.format(item.left)),
      make("td","numeric" + (item.amount ? " left-amount" : ""),"¥" + money.format(item.amount)),
      make("td","row-date",item.switchDay || "待排产"),make("td","row-date",item.orderDay || "待排产"),action);
    return tr;
  }
  function render() {
    const rows = Planner.calculate(plan);
    const stock = rows.reduce((sum,r) => sum+r.stock,0);
    const left = rows.reduce((sum,r) => sum+r.left,0);
    const cents = rows.reduce((sum,r) => sum + BigInt(Math.round(r.amount * 100)), 0n);
    const fraction = (cents % 100n).toString().padStart(2,"0");
    const affected = rows.filter(r => r.left > 0).length;
    const noDate = rows.filter(r => !r.switchDay).length;
    el("risk-amount").textContent = format.format(cents / 100n) + (fraction === "00" ? "" : "." + fraction);
    el("left-total").textContent = format.format(left);
    el("risk-caption").textContent = affected + " 款产品预计有旧包材剩余";
    el("risk-meter").style.width = (stock ? 100*left/stock : 0) + "%";
    el("used-caption").textContent = "预计耗用 " + format.format(stock - left) + " 件";
    el("cutoff").value = plan.cutoff;
    el("source").textContent = captions[source];
    el("all-count").textContent = rows.length;
    el("left-count").textContent = affected;
    el("pending-count").textContent = noDate;
    const q = el("search").value.toLowerCase().trim();
    const visible = rows.filter(r => (filter === "all" || (filter === "left" ? r.left > 0 : !r.switchDay))
      && (r.id.toLowerCase().includes(q) || r.name.toLowerCase().includes(q)));
    const fragment = document.createDocumentFragment();
    for (const item of visible) fragment.append(renderRow(item));
    el("products").replaceChildren(fragment);
    el("empty").hidden = visible.length !== 0;
    el("table-count").textContent = "显示 " + visible.length + " / " + rows.length + " 款产品";
  }

  el("search").addEventListener("input", render);
  document.querySelectorAll(".filter").forEach(button => button.addEventListener("click", () => {
    filter = button.dataset.filter;
    document.querySelectorAll(".filter").forEach(b => {
      b.classList.toggle("active", b === button);
      b.setAttribute("aria-pressed",String(b === button));
    });
    render();
  }));
  el("cutoff").addEventListener("change", () => {
    try { save({...plan, cutoff:el("cutoff").value}, "edited", "已按新截止日重新测算"); }
    catch (error) { el("cutoff").value = plan.cutoff; notify(error.message); }
  });
  el("add-product").addEventListener("click", () => edit(null));
  document.querySelector(".help-trigger").addEventListener("click", () => el("help").showModal());
  document.querySelectorAll("[data-close]").forEach(button => button.addEventListener("click", () => button.closest("dialog").close()));
  const sampleTemplate = () => Workbook.template({...Planner.sample(), cutoff:plan.cutoff});
  el("template").addEventListener("click",sampleTemplate);
  el("template-dialog").addEventListener("click",sampleTemplate);
  el("export").addEventListener("click", () => Workbook.exportResults(plan));
  el("import").addEventListener("click", () => {
    ++readVersion;
    pending = null;
    el("import-file").value = "";
    el("apply-import").disabled = true;
    el("import-error").textContent = el("import-preview").textContent = "";
    el("import-dialog").showModal();
  });
  el("import-file").addEventListener("change",async event => {
    const version = ++readVersion;
    const file = event.target.files[0];
    pending = null;
    el("apply-import").disabled = true;
    el("import-error").textContent = "";
    el("import-preview").textContent = file ? "正在读取工作簿…" : "";
    if (!file) return;
    try {
      if (file.size > 10 * 1024 * 1024) throw Error("请选择 10 MB 以内的工作簿");
      const buffer = await file.arrayBuffer();
      if (version !== readVersion) return;
      pending = Workbook.read(buffer, plan.cutoff);
      el("import-preview").textContent = "将导入 " + pending.items.length + " 款产品、" + pending.batches.length + " 条排产，截止日为 " + pending.cutoff + "。";
      el("apply-import").disabled = false;
    } catch (error) {
      if (version !== readVersion) return;
      el("import-preview").textContent = "";
      el("import-error").textContent = error.message;
    }
  });
  el("apply-import").addEventListener("click", () => {
    if (!pending) return;
    el("search").value = "";
    document.querySelector('[data-filter="all"]').click();
    save(pending,"imported","Excel 工作簿已导入");
    el("import-dialog").close();
    pending = null;
  });
  render();
})();
