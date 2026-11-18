(function (root) {
  "use strict";

  function element(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
  }

  function button(text, className, action) {
    const control = element("button", className, text);
    control.type = "button";
    control.addEventListener("click", action);
    return control;
  }

  function field(container, text, value, type, options = {}) {
    const label = element("label", "field");
    const input = element("input");
    input.type = type;
    input.required = true;
    Object.assign(input, options);
    input.value = value ?? "";
    label.append(element("span", "", text), input);
    container.append(label);
    return input;
  }

  function open(plan, originalId, onSave) {
    const editing = originalId != null;
    const product = editing ? plan.items.find(item => item.id === originalId) : {};
    if (!product) throw Error("找不到要编辑的产品");
    const originalBatches = editing ? plan.batches.filter(batch => batch.id === originalId) : [];

    let dialog = document.getElementById("product-editor");
    if (!dialog) {
      dialog = element("dialog");
      dialog.id = "product-editor";
      dialog.setAttribute("aria-labelledby", "product-editor-title");
      document.body.append(dialog);
    }
    if (dialog.open) dialog.close();
    const form = element("form");
    const header = element("header", "dialog-head");
    const title = element("h2", "", editing ? "编辑产品与排产" : "新增产品");
    title.id = "product-editor-title";
    const closeButton = button("×", "icon-button", () => dialog.close());
    closeButton.setAttribute("aria-label", "关闭编辑窗口");
    header.append(title, closeButton);

    const body = element("div", "dialog-body");
    const grid = element("div", "form-grid");
    const idInput = field(grid, "产品编号", product.id, "text", {maxLength: 100, placeholder: "如 PK-005"});
    const nameInput = field(grid, "产品名称", product.name, "text", {maxLength: 100});
    const stockInput = field(grid, "旧包材库存（件）", product.stock, "number", {min: 0, max: 1000000000, step: 1});
    const priceInput = field(grid, "包材单价（元）", product.price, "number", {min: 0, max: 10000, step: 0.0001});
    const leadInput = field(grid, "印刷交期（自然日）", product.lead, "number", {min: 0, max: 3650, step: 1});

    const batchesSection = element("section");
    batchesSection.append(element("h3", "", "生产排期"));
    const batchList = element("div");
    const emptyHint = element("p", "subtle", "还没有生产批次，可先保存库存。");
    let batchFields = [];
    function addBatch(batch = {}) {
      const row = element("div", "batch-row");
      const day = field(row, "生产日期", batch.day, "date", {min: "1901-01-01", max: "9999-12-31"});
      const qty = field(row, "包材用量（件）", batch.qty, "number", {min: 1, max: 1000000000, step: 1});
      const entry = {day, qty};
      const removeButton = button("移除", "button subtle", () => {
        const index = batchFields.indexOf(entry);
        batchFields.splice(index, 1);
        row.remove();
        emptyHint.hidden = batchFields.length > 0;
        (batchFields[index]?.day || batchFields[index - 1]?.day || addButton).focus();
      });
      removeButton.setAttribute("aria-label", "移除此排产批次");
      row.append(removeButton);
      batchFields.push(entry);
      batchList.append(row);
      emptyHint.hidden = true;
      return day;
    }
    const addButton = button("＋ 添加批次", "button", () => addBatch().focus());
    originalBatches.forEach(addBatch);
    batchesSection.append(batchList, emptyHint, addButton);
    body.append(grid, batchesSection);

    const error = element("p", "error");
    error.setAttribute("role", "alert");
    error.tabIndex = -1;
    error.hidden = true;
    body.append(error);

    function commit(items, batches) {
      try {
        const next = root.Planner.validate({cutoff: plan.cutoff, items, batches});
        onSave(next);
        dialog.close();
      } catch (problem) {
        error.textContent = problem.message;
        error.hidden = false;
        error.focus();
      }
    }
    form.addEventListener("submit", event => {
      event.preventDefault();
      const id = idInput.value.trim();
      if (plan.items.some(item => item.id === id && item.id !== originalId)) {
        idInput.setCustomValidity("这个产品编号已被使用，请更换编号。");
        idInput.reportValidity();
        return;
      }
      const item = {id, name: nameInput.value, stock: stockInput.value, price: priceInput.value, lead: leadInput.value};
      const items = editing ? plan.items.map(existing => existing.id === originalId ? item : existing) : [...plan.items, item];
      const batches = plan.batches.filter(batch => batch.id !== originalId);
      for (const entry of batchFields) batches.push({id, day: entry.day.value, qty: entry.qty.value});
      commit(items, batches);
    });
    idInput.addEventListener("input", () => idInput.setCustomValidity(""));
    const footer = element("footer", "dialog-foot");
    if (editing) {
      const confirmation = element("div");
      confirmation.hidden = true;
      const question = element("p", "", "确定删除“" + product.name + "”及其 " + originalBatches.length + " 条排产吗？");
      const confirmDelete = button("确认删除", "button danger", () => {
        commit(plan.items.filter(item => item.id !== originalId), plan.batches.filter(batch => batch.id !== originalId));
      });
      const abandonDelete = button("保留产品", "button", () => {
        confirmation.hidden = true;
        footer.hidden = false;
        deleteButton.focus();
      });
      confirmation.append(question, abandonDelete, confirmDelete);
      const deleteButton = button("删除产品", "button danger", () => {
        confirmation.hidden = false;
        footer.hidden = true;
        abandonDelete.focus();
      });
      body.append(confirmation);
      footer.append(deleteButton);
    }
    const save = element("button", "button primary", "保存并测算");
    save.type = "submit";
    footer.append(button("取消", "button", () => dialog.close()), save);
    form.append(header, body, footer);
    dialog.replaceChildren(form);
    dialog.showModal();
    (editing ? nameInput : idInput).focus();
  }

  root.Editor = {open};
})(globalThis);
