(function (root) {
  "use strict";
  const DEFAULT_DATE = "2027-03-16";
  const LIMIT = 1000000000;
  const inventoryColumns = ["产品编号", "产品名称", "旧包材库存", "单价(元)", "印刷交期(天)"];
  const scheduleColumns = ["产品编号", "生产日期", "包材用量"];
  function date(value, label = "日期") {
    const s = String(value ?? "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s) || s < "1901-01-01" || s > "9999-12-31") throw Error(label + "须填写 YYYY-MM-DD 格式的有效日期");
    const d = new Date(s + "T00:00:00Z");
    if (!Number.isFinite(d.getTime()) || d.toISOString().slice(0, 10) !== s) throw Error(label + "无效");
    return s;
  }
  function number(value, label, max, places = 0, min = 0) {
    const s = String(value ?? "").trim();
    const n = Number(s);
    const pattern = places ? new RegExp("^\\d+(?:\\.\\d{1," + places + "})?$") : /^\d+$/;
    if (!pattern.test(s) || !Number.isFinite(n) || n < min || n > max) throw Error(label + "须在 " + min + " 至 " + max + " 之间" + (places ? "，最多 " + places + " 位小数" : "，且为整数"));
    return n;
  }
  function required(value, label) {
    const s = String(value ?? "").trim();
    if (!s || s.length > 100) throw Error(label + "须填写 1 至 100 个字符");
    return s;
  }
  function validate(input) {
    const cutoff = date(input.cutoff, "换版截止日");
    if (!Array.isArray(input.items) || input.items.length > 2000) throw Error("最多支持 2,000 款产品");
    const ids = new Set();
    const items = input.items.map((item, i) => {
      const where = "库存第 " + (i + 2) + " 行：";
      const id = required(item.id, where + "产品编号");
      if (ids.has(id)) throw Error(where + "产品编号重复：" + id);
      ids.add(id);
      return {id, name: required(item.name, where + "产品名称"), stock: number(item.stock, where + "库存", LIMIT),
        price: number(item.price, where + "单价", 10000, 4), lead: number(item.lead, where + "交期", 3650)};
    });
    if (!Array.isArray(input.batches) || input.batches.length > 50000) throw Error("最多支持 50,000 条排产");
    const batches = input.batches.map((batch, i) => {
      const where = "排期第 " + (i + 2) + " 行：";
      const id = required(batch.id, where + "产品编号");
      if (!ids.has(id)) throw Error(where + "找不到产品编号 " + id);
      return {id, day: date(batch.day, where + "生产日期"), qty: number(batch.qty, where + "用量", LIMIT, 0, 1)};
    });
    return {cutoff, items, batches};
  }
  function calculate(plan) {
    const byId = new Map(plan.items.map(item => [item.id, []]));
    for (const batch of plan.batches) byId.get(batch.id).push(batch);
    return plan.items.map(item => {
      const batches = byId.get(item.id).sort((a,b) => a.day.localeCompare(b.day));
      let left = item.stock, switchDay = null, firstNew = 0;
      for (const batch of batches) {
        const used = batch.day < plan.cutoff ? Math.min(left, batch.qty) : 0;
        left -= used;
        if (batch.qty > used) {
          if (!switchDay) switchDay = batch.day;
          if (batch.day === switchDay) firstNew += batch.qty - used;
        }
      }
      let orderDay = null;
      if (switchDay) {
        const d = new Date(switchDay + "T00:00:00Z");
        d.setUTCDate(d.getUTCDate() - item.lead);
        orderDay = d.toISOString().slice(0,10);
      }
      // 单价保留万分位。较大库存下，先用整数四舍五入到分以避免浮点乘法跨过舍入边界。
      const priceUnits = BigInt(Math.round(item.price * 10000));
      const cents = (BigInt(left) * priceUnits + 50n) / 100n;
      return {...item, left, used:item.stock-left, amount: Number(cents) / 100,
        switchDay, orderDay, firstNew, count:batches.length};
    });
  }
  function sample() {
    return {cutoff:DEFAULT_DATE,items:[
      {id:"PK-001",name:"燕麦饼干 · 200g",stock:30000,price:0.35,lead:21},
      {id:"PK-002",name:"原味苏打 · 160g",stock:18000,price:0.28,lead:14},
      {id:"PK-003",name:"黑芝麻糊 · 400g",stock:12500,price:0.62,lead:28},
      {id:"PK-004",name:"冻干草莓 · 50g",stock:9000,price:0.46,lead:18}],
      batches:[
        {id:"PK-001",day:"2027-02-18",qty:12000},{id:"PK-001",day:"2027-03-10",qty:10000},{id:"PK-001",day:"2027-03-18",qty:6000},
        {id:"PK-002",day:"2027-03-01",qty:15000},{id:"PK-002",day:"2027-03-12",qty:5000},
        {id:"PK-003",day:"2027-03-05",qty:9500},{id:"PK-003",day:"2027-03-16",qty:4000},
        {id:"PK-004",day:"2027-03-09",qty:9000},{id:"PK-004",day:"2027-03-24",qty:4000}]};
  }
  const api = {DEFAULT_DATE, inventoryColumns, scheduleColumns, validate, calculate, sample, date};
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Planner = api;
})(globalThis);
