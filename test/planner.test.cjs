const test = require("node:test");
const assert = require("node:assert/strict");
const { validate, calculate, sample, date, DEFAULT_DATE } = require("../dist/planner.js");

const item = { id: "A", name: "样品", stock: 10, price: 0.35, lead: 14 };
const batch = (day, qty, id = "A") => ({ id, day, qty });
const plan = (batches, items = [item], cutoff = DEFAULT_DATE) => validate({ cutoff, items, batches });

test("示例中的旧袋耗用、余量、金额与下单日", () => {
  const rows = calculate(validate(sample()));
  assert.equal(rows.length, 4);
  const [oat] = rows;
  assert.deepEqual(
    [oat.used, oat.left, oat.amount, oat.switchDay, oat.orderDay, oat.firstNew],
    [22000, 8000, 2800, "2027-03-18", "2027-02-25", 6000]
  );
  assert.equal(rows.reduce((sum, row) => sum + row.left, 0), 11000);
  assert.equal(rows.reduce((sum, row) => sum + row.amount, 0), 4660);
});

test("按日期排序、同日合并首次新包材需求，并保留原排产顺序", () => {
  const input = plan([
    batch("2027-03-10", 4), batch("2027-03-05", 8),
    batch("2027-03-10", 6), batch("2027-03-11", 12)
  ]);
  const [row] = calculate(input);
  assert.equal(input.batches[0].day, "2027-03-10");
  assert.deepEqual(
    [row.used, row.left, row.switchDay, row.orderDay, row.firstNew],
    [10, 0, "2027-03-10", "2027-02-24", 8]
  );
});

test("截止当天必须使用新版，延后截止日期可重新试算", () => {
  const batches = [batch("2027-03-15", 3), batch("2027-03-16", 6), batch("2027-03-18", 4)];
  const [early] = calculate(plan(batches));
  const [late] = calculate(plan(batches, [item], "2027-03-17"));
  assert.deepEqual([early.left, early.firstNew, early.switchDay], [7, 6, "2027-03-16"]);
  assert.deepEqual([late.left, late.firstNew, late.switchDay], [1, 4, "2027-03-18"]);
});

test("无排产或库存恰好耗完时，等待后续排产确定首次新版日期", () => {
  const data = plan([batch("2027-03-15", 10)], [item, { ...item, id: "B" }]);
  const [allUsed, noBatches] = calculate(data);
  assert.deepEqual([allUsed.left, allUsed.switchDay, allUsed.orderDay, allUsed.firstNew], [0, null, null, 0]);
  assert.deepEqual([noBatches.left, noBatches.count, noBatches.switchDay], [10, 0, null]);
});

test("按产品分配库存，正确跨越闰日和零交期", () => {
  const rows = calculate(plan(
    [batch("2024-03-01", 2), batch("2024-03-01", 6, "B")],
    [{ ...item, stock: 0, lead: 1 }, { ...item, id: "B", stock: 0, lead: 0 }]
  ));
  assert.deepEqual(rows.map(row => [row.orderDay, row.firstNew]),
    [["2024-02-29", 2], ["2024-03-01", 6]]);
});

test("逐款按分四舍五入，大数量和四位小数单价仍准确", () => {
  const rows = calculate(plan([], [
    { ...item, stock: 1, price: 0.005 },
    { ...item, id: "B", stock: 1, price: 0.0049 },
    { ...item, id: "C", stock: 999999999, price: 9999.9959 }
  ]));
  assert.deepEqual(rows.map(row => row.amount), [0.01, 0, 9999995890000]);
});

test("拒绝无效日期、重复产品、跨表缺失产品和无效数量", () => {
  assert.equal(date("2024-02-29"), "2024-02-29");
  assert.throws(() => date("2027-02-29"), /日期无效/);
  assert.throws(() => date("2027-3-16"), /YYYY-MM-DD/);
  assert.throws(() => plan([], [item, { ...item, id: " A " }]), /编号重复/);
  assert.throws(() => plan([batch("2027-02-01", 2, "B")]), /排期第 2 行.*找不到产品/);
  for (const qty of [0, -1, 1.2, ""]) assert.throws(() => plan([batch("2027-02-01", qty)]), /用量/);
  assert.throws(() => plan([], [{ ...item, price: 0.00001 }]), /单价/);
});
