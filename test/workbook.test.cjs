const test = require("node:test");
const assert = require("node:assert/strict");
const xlsx = require("../dist/vendor/xlsx.mini.min.js");
const planner = require("../dist/planner.js");
const workbook = require("../dist/workbook.js");

function bytes(book) {
  return xlsx.write(book, {type:"buffer",bookType:"xlsx",compression:true});
}
function replaceSheet(book, name, rows) {
  book.Sheets[name] = xlsx.utils.aoa_to_sheet(rows);
}
test("template and report both import losslessly and preserve user text", () => {
  const plan = planner.sample();
  plan.items[0].name = "=测试公式";
  for (const book of [workbook.template(plan), workbook.exportResults(plan)]) {
    assert.deepEqual(workbook.read(bytes(book)), plan);
    assert.equal(book.Sheets["包材库存"].B2.t, "s");
  }
  const results = workbook.exportResults(plan).Sheets["换版测算"];
  assert.equal(results.E2.v, 8000);
  assert.equal(results.F2.v, 2800);
  assert.equal(results.I2.v, "2027-02-25");
});
test("real Excel dates import with either workbook epoch", () => {
  for (const date1904 of [false, true]) {
    const book = workbook.template();
    book.Workbook = { WBProps:{date1904} };
    const serial = (Date.UTC(2027, 2, 16) - Date.UTC(1899, 11, 30)) / 86400000 - (date1904 ? 1462 : 0);
    book.Sheets["生产排期"].B2 = {t:"n",v:serial,z:"yyyy-mm-dd"};
    book.Sheets["参数"].A2 = {t:"n",v:serial,z:"yyyy-mm-dd"};
    const actual = workbook.read(bytes(book));
    assert.equal(actual.cutoff, "2027-03-16");
    assert.equal(actual.batches[0].day, actual.cutoff);
  }
});
test("required headers can be reordered and missing optional sheet keeps current cutoff", () => {
  const book = workbook.template();
  replaceSheet(book, "生产排期", [["生产日期","包材用量","产品编号"],["2027/2/9",12,"PK-001"]]);
  book.SheetNames = book.SheetNames.filter(name => name !== "参数");
  delete book.Sheets["参数"];
  const result = workbook.read(bytes(book), "2027-01-01");
  assert.equal(result.cutoff, "2027-01-01");
  assert.deepEqual(result.batches, [{id:"PK-001",day:"2027-02-09",qty:12}]);
});
test("invalid imports report missing or duplicate column, dangling SKU or missing tab", () => {
  const book = workbook.template();
  replaceSheet(book, "生产排期", [["产品编号", "生产日期"]]);
  assert.throws(() => workbook.read(bytes(book)), /缺少列：包材用量/);
  replaceSheet(book, "生产排期", [["产品编号","生产日期","包材用量","包材用量"]]);
  assert.throws(() => workbook.read(bytes(book)), /表头重复：包材用量/);
  replaceSheet(book, "生产排期", [planner.scheduleColumns,[],["MISSING","2027-03-10",10]]);
  assert.throws(() => workbook.read(bytes(book)), /排期第 3 行：找不到产品编号 MISSING/);
  book.SheetNames = ["包材库存"];
  assert.throws(() => workbook.read(bytes(book)), /缺少工作表「生产排期」/);
});
test("formatted numeric product IDs preserve leading zeroes", () => {
  const plan = {cutoff:"2027-03-16",items:[{id:"0001",name:"样品",stock:3,price:1,lead:0}],batches:[{id:"0001",day:"2027-02-02",qty:1}]};
  const book = workbook.template(plan);
  book.Sheets["包材库存"].A2 = {t:"n", v:1, z:"0000"};
  assert.deepEqual(workbook.read(bytes(book)), plan);
});
