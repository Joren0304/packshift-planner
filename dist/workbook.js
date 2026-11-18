(function (root) {
  "use strict";
  const inNode = typeof module !== "undefined" && module.exports;
  const xlsx = inNode ? require("./vendor/xlsx.mini.min.js") : root.XLSX;
  const planner = inNode ? require("./planner.js") : root.Planner;
  const PARAMETER = "换版截止日";
  const MAX_ROWS = 50001;

  function append(book, name, headers, rows, widths) {
    const sheet = xlsx.utils.aoa_to_sheet([headers, ...rows]);
    sheet["!cols"] = widths.map(wch => ({wch}));
    sheet["!autofilter"] = {ref: sheet["!ref"]};
    xlsx.utils.book_append_sheet(book, sheet, name);
  }

  function appendInputs(book, plan) {
    append(book, "包材库存", planner.inventoryColumns,
      plan.items.map(p => [p.id, p.name, p.stock, p.price, p.lead]), [21, 30, 18, 15, 18]);
    append(book, "生产排期", planner.scheduleColumns,
      plan.batches.map(p => [p.id, p.day, p.qty]), [21, 18, 18]);
    append(book, "参数", [PARAMETER], [[plan.cutoff]], [20]);
  }

  function download(book, filename) {
    if (!inNode) xlsx.writeFile(book, filename, {bookType:"xlsx", compression:true});
    return book;
  }

  function template(plan = planner.sample()) {
    const book = xlsx.utils.book_new();
    appendInputs(book, planner.validate(plan));
    return download(book, "换版台-排产模板.xlsx");
  }

  function exportResults(input) {
    const plan = planner.validate(input);
    const book = xlsx.utils.book_new();
    const results = planner.calculate(plan);
    append(book, "换版测算", ["产品编号", "产品名称", "旧包材库存", "预计使用旧包材", "预计剩余旧包材",
      "积压金额(元)", "首次使用新包材日期", "当日新包材用量", "最迟下单日"],
      results.map(p => [p.id, p.name, p.stock, p.used, p.left, p.amount, p.switchDay, p.firstNew, p.orderDay]),
      [21, 30, 18, 20, 20, 18, 24, 20, 18]);
    appendInputs(book, plan);
    return download(book, "换版台-测算结果-" + plan.cutoff + ".xlsx");
  }

  function valueOf(cell, where) {
    if (cell?.t === "e") throw Error(where + "含有 Excel 错误值，请检查该单元格");
    if (cell?.f && cell.v == null) throw Error(where + "的公式缺少计算结果，请先在 Excel 中计算并保存");
    return cell?.v;
  }

  function excelDate(value, date1904, where) {
    if (typeof value === "number") {
      const parts = xlsx.SSF.parse_date_code(value, {date1904});
      if (!parts) throw Error(where + "的 Excel 日期无效");
      return planner.date([String(parts.y).padStart(4, "0"), String(parts.m).padStart(2, "0"), String(parts.d).padStart(2, "0")].join("-"), where);
    }
    const s = String(value ?? "").trim();
    const match = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    return planner.date(match ? [match[1], match[2].padStart(2, "0"), match[3].padStart(2, "0")].join("-") : s, where);
  }

  function records(book, name, headers) {
    const sheet = book.Sheets[name];
    if (!sheet) throw Error("缺少工作表「" + name + "」，请使用 Excel 模板");
    if (!sheet["!ref"]) throw Error("工作表「" + name + "」缺少表头");
    const range = xlsx.utils.decode_range(sheet["!fullref"] || sheet["!ref"]);
    if (range.e.r >= MAX_ROWS || range.e.c >= 128) {
      throw Error("工作表「" + name + "」超出 50,001 行或 128 列，请删除多余行列");
    }
    const positions = new Map();
    for (let c = range.s.c; c <= range.e.c; c++) {
      const key = xlsx.utils.encode_col(c);
      const header = String(valueOf(sheet[key + "1"], name + "表头") ?? "").trim();
      if (!header) continue;
      if (positions.has(header)) throw Error("工作表「" + name + "」表头重复：" + header);
      positions.set(header, key);
    }
    const missing = headers.filter(header => !positions.has(header));
    if (missing.length) throw Error("工作表「" + name + "」缺少列：" + missing.join("、"));
    const data = [];
    for (let line = 2; line <= range.e.r + 1; line++) {
      const values = headers.map(header => {
        const cell = sheet[positions.get(header) + line];
        const value = valueOf(cell, name + "第 " + line + " 行" + header);
        if (header === "产品编号" && cell?.t === "n" && /^0+$/.test(cell.z || "")) return cell.w || String(value);
        return value;
      });
      if (values.some(value => value != null && String(value).trim() !== "")) data.push({line, values});
    }
    return data;
  }

  function read(buffer, fallbackCutoff = planner.DEFAULT_DATE) {
    if (!buffer || buffer.byteLength > 10 * 1024 * 1024) throw Error("请选择 10 MB 以内的 Excel 文件");
    let book;
    try {
      book = xlsx.read(buffer, {type:"array", cellDates:false, cellNF:true, sheets:["包材库存", "生产排期", "参数"], sheetRows:MAX_ROWS + 1});
    } catch {
      throw Error("无法读取 Excel 工作簿，请用 Excel 或 WPS 另存为未加密的 .xlsx 文件");
    }
    const date1904 = !!book.Workbook?.WBProps?.date1904;
    const stock = records(book, "包材库存", planner.inventoryColumns);
    const schedule = records(book, "生产排期", planner.scheduleColumns);
    let cutoff = fallbackCutoff;
    if (book.Sheets["参数"]) {
      const parameters = records(book, "参数", [PARAMETER]);
      if (parameters.length !== 1) throw Error("参数表须填写一个换版截止日");
      cutoff = excelDate(parameters[0].values[0], date1904, PARAMETER);
    }
    const items = stock.map(({values:[id,name,stock,price,lead]}) => ({id, name, stock, price, lead}));
    const batches = schedule.map(({line,values:[id,day,qty]}) => ({id, day:excelDate(day, date1904, "生产排期第 " + line + " 行生产日期"), qty}));
    try {
      return planner.validate({cutoff, items, batches});
    } catch (error) {
      const message = error.message.replace(/^(库存|排期)第 (\d+) 行/, (_, type, number) => {
        const record = (type === "库存" ? stock : schedule)[Number(number) - 2];
        return type + "第 " + (record?.line ?? number) + " 行";
      });
      throw Error(message);
    }
  }

  const api = {template, exportResults, read};
  if (inNode) module.exports = api;
  else root.Workbook = api;
})(globalThis);
