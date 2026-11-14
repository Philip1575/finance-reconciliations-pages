"use strict";

// Rolls the daily Fed Rec forward one business day, the way the preparer does it by hand: open yesterday's workbook,
// save it as today, add today's rows below yesterday's (formulas dragged down, labels copied, amounts left for today)
// and set the date. The existing daily fill then completes today's rows.

const fedRollCategories = ["ach in", "ach out", "wire out", "wire in", "fednow in", "fednow out"];

// Shifts the relative cell references in a formula by rows and columns ($-fixed parts stay put).
function fedShiftFormula(text, rows, columns = 0) {
  return text.replace(/("[^"]*")|(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])/g, (all, quoted, colFixed, column, rowFixed, row) => {
    if (quoted) return quoted;
    const newColumn = colFixed || !columns ? column : csvColumn(columnNumber(column) + columns);
    const newRow = rowFixed ? row : String(Number(row) + rows);
    return colFixed + newColumn + rowFixed + newRow;
  });
}

const fedSerial = iso => (Date.parse(iso + "T00:00:00Z") - Date.UTC(1899, 11, 30)) / 86400000;

function fedNextBusinessDay(iso) {
  let day = new Date(iso + "T00:00:00Z");
  do day = new Date(day.getTime() + 86400000); while ([0, 6].includes(day.getUTCDay()));
  return day.toISOString().slice(0, 10);
}

// A worksheet opened for row-level edits: cells by row, shared formulas resolved to their own text.
async function fedSheetEditor(workbook, sheet) {
  const path = excelWorksheetPath(workbook, sheet);
  const doc = await readXml(workbook.zip, path);
  const sheetData = doc.getElementsByTagNameNS(spreadsheetNs, "sheetData")[0];
  const rows = new Map(Array.from(sheetData.children).filter(node => node.localName === "row").map(node => [Number(node.getAttribute("r")), node]));
  const strings = workbook.sharedStrings ?? [];
  const masters = new Map();
  for (const f of Array.from(doc.getElementsByTagNameNS(spreadsheetNs, "f"))) {
    if (f.getAttribute("t") === "shared" && f.textContent) masters.set(f.getAttribute("si"), { address: f.parentNode.getAttribute("r"), text: f.textContent });
  }
  const parse = address => { const [, column, row] = /^([A-Z]+)(\d+)$/.exec(address); return { column, row: Number(row) }; };
  const cellsOf = row => Array.from(rows.get(row)?.children ?? []).filter(node => node.localName === "c");
  const cell = address => cellsOf(parse(address).row).find(node => node.getAttribute("r") === address);
  const formulaOf = node => {
    const f = Array.from(node.children).find(child => child.localName === "f");
    if (!f) return undefined;
    if (f.textContent) return f.textContent;
    const master = masters.get(f.getAttribute("si"));
    if (!master) return undefined;
    const from = parse(master.address), to = parse(node.getAttribute("r"));
    return fedShiftFormula(master.text, to.row - from.row, columnNumber(to.column) - columnNumber(from.column));
  };
  const valueOf = node => {
    if (!node) return undefined;
    const v = Array.from(node.children).find(child => child.localName === "v");
    const type = node.getAttribute("t");
    if (type === "s") return strings[Number(v?.textContent)];
    if (type === "inlineStr") return Array.from(node.getElementsByTagNameNS(spreadsheetNs, "t")).map(t => t.textContent).join("");
    if (type === "str") return v?.textContent;
    return v ? Number(v.textContent) : undefined;
  };
  // Shared formulas become each cell's own formula, so rows can be moved or overwritten without breaking a shared group.
  for (const f of Array.from(doc.getElementsByTagNameNS(spreadsheetNs, "f")).filter(node => node.getAttribute("t") === "shared")) {
    const text = formulaOf(f.parentNode);
    if (text === undefined) continue;
    f.textContent = text;
    for (const name of ["t", "si", "ref"]) f.removeAttribute(name);
  }
  const element = (parent, name) => doc.createElementNS(spreadsheetNs, parent.prefix ? parent.prefix + ":" + name : name);
  const rowNode = row => {
    if (rows.has(row)) return rows.get(row);
    const node = element(sheetData, "row");
    node.setAttribute("r", String(row));
    sheetData.insertBefore(node, [...rows].filter(([r]) => r > row).sort((a, b) => a[0] - b[0])[0]?.[1] ?? null);
    rows.set(row, node);
    return node;
  };
  const put = (address, build) => {
    const { column, row } = parse(address);
    const parent = rowNode(row);
    const existing = cell(address);
    const node = element(parent, "c");
    node.setAttribute("r", address);
    build(node);
    if (existing) parent.replaceChild(node, existing);
    else parent.insertBefore(node, cellsOf(row).find(other => columnNumber(parse(other.getAttribute("r")).column) > columnNumber(column)) ?? null);
  };
  const setNumber = (address, value, style) => put(address, node => {
    if (style) node.setAttribute("s", style);
    const v = element(node, "v");
    v.textContent = String(value);
    node.append(v);
  });
  const setFormula = (address, text, style) => put(address, node => {
    if (style) node.setAttribute("s", style);
    const f = element(node, "f");
    f.textContent = text;
    node.append(f);
  });
  const setText = (address, text, style) => put(address, node => {
    if (style) node.setAttribute("s", style);
    node.setAttribute("t", "inlineStr");
    const is = element(node, "is"), t = element(is, "t");
    t.textContent = text;
    is.append(t);
    node.append(is);
  });
  const setBlank = (address, style) => put(address, node => { if (style) node.setAttribute("s", style); });
  // Copies a row the way dragging it down does: labels as they are, formulas shifted, amounts left blank, dates set.
  // Cells already in the destination row stay (merge), except formula cells, which are refreshed.
  const copyRow = (from, to, dates = {}, merge = false, zero = false) => {
    const source = rows.get(from);
    if (!source) return;
    const target = rowNode(to);
    for (const name of ["ht", "customHeight", "s", "customFormat"]) if (source.hasAttribute(name) && !target.hasAttribute(name)) target.setAttribute(name, source.getAttribute(name));
    for (const node of cellsOf(from)) {
      const { column } = parse(node.getAttribute("r"));
      const address = column + to, style = node.getAttribute("s");
      const current = cell(address);
      if (column in dates) {
        if (dates[column] !== undefined && fedIsDate(valueOf(node) ?? (merge ? valueOf(current) : undefined)) || merge) setNumber(address, dates[column], style);
        else if (!merge) setBlank(address, style);
        continue;
      }
      const formula = formulaOf(node);
      if (formula !== undefined) { setFormula(address, fedShiftFormula(formula, to - from), style); continue; }
      if (merge && current && valueOf(current) !== undefined) continue;
      const value = valueOf(node);
      if (typeof value === "string") setText(address, value, style);
      else if (zero && typeof value === "number" && !fedIsDate(value) && fedZeroSide(column, dates)) setNumber(address, 0, style);
      else setBlank(address, style);
    }
    // A date column the source row leaves empty still gets today's date, formatted like the date above it.
    for (const [column, value] of Object.entries(dates)) {
      if (value === undefined || cellsOf(from).some(node => parse(node.getAttribute("r")).column === column)) continue;
      const above = [from - 1, from - 2, from - 3].map(row => cell(column + row)).find(node => node && fedIsDate(valueOf(node)));
      setNumber(column + to, value, above?.getAttribute("s"));
    }
  };
  // Moves a whole row (values, labels, dates and shifted formulas) to another row.
  const moveRow = (from, to) => {
    if (from === to) return;
    for (const node of cellsOf(from)) {
      const { column } = parse(node.getAttribute("r")), address = column + to, style = node.getAttribute("s");
      const formula = formulaOf(node), value = valueOf(node);
      if (formula !== undefined) setFormula(address, fedShiftFormula(formula, to - from), style);
      else if (typeof value === "string") setText(address, value, style);
      else if (typeof value === "number") setNumber(address, value, style);
      else setBlank(address, style);
    }
  };
  // Empties a row, keeping its formats and (optionally) its formulas.
  const clearRow = (row, keepFormulas) => {
    for (const node of cellsOf(row)) {
      const address = node.getAttribute("r"), style = node.getAttribute("s"), formula = formulaOf(node);
      if (keepFormulas && formula !== undefined) setFormula(address, formula, style);
      else setBlank(address, style);
    }
  };
  const save = () => workbook.zip.file(path, new XMLSerializer().serializeToString(doc));
  return { rows, cellsOf, cell, valueOf, formulaOf, copyRow, moveRow, clearRow, setNumber, setText, setFormula, save, parse };
}

// In VIZO MATCHUP a row's amounts default to 0 only on the side (VIZO A-D, Wisdom E-H) that gets today's date.
const fedZeroSide = (column, dates) => columnNumber(column) < columnNumber("E") ? "A" in dates : columnNumber(column) < columnNumber("K") ? "E" in dates : "K" in dates;

const fedMonthOf = serial => new Date(Date.UTC(1899, 11, 30) + serial * 86400000).toISOString().slice(0, 7);

const fedIsDate = value => typeof value === "number" && Number.isInteger(value) && value > 40000 && value < 60000;

// Adds today's MATCHUP rows: an ACH In / ACH Out pair and one row each for wires and FedNow, below yesterday's.
function fedRollMatchup(sheet, date) {
  const serial = fedSerial(date);
  const label = row => String(sheet.valueOf(sheet.cell("B" + row)) ?? "").trim().toLowerCase();
  const lastRow = new Map();
  for (const [row] of [...sheet.rows].sort((a, b) => a[0] - b[0])) {
    const name = label(row);
    if (fedRollCategories.includes(name) && fedIsDate(sheet.valueOf(sheet.cell("A" + row)))) lastRow.set(name, row);
  }
  requireReport(lastRow.has("ach out") && lastRow.has("wire out") && lastRow.has("wire in"), "Yesterday's MATCHUP sheet has no dated ACH and wire rows.");
  const last = sheet.valueOf(sheet.cell("A" + lastRow.get("ach out")));
  requireReport(serial > last, "Yesterday's workbook already has rows for " + date + " or later.");
  // First business day of a new month: each block keeps only last month's final day, moved to the top of the block.
  if (fedMonthOf(last) !== date.slice(0, 7)) {
    const firstRow = new Map();
    for (const [row] of [...sheet.rows].sort((a, b) => a[0] - b[0])) {
      const name = label(row);
      if (fedRollCategories.includes(name) && fedIsDate(sheet.valueOf(sheet.cell("A" + row))) && !firstRow.has(name)) firstRow.set(name, row);
    }
    const achTop = firstRow.get("ach in"), achOut = lastRow.get("ach out");
    sheet.moveRow(achOut - 1, achTop);
    sheet.moveRow(achOut, achTop + 1);
    for (let row = achTop + 2; row <= achOut; row++) sheet.clearRow(row, false);
    lastRow.set("ach in", achTop).set("ach out", achTop + 1);
    for (const name of ["wire out", "wire in", "fednow in", "fednow out"]) {
      if (!lastRow.has(name)) continue;
      const top = firstRow.get(name), bottom = lastRow.get(name);
      sheet.moveRow(bottom, top);
      for (let row = top + 1; row <= bottom; row++) sheet.clearRow(row, true);
      lastRow.set(name, top);
    }
  }
  const free = row => sheet.cellsOf(row).every(node => {
    const value = sheet.valueOf(node);
    return value === undefined || value === "" || sheet.formulaOf(node) !== undefined;
  });
  const plans = [];
  const achOut = lastRow.get("ach out");
  plans.push([achOut - 1, achOut + 1], [achOut, achOut + 2]);
  // Each block gets today's row unless it already has one.
  for (const name of ["wire out", "wire in", "fednow in", "fednow out"]) {
    if (lastRow.has(name) && sheet.valueOf(sheet.cell("A" + lastRow.get(name))) < serial) plans.push([lastRow.get(name), lastRow.get(name) + 1]);
  }
  for (const [, to] of plans) requireReport(free(to), "MATCHUP has no free row at " + to + " for today's entries. Add rows below the last day, then try again.");
  for (const [from, to] of plans) sheet.copyRow(from, to, { A: serial, E: serial });
  return plans.map(([, to]) => to);
}

// Adds today's VIZO MATCHUP rows: in each block, each date column gets today's date in the row after its last date (the
// mobile block's Wisdom dates run a day behind, so its two date columns land on different rows).
function fedRollVizo(sheet, date) {
  const serial = fedSerial(date);
  const sorted = [...sheet.rows.keys()].sort((a, b) => a - b);
  const heads = sorted.filter(row => /^date$/i.test(String(sheet.valueOf(sheet.cell("A" + row)) ?? "").trim()));
  const added = [];
  heads.forEach((head, index) => {
    const end = heads[index + 1] ?? Infinity;
    const block = sorted.filter(row => row > head && row < end);
    // Today's date goes on one new row, in every date column the block uses. Blocks whose Wisdom dates ran a row behind
    // (mobile) line up again: Wisdom deposits are compared with the next day's VIZO row, the same as branch deposits.
    const targets = new Map();
    const columns = ["A", "E", "K"].filter(column => block.some(row => fedIsDate(sheet.valueOf(sheet.cell(column + row)))));
    const lastDated = Math.max(0, ...columns.flatMap(column => block.filter(row => fedIsDate(sheet.valueOf(sheet.cell(column + row))))));
    const latest = Math.max(0, ...columns.flatMap(column => block.map(row => sheet.valueOf(sheet.cell(column + row))).filter(fedIsDate)));
    let lastRowInBlock = lastDated;
    if (columns.length && latest < serial && fedMonthOf(latest) !== date.slice(0, 7)) {
      // New month: keep the rows of last month's final day at the top of the block.
      const keep = block.filter(row => columns.some(column => sheet.valueOf(sheet.cell(column + row)) === latest));
      const top = block[0], from = Math.min(...keep);
      for (let row = from; row <= lastDated; row++) sheet.moveRow(row, top + row - from);
      for (let row = top + lastDated - from + 1; row <= lastDated; row++) sheet.clearRow(row, true);
      lastRowInBlock = top + lastDated - from;
    }
    if (columns.length && latest < serial) {
      const to = lastRowInBlock + 1;
      requireReport(to < end && !/^date$/i.test(String(sheet.valueOf(sheet.cell("A" + to)) ?? "")), "VIZO MATCHUP has no free row at " + to + ". Add rows to that block, then try again.");
      targets.set(to, Object.fromEntries(columns.map(column => [column, serial])));
    }
    for (const [to, dates] of [...targets].sort((a, b) => a[0] - b[0])) {
      const all = { A: undefined, E: undefined, K: undefined, ...dates };
      for (const column of ["A", "E", "K"]) if (!(column in dates)) delete all[column];
      // Date columns not being set keep what the row already has.
      sheet.copyRow(to - 1, to, Object.fromEntries(Object.entries(dates)), true, true);
      added.push(to);
    }
  });
  return added;
}

// Today's workbook from yesterday's: the new date in MATCHUP!D6 and today's rows added. Returns the original file when it
// is already today's workbook.
async function fedRollForward(file, date) {
  const workbook = await openExcel(file, true);
  const strings = await readXml(workbook.zip, "xl/sharedStrings.xml").catch(() => undefined);
  workbook.sharedStrings = strings ? Array.from(strings.getElementsByTagNameNS(spreadsheetNs, "si"),
    si => Array.from(si.getElementsByTagNameNS(spreadsheetNs, "t")).map(t => t.textContent).join("")) : [];
  const matchup = await fedSheetEditor(workbook, "MATCHUP");
  const current = matchup.valueOf(matchup.cell("D6"));
  if (fedIsDate(current) && current === fedSerial(date)) return { file, rolled: false };
  const achRows = [...matchup.rows.keys()].filter(row => String(matchup.valueOf(matchup.cell("B" + row)) ?? "").trim().toLowerCase() === "ach out"
    && matchup.valueOf(matchup.cell("A" + row)) === fedSerial(date));
  if (achRows.length) return { file, rolled: false };
  const rows = fedRollMatchup(matchup, date);
  matchup.setNumber("D6", fedSerial(date), matchup.cell("D6")?.getAttribute("s"));
  matchup.save();
  let vizoRows = [];
  if (workbook.sheets.has("VIZO MATCHUP")) {
    const vizo = await fedSheetEditor(workbook, "VIZO MATCHUP");
    vizoRows = fedRollVizo(vizo, date);
    vizo.save();
  }
  workbook.zip.remove("xl/calcChain.xml");
  const doc = await readXml(workbook.zip, "xl/workbook.xml");
  let calc = doc.getElementsByTagNameNS(spreadsheetNs, "calcPr")[0];
  if (!calc) {
    calc = doc.createElementNS(spreadsheetNs, doc.documentElement.prefix ? doc.documentElement.prefix + ":calcPr" : "calcPr");
    doc.documentElement.append(calc);
  }
  calc.setAttribute("fullCalcOnLoad", "1");
  workbook.zip.file("xl/workbook.xml", new XMLSerializer().serializeToString(doc));
  let types = await workbook.zip.file("[Content_Types].xml").async("string");
  types = types.replace(/<Override[^>]*calcChain[^>]*\/>/, "");
  workbook.zip.file("[Content_Types].xml", types);
  const blob = await workbook.zip.generateAsync({ type: "blob", compression: "DEFLATE" });
  const name = file.name.replace(/\d{6}/, date.slice(5, 7) + date.slice(8, 10) + date.slice(2, 4));
  return { file: new File([blob], name === file.name ? date + " " + file.name : name), rolled: true, rows, vizoRows };
}

// The day's items for the cover, as the preparer adds them: a dated "ACH is off" (or wire / FedNow) line for any
// difference left in MATCHUP, and Fed-only statement items added to their existing line with the day's amount appended
// to the formula so the month's breakdown stays visible.
function fedCoverItems(result) {
  const [month, day] = [Number(result.date.slice(5, 7)), Number(result.date.slice(8, 10))];
  const items = [];
  if (result.adjustedAch) items.push({ label: "ACH is off " + month + "/" + day, amount: result.adjustedAch, section: /^ach is off/i, review: true });
  for (const row of result.comparisons ?? []) {
    if (row.difference) items.push({ label: row.category + " is off " + month + "/" + day, amount: row.difference, section: /^ach is off/i, review: true });
  }
  // LendKey and Student Choice: the day's debits are appended to the "debits" line as -amount and credits to the
  // "credits" line as +amount, until the month-end LendKey and Student Choice recs clear them.
  const lines = { LendKey: /^(lk|lend\s*key)\s+/i, "Student Choice": /^(sc|student choice)\s+/i };
  for (const [provider, prefix] of Object.entries(lines)) {
    for (const direction of ["debit", "credit"]) {
      const amount = (result.finance?.included ?? []).filter(item => item.finance === provider && (item.amount < 0) === (direction === "debit"))
        .reduce((sum, item) => sum + item.amount, 0);
      if (amount) items.push({ label: provider + " " + direction + "s", amount, match: new RegExp(prefix.source + direction + "s?\\s*$", "i"), review: false });
    }
  }
  for (const row of (result.activity ?? []).filter(item => item.source === "statement" && item.category === "Other · review" && item.amount)) {
    const words = row.description.toUpperCase().match(/[A-Z]{3,}/g) ?? [];
    items.push({ label: row.description.trim() + " " + month + "/" + day, amount: row.amount, words, review: false });
  }
  return items;
}

async function fedApplyCoverItems(blob, name, items) {
  if (!items.length) return { blob, placed: [] };
  const workbook = await openExcel(new File([blob], name), true);
  const strings = await readXml(workbook.zip, "xl/sharedStrings.xml").catch(() => undefined);
  workbook.sharedStrings = strings ? Array.from(strings.getElementsByTagNameNS(spreadsheetNs, "si"),
    si => Array.from(si.getElementsByTagNameNS(spreadsheetNs, "t")).map(t => t.textContent).join("")) : [];
  const cover = await fedSheetEditor(workbook, "Cover");
  const text = row => String(cover.valueOf(cover.cell("E" + row)) ?? "").trim();
  // Line items sit between the Wisdom balance and the cover-plug / total rows.
  const total = [...cover.rows.keys()].find(row => /^adjusted book balance$/i.test(String(cover.valueOf(cover.cell("C" + row)) ?? "").trim()));
  requireReport(total, "The Cover has no Adjusted Book Balance row.");
  const lines = [...Array(Math.max(0, total - 16)).keys()].map(k => 14 + k).filter(row => row < total - 3);
  const empty = row => !text(row) && cover.valueOf(cover.cell("K" + row)) === undefined && cover.formulaOf(cover.cell("K" + row) ?? { children: [] }) === undefined;
  const placed = [];
  const fmt = amount => (amount / 100).toFixed(2);
  for (const item of items) {
    let row;
    if (item.match) {
      row = lines.find(r => item.match.test(text(r)));
      requireReport(row, "The Cover has no " + item.label + " line.");
    } else if (item.words?.length) {
      row = lines.find(r => { const label = text(r).toUpperCase(); return label && item.words.filter(word => label.includes(word)).length >= Math.min(2, item.words.length); });
    }
    if (row) {
      // Append today's amount to the existing line's formula.
      const node = cover.cell("K" + row), formula = node && cover.formulaOf(node), value = node && cover.valueOf(node);
      const base = formula ?? (typeof value === "number" ? String(value) : "");
      const expression = (base ? base : "") + (item.amount < 0 ? "-" : base ? "+" : "") + fmt(Math.abs(item.amount));
      cover.setFormula("K" + row, expression.replace(/^\+/, ""), node?.getAttribute("s"));
    } else {
      const anchor = item.section ? lines.filter(r => item.section.test(text(r))).at(-1) : undefined;
      row = lines.find(r => r > (anchor ?? lines[0]) && empty(r));
      requireReport(row, "The Cover has no empty line-item row for " + item.label + ". Add a row, then try again.");
      const labelStyle = cover.cell("E" + (anchor ?? row))?.getAttribute("s"), amountStyle = cover.cell("K" + (anchor ?? row))?.getAttribute("s");
      cover.setText("E" + row, item.label, labelStyle);
      cover.setNumber("K" + row, fmt(item.amount), amountStyle);
    }
    placed.push({ ...item, row });
  }
  cover.save();
  return { blob: await workbook.zip.generateAsync({ type: "blob", compression: "DEFLATE" }), placed };
}
