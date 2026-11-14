"use strict";

// Cashier's checks: last month's reconciliation plus the 903 draft history become this month's reconciliation. The
// month's 903 activity is matched (checks and journal entries written against the drafts and journals that cleared
// them), cleared items come off the outstanding list, new ones go on, and the cover gets the month and the 903 balance.

const ckField = id => document.getElementById(id);
const ckFiles = {};
let ckResult, ckRevision = 0;
const ckMoney = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const ckMonthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function ckStatus(message, error = false) {
  ckField("status").textContent = message;
  ckField("status").classList.toggle("error", error);
}

const ckSerial = iso => (Date.parse(iso + "T00:00:00Z") - Date.UTC(1899, 11, 30)) / 86400000;
const ckText = value => String(value ?? "").replace(/\s+/g, " ").trim();
function ckDate(value, date1904) {
  if (value === undefined || value === null || value === "") return undefined;
  try { return sourceDate(typeof value === "number" ? Math.floor(value) : value, date1904); } catch { return undefined; }
}
const ckRows = cells => [...new Set([...cells.keys()].map(address => Number(/\d+$/.exec(address)[0])))].sort((a, b) => a - b);

// The Portico draft history: post date, time, transaction code, reversal flag, amount, balance after and draft number.
async function ckReadHistory(file) {
  let cells, date1904 = false;
  if (/\.csv$/i.test(file.name)) cells = readCsvCells(await file.text());
  else {
    const book = await openExcel(file, true);
    date1904 = book.date1904;
    cells = await readCells(book, [...book.sheets.keys()][0]);
  }
  const rows = ckRows(cells);
  const header = rows.find(row => row <= 10 && [...cells].some(([address, value]) => Number(/\d+$/.exec(address)[0]) === row && /^post date$/i.test(ckText(value))));
  requireReport(header, "The draft history needs a Post Date heading.");
  const columns = new Map([...cells].filter(([address]) => Number(/\d+$/.exec(address)[0]) === header)
    .map(([address, value]) => [ckText(value).toLowerCase(), address.replace(/\d+$/, "")]));
  const column = (...names) => names.map(name => columns.get(name)).find(Boolean);
  const at = { date: column("post date"), time: column("time"), code: column("txn"), reversal: column("r"), journal: column("can"),
    amount: column("transaction amount", "amount"), balance: column("after balance", "balance"), draft: column("draft number"),
    description: column("history code description"), statement: column("statement description") };
  for (const name of ["date", "code", "amount", "balance"]) requireReport(at[name], "The draft history needs a " + name + " column.");
  const items = [];
  for (const row of rows.filter(row => row > header)) {
    const date = ckDate(cells.get(at.date + row), date1904);
    if (!date) continue;
    items.push({ row, date, time: ckText(cells.get(at.time + row)), code: ckText(cells.get(at.code + row)).toUpperCase(),
      reversal: /^y/i.test(ckText(cells.get(at.reversal + row))), journal: Number(cells.get(at.journal + row)) || 0,
      amount: cents(cells.get(at.amount + row)), balance: cents(cells.get(at.balance + row)), draft: cells.get(at.draft + row) ?? "",
      description: [ckText(cells.get(at.description + row)), ckText(cells.get(at.statement + row))].filter(Boolean).join(" · ") });
  }
  requireReport(items.length, "The draft history has no dated rows.");
  // Oldest first: date, then time, then the journal (CAN) number.
  items.sort((a, b) => a.date.localeCompare(b.date) || a.time.localeCompare(b.time) || a.journal - b.journal);
  return { file: file.name, items };
}

// Last month's workbook: the latest CC Worksheet block (outstanding items on the left, the month's written and cleared
// items on the right), the Checks Outstanding list and the cover cells.
async function ckReadBook(file) {
  const book = await openExcel(file, true);
  for (const sheet of ["Reconciliation", "CC Worksheet", "Checks Outstanding"]) requireReport(book.sheets.has(sheet), "The workbook has no " + sheet + " sheet.");
  const cc = await readCells(book, "CC Worksheet");
  const rows = ckRows(cc);
  const titles = rows.filter(row => /^other outstanding$/i.test(ckText(cc.get("A" + row))));
  requireReport(titles.length, "The CC Worksheet has no OTHER OUTSTANDING block.");
  const title = titles.at(-1), head = title - 1;
  const headDate = ckDate(cc.get("A" + head), book.date1904);
  requireReport(headDate, "The CC Worksheet block above row " + title + " has no month date.");
  const rightColumn = [...cc].find(([address, value]) => Number(/\d+$/.exec(address)[0]) === title && /^written and cleared$/i.test(ckText(value)))?.[0].replace(/\d+$/, "");
  requireReport(rightColumn, "The CC Worksheet block needs a WRITTEN AND CLEARED heading.");
  const right = [0, 1, 2, 3].map(k => csvColumn(columnNumber(rightColumn) + k));
  const body = rows.filter(row => row > title);
  const left = body.filter(row => typeof cc.get("D" + row) === "number").map(row => ({
    row, cells: ["A", "B", "C", "D"].map(column => cc.get(column + row)), date: ckDate(cc.get("A" + row), book.date1904),
    ref: cc.get("B" + row) ?? "", code: ckText(cc.get("C" + row)).toUpperCase(), amount: cents(cc.get("D" + row)), carried: true }));
  const previous = rows.filter(row => row < head);
  const gap = Math.max(0, head - (previous.at(-1) ?? head - 1) - 1);
  const outstandingCells = await readCells(book, "Checks Outstanding");
  const outRows = ckRows(outstandingCells);
  const totalRow = outRows.find(row => /^total checks outstanding$/i.test(ckText(outstandingCells.get("B" + row))));
  requireReport(totalRow, "The Checks Outstanding sheet needs its Total Checks Outstanding row.");
  const outstanding = [], years = [];
  for (const row of outRows.filter(row => row > 3 && row < totalRow)) {
    const b = outstandingCells.get("B" + row), d = outstandingCells.get("D" + row);
    if (typeof d === "number") outstanding.push({ row, date: ckDate(b, book.date1904), amount: cents(d), code: ckText(outstandingCells.get("E" + row)) });
    else if (typeof b === "number" && Number.isInteger(b) && b >= 1990 && b <= 2200) years.push({ row, year: b });
    else if (typeof b === "number" || /^\d{4}$/.test(ckText(b))) {
      const date = ckDate(b, book.date1904);
      years.push({ row, year: /^\d{4}$/.test(ckText(b)) ? Number(ckText(b)) : Number(date?.slice(0, 4)) });
    }
  }
  const cover = await readCells(book, "Reconciliation");
  const coverRow = label => ckRows(cover).find(row => label.test(ckText(cover.get("A" + row))));
  return { file, book, head, title, headDate, right, left, gap, outstanding, years, totalRow, outCells: outstandingCells,
    coverTitle: ckText(cover.get("A2")), balanceRow: coverRow(/^balance per 903 account$/i), completedRow: coverRow(/^reconciliation completed by:?$/i),
    format: (await readWorkbookFormatting(book, ["CC Worksheet"]))?.sheets["CC Worksheet"], bodyRows: body, lastRow: rows.at(-1),
    rightRows: body.filter(row => typeof cc.get(right[3] + row) === "number"), leftRows: left.map(item => item.row) };
}

// The month's 903 activity in order: amounts in (checks and journal entries written) open an item; amounts out (drafts
// and journals) clear the open item of the same amount, the month's own items first, then the oldest carried item; two
// same-day items can clear together. Anything out with nothing to clear is flagged.
function ckReconcile(book, history, month) {
  const [year, number] = month.split("-").map(Number);
  const first = month + "-01", last = new Date(Date.UTC(year, number, 0)).toISOString().slice(0, 10);
  requireReport(book.headDate < first, "Last month's workbook already has a block for " + month + " or later. Use last month's reconciliation.");
  const monthRows = history.items.filter(item => item.date >= first && item.date <= last);
  const before = history.items.filter(item => item.date < first);
  const open = book.left.map(item => ({ ...item }));
  const pairs = [], issues = [], takenRows = new Set();
  for (const item of monthRows) {
    if (item.amount > 0) { open.push({ date: item.date, ref: "", code: item.code, amount: item.amount, carried: false, history: item }); continue; }
    if (!item.amount) continue;
    const want = -item.amount;
    const pick = () => {
      const exact = open.filter(entry => entry.amount === want);
      return exact.find(entry => !entry.carried) ?? exact[0];
    };
    let written = pick();
    if (written) {
      open.splice(open.indexOf(written), 1);
      pairs.push({ written: [written], cleared: item });
      continue;
    }
    let pair;
    for (let a = 0; a < open.length && !pair; a++) for (let b = a + 1; b < open.length && !pair; b++) {
      if (open[a].date === open[b].date && open[a].amount + open[b].amount === want) pair = [open[a], open[b]];
    }
    if (pair) {
      for (const entry of pair) open.splice(open.indexOf(entry), 1);
      pairs.push({ written: pair, cleared: item, note: "Two items written the same day cleared together" });
      continue;
    }
    // An older check on the Checks Outstanding list (not in the worksheet's recent list): the draft number first, then the
    // oldest check of that amount.
    const listed = book.outstanding.filter(row => !takenRows.has(row.row) && row.amount === want);
    const byNumber = listed.find(row => item.draft && String(row.code) === String(item.draft));
    const old = byNumber ?? listed[0];
    if (old) {
      takenRows.add(old.row);
      pairs.push({ written: [{ date: old.date, code: String(old.code || "CK"), amount: old.amount, carried: true, outRow: old.row }], cleared: item,
        note: "Clears a check from the Checks Outstanding list" });
      continue;
    }
    pairs.push({ written: [], cleared: item, review: true });
    issues.push({ date: item.date, amount: item.amount, text: item.code + " " + (item.draft ? "draft " + item.draft + " " : "") + ckMoney.format(item.amount / 100)
      + " has no outstanding check or journal entry of that amount" });
  }
  const balanceRow = monthRows.at(-1) ?? before.at(-1);
  requireReport(balanceRow, "The draft history has no balance at the end of " + month + ".");
  const balance = balanceRow.balance;
  const total = open.reduce((sum, entry) => sum + entry.amount, 0);
  // Checks Outstanding rows to clear and to add.
  const used = new Set();
  const clearRows = [];
  for (const entry of pairs.flatMap(pair => pair.written).filter(entry => entry.carried)) {
    if (entry.outRow) { used.add(entry.outRow); clearRows.push(entry.outRow); continue; }
    const row = book.outstanding.find(item => !used.has(item.row) && item.amount === entry.amount && (!entry.date || !item.date || item.date === entry.date))
      ?? book.outstanding.find(item => !used.has(item.row) && item.amount === entry.amount);
    if (row) { used.add(row.row); clearRows.push(row.row); }
    else issues.push({ date: entry.date, amount: entry.amount, text: "Cleared item " + ckMoney.format(entry.amount / 100) + " was not found on Checks Outstanding; remove it there by hand" });
  }
  const added = open.filter(entry => !entry.carried);
  const listTotal = book.outstanding.filter(item => !clearRows.includes(item.row)).reduce((sum, item) => sum + item.amount, 0)
    + added.reduce((sum, entry) => sum + entry.amount, 0);
  const listCount = book.outstanding.length - clearRows.length + added.length;
  if (listTotal !== balance) issues.push({ date: last, amount: balance - listTotal, text: "The Checks Outstanding list will not add up to the 903 balance" });
  return { month, first, last, monthRows, pairs, open, added, clearRows, balance, total, listTotal, listCount, issues };
}

function ckShow(result) {
  const table = (rows, headers) => {
    const element = document.createElement("table");
    if (headers) element.createTHead().insertRow().append(...headers.map(text => Object.assign(document.createElement("th"), { textContent: text })));
    const body = element.createTBody();
    for (const values of rows) {
      const row = body.insertRow();
      values.forEach((value, i) => {
        const cell = i === 0 && !headers ? Object.assign(document.createElement("th"), { scope: "row" }) : row.insertCell();
        cell.textContent = value;
        if (i === 0 && !headers) row.append(cell);
      });
    }
    return element;
  };
  const written = result.monthRows.filter(item => item.amount > 0).length;
  ckField("summary").replaceChildren(table([
    ["903 balance at month end", ckMoney.format(result.balance / 100)],
    ["Checks Outstanding after this month", result.listCount + " · " + ckMoney.format(result.listTotal / 100)],
    ["Difference", ckMoney.format((result.balance - result.listTotal) / 100)],
    ["Written this month", written],
    ["Cleared this month", result.pairs.filter(pair => pair.written.length).length + " (" + result.pairs.filter(pair => pair.written.some(entry => entry.carried)).length + " from earlier months)"],
    ["New outstanding", result.added.length],
    ["Items to review", result.issues.length],
  ]));
  ckField("details").replaceChildren(table(
    [...result.pairs.map(pair => [pair.cleared.date, pair.cleared.code + (pair.cleared.draft ? " " + pair.cleared.draft : ""), ckMoney.format(pair.cleared.amount / 100),
      pair.written.map(entry => (entry.date ?? "") + " " + entry.code + " " + ckMoney.format(entry.amount / 100)).join(" + ") || "—",
      pair.review ? "Review" : pair.note ?? (pair.written.some(entry => entry.carried) ? "Clears an earlier month" : "Cleared")]),
    ...result.added.map(entry => [entry.date, entry.code, ckMoney.format(entry.amount / 100), "—", "New outstanding"]),
    ...result.issues.map(issue => [issue.date, "", ckMoney.format(issue.amount / 100), "", issue.text])],
    ["Date", "Item", "Amount", "Written item", "Result"]));
  ckField("results").hidden = false;
}

// The month's CC Worksheet block, laid out and formatted like last month's: the month date, the OTHER OUTSTANDING /
// WRITTEN AND CLEARED headings, outstanding items on the left and the written and cleared pairs on the right.
function ckBlock(book, result) {
  const grid = new Map(), roles = new Map(), cellRoles = new Map(), marks = [];
  const put = (r, column, xml) => { const cells = grid.get(r) ?? new Map(); cells.set(column, xml); grid.set(r, cells); };
  const value = (r, column, raw) => {
    if (raw === undefined || raw === null || raw === "") return;
    if (typeof raw === "number") put(r, column, `<c r="${column}${r}" s="0"><v>${raw}</v></c>`);
    else put(r, column, `<c r="${column}${r}" s="0" t="inlineStr"><is><t xml:space="preserve">${xmlText(String(raw))}</t></is></c>`);
  };
  let r = 1 + Math.max(1, book.gap);
  const head = r;
  value(r, "A", ckSerial(result.first));
  roles.set(r++, "head");
  value(r, "A", "OTHER OUTSTANDING");
  value(r, book.right[0], "WRITTEN AND CLEARED");
  roles.set(r++, "title");
  const start = r;
  const leftRows = result.open.map(entry => entry.carried ? entry.cells
    : [ckSerial(entry.date), entry.ref || "", entry.code, entry.amount / 100]);
  leftRows.forEach((cells, k) => ["A", "B", "C", "D"].forEach((column, i) => value(start + k, column, cells[i])));
  // Pairs written this month first, then those clearing earlier months, each followed by a blank row.
  const rightRows = [];
  const ordered = [...result.pairs].sort((x, y) => Number(x.written.some(entry => entry.carried)) - Number(y.written.some(entry => entry.carried)));
  ordered.forEach((pair, index) => {
    if (index) rightRows.push({ cells: [] });
    for (const entry of pair.written) rightRows.push({ cells: [entry.date ? ckSerial(entry.date) : entry.cells?.[0], "", entry.code, entry.amount / 100] });
    rightRows.push({ cells: [ckSerial(pair.cleared.date), pair.cleared.draft || "", pair.cleared.code, pair.cleared.amount / 100], review: pair.review });
  });
  rightRows.forEach((entry, k) => {
    entry.cells.forEach((raw, i) => value(start + k, book.right[i], raw));
    if (entry.cells.length) for (const column of book.right) cellRoles.set(column + (start + k), "right");
    if (entry.review) for (const column of book.right) marks.push({ row: start + k, column });
  });
  const end = start + Math.max(leftRows.length, rightRows.length, 1) - 1;
  for (let k = start; k <= end; k++) roles.set(k, "body");
  const columns = Array.from({ length: columnNumber(book.right[3]) }, (_, k) => csvColumn(k + 1));
  // Outstanding items take last month's left-side format and written/cleared items its right-side format.
  for (let k = 0; k < leftRows.length; k++) for (const column of ["A", "B", "C", "D"]) cellRoles.set(column + (start + k), "left");
  const roleRows = { head: book.head, title: book.title, body: book.bodyRows };
  if (book.leftRows.length) roleRows.left = book.leftRows;
  if (book.rightRows.length) roleRows.right = book.rightRows;
  const styled = styledBlock({ grid, roles, cellRoles }, book.format, roleRows, columns, []);
  return { append: { sheet: "CC Worksheet", grid: styled.styled, merges: styled.merges, heights: styled.heights, end, after: book.lastRow }, marks, head };
}

async function ckBuild(book, result) {
  const audit = reconciliationAudit();
  const [year, number] = result.month.split("-");
  const updates = [["Reconciliation", "A2", "For " + ckMonthNames[Number(number) - 1] + " " + year]];
  if (book.balanceRow) updates.push(["Reconciliation", "E" + book.balanceRow, result.balance / 100]);
  if (book.completedRow) updates.push(["Reconciliation", "E" + book.completedRow, audit.by], ["Reconciliation", "E" + (book.completedRow + 1), { date: audit.on }]);
  for (const row of result.clearRows) for (const column of ["B", "C", "D", "E", "F"]) if (book.outCells.has(column + row)) updates.push(["Checks Outstanding", column + row, null]);
  // New outstanding items go in their year's section of Checks Outstanding, in the first empty rows.
  const taken = new Set();
  const notes = [];
  for (const entry of result.added) {
    const section = book.years.filter(item => item.year === Number(entry.date.slice(0, 4))).at(-1);
    const next = section && (book.years.find(item => item.row > section.row)?.row ?? book.totalRow);
    let row;
    for (let r = section ? section.row + 1 : Infinity; r < next; r++) {
      if (taken.has(r) || book.outCells.get("D" + r) !== undefined && !result.clearRows.includes(r) || ckText(book.outCells.get("B" + r))) continue;
      row = r;
      break;
    }
    if (!row) { notes.push("Add " + entry.code + " " + ckMoney.format(entry.amount / 100) + " (" + entry.date + ") to Checks Outstanding by hand: no empty row in its year."); continue; }
    taken.add(row);
    updates.push(["Checks Outstanding", "B" + row, { date: entry.date }], ["Checks Outstanding", "D" + row, entry.amount / 100], ["Checks Outstanding", "E" + row, entry.code]);
  }
  const block = ckBlock(book, result);
  let blob = await fillExcelTemplate(book.file, updates, [block.append]);
  const marks = block.marks.map(mark => ({ sheet: "CC Worksheet", address: mark.column + (block.append.startRow + mark.row - 1), color: "FFFF7C80" }));
  if (marks.length) blob = await excelFillCells(blob, book.file.name, marks);
  return { blob, notes };
}

for (const kind of ["book", "history"]) {
  const input = ckField(kind + "-file"), drop = ckField(kind + "-drop");
  const select = files => {
    ckFiles[kind] = files[0];
    ckField(kind + "-status").textContent = files[0]?.name ?? "";
    ckResult = undefined;
    ckField("results").hidden = true;
  };
  input.addEventListener("change", () => select(input.files));
  drop.addEventListener("dragover", event => { event.preventDefault(); drop.classList.add("active"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("active"));
  drop.addEventListener("drop", event => { event.preventDefault(); drop.classList.remove("active"); input.value = ""; select(event.dataTransfer.files); });
}
document.addEventListener("dragover", event => event.preventDefault());
document.addEventListener("drop", event => event.preventDefault());

ckField("run").addEventListener("click", async () => {
  const current = ++ckRevision;
  ckResult = undefined;
  ckField("results").hidden = true;
  try {
    const month = ckField("month").value;
    requireReport(month && ckField("confirm").checked, "Choose the month and confirm the source files.");
    requireReport(ckFiles.book && ckFiles.history, "Load last month's reconciliation and the 903 draft history.");
    const [book, history] = await Promise.all([ckReadBook(ckFiles.book), ckReadHistory(ckFiles.history)]);
    if (current !== ckRevision) return;
    const result = ckReconcile(book, history, month);
    ckResult = { book, result };
    ckShow(result);
    ckStatus(result.issues.length ? "Review the items listed below before using the workbook." : "");
  } catch (error) {
    if (current === ckRevision) ckStatus(error.message, true);
  }
});

ckField("download").addEventListener("click", async event => {
  const state = ckResult;
  if (!state) return;
  event.currentTarget.disabled = true;
  try {
    const { blob, notes } = await ckBuild(state.book, state.result);
    if (notes.length) ckStatus(notes.join(" "), true);
    // Named for the month it is prepared in, like the saved workbooks (September's reconciliation is 10-2026).
    const [year, number] = state.result.month.split("-").map(Number);
    const prepared = String(number % 12 + 1).padStart(2, "0") + "-" + (number === 12 ? year + 1 : year);
    const name = /^\d{2}-\d{4}/.test(state.book.file.name) ? state.book.file.name.replace(/^\d{2}-\d{4}/, prepared) : prepared + " " + state.book.file.name;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = name;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 30000);
  } catch (error) {
    ckStatus(error.message, true);
  } finally {
    event.currentTarget.disabled = false;
  }
});
