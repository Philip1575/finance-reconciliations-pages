"use strict";

const spField = id => document.getElementById(id);
const suspenseAccounts = {
  share: { name: "Share draft", source: "trial", note: "Compare the Wisdom ending balance with the expected month-end balance." },
  transfer: { name: "ATM transfer", source: "history", note: "Review the timing difference between the ATM statement date and the CUSA close." },
  ach: { name: "ACH settlement", source: "history", note: "Review the CUSA closing balance and any outstanding settlement items." },
  atm: { name: "ATM exception", source: "history", note: "If a difference remains, compare the Daily VIA Offload and IF-67 reports." },
  teller: { name: "Teller operation", source: "history", note: "Confirm the expected balance and any reviewed items from the existing reconciliation." },
  credit: { name: "Exception credit", source: "history", note: "Review the closing balance and any recurring entries for the month." },
  visa: { name: "Visa provisional credit", source: "history", note: "Use this legacy account for historical periods as applicable." },
};
const suspenseFiles = {};
let suspenseResult, suspenseRevision = 0, profileRevision = 0, ledgerLoad = 0, suspenseLedger;
let suspenseProfile = new Map();
let suspenseNames = {};
let batchResult;
let batchExporting = false;
let batchRevision = 0, batchLoad = 0, batchLedger;
let suspenseTemplate, templateRevision = 0;
const suspenseDifferenceFormula = "K10+SUM(K15:K22)-K23";
addTrialBalanceDetection("", () => suspenseLedger ? [{ book: suspenseLedger }] : [],
  ["account-column", "balance-column", "first", "sheet"]);

function spStatus(message, error = false) {
  spField("status").textContent = message;
  spField("status").classList.toggle("error", error);
}

function invalidateSuspense() {
  suspenseRevision++;
  suspenseResult = undefined;
  spField("confirm").checked = false;
  spField("results").hidden = true;
  spStatus("");
}

function spSourceFields() {
  const trial = spField("source-type").value === "trial";
  spField("trial-fields").hidden = !trial;
  spField("history-fields").hidden = trial;
}

function spSelectAccount() {
  const mode = spField("mode").value;
  const profile = suspenseAccounts[mode];
  spField("account").value = suspenseProfile.get(mode) ?? "";
  spField("account-note").textContent = profile.note;
  spField("source-type").value = profile.source;
  spField("timing-fields").hidden = mode !== "transfer";
  spField("trace-fields").hidden = mode !== "atm";
  spSourceFields();
}

spField("profile").addEventListener("change", async () => {
  const revision = ++profileRevision;
  suspenseProfile = new Map();
  suspenseNames = {};
  invalidateBatch();
  spField("account").value = "";
  invalidateSuspense();
  const file = spField("profile").files[0];
  spField("profile-status").textContent = "";
  if (!file) return;
  try {
    requireReport(/\.json$/i.test(file.name) && file.size <= 65536, "Choose the local suspense profile in JSON format.");
    const { accounts, names = {} } = JSON.parse(await file.text());
    const entries = Object.entries(accounts ?? {});
    requireReport(entries.length && entries.every(([key, value]) => Object.hasOwn(suspenseAccounts, key) && typeof value === "string" && /^\d{6}$/.test(value))
      && new Set(entries.map(([, value]) => value)).size === entries.length, "The profile needs distinct account numbers for recognized suspense workflows.");
    requireReport(names && typeof names === "object" && Object.entries(names).every(([key, name]) =>
      Object.hasOwn(suspenseAccounts, key) && typeof name === "string" && name.trim() && name.length <= 80),
    "Optional suspense account names must be short text labels for recognized workflows.");
    if (revision !== profileRevision) return;
    suspenseProfile = new Map(entries);
    suspenseNames = names;
    spField("account").value = suspenseProfile.get(spField("mode").value) ?? "";
    spField("profile-status").textContent = "Local account mappings loaded.";
    invalidateSuspense();
  } catch (error) {
    if (revision === profileRevision) spField("profile-status").textContent = error.message;
  }
});

async function selectSuspenseFile(kind, file) {
  suspenseFiles[kind] = file;
  invalidateSuspense();
  const status = spField(kind + "-status");
  status.textContent = file?.name ?? "";
  if (kind !== "ledger") return;
  const current = ++ledgerLoad;
  suspenseLedger = undefined;
  spField("ledger-map").hidden = true;
  if (!file) return;
  try {
    const book = await openExcel(file, true);
    if (current !== ledgerLoad) return;
    suspenseLedger = book;
    spField("sheet").replaceChildren(...[...book.sheets.keys()].map(name => new Option(name, name)));
    spField("ledger-map").hidden = false;
  } catch (error) {
    if (current === ledgerLoad) status.textContent = error.message;
  }
}

async function spEnding(month, account) {
  const sheet = spField("sheet").value;
  const balance = spField("balance-column").value.trim().toUpperCase();
  const valid = column => /^[A-Z]{1,3}$/.test(column) && columnNumber(column) <= 16384;
  requireReport(valid(balance), "Choose the closing balance column.");
  if (spField("source-type").value === "trial") {
    const first = spField("first").valueAsNumber, last = spField("last");
    const matches = await readWisdom(suspenseLedger, sheet, spField("account-column").value.trim().toUpperCase(), balance,
      first, [account], last.value ? last.valueAsNumber : 1048576, true, month);
    const match = matches.get(account);
    return { value: cents(match.amount), address: sheet + "!" + match.address, date: "" };
  }
  const row = spField("closing-row").valueAsNumber;
  const dateColumn = spField("date-column").value.trim().toUpperCase();
  requireReport(Number.isInteger(row) && row >= 1 && row <= 1048576 && (!dateColumn || (valid(dateColumn) && dateColumn !== balance)),
    "Check the verified closing row and date column.");
  const cells = await readCells(suspenseLedger, sheet);
  const date = dateColumn ? sourceDate(cells.get(dateColumn + row), suspenseLedger.date1904) : "";
  requireReport(!date || date.slice(0, 7) === month, "The selected closing balance row is outside the reconciliation month.");
  const address = balance + row;
  try { return { value: cents(cells.get(address)), address: sheet + "!" + address, date }; }
  catch { throw new Error("The closing balance at " + sheet + "!" + address + " is missing or invalid."); }
}

async function spTrace(file, month) {
  const book = await openExcel(file, true), tables = [];
  for (const sheet of book.sheets.keys()) {
    const cells = await readCells(book, sheet), headers = new Map();
    for (const [address, raw] of cells) {
      const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
      if (!match || Number(match[2]) > 50 || typeof raw !== "string") continue;
      const value = raw.trim().toLowerCase();
      const key = value === "date" || value === "transaction date" ? "date" : value === "reference" || value === "trace" ? "reference"
        : value === "signed amount" || value === "amount" ? "amount" : "";
      if (!key) continue;
      const row = Number(match[2]), fields = headers.get(row) ?? {};
      fields[key] = Object.hasOwn(fields, key) ? "" : match[1];
      headers.set(row, fields);
    }
    const header = [...headers].find(([, f]) => f.date && f.reference && f.amount);
    if (header) tables.push({ sheet, cells, header });
  }
  requireReport(tables.length === 1, file.name + " needs one table with Date, Reference and Signed Amount.");
  const { sheet, cells, header: [first, fields] } = tables[0];
  const rows = new Set();
  for (const [address, value] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (match && Object.values(fields).includes(match[1]) && Number(match[2]) > first && String(value ?? "").trim()) rows.add(Number(match[2]));
  }
  return [...rows].sort((a, b) => a - b).flatMap(row => {
    try {
      const date = sourceDate(cells.get(fields.date + row), book.date1904);
      if (date.slice(0, 7) !== month) return [];
      const reference = String(cells.get(fields.reference + row) ?? "").trim();
      return [{ date, reference, amount: cents(cells.get(fields.amount + row)), sheet, row }];
    } catch (error) {
      throw new Error(file.name + " · " + sheet + " row " + row + ": " + error.message);
    }
  });
}

function compareSuspenseTraces(via, fiserv) {
  const groups = new Map(), review = [];
  for (const [side, entries] of [["via", via], ["fiserv", fiserv]]) {
    for (const item of entries) {
      if (!item.reference) {
        review.push([item.date, "", side === "via" ? [item] : [], side === "fiserv" ? [item] : [], "Missing transaction reference"]);
        continue;
      }
      const key = JSON.stringify([item.date, item.reference]);
      const group = groups.get(key) ?? { date: item.date, reference: item.reference, via: [], fiserv: [] };
      group[side].push(item);
      groups.set(key, group);
    }
  }
  for (const { date, reference, via: left, fiserv: right } of groups.values()) {
    const reason = left.length > 1 || right.length > 1 ? "Repeated references · review"
      : !left.length ? "IF-67 only" : !right.length ? "VIA Offload only"
        : left[0].amount !== right[0].amount ? "Amounts differ" : "Date, reference and amount agree";
    review.push([date, reference, left, right, reason]);
  }
  return review.sort((a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]));
}

function showSuspense(result) {
  const money = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
  const fields = [
    ["Closing ledger balance", money.format(result.ending.value / 100)],
    ...(result.mode === "transfer" ? [["Proposed ATM timing item", money.format(result.timing / 100)]] : []),
    ["Adjusted balance", money.format(result.adjusted / 100)],
    ["Expected balance", money.format(result.expected / 100)],
    ["Difference", money.format(result.difference / 100)],
  ];
  if (result.trace) fields.push(["Trace groups to review", result.trace.filter(row => row[4] !== "Date, reference and amount agree").length]);
  const table = document.createElement("table"), body = table.createTBody();
  for (const [name, value] of fields) {
    const row = body.insertRow(), heading = document.createElement("th");
    heading.scope = "row";
    heading.textContent = name;
    row.append(heading);
    row.insertCell().textContent = value;
  }
  spField("summary").replaceChildren(table);
  spField("download-template").hidden = !suspenseTemplate;
  spField("results").hidden = false;
}

function suspenseAmounts(ending, timing, adjustment, expected) {
  const adjusted = ending + timing + adjustment, difference = adjusted - expected;
  requireReport([adjusted, difference].every(Number.isSafeInteger), "The amounts exceed the supported range.");
  return { adjusted, difference };
}

function suspenseCover(result, audit, cells) {
  const { label, date, number, row, worksheet } = cells;
  const month = new Intl.DateTimeFormat("en-US", { month: "long", year: "numeric", timeZone: "UTC" })
    .format(new Date(result.month + "-01T00:00:00Z"));
  const rows = [
    row(4, label("B4", result.account + " " + result.coverName.toUpperCase() + " RECONCILIATION", 1)),
    row(5, label("B5", "For the Month of " + month)),
    row(10, label("B10", "BALANCE PER WISDOM G/L", 1) + number("K10", result.ending.value / 100)),
    row(14, label("C14", "Item Description", 1) + label("H14", "Date", 1) + label("J14", "Will Clear", 1) + label("K14", "Amount", 1)),
  ];
  if (result.timingDate) rows.push(row(15, label("C15", "ATM closing cutoff timing item") + date("H15", result.timingDate)
    + number("K15", result.timing / 100)));
  if (result.adjustment) rows.push(row(16, label("C16", result.description) + number("K16", result.adjustment / 100)));
  rows.push(
    row(23, label("B23", "Expected balance") + number("K23", result.expected / 100)),
    row(25, label("B25", "Difference", 1) + `<c r="K25" s="2"><f>${suspenseDifferenceFormula}</f><v>${result.difference / 100}</v></c>`),
    ...(result.timingPending ? [row(27, label("B27", "Review the ATM cutoff before filing."))] : []),
    row(34, label("G34", "Prepared By") + label("K34", audit.by)),
    row(35, label("G35", "Completed on") + date("K35", audit.on)),
    row(37, label("G37", "Verified By")),
  );
  return ["Reconciliation", worksheet("A1:K37", [[1, 1, 8], [2, 2, 29], [3, 3, 36], [4, 6, 8],
    [7, 7, 17], [8, 8, 13], [9, 9, 8], [10, 10, 14], [11, 11, 21]], rows.join(""),
  ["B4:K4", "B5:K5", "B10:J10", "C14:G14", "C15:G15", "C16:G16", "B23:J23", "B25:J25", "B27:K27", "G34:J34", "G35:J35", "G37:J37"])];
}

async function readSuspenseTemplate(file) {
  const workbook = await openExcel(file, true);
  requireReport(workbook.sheets.size === 1, "Use a workbook containing only the shared suspense cover.");
  const sheet = [...workbook.sheets.keys()][0];
  const formulas = new Map(), cells = await readCells(workbook, sheet, formulas);
  const text = address => String(cells.get(address) ?? "").trim();
  requireReport(/reconciliation$/i.test(text("B4")) && /^for the month of /i.test(text("B5"))
    && /^balance per wisdom g\/l$/i.test(text("B10")) && /^difference$/i.test(text("B25"))
    && ["Item Description", "Date", "Will Clear", "Amount"].every((title, i) => text("CHJK"[i] + "14") === title),
    "Use the shared suspense cover with the Wisdom balance and in-transit items.");
  const audit = completionTarget(sheet, cells, formulas, "G34", "G37", "K34", "K35");
  const difference = formulas.get("K25");
  requireReport(audit && (!difference || String(difference).replace(/[\s$]/g, "").toUpperCase() === suspenseDifferenceFormula),
    "Check the completion fields and difference formula on the shared cover.");
  const clear = templateTableUpdates(sheet, { first: 15, last: 22 }, "CHJK", [], cells, formulas);
  requireReport(!["B4", "B5", "K10", "B23", "K23", "B27", "K37", "K38"].some(cell => formulas.has(cell))
    && ["", "Expected balance"].includes(text("B23"))
    && ["", "Review the ATM cutoff before filing."].includes(text("B27")),
    "Check the balance and review inputs on the shared cover.");
  return { file, sheet, clear, difference, extension: file.name.split(".").at(-1) };
}

function fillSuspenseTemplate(result, audit, template) {
  const { sheet } = template;
  const month = new Intl.DateTimeFormat("en-US", { month: "long", year: "numeric", timeZone: "UTC" })
    .format(new Date(result.month + "-01T00:00:00Z"));
  const values = new Map(template.clear.map(([, cell, value]) => [cell, value]));
  for (const [cell, value] of [
    ["B4", result.account + " " + result.coverName.toUpperCase() + " RECONCILIATION"],
    ["B5", "For the Month of " + month], ["K10", result.ending.value / 100],
    ["B23", result.expected ? "Expected balance" : null], ["K23", result.expected ? result.expected / 100 : null],
    ["B27", result.timingPending ? "Review the ATM cutoff before filing." : null],
    ["K34", audit.by], ["K35", { date: audit.on }], ["K37", null], ["K38", null],
  ]) values.set(cell, value);
  values.set("K25", { formula: template.difference ?? suspenseDifferenceFormula, cached: result.difference / 100 });
  if (result.timingDate) {
    values.set("C15", "ATM closing cutoff timing item");
    values.set("H15", { date: result.timingDate });
    values.set("K15", result.timing / 100);
  }
  if (result.adjustment) {
    values.set("C16", result.description);
    values.set("K16", result.adjustment / 100);
  }
  return fillExcelTemplate(template.file, [...values].map(([cell, value]) => [sheet, cell, value]));
}

async function buildSuspense(result, audit = reconciliationAudit()) {
  const cells = excelOutputCells();
  const { label, number, row, worksheet } = cells;
  function table(name, headings, entries, widths) {
    const cols = "ABCDEFGH".slice(0, headings.length);
    const body = row(1, headings.map((title, i) => label(cols[i] + "1", title, 1)).join(""))
      + entries.map((entry, i) => row(i + 2, entry.map((value, j) => {
        const cell = cols[j] + (i + 2);
        return typeof value === "number" ? number(cell, value) : label(cell, value);
      }).join(""))).join("");
    return [name, worksheet("A1:" + cols.at(-1) + (entries.length + 1),
      widths.map((width, i) => [i + 1, i + 1, width]), body)];
  }
  const balanceRow = (index, title, amount) => row(index, label("A" + index, title) + number("B" + index, amount / 100));
  const coverRows = [
    row(1, label("A1", result.name + " reconciliation", 1)),
    row(2, label("A2", "Period") + label("B2", result.month)),
    row(3, label("A3", "G/L account") + label("B3", result.account)),
    balanceRow(4, "Closing ledger balance", result.ending.value),
    balanceRow(5, result.mode === "transfer" ? "Proposed ATM timing item" : "", result.timing),
    balanceRow(6, "Other reviewed items", result.adjustment),
    row(7, label("A7", "Adjusted balance") + `<c r="B7" s="2"><f>SUM(B4:B6)</f><v>${result.adjusted / 100}</v></c>`),
    balanceRow(8, "Expected balance", result.expected),
    row(9, label("A9", "Difference") + `<c r="B9" s="2"><f>B7-B8</f><v>${result.difference / 100}</v></c>`),
  ];
  const lineItems = [];
  if (result.timingDate) lineItems.push([result.timingDate, "ATM closing cutoff timing item · verify", result.timing / 100]);
  if (result.adjustment) lineItems.push(["", result.description, result.adjustment / 100]);
  const sheets = [
    suspenseCover(result, audit, cells),
    [result.name + " cover", worksheet("A1:B9", [[1, 1, 38], [2, 2, 23]], coverRows.join(""))],
    table("Line items", ["Date", "Description", "Amount"], lineItems, [15, 61, 19]),
    table("Sources", ["Field", "Value"], [
      ["Ledger file", result.ledgerFile], ["Closing balance cell", result.ending.address],
      ["Closing balance date", result.ending.date], ["VIA Offload file", result.viaFile], ["IF-67 file", result.fiservFile],
    ], [27, 54]),
  ];
  if (result.trace) {
    const total = entries => entries.reduce((sum, item) => sum + item.amount, 0) / 100;
    const refs = entries => entries.map(item => item.sheet + " row " + item.row).join(", ");
    sheets.push(table("Trace review", ["Date", "Reference", "VIA amount", "IF-67 amount", "VIA source rows", "IF-67 source rows", "Review"],
      result.trace.map(([date, reference, via, fiserv, status]) => [date, reference, via.length ? total(via) : "",
        fiserv.length ? total(fiserv) : "", refs(via), refs(fiserv), status]), [15, 22, 16, 16, 31, 31, 43]));
  }
  return createExcelWorkbook(sheets, audit);
}

for (const kind of ["ledger", "via", "fiserv"]) {
  const input = spField(kind + "-file"), drop = spField(kind + "-drop");
  input.addEventListener("change", () => selectSuspenseFile(kind, input.files[0]));
  drop.addEventListener("dragover", event => { event.preventDefault(); drop.classList.add("active"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("active"));
  drop.addEventListener("drop", event => {
    event.preventDefault();
    drop.classList.remove("active");
    if (event.dataTransfer.files.length !== 1) return spStatus("Drop one workbook at a time.", true);
    input.value = "";
    selectSuspenseFile(kind, event.dataTransfer.files[0]);
  });
}
document.addEventListener("dragover", event => event.preventDefault());
document.addEventListener("drop", event => event.preventDefault());
spField("mode").addEventListener("change", spSelectAccount);
spField("source-type").addEventListener("change", spSourceFields);
spSelectAccount();
spField("month").addEventListener("input", () => {
  const month = spField("month").value;
  if (!month) return;
  const [year, m] = month.split("-").map(Number);
  spField("timing-date").value = new Date(Date.UTC(year, m - 1, 0)).toISOString().slice(0, 10);
});
for (const input of document.querySelectorAll("input:not([type=file]):not(#confirm):not([id^=batch-]), select:not([id^=batch-])")) {
  input.addEventListener("input", invalidateSuspense);
}
spField("confirm").addEventListener("change", () => { if (!spField("confirm").checked) invalidateSuspense(); });

spField("cover-template").addEventListener("change", async event => {
  const current = ++templateRevision, file = event.currentTarget.files[0];
  const status = spField("cover-template-status");
  suspenseTemplate = undefined;
  invalidateSuspense();
  invalidateBatch();
  status.classList.remove("error");
  status.textContent = file ? "Checking the shared cover…" : "";
  if (!file) return;
  try {
    const template = await readSuspenseTemplate(file);
    if (current !== templateRevision) return;
    suspenseTemplate = template;
    status.textContent = "Shared cover loaded. Batch downloads will use its layout.";
  } catch (error) {
    if (current !== templateRevision) return;
    status.classList.add("error");
    status.textContent = error.message;
  }
});

spField("run").addEventListener("click", async () => {
  const current = ++suspenseRevision;
  suspenseResult = undefined;
  spField("results").hidden = true;
  try {
    const mode = spField("mode").value, name = suspenseAccounts[mode].name;
    const month = spField("month").value, account = normalizeAccount(spField("account").value);
    requireReport(month && account && suspenseLedger && spField("confirm").checked, "Choose the month, account and ledger workbook, then verify the inputs.");
    const timingDate = mode === "transfer" ? spField("timing-date").value : "";
    requireReport(mode !== "transfer" || timingDate, "Enter the reviewed timing item date.");
    const adjustment = cents(spField("adjustment").value), expected = cents(spField("expected").value);
    const description = spField("adjustment-description").value.trim();
    requireReport(!adjustment || description, "Describe the other reviewed line items.");
    const ending = await spEnding(month, account);
    if (current !== suspenseRevision) return;
    const timing = mode === "transfer" ? -ending.value : 0;
    const totals = suspenseAmounts(ending.value, timing, adjustment, expected);
    let trace;
    if (mode === "atm" && (suspenseFiles.via || suspenseFiles.fiserv)) {
      requireReport(suspenseFiles.via && suspenseFiles.fiserv, "Load both VIA Offload and IF-67 to compare ATM traces.");
      const [via, fiserv] = await Promise.all([spTrace(suspenseFiles.via, month), spTrace(suspenseFiles.fiserv, month)]);
      requireReport(via.length || fiserv.length, "The trace exports have no rows for the selected month.");
      trace = compareSuspenseTraces(via, fiserv);
    }
    if (current !== suspenseRevision) return;
    suspenseResult = { mode, name, coverName: suspenseNames[mode] ?? name, month, account, timingDate, timing, adjustment, expected, description, ending, ...totals, trace,
      ledgerFile: suspenseFiles.ledger.name, viaFile: mode === "atm" ? suspenseFiles.via?.name ?? "" : "",
      fiservFile: mode === "atm" ? suspenseFiles.fiserv?.name ?? "" : "" };
    showSuspense(suspenseResult);
    spStatus(mode === "transfer" ? "Verify the timing item before using the adjusted balance."
      : mode === "atm" && totals.difference ? "An ATM difference remains. Review the VIA Offload and IF-67 activity." : "");
  } catch (error) {
    if (current === suspenseRevision) spStatus(error.message, true);
  }
});

async function downloadSuspenseSingle(button, filled) {
  const result = suspenseResult, template = suspenseTemplate;
  if (!result) return;
  button.disabled = true;
  try {
    const blob = filled ? await fillSuspenseTemplate(result, reconciliationAudit(), template) : await buildSuspense(result);
    if (result !== suspenseResult) return;
    downloadSuspense(blob, "Suspense " + result.name + " " + result.month + (filled ? " filled." + template.extension : ".xlsx"));
  } catch (error) {
    if (result === suspenseResult) spStatus(error.message, true);
  } finally {
    button.disabled = false;
  }
}

spField("download").addEventListener("click", event => downloadSuspenseSingle(event.currentTarget, false));
spField("download-template").addEventListener("click", event => downloadSuspenseSingle(event.currentTarget, true));

function downloadSuspense(blob, filename) {
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(link.href), 30000);
}

function batchStatus(message, error = false) {
  spField("batch-status").textContent = message;
  spField("batch-status").classList.toggle("error", error);
}

function invalidateBatch() {
  batchRevision++;
  batchResult = undefined;
  spField("batch-confirm").checked = false;
  spField("batch-download").disabled = true;
  spField("batch-results").hidden = true;
  batchStatus("");
}

function makeBatchAccount(month, account, mode, ending, ledgerFile, expected = 0, adjustment = 0, description = "", timingDate = "") {
  const timingPending = mode === "transfer" && ending.value !== 0 && !timingDate;
  const timing = timingDate ? -ending.value : 0;
  const name = suspenseAccounts[mode].name;
  return { month, mode, account, name, coverName: suspenseNames[mode] ?? name, ending,
    expected, adjustment, description, timing, timingDate, timingPending,
    ...suspenseAmounts(ending.value, timing, adjustment, expected), ledgerFile, viaFile: "", fiservFile: "" };
}

function readSuspenseBatch(month, text) {
  const accounts = new Map([...suspenseProfile].map(([mode, account]) => [account, mode]));
  const matches = new Map();
  let ignored = 0;
  for (const [index, line] of text.replace(/^\uFEFF/, "").split(/\r\n|\n|\r/).entries()) {
    if (!line.trim()) continue;
    const [rawAccount, balance, rawExpected, rawAdjustment, description = "", rawDate] = line.split("\t").map(value => value.trim());
    const account = normalizeAccount(rawAccount), mode = accounts.get(account);
    if (!mode) {
      if (account) ignored++;
      continue;
    }
    try {
      requireReport(!matches.has(account), "The suspense G/L appears more than once.");
      const value = cents(balance), expected = rawExpected ? cents(rawExpected) : 0;
      const adjustment = rawAdjustment ? cents(rawAdjustment) : 0;
      requireReport(!adjustment || description, "Describe the signed adjustment.");
      requireReport(!rawDate || mode === "transfer", "Use the timing date only for the ATM transfer account.");
      requireReport(!rawDate || /^\d{4}-\d{2}-\d{2}$/.test(rawDate), "Enter the reviewed ATM timing date in YYYY-MM-DD format.");
      const timingDate = rawDate ? sourceDate(rawDate) : "";
      matches.set(account, makeBatchAccount(month, account, mode, { value, address: "Pasted row " + (index + 1), date: "" },
        "Pasted G/L ending balances", expected, adjustment, description, timingDate));
    } catch (error) {
      throw new Error("Row " + (index + 1) + " · " + account + " · " + error.message);
    }
  }
  const missing = [...accounts].filter(([account]) => !matches.has(account));
  return { matches, missing, ignored };
}

async function suggestSuspenseBatchColumns() {
  const revision = batchRevision, ledger = batchLedger;
  const sheet = spField("batch-sheet").value;
  const status = spField("batch-file-status");
  if (!ledger) return;
  status.textContent = ledger.file.name;
  spField("batch-account-column").value = "";
  spField("batch-balance-column").value = "";
  spField("batch-first").value = "2";
  spField("batch-last").value = "";
  try {
    const columns = await detectTrialBalanceColumns(ledger.book, sheet);
    if (revision !== batchRevision) return;
    if (!columns) return;
    spField("batch-account-column").value = columns.account;
    spField("batch-balance-column").value = columns.balance;
    spField("batch-first").value = columns.first;
    status.textContent = ledger.file.name + " · Account and ending balance columns detected.";
  } catch (error) {
    if (revision === batchRevision) status.textContent = error.message;
  }
}

async function selectSuspenseBatchFile(file) {
  const current = ++batchLoad;
  const status = spField("batch-file-status");
  batchLedger = undefined;
  invalidateBatch();
  spField("batch-map").hidden = true;
  status.textContent = file ? "Opening " + file.name + "…" : "";
  if (!file) return;
  try {
    const book = await openExcel(file, true);
    if (current !== batchLoad) return;
    batchLedger = { file, book };
    spField("batch-sheet").replaceChildren(...[...book.sheets.keys()].map(name => new Option(name, name)));
    spField("batch-map").hidden = false;
    await suggestSuspenseBatchColumns();
  } catch (error) {
    if (current === batchLoad) status.textContent = error.message;
  }
}

async function readSuspenseBatchWorkbook(month) {
  requireReport(batchLedger, "Load the month-end trial balance workbook.");
  const { book, file } = batchLedger;
  const sheet = spField("batch-sheet").value, last = spField("batch-last");
  const cells = await readCells(book, sheet);
  const native = /^consolidated trial balance report$/i.test(String(cells.get("B1") ?? "").trim());
  const accountColumn = spField("batch-account-column").value.trim().toUpperCase();
  const balanceColumn = spField("batch-balance-column").value.trim().toUpperCase();
  const accounts = new Map([...suspenseProfile].map(([mode, account]) => [account, mode]));
  const rows = await readWisdom(book, sheet, accountColumn, balanceColumn, spField("batch-first").valueAsNumber,
    [...accounts.keys()], last.value ? last.valueAsNumber : 1048576, false, month);
  const matches = new Map([...rows].map(([account, entry]) => {
    const address = sheet + "!" + entry.address;
    let value;
    try { value = cents(entry.amount); }
    catch { throw new Error("Use a balance with at most two decimals at " + address + "."); }
    const mode = accounts.get(account);
    const result = makeBatchAccount(month, account, mode, { value, address, date: "" }, file.name);
    const reportedName = native && accountColumn === "A" ? String(cells.get("E" + entry.address.match(/\d+$/)[0]) ?? "").trim() : "";
    if (reportedName && !suspenseNames[mode]) result.coverName = reportedName;
    return [account, result];
  }));
  return { matches, missing: [...accounts].filter(([account]) => !matches.has(account)) };
}

function showSuspenseBatch(result) {
  const table = document.createElement("table"), headings = table.createTHead().insertRow(), body = table.createTBody();
  for (const heading of ["Suspense account", "G/L", "Ending balance", "Expected", "Difference", "Review"]) {
    const cell = document.createElement("th");
    cell.scope = "col";
    cell.textContent = heading;
    headings.append(cell);
  }
  const money = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
  for (const item of result.matches.values()) {
    const status = item.timingPending ? "Review ATM cutoff" : item.difference ? "Difference remains" : "Balances agree";
    const values = [item.coverName, item.account, ...[item.ending.value, item.expected, item.difference].map(value => money.format(value / 100)), status];
    const row = body.insertRow();
    for (const value of values) row.insertCell().textContent = value;
  }
  for (const [account, mode] of result.missing) {
    const row = body.insertRow();
    for (const value of [suspenseNames[mode] ?? suspenseAccounts[mode].name, account, "", "", "", "Missing balance"]) row.insertCell().textContent = value;
  }
  spField("batch-summary").replaceChildren(table);
  spField("batch-results").hidden = false;
}

spField("batch-balances").addEventListener("input", invalidateBatch);
spField("month").addEventListener("input", invalidateBatch);
spField("batch-source").addEventListener("input", () => {
  const workbook = spField("batch-source").value === "workbook";
  spField("batch-workbook").hidden = !workbook;
  spField("batch-paste").hidden = workbook;
  invalidateBatch();
});
spField("batch-sheet").addEventListener("change", () => {
  invalidateBatch();
  suggestSuspenseBatchColumns();
});
for (const id of ["batch-account-column", "batch-balance-column", "batch-first", "batch-last"]) {
  spField(id).addEventListener("input", invalidateBatch);
}
const batchFile = spField("batch-file"), batchDrop = spField("batch-drop");
batchFile.addEventListener("change", () => selectSuspenseBatchFile(batchFile.files[0]));
batchDrop.addEventListener("dragover", event => { event.preventDefault(); batchDrop.classList.add("active"); });
batchDrop.addEventListener("dragleave", () => batchDrop.classList.remove("active"));
batchDrop.addEventListener("drop", event => {
  event.preventDefault();
  batchDrop.classList.remove("active");
  if (event.dataTransfer.files.length !== 1) return batchStatus("Drop one trial balance workbook.", true);
  batchFile.value = "";
  selectSuspenseBatchFile(event.dataTransfer.files[0]);
});
spField("batch-confirm").addEventListener("change", () => {
  spField("batch-download").disabled = batchExporting || !batchResult?.matches.size || !spField("batch-confirm").checked;
});
spField("batch-run").addEventListener("click", async () => {
  invalidateBatch();
  const revision = batchRevision;
  try {
    const month = spField("month").value, text = spField("batch-balances").value;
    requireReport(/^\d{4}-\d{2}$/.test(month), "Choose the reconciliation month.");
    requireReport(suspenseProfile.size, "Load the local G/L profile for the suspense accounts.");
    requireReport(!spField("cover-template").files.length || suspenseTemplate, "Load a valid shared cover or clear the template selection.");
    const workbook = spField("batch-source").value === "workbook";
    requireReport(workbook || (text.trim() && text.length <= 5_000_000), "Paste up to 5 MB of G/L ending balances.");
    const balances = workbook ? await readSuspenseBatchWorkbook(month) : readSuspenseBatch(month, text);
    if (revision !== batchRevision) return;
    batchResult = { month, ...balances };
    showSuspenseBatch(batchResult);
    const { matches, missing, ignored } = batchResult;
    const review = [...matches.values()].filter(item => item.difference || item.timingPending).length;
    const ignoredNote = ignored === undefined ? "" : " · " + ignored + " other G/L rows ignored";
    batchStatus(matches.size + " matched · " + missing.length + " missing · " + review + " need review" + ignoredNote + ".", !matches.size);
  } catch (error) {
    if (revision === batchRevision) batchStatus(error.message, true);
  }
});
spField("batch-download").addEventListener("click", async () => {
  const result = batchResult, button = spField("batch-download"), template = suspenseTemplate;
  if (batchExporting || !result?.matches.size || !spField("batch-confirm").checked) return;
  batchExporting = true;
  button.disabled = true;
  try {
    const audit = reconciliationAudit();
    const zip = new JSZip();
    for (const item of result.matches.values()) {
      const workbook = template ? await fillSuspenseTemplate(item, audit, template) : await buildSuspense(item, audit);
      const fileName = item.coverName.replace(/[<>:"/\\|?*\u0000-\u001F]/g, " ").replace(/\s+/g, " ").trim();
      zip.file(item.account + " " + fileName + " " + item.month + "." + (template?.extension ?? "xlsx"), await workbook.arrayBuffer());
    }
    const blob = await zip.generateAsync({ type: "blob" });
    if (result === batchResult && spField("batch-confirm").checked) downloadSuspense(blob, "Suspense " + result.month + ".zip");
  } catch (error) {
    if (result === batchResult) batchStatus(error.message, true);
  } finally {
    batchExporting = false;
    button.disabled = !batchResult?.matches.size || !spField("batch-confirm").checked;
  }
});
