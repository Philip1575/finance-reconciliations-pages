"use strict";

const cdField = id => document.getElementById(id);
const cdFiles = { principal: [], receivable: [], interest: [], ledger: [], provider: [], "provider-receivable": [] };
let cdResult, cdRevision = 0, ledgerRevision = 0, cdLedger;
addTrialBalanceDetection("ledger", () => cdLedger ? [{ book: cdLedger }] : []);

function cdStatus(message, error = false) {
  cdField("status").textContent = message;
  cdField("status").classList.toggle("error", error);
}

function invalidateCd() {
  cdRevision++;
  cdResult = undefined;
  cdField("confirm").checked = false;
  cdField("results").hidden = true;
  cdStatus("");
}

async function selectCdFiles(kind, files) {
  cdFiles[kind] = Array.from(files);
  const status = cdField(kind + "-status");
  status.textContent = files.length === 1 ? files[0].name : files.length ? files.length + " files selected" : "";
  invalidateCd();
  if (kind !== "ledger") return;
  const current = ++ledgerRevision;
  cdLedger = undefined;
  cdField("ledger-map").hidden = true;
  if (!files.length) return;
  try {
    const book = await openExcel(files[0], true);
    if (current !== ledgerRevision) return;
    cdLedger = book;
    cdField("ledger-sheet").replaceChildren(...[...book.sheets.keys()].map(name => new Option(name, name)));
    cdField("ledger-map").hidden = false;
  } catch (error) {
    if (current === ledgerRevision) status.textContent = error.message;
  }
}

async function cdRoster(file, kind) {
  const book = await openExcel(file, true);
  const principal = kind === "principal";
  const balanceAddress = principal ? "I8" : "H8";
  const candidates = [];
  for (const name of book.sheets.keys()) {
    const formulas = new Map();
    const cells = await readCells(book, name, formulas);
    const range = /^SUM\(H(\d+):H(\d+)\)$/i.exec(formulas.get(balanceAddress) ?? "");
    if (range && /balance per .+g\/l/i.test(String(cells.get("B8") ?? ""))) candidates.push({ name, cells, formulas, range });
  }
  requireReport(candidates.length === 1, file.name + " does not have one recognized " + kind + " template.");
  const { name, cells, formulas, range } = candidates[0];
  const first = Number(range[1]), last = Number(range[2]);
  requireReport(first > 8 && first <= last && last < 1048576, "The template balance total has an invalid account range.");
  let adjustment;
  if (principal) {
    requireReport((formulas.get("I9") ?? "").toUpperCase() === "SUM(I" + first + ":I" + last + ")",
      "The provider total does not cover the same rows as the ledger total.");
    const difference = [...formulas.values()].map(formula => /^\+?I8-I9-I(\d+)$/i.exec(formula)).find(Boolean);
    requireReport(difference && Number(difference[1]) > last, "The principal difference formula has no recognized line-items cell.");
    adjustment = "I" + difference[1];
  }
  const entries = [];
  for (let row = first; row <= last; row++) {
    const accountValue = cells.get("B" + row), issuer = String(cells.get("C" + row) ?? "").trim();
    if (!String(accountValue ?? "").trim() && !issuer) continue;
    const account = normalizeAccount(accountValue);
    requireReport(account && issuer, file.name + " row " + row + " needs a G/L account and issuer.");
    requireReport(!formulas.has("H" + row) && (!principal || !formulas.has("I" + row)), "Check the balance input formulas in row " + row + ".");
    const saved = cells.get("H" + row);
    entries.push({ row, account, issuer, saved: typeof saved === "number" ? cents(saved) : undefined });
  }
  requireReport(entries.length && new Set(entries.map(entry => entry.account)).size === entries.length, "The template needs distinct G/L accounts.");
  const completion = completionTarget(name, cells, formulas, "F27", "F31", principal ? "I27" : "H27", principal ? "I28" : "H28");
  return { kind, name, file: file.name, template: file, entries, adjustment, completion, periodLabel: String(cells.get("B5") ?? "").trim() };
}

function cdIssuer(value) {
  return String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
}

async function cdProvider(file, kind = "principal") {
  const amountLabels = kind === "principal" ? ["principal balance", "ending principal", "principal"]
    : ["accrued interest", "accrued interest balance", "ending accrued interest", "interest receivable"];
  const amountName = kind === "principal" ? "principal balance" : "accrued interest";
  const choices = [];
  const addTable = (name, cells) => {
    const headers = new Map();
    for (const [address, value] of cells) {
      const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
      if (!match || Number(match[2]) > 50 || typeof value !== "string") continue;
      const row = Number(match[2]), columns = headers.get(row) ?? {};
      const title = cdIssuer(value);
      const key = ["issuer", "institution", "bank name"].includes(title) ? "issuer" : ["g/l", "g/l account", "gl account"].includes(title) ? "account"
        : amountLabels.includes(title) ? "amount"
          : ["position id", "position identifier", "cd number", "certificate number"].includes(title) ? "position" : "";
      if (key) columns[key] = Object.hasOwn(columns, key) ? "" : match[1];
      headers.set(row, columns);
    }
    const header = [...headers].find(([, columns]) => columns.amount && (columns.account || columns.issuer));
    if (header) choices.push({ name, cells, header });
  };
  if (/\.csv$/i.test(file.name)) {
    requireReport(file.size <= 20 * 1024 * 1024, "The CSV exceeds the 20 MB import limit.");
    addTable("CSV", readCsvCells(await file.text()));
  } else {
    const book = await openExcel(file, true);
    for (const name of book.sheets.keys()) addTable(name, await readCells(book, name));
  }
  requireReport(choices.length === 1, file.name + " needs one sheet with issuer or G/L and " + amountName + " headers.");
  const { name, cells, header: [first, columns] } = choices[0];
  const result = [];
  const address = new RegExp("^" + columns.amount + "([1-9]\\d*)$");
  for (const [cell, raw] of cells) {
    const row = Number(address.exec(cell)?.[1]);
    if (!(row > first) || !String(raw ?? "").trim()) continue;
    const rawAccount = columns.account ? String(cells.get(columns.account + row) ?? "").trim() : "";
    const account = normalizeAccount(rawAccount);
    requireReport(!rawAccount || account, file.name + " row " + row + " has an invalid G/L account.");
    const issuer = columns.issuer ? String(cells.get(columns.issuer + row) ?? "").trim() : "";
    const position = columns.position ? String(cells.get(columns.position + row) ?? "").trim() : "";
    requireReport(account || issuer, file.name + " row " + row + " has an amount without an issuer or account.");
    result.push({ file: file.name, source: name + "!" + cell, account, issuer, position, amount: cents(raw) });
  }
  requireReport(result.length, "No " + amountName + " rows were found in " + file.name + ".");
  return result;
}

function cdApplyBalances(roster, ledger, provider) {
  const used = new Set();
  const entries = roster.entries.map(entry => {
    const balance = ledger.get(entry.account);
    let statement;
    const matches = provider.filter(item => item.account ? item.account === entry.account : cdIssuer(item.issuer) === cdIssuer(entry.issuer));
    if (matches.length) {
      requireReport(matches.every(item => !used.has(item)), "Check ambiguous provider balances for " + entry.issuer + ".");
      if (matches.length > 1) {
        const ids = matches.map(item => item.position.toLowerCase());
        requireReport(ids.every(Boolean) && new Set(ids).size === matches.length,
          "Use distinct position IDs for multiple provider rows for " + entry.issuer + ".");
      }
      const amount = matches.reduce((sum, item) => sum + item.amount, 0);
      requireReport(Number.isSafeInteger(amount), "A provider balance exceeds the supported amount range.");
      statement = { amount, positions: matches, file: matches.length === 1 ? matches[0].file : "See Provider " + roster.kind + " rows",
        source: matches.length === 1 ? matches[0].source : matches.length + " positions" };
      matches.forEach(item => used.add(item));
    }
    const ledgerAmount = cents(balance.amount);
    const variance = statement ? ledgerAmount - statement.amount : undefined;
    const change = entry.saved === undefined ? undefined : ledgerAmount - entry.saved;
    requireReport([variance, change].every(value => value === undefined || Number.isSafeInteger(value)), "A balance difference exceeds the supported amount range.");
    return { ...entry, ledger: ledgerAmount, ledgerSource: balance.address, statement, variance, change };
  });
  const balance = entries.reduce((sum, entry) => sum + entry.ledger, 0);
  const missingProvider = entries.filter(entry => !entry.statement);
  const complete = !missingProvider.length;
  const statement = complete ? entries.reduce((sum, entry) => sum + entry.statement.amount, 0) : undefined;
  const variance = complete ? balance - statement : undefined;
  requireReport([balance, statement, variance].every(value => value === undefined || Number.isSafeInteger(value)),
    "A balance total exceeds the supported amount range.");
  return { ...roster, entries, balance, statement, variance, complete, missingProvider,
    hasProvider: Boolean(provider.length), unused: provider.filter(item => !used.has(item)) };
}

function cdProviderCoverage(group) {
  const total = group.entries.length, received = total - group.missingProvider.length;
  return received + " of " + total + " mapped balances received";
}

function cdFormula(text) {
  return String(text ?? "").replace(/\s+/g, "").toUpperCase();
}

function cdMatchesFormula(formulas, column, row, expected) {
  const stored = formulas.get(column + row);
  if (stored && typeof stored === "object") {
    return cdFormula(stored.expression) === expected(Number(stored.sharedFrom.slice(column.length)));
  }
  return cdFormula(stored) === expected(row);
}

async function cdInterest(file, entry) {
  const book = await openExcel(file, true), formulas = new Map();
  const cells = await readCells(book, "CDs", formulas);
  const headers = [["C4", "Purchase Date"], ["D4", "Maturity Date"], ["E4", "Amount"], ["F4", "Interest Rate"]];
  requireReport(headers.every(([cell, text]) => String(cells.get(cell) ?? "").trim() === text), "Check the columns in the CDs schedule.");
  requireReport(cdFormula(formulas.get("Q3")) === "DATE(O3,M3,L3)" && cdFormula(formulas.get("S3")) === "DATE(O3,M3+2,0)",
    "The CDs schedule uses a different entry-date or accrual-end formula.");
  const totals = [...formulas].filter(([cell, formula]) => /^M\d+$/.test(cell) && /^SUM\(M5:M\d+\)$/.test(cdFormula(formula)));
  requireReport(totals.length === 1, "The CDs schedule needs one recognized interest total.");
  const last = Number(cdFormula(totals[0][1]).match(/:M(\d+)/)[1]);
  const [year, month, day] = entry.split("-").map(Number);
  const posting = Date.UTC(year, month - 1, day), nextEnd = Date.UTC(year, month + 1, 0);
  const monthDays = new Date(nextEnd).getUTCDate(), milliseconds = 86400000;
  const rows = [], inactive = [], exceptions = [];
  for (let r = 5; r <= last; r++) {
    const issuer = String(cells.get("B" + r) ?? "").trim();
    const raw = ["C", "D", "E", "F"].map(column => cells.get(column + r));
    if (raw.every(value => !String(value ?? "").trim()) && (!cells.get("M" + r) || cells.get("M" + r) === 0)) continue;
    const [startRaw, endRaw, principalRaw, rateRaw] = raw;
    try {
      const purchase = sourceDate(startRaw, book.date1904), maturity = sourceDate(endRaw, book.date1904);
      const end = Date.parse(maturity + "T00:00:00Z"), start = Date.parse(purchase + "T00:00:00Z");
      requireReport(end > start, "Maturity must be after purchase.");
      if (!formulas.has("M" + r) && cells.get("M" + r) === undefined) {
        requireReport(end <= posting, "The monthly interest cell is blank for a position that has not matured.");
        inactive.push({ row: r, issuer, maturity });
        continue;
      }
      const expected = n => `ROUND(IF($S$3-D${n}>0,MAX(I${n}/G${n}*(D${n}-$Q$3),0),IF($Q$3-C${n}<0,0,IF($M$3=MONTH(C${n}),IF($O$3=YEAR(C${n}),R${n}+(I${n}/G${n})*($Q$3-C${n}),R${n}),R${n}))),2)`;
      requireReport(cdMatchesFormula(formulas, "M", r, expected), "The monthly interest formula differs from the supported template.");
      if (end <= posting || (start > posting && end >= nextEnd)) {
        rows.push({ row: r, issuer, purchase, maturity, days: 0, amount: 0 });
        continue;
      }
      requireReport(start <= posting, "Purchase follows the entry date. Review the partial-month calculation.");
      requireReport(cdMatchesFormula(formulas, "G", r, n => `D${n}-C${n}`)
        && cdMatchesFormula(formulas, "I", r, n => `E${n}*F${n}*G${n}/365`)
        && cdMatchesFormula(formulas, "R", r, n => `I${n}/G${n}*DAY($S$3)`), "The daily interest calculation differs from the supported template.");
      const principal = cents(principalRaw), rate = readBalance(rateRaw);
      requireReport(principal >= 0 && Number.isFinite(rate) && rate >= 0, "Check the principal and annual interest rate.");
      const samePurchaseMonth = purchase.slice(0, 7) === entry.slice(0, 7);
      const days = end < nextEnd ? (end - posting) / milliseconds : monthDays + (samePurchaseMonth ? (posting - start) / milliseconds : 0);
      const amount = Math.round(principal * rate * days / 365 + 1e-8);
      requireReport(Number.isSafeInteger(amount), "The estimated interest exceeds the supported amount range.");
      rows.push({ row: r, issuer, purchase, maturity, days, amount });
    } catch (error) {
      exceptions.push({ row: r, issuer, reason: error.message });
    }
  }
  const total = rows.reduce((sum, row) => sum + row.amount, 0);
  requireReport(Number.isSafeInteger(total), "Interest exceeds the supported amount range.");
  return { file: file.name, template: file, entry, year, month, day, rows, inactive, exceptions, total };
}

function showCd(result) {
  const money = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
  const values = [];
  if (result.principal) values.push(["Principal ledger total", money.format(result.principal.balance / 100)]);
  if (result.principal?.complete) values.push(["Mapped principal difference after reviewed items", money.format(result.difference / 100)],
    ["Unmapped provider rows", result.principal.unused.length]);
  else if (result.principal) values.push(["Principal comparison", "Awaiting provider balances · " + cdProviderCoverage(result.principal)]);
  if (result.receivable) values.push(["Receivable ledger total", money.format(result.receivable.balance / 100)]);
  if (result.receivable?.complete) values.push(["Receivables less provider accrual", money.format(result.receivable.variance / 100)],
    ["Unmapped provider accrued interest rows", result.receivable.unused.length]);
  else if (result.receivable) values.push(["Receivable comparison", "Awaiting provider accrued interest · " + cdProviderCoverage(result.receivable)]);
  const balanceGroups = [result.principal, result.receivable].filter(Boolean);
  if (balanceGroups.length) values.push(["Wisdom inputs differing from saved templates",
    balanceGroups.reduce((count, group) => count + group.entries.filter(entry => entry.change).length, 0)]);
  if (result.interest) values.push(["Calculated interest for supported rows", money.format(result.interest.total / 100)],
    ["Matured rows with blank interest", result.interest.inactive.length],
    ["Interest rows to review", result.interest.exceptions.length]);
  const table = document.createElement("table"), body = table.createTBody();
  for (const [name, value] of values) {
    const row = body.insertRow(), heading = document.createElement("th");
    heading.scope = "row";
    heading.textContent = name;
    row.append(heading);
    row.insertCell().textContent = value;
  }
  cdField("summary").replaceChildren(table);
  for (const kind of ["principal", "receivable", "interest"]) cdField("download-" + kind).hidden = !result[kind];
  cdField("principal-draft-note").hidden = !result.principal || result.principal.complete;
  cdField("results").hidden = false;
}

function cdTemplateUpdates(result, kind, audit) {
  const group = result[kind];
  if (kind === "interest") return [["CDs", "L3", group.day], ["CDs", "M3", group.month], ["CDs", "O3", group.year]];
  const period = new Intl.DateTimeFormat("en-US", { month: "long", year: "numeric", timeZone: "UTC" })
    .format(new Date(result.month + "-01T00:00:00Z"));
  const statementComplete = kind === "principal" && group.complete;
  const updates = [[group.name, "B5", "For the Month of " + period]];
  for (const entry of group.entries) {
    updates.push([group.name, "H" + entry.row, entry.ledger / 100]);
    if (statementComplete) updates.push([group.name, "I" + entry.row, entry.statement.amount / 100]);
  }
  if (group.completion) {
    updates.push([group.name, group.completion.initials, audit.by],
      [group.name, group.completion.date, { date: audit.on }]);
  }
  if (statementComplete) updates.push([group.name, group.adjustment, result.adjustment / 100]);
  return updates;
}

async function buildCd(result) {
  const audit = reconciliationAudit();
  const { label, number, date, row, worksheet } = excelOutputCells();
  function table(name, headings, data, widths) {
    const cols = "ABCDEFGHI".slice(0, headings.length);
    const body = row(1, headings.map((title, i) => label(cols[i] + "1", title, 1)).join(""))
      + data.map((entry, i) => row(i + 2, entry.map((value, j) => {
        const cell = cols[j] + (i + 2);
        return typeof value === "number" ? number(cell, value) : typeof value === "object" ? date(cell, value.date) : label(cell, value);
      }).join(""))).join("");
    return [name, worksheet("A1:" + cols.at(-1) + (data.length + 1),
      widths.map((width, i) => [i + 1, i + 1, width]), body)];
  }
  const summary = [["Period", result.month]];
  const paste = [];
  const sheets = [];
  const savedLedgerRows = [];
  const missingProviderRows = [];
  for (const group of [result.principal, result.receivable].filter(Boolean)) {
    const principal = group.kind === "principal";
    summary.push([group.kind + " template", group.file], [group.kind + " ledger total", group.balance / 100]);
    paste.push(...cdTemplateUpdates(result, group.kind, audit).map(update => [group.file, ...update]));
    for (const entry of group.entries) {
      savedLedgerRows.push([group.file, group.periodLabel, group.name + "!H" + entry.row, entry.issuer,
        entry.saved === undefined ? "" : entry.saved / 100, entry.ledger / 100,
        entry.change === undefined ? "" : entry.change / 100, result.ledgerSheet + "!" + entry.ledgerSource]);
    }
    if (!group.complete) {
      summary.push([principal ? "Principal comparison" : "Receivable comparison",
        (principal ? "Awaiting provider balances" : "Awaiting provider accrued interest") + " · " + cdProviderCoverage(group)]);
      missingProviderRows.push(...group.missingProvider.map(entry => [group.kind, entry.issuer, entry.account, entry.ledger / 100]));
    }
    if (principal && group.complete) {
      summary.push(["Mapped provider principal total", group.statement / 100], ["Reviewed line items", result.adjustment / 100],
        ["Mapped principal difference", result.difference / 100]);
    }
    if (principal && group.hasProvider) {
      summary.push(["Unmapped provider rows", group.unused.length]);
      sheets.push(table("Unmapped provider rows", ["File", "Source", "Issuer", "Account", "Principal"],
        group.unused.map(item => [item.file, item.source, item.issuer, item.account, item.amount / 100]), [35, 23, 38, 18, 20]));
      sheets.push(table("Provider principal rows", ["Template row", "Issuer", "G/L", "Position ID", "Principal", "File", "Source"],
        group.entries.flatMap(entry => (entry.statement?.positions ?? []).map(item => [entry.row, entry.issuer, entry.account,
          item.position, item.amount / 100, item.file, item.source])), [15, 34, 15, 20, 18, 32, 24]));
    }
    if (!principal && group.hasProvider) {
      if (group.complete) summary.push(["Mapped provider accrued interest", group.statement / 100], ["Receivables less provider accrual", group.variance / 100]);
      summary.push(["Unmapped provider accrued interest rows", group.unused.length]);
      sheets.push(table("Receivable comparison", ["Issuer", "G/L", "Wisdom balance", "Provider accrued interest", "Difference", "Provider source"],
        group.entries.filter(entry => entry.statement).map(entry => [entry.issuer, entry.account, entry.ledger / 100, entry.statement.amount / 100,
          entry.variance / 100, entry.statement.file + " · " + entry.statement.source]), [36, 16, 19, 23, 18, 55]));
      sheets.push(table("Provider receivable rows", ["Template row", "Issuer", "G/L", "Position ID", "Accrued interest", "File", "Source"],
        group.entries.flatMap(entry => (entry.statement?.positions ?? []).map(item => [entry.row, entry.issuer, entry.account,
          item.position, item.amount / 100, item.file, item.source])), [15, 34, 15, 20, 18, 32, 24]));
      if (group.unused.length) sheets.push(table("Unmapped accrued interest", ["File", "Source", "Issuer", "Account", "Accrued interest"],
        group.unused.map(item => [item.file, item.source, item.issuer, item.account, item.amount / 100]), [35, 23, 38, 18, 21]));
    }
    sheets.push(table(principal ? "Principal details" : "Receivable details",
      ["Issuer", "G/L", "Template row", "Ledger balance", "Ledger cell", "Provider balance", "Provider file", "Provider cell", "Wisdom less provider"],
      group.entries.map(entry => [entry.issuer, entry.account, entry.row, entry.ledger / 100, result.ledgerSheet + "!" + entry.ledgerSource,
        entry.statement ? entry.statement.amount / 100 : "", entry.statement?.file ?? "", entry.statement?.source ?? "",
        entry.variance === undefined ? "" : entry.variance / 100]), [36, 16, 15, 20, 26, 19, 32, 21, 22]));
  }
  if (missingProviderRows.length) sheets.push(table("Missing provider balances",
    ["Reconciliation", "Issuer", "G/L", "Wisdom balance"], missingProviderRows, [20, 38, 16, 20]));
  if (savedLedgerRows.length) sheets.push(table("Saved ledger comparison",
    ["Template", "Template period", "Destination", "Issuer", "Saved amount", "Wisdom amount", "Change", "Wisdom cell"],
    savedLedgerRows, [29, 28, 19, 35, 18, 18, 18, 20]));
  if (result.interest) {
    const interest = result.interest;
    summary.push(["Interest template", interest.file], ["Interest entry date", interest.entry],
      ["Calculated interest for supported rows", interest.total / 100], ["Interest rows to review", interest.exceptions.length],
      ["Matured rows with blank interest", interest.inactive.length]);
    paste.push(...cdTemplateUpdates(result, "interest", audit).map(update => [interest.file, ...update]));
    sheets.push(table("Interest review", ["CDs row", "Issuer", "Purchase", "Maturity", "Accrual days", "Calculated interest"],
      interest.rows.map(item => [item.row, item.issuer, item.purchase, item.maturity, item.days, item.amount / 100]), [12, 38, 16, 16, 16, 24]));
    sheets.push(table("Interest exceptions", ["CDs row", "Issuer", "Review"],
      interest.exceptions.map(item => [item.row, item.issuer, item.reason]), [12, 38, 62]));
    if (interest.inactive.length) sheets.push(table("Matured blank interest", ["CDs row", "Issuer", "Maturity"],
      interest.inactive.map(item => [item.row, item.issuer, item.maturity]), [12, 38, 16]));
  }
  return createExcelWorkbook([table("Summary", ["Field", "Value"], summary, [45, 43]),
    table("Paste cells", ["Workbook", "Worksheet", "Cell", "Value"], paste, [36, 20, 14, 34]), ...sheets], audit);
}

for (const kind of Object.keys(cdFiles)) {
  const input = cdField(kind + "-file"), drop = cdField(kind + "-drop");
  input.addEventListener("change", () => selectCdFiles(kind, input.files));
  drop.addEventListener("dragover", event => { event.preventDefault(); drop.classList.add("active"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("active"));
  drop.addEventListener("drop", event => {
    event.preventDefault();
    drop.classList.remove("active");
    if (!event.dataTransfer.files.length || (!kind.startsWith("provider") && event.dataTransfer.files.length !== 1)) return cdStatus("Drop one workbook for " + kind + ".", true);
    input.value = "";
    selectCdFiles(kind, event.dataTransfer.files);
  });
}
document.addEventListener("dragover", event => event.preventDefault());
document.addEventListener("drop", event => event.preventDefault());
cdField("month").addEventListener("input", () => {
  const month = cdField("month").value;
  if (!month) return;
  const [year, m] = month.split("-").map(Number);
  cdField("entry").value = new Date(Date.UTC(year, m, 0)).toISOString().slice(0, 10);
});
for (const input of document.querySelectorAll("input:not([type=file]):not(#confirm), select")) input.addEventListener("input", invalidateCd);
cdField("confirm").addEventListener("change", () => { if (!cdField("confirm").checked) invalidateCd(); });

cdField("run").addEventListener("click", async () => {
  const current = ++cdRevision;
  cdResult = undefined;
  cdField("results").hidden = true;
  try {
    const month = cdField("month").value, entry = cdField("entry").value;
    requireReport(month && cdField("confirm").checked, "Choose the month and verify the source inputs.");
    requireReport(cdFiles.principal.length || cdFiles.receivable.length || cdFiles.interest.length, "Load at least one monthly template.");
    requireReport(!cdFiles.interest.length || entry.slice(0, 7) === month, "Choose an interest entry date in the selected month.");
    const [principal, receivable, interest] = await Promise.all([
      cdFiles.principal.length ? cdRoster(cdFiles.principal[0], "principal") : undefined,
      cdFiles.receivable.length ? cdRoster(cdFiles.receivable[0], "receivable") : undefined,
      cdFiles.interest.length ? cdInterest(cdFiles.interest[0], entry) : undefined,
    ]);
    if (current !== cdRevision) return;
    let p, r, adjustment, difference, ledgerSheet;
    if (principal || receivable) {
      requireReport(cdLedger, "Load the Wisdom trial balance for the balance templates.");
      ledgerSheet = cdField("ledger-sheet").value;
      const accounts = [...new Set([principal, receivable].filter(Boolean).flatMap(group => group.entries.map(entry => entry.account)))];
      const [ledger, provider, accrued] = await Promise.all([
        readWisdom(cdLedger, ledgerSheet, cdField("ledger-account").value.trim().toUpperCase(), cdField("ledger-balance").value.trim().toUpperCase(),
          cdField("ledger-first").valueAsNumber, accounts, cdField("ledger-last").value ? cdField("ledger-last").valueAsNumber : 1048576,
          true, month),
        principal ? Promise.all(cdFiles.provider.map(file => cdProvider(file))).then(groups => groups.flat()) : [],
        receivable ? Promise.all(cdFiles["provider-receivable"].map(file => cdProvider(file, "receivable"))).then(groups => groups.flat()) : [],
      ]);
      if (current !== cdRevision) return;
      if (principal) {
        p = cdApplyBalances(principal, ledger, provider);
        if (p.complete) {
          adjustment = cents(cdField("adjustment").value);
          difference = p.balance - p.statement - adjustment;
          requireReport(Number.isSafeInteger(difference), "The principal difference exceeds the supported amount range.");
        }
      }
      if (receivable) r = cdApplyBalances(receivable, ledger, accrued);
    }
    cdResult = { month, principal: p, receivable: r, interest, adjustment, difference, ledgerSheet };
    showCd(cdResult);
    cdStatus("");
  } catch (error) {
    if (current === cdRevision) cdStatus(error.message, true);
  }
});

async function downloadCd(button, kind) {
  const result = cdResult;
  if (!result) return;
  button.disabled = true;
  try {
    const group = result[kind];
    const blob = group
      ? await fillExcelTemplate(group.template, cdTemplateUpdates(result, kind, reconciliationAudit()))
      : await buildCd(result);
    if (cdResult !== result) return;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    const draft = kind === "principal" && !group.complete ? " draft" : "";
    link.download = group
      ? "CD " + kind + " " + result.month + draft + "." + group.file.split(".").at(-1)
      : "CD investments " + result.month + ".xlsx";
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 30000);
  } catch (error) {
    if (cdResult === result) cdStatus(error.message, true);
  } finally {
    button.disabled = false;
  }
}

cdField("download").addEventListener("click", event => downloadCd(event.currentTarget));
for (const kind of ["principal", "receivable", "interest"]) {
  cdField("download-" + kind).addEventListener("click", event => downloadCd(event.currentTarget, kind));
}
