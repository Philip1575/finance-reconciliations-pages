"use strict";

const checkField = id => document.getElementById(id);
let draftFile, priorFile, draftBook, checkResult;
const checkCoverFiles = {};
let checkRevision = 0, draftRevision = 0;
const draftCheckHeaders = {
  date: ["date", "effective date", "transaction date"],
  code: ["txn", "txn code", "tran code", "transaction code"],
  reference: ["reference", "reference number", "check number", "check reference", "draft number"],
  amount: ["amount", "signed amount", "transaction amount"],
};

function invalidateChecks() {
  checkRevision++;
  checkResult = undefined;
  checkField("results").hidden = true;
  checkField("status").textContent = "";
}

function checkStatus(message, error = false) {
  const status = checkField("status");
  status.textContent = message;
  status.classList.toggle("error", error);
}

async function selectDraft(file) {
  const revision = ++draftRevision;
  draftFile = file;
  draftBook = undefined;
  invalidateChecks();
  checkField("draft-mapping").hidden = true;
  checkField("draft-status").textContent = file ? "Opening " + file.name + "…" : "";
  if (!file) return;
  try {
    const workbook = await openExcel(file, true);
    if (revision !== draftRevision) return;
    draftBook = workbook;
    checkField("draft-sheet").replaceChildren(...Array.from(workbook.sheets.keys(), name => new Option(name, name)));
    try { await suggestDraftCheckColumns(); }
    catch (error) {
      if (revision === draftRevision) checkField("draft-status").textContent = error.message;
    }
    if (revision !== draftRevision) return;
    checkField("draft-mapping").hidden = false;
  } catch (error) {
    if (revision === draftRevision) checkField("draft-status").textContent = error.message;
  }
}

async function suggestDraftCheckColumns() {
  const workbook = draftBook, revision = checkRevision, sheet = checkField("draft-sheet").value;
  if (!workbook) return;
  const status = checkField("draft-status");
  status.textContent = draftFile.name;
  const cells = await readCells(workbook, sheet), headers = new Map();
  if (revision !== checkRevision || draftBook !== workbook || checkField("draft-sheet").value !== sheet) return;
  for (const [address, value] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (!match || Number(match[2]) > 50 || typeof value !== "string") continue;
    const label = value.trim().toLowerCase().replace(/\s+/g, " ");
    const field = Object.keys(draftCheckHeaders).find(key => draftCheckHeaders[key].includes(label));
    if (!field) continue;
    const row = Number(match[2]), columns = headers.get(row) ?? {};
    columns[field] = Object.hasOwn(columns, field) ? "" : match[1];
    headers.set(row, columns);
  }
  const matches = [...headers].filter(([, columns]) => Object.keys(draftCheckHeaders).every(key => columns[key]));
  if (matches.length !== 1) return;
  const [[row, columns]] = matches;
  for (const [field, column] of Object.entries(columns)) checkField(field + "-column").value = column;
  checkField("first-row").value = row + 1;
  status.textContent = draftFile.name + " · Draft history columns detected. Verify the worksheet and row range.";
}

function selectPriorChecks(file) {
  priorFile = file;
  checkField("prior-status").textContent = file?.name ?? "";
  if (file) checkField("empty-prior").checked = false;
  invalidateChecks();
}

function selectCheckCover(kind, file) {
  checkCoverFiles[kind] = file;
  checkField(kind + "-status").textContent = file?.name ?? "";
  if (kind === "closing") {
    checkField("closing-mapping").hidden = !file;
    checkField("balance").disabled = Boolean(file);
  } else {
    checkField("adjustments").disabled = Boolean(file);
  }
  invalidateChecks();
}

async function readCheckClosing(month, file) {
  if (!file) {
    const manual = checkField("balance").value.trim();
    if (manual) return { value: cents(manual), source: "Entered manually" };
    const sheet = checkField("draft-sheet").value;
    const cells = await readCells(draftBook, sheet);
    const header = column => String(cells.get(column + "1") ?? "").trim().replace(/\s+/g, " ").toLowerCase();
    requireReport(/^Acct_Num_\d+$/i.test(sheet) && ["effective date", "txn", "transaction amount", "after balance", "draft number"]
      .every((name, i) => header("EHKLM"[i]) === name),
    "Enter the closing balance or upload a CUSA closing-balance workbook.");
    const rows = [...cells].flatMap(([address, value]) => {
      const row = Number(/^E([1-9]\d*)$/.exec(address)?.[1]);
      return row > 1 && String(value ?? "").trim() ? [{
        row, date: sourceDate(value, draftBook.date1904), code: String(cells.get("H" + row) ?? "").trim().toUpperCase(),
        amount: cents(cells.get("K" + row)), balance: cents(cells.get("L" + row)),
      }] : [];
    }).sort((a, b) => a.row - b.row);
    const closing = rows.find(item => item.date.startsWith(month));
    requireReport(closing && rows.every((item, i) => i === 0
      || (rows[i - 1].date >= item.date && item.balance + rows[i - 1].amount === rows[i - 1].balance)),
    "Check the dates and running balances in the CUSA history, or supply a verified month-end balance.");
    return { value: closing.balance, source: draftFile.name + " · " + sheet + "!L" + closing.row, accountHistory: rows };
  }
  const book = await openExcel(file, true);
  const sheet = checkField("closing-sheet").value.trim() || (book.sheets.size === 1 ? [...book.sheets.keys()][0] : "");
  const cell = checkField("closing-cell").value.trim().toUpperCase();
  const dateColumn = checkField("closing-date").value.trim().toUpperCase();
  const location = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(cell);
  requireReport(book.sheets.has(sheet) && location && columnNumber(location[1]) <= 16384 && Number(location[2]) <= 1048576,
    "Choose the closing-balance worksheet and its verified ending cell.");
  requireReport(!dateColumn || (/^[A-Z]{1,3}$/.test(dateColumn) && columnNumber(dateColumn) <= 16384 && dateColumn !== location[1]),
    "Use a date column different from the balance column.");
  const cells = await readCells(book, sheet);
  if (dateColumn) {
    const date = sourceDate(cells.get(dateColumn + location[2]), book.date1904);
    requireReport(date.slice(0, 7) === month, "The ending balance date is outside the selected month.");
  }
  return { value: cents(cells.get(cell)), source: file.name + " · " + sheet + "!" + cell };
}

function checkMonthLabel(month) {
  return new Intl.DateTimeFormat("en-US", { month: "long", year: "numeric", timeZone: "UTC" })
    .format(new Date(month + "-01T00:00:00Z"));
}

async function readCheckCover(month, file) {
  if (!file) return { adjustments: cents(checkField("adjustments").value), items: [] };
  const book = await openExcel(file, true), formulas = new Map();
  const cells = await readCells(book, "Reconciliation", formulas);
  const period = checkMonthLabel(month);
  requireReport(String(cells.get("A2") ?? "").trim().toLowerCase() === "for " + period.toLowerCase(),
    "Set Reconciliation!A2 to the selected month in the prepared workbook.");
  const text = address => String(cells.get(address) ?? "").trim().toLowerCase();
  const formula = address => String(formulas.get(address) ?? "").replace(/[\s$]/g, "").toUpperCase();
  requireReport(text("A1") === "cashier's check reconciliation" && /^balance per .+ account$/.test(text("A5"))
    && text("A17") === "reconciled balance" && !formulas.has("E5")
    && ["SUM(E13:E16,E5)", "SUM(E5,E13:E16)", "E5+SUM(E13:E16)"].includes(formula("E17"))
    && formula("E18") === "E9-E17", "The prepared Reconciliation sheet differs from the supported cover.");
  const items = [];
  for (let row = 13; row <= 16; row++) {
    const address = "E" + row, raw = cells.get(address);
    if (raw === undefined && !formulas.has(address)) continue;
    requireReport(raw !== undefined, "Recalculate and save the line item at Reconciliation!" + address + ".");
    items.push([address, String(cells.get("A" + row) ?? ""), cents(raw)]);
  }
  const outstanding = book.sheets.has("Checks Outstanding") ? await readCheckOutstandingSheet(book) : undefined;
  const destination = outstanding?.range;
  const linked = destination && formula("E9").replace(/^\+/, "") === "'CHECKSOUTSTANDING'!D" + destination.totalRow
    && ["Date Written", "Account Number", "Check Amount", "Check Number", "Reason"]
      .every((title, i) => String(outstanding.cells.get("BCDEF"[i] + "3") ?? "").trim().toLowerCase() === title.toLowerCase());
  const auditTarget = completionTarget("Reconciliation", cells, formulas, "A26", "A29", "E26", "E27");
  return { adjustments: items.reduce((sum, [, , amount]) => sum + amount, 0), items, file: file.name, template: file,
    destination, outstanding: linked ? outstanding : undefined, auditTarget };
}

async function readDraftChecks(workbook, sheet, columns, first, last, month) {
  const cells = await readCells(workbook, sheet);
  const rows = new Set();
  for (const [address, value] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (match && columns.includes(match[1]) && String(value ?? "").trim()) {
      const row = Number(match[2]);
      if (row >= first && row <= last) rows.add(row);
    }
  }
  return Array.from(rows).sort((a, b) => a - b).flatMap(row => {
    const [dateCell, codeCell, referenceCell, amountCell] = columns.map(column => cells.get(column + row));
    try {
      const date = sourceDate(dateCell, workbook.date1904);
      if (date.slice(0, 7) !== month) return [];
      const code = String(codeCell ?? "").trim().toUpperCase();
      const reference = String(referenceCell ?? "").trim();
      const amount = cents(amountCell);
      return [{ date, code, reference, amount, row, source: "Draft history",
        kind: code === "SU" && amount > 0 ? "issued" : code === "DF" && amount < 0 ? "clearing" : "review" }];
    } catch (error) {
      throw new Error(sheet + " row " + row + ": " + error.message);
    }
  });
}

function draftReviewRows(draft, text, label) {
  const updates = new Map();
  const byRow = new Map(draft.map(item => [item.row, item]));
  for (const [index, line] of text.trim().split(/\r?\n/).entries()) {
    if (!line.trim()) continue;
    const [rowText, value, ...extra] = line.split("\t").map(value => value.trim());
    if (index === 0 && /^(?:source|worksheet) row$/i.test(rowText)) continue;
    const row = Number(rowText);
    requireReport(/^[1-9]\d*$/.test(rowText) && Number.isSafeInteger(row) && value && !extra.length,
      label + " row " + (index + 1) + " needs a draft worksheet row number and a reviewed value.");
    requireReport(!updates.has(row), "More than one " + label.toLowerCase() + " was entered for draft row " + row + ".");
    requireReport(byRow.has(row), label + " row " + row + " is outside the selected draft history range or month.");
    updates.set(row, { item: byRow.get(row), value });
  }
  return updates.values();
}

async function readCheckOutstandingSheet(workbook) {
  const formulas = new Map();
  const cells = await readCells(workbook, "Checks Outstanding", formulas);
  const ranges = [...formulas].flatMap(([address, expression]) => {
    const row = Number(/^D([1-9]\d*)$/.exec(address)?.[1]);
    if (!row || !/^total\b/i.test(String(cells.get("B" + row) ?? "").trim()) || String(cells.get("E" + row) ?? "").trim()) return [];
    const match = /^SUM\(D([1-9]\d*):D([1-9]\d*)\)$/i.exec(String(expression).replace(/[\s$]/g, ""));
    if (!match) return [];
    const [first, last] = match.slice(1).map(Number);
    requireReport(first >= 4 && first <= last && last < row, "The Checks Outstanding total has an invalid range.");
    const saved = cells.get(address);
    return [{ first, last, totalRow: row, savedTotal: typeof saved === "number" ? cents(saved) : undefined }];
  });
  requireReport(ranges.length <= 1, "The Checks Outstanding worksheet has several possible total ranges.");
  return { cells, formulas, range: ranges[0] };
}

async function readPriorChecks(file, month) {
  if (!file) return { entries: [], outside: [] };
  const workbook = await openExcel(file, true);
  if (workbook.sheets.has("Reconciliation")) {
    const cover = await readCells(workbook, "Reconciliation");
    if (/^cashier's check reconciliation(?: · proposed)?$/i.test(String(cover.get("A1") ?? "").trim())) {
      const priorMonth = dateOffset(month + "-01", -1).slice(0, 7);
      const period = String(cover.get("A2") ?? "").trim().toLowerCase();
      requireReport([priorMonth, checkMonthLabel(priorMonth).toLowerCase()].some(value => period === "for " + value),
        "The opening list must come from the " + checkMonthLabel(priorMonth) + " reconciliation. Check the period on its cover.");
    }
  }
  const { cells, formulas, range } = await readCheckOutstandingSheet(workbook);
  const entries = [], outside = [];
  for (const [address, value] of cells) {
    const match = /^D([1-9]\d*)$/.exec(address);
    if (!match || Number(match[1]) < 4) continue;
    const row = Number(match[1]);
    const values = ["B", "C", "E", "F"].map(column => cells.get(column + row));
    if (formulas.has(address) && values.slice(1).every(value => !String(value ?? "").trim())
      && (!String(values[0] ?? "").trim() || /^total\b/i.test(String(values[0]).trim()))) continue;
    const [rawDate, account, reference, reason] = values;
    const amount = cents(value);
    if (amount === 0) continue;
    let date;
    try { date = sourceDate(rawDate, workbook.date1904); } catch { date = ""; }
    const item = {
      date, written: date || String(rawDate ?? ""), account: String(account ?? ""), reason: String(reason ?? ""),
      reference: String(reference ?? "").trim(), amount, row, source: "Prior outstanding", kind: "prior", code: "",
    };
    if (range && (row < range.first || row > range.last)) {
      if (item.reference || item.date) outside.push(item);
    } else entries.push(item);
  }
  return { entries, outside, range };
}

function reconcileChecks(prior, draft) {
  const groups = new Map();
  const outstanding = new Set(prior);
  const matches = [];
  const exceptions = [];
  for (const item of [...prior, ...draft]) {
    if (item.disposition === "exclude") continue;
    if (item.disposition === "carry") outstanding.add(item);
    if (!item.reference) {
      if (item.disposition !== "carry") exceptions.push([item, "Missing check reference"]);
      continue;
    }
    if (!groups.has(item.reference)) groups.set(item.reference, []);
    groups.get(item.reference).push(item);
  }
  for (const group of groups.values()) {
    if (group.some(item => item.disposition === "carry")) {
      for (const item of group) {
        if (item.disposition !== "carry") exceptions.push([item, "Reference also has a reviewed carryover"]);
      }
      continue;
    }
    const opening = group.filter(item => item.kind === "prior");
    const issued = group.filter(item => item.kind === "issued");
    const cleared = group.filter(item => item.kind === "clearing");
    const written = [...opening, ...issued];
    const repeated = new Set(group.map(item => JSON.stringify([item.source, item.date, item.code, item.amount]))).size !== group.length;
    const mixedIssue = new Set(opening.map(item => JSON.stringify([item.date, item.account]))).size > 1 || (opening.length && issued.length)
      || new Set(issued.map(item => item.date)).size > 1;
    const unresolved = group.some(item => item.kind === "review" || !item.date
      || (item.kind === "prior" && item.amount < 0)) || repeated || mixedIssue;
    if (!unresolved && written.length && !cleared.length) {
      issued.forEach(item => outstanding.add(item));
      continue;
    }
    const total = group.reduce((sum, item) => sum + item.amount, 0);
    const chronological = cleared.every(clear => written.every(issue => clear.date >= issue.date));
    if (!unresolved && written.length && cleared.length && chronological && total === 0) {
      opening.forEach(item => outstanding.delete(item));
      matches.push(group);
      continue;
    }
    const reason = unresolved ? "Review duplicate, date, code or signed carryover"
      : !written.length ? "No matching open check" : !chronological ? "Clearing date precedes issue date" : "Amounts do not balance";
    for (const item of group) exceptions.push([item, reason]);
  }
  return { outstanding: [...outstanding], matches, exceptions };
}

function suggestCheckFunding(draft, accountHistory, month) {
  const earlier = (accountHistory ?? []).filter(item => item.date < month + "-01");
  const activity = [...earlier, ...draft.filter(item => item.disposition !== "exclude")];
  const funding = new Map(), debits = new Map();
  for (const item of activity) {
    const index = item.code === "JV" && item.amount > 0 ? funding
      : ["DF", "JV"].includes(item.code) && item.amount < 0 ? debits : undefined;
    if (!index) continue;
    const amount = Math.abs(item.amount), rows = index.get(amount) ?? [];
    rows.push(item);
    index.set(amount, rows);
  }
  return draft.flatMap(item => {
    const amount = -item.amount, candidates = funding.get(amount) ?? [];
    const matching = candidates.filter(candidate => candidate.date <= item.date);
    if (debits.get(amount)?.length !== 1 || debits.get(amount)[0] !== item || matching.length !== 1) return [];
    return [{ debit: item, funding: matching[0], type: item.code === "DF" ? "Possible funding for draft clearing" : "Possible journal offset" }];
  });
}

function displayChecks(result) {
  const table = document.createElement("table");
  const currency = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
  const values = [
    ["Opening checks in the prior total", result.priorCount],
    ["Outstanding paste begins at", "B" + result.pasteRange.first],
    ["Matched check references", result.matches.length],
    ["Rows needing review", result.exceptions.length],
    ["Proposed outstanding", currency.format(result.total / 100)],
    ["Closing balance plus reviewed line items", currency.format(result.adjusted / 100)],
    ["Proposed difference", currency.format(result.difference / 100)],
  ];
  const corrected = result.history.filter(item => item.originalReference !== undefined).length;
  if (result.fundingReview.length) values.push(["Journal funding or offset candidates", result.fundingReview.length]);
  if (corrected) values.push(["Reviewed reference corrections applied", corrected]);
  const carry = result.history.filter(item => item.disposition === "carry").length;
  if (carry) values.push(["Reviewed issues carried forward", carry]);
  const excluded = result.history.filter(item => item.disposition === "exclude").length;
  if (excluded) values.push(["Draft rows excluded by review", excluded]);
  const unresolvedIssued = result.exceptions.filter(([item]) => item.kind === "issued").length;
  if (unresolvedIssued) values.push(["New issues awaiting review before carryover", unresolvedIssued]);
  if (result.pasteRange.overflow) values.push(["Destination total", "Extend the total to cover every new row"]);
  const body = table.createTBody();
  for (const [label, value] of values) {
    const row = body.insertRow();
    const heading = document.createElement("th");
    heading.scope = "row";
    heading.textContent = label;
    row.append(heading);
    row.insertCell().textContent = value;
  }
  checkField("summary").replaceChildren(table);
  checkField("download-template").hidden = !result.cover.template;
  const note = checkField("template-download-note");
  note.hidden = !result.cover.template;
  note.textContent = result.cover.outstanding && !result.pasteRange.overflow
    ? "The filled copy updates the cover and replaces the outstanding list. Use the review Excel for draft history and exceptions."
    : "The filled copy updates the cover. Use the review Excel to paste outstanding rows and verify the total range.";
  checkField("results").hidden = false;
}

async function buildChecks(result) {
  const audit = reconciliationAudit();
  const { label, number, date, row, worksheet } = excelOutputCells();
  const formula = (address, expression, value) =>
    `<c r="${address}" s="2"><f>${xmlText(expression)}</f><v>${value / 100}</v></c>`;
  const { first, end, totalRow } = result.pasteRange;
  const summary = [
    row(1, label("A1", "Cashier's Check Reconciliation · proposed", 1)),
    row(2, label("A2", "For " + result.month)),
    row(5, label("A5", "Closing account balance") + number("E5", result.balance / 100)),
    row(9, label("A9", "Proposed checks outstanding") + formula("E9", "'Checks Outstanding'!D" + totalRow, result.total)),
    row(13, label("A13", "Reviewed line item total") + number("E13", result.adjustments / 100)),
    row(17, label("A17", "Reconciled balance") + formula("E17", "E5+E13", result.adjusted)),
    row(18, label("A18", "Proposed difference") + formula("E18", "E9-E17", result.difference)),
    row(21, label("A21", "Draft history") + label("B21", result.draft)),
    row(22, label("A22", "Prior outstanding") + label("B22", result.prior)),
    row(23, label("A23", "Closing balance source") + label("B23", result.closingSource)),
    ...(result.cover.file ? [row(24, label("A24", "Prepared reconciliation") + label("B24", result.cover.file))] : []),
    row(25, label("A25", "First outstanding paste row") + number("B25", first)),
    row(26, label("A26", "Opening outstanding total") + number("B26", result.openingTotal / 100)),
    ...(result.priorSaved !== undefined ? [row(27, label("A27", "Saved prior total") + number("B27", result.priorSaved / 100))] : []),
    row(29, label("A29", "Completed by") + label("E29", audit.by)),
    row(30, label("A30", "Completed on") + date("E30", audit.on)),
  ].join("");
  const checkList = (title, items, end, totalRow, total) => [
    row(1, label("B1", title, 1)),
    row(3, label("B3", "Date Written", 1) + label("C3", "Account Number", 1)
      + label("D3", "Check Amount", 1) + label("E3", "Check Number", 1) + label("F3", "Reason", 1)),
    ...items.map((item, i) => {
      const r = i + first;
      return row(r, date("B" + r, item.written ?? item.date) + label("C" + r, item.account ?? "")
        + number("D" + r, item.amount / 100) + label("E" + r, item.reference) + label("F" + r, item.reason ?? ""));
    }),
    row(totalRow, label("B" + totalRow, "Total") + formula("D" + totalRow, "SUM(D" + first + ":D" + end + ")", total)),
  ].join("");
  const outstanding = checkList("Proposed checks outstanding", result.outstanding, end, totalRow, result.total);
  const openingEnd = Math.max(first, first + result.opening.length - 1);
  const opening = checkList("Opening checks from the prior month", result.opening, openingEnd, openingEnd + 1, result.openingTotal);
  const columnHeaders = (names, columns, r = 1) => names.map((name, i) => label(columns[i] + r, name, 1)).join("");
  const includedHistory = result.history.filter(item => item.disposition !== "exclude");
  const history = [
    row(1, label("A1", "Draft history · " + result.month, 1)),
    row(2, columnHeaders(["Date", "Reference", "Txn", "Amount"], "ABCD", 2)
      + columnHeaders(["Date", "Reference", "Txn", "Amount"], "FGHI", 2)),
    ...includedHistory.map((item, i) => {
      const r = i + 3;
      const columns = item.amount >= 0 ? "ABCD" : "FGHI";
      return row(r, date(columns[0] + r, item.date) + label(columns[1] + r, item.reference)
        + label(columns[2] + r, item.code) + number(columns[3] + r, item.amount / 100));
    }),
  ].join("");
  function auditRows(entries) {
    return row(1, columnHeaders(["Source", "Source row", "Date", "Reference", "Code", "Amount", "Result"], "ABCDEFG"))
      + entries.map(([item, outcome], i) => {
        const r = i + 2;
        return row(r, label("A" + r, item.source) + number("B" + r, item.row) + label("C" + r, item.written ?? item.date)
          + label("D" + r, item.reference) + label("E" + r, item.code) + number("F" + r, item.amount / 100) + label("G" + r, outcome));
      }).join("");
  }
  const matched = result.matches.flatMap((group, i) => group.map(item => [item, "Match " + (i + 1)]));
  const pasteRows = [
    row(1, columnHeaders(["Sheet", "Cell", "Input", "Value", "Source"], "ABCDE")),
    row(2, label("A2", "Reconciliation") + label("B2", "E5") + label("C2", "Closing account balance")
      + number("D2", result.balance / 100) + label("E2", result.closingSource)),
  ];
  const lineRows = [row(1, columnHeaders(["Reviewed cover cell", "Description", "Amount"], "ABC")),
    ...result.cover.items.map(([address, description, amount], i) => {
      const r = i + 2;
      return row(r, label("A" + r, address) + label("B" + r, description) + number("C" + r, amount / 100));
    })];
  const corrections = result.history.filter(item => item.originalReference !== undefined);
  const correctionRows = [row(1, columnHeaders(["Draft worksheet", "Source row", "Date", "Original reference", "Reviewed reference"], "ABCDE")),
    ...corrections.map((item, i) => {
      const r = i + 2;
      return row(r, label("A" + r, result.draftSheet) + number("B" + r, item.row) + date("C" + r, item.date)
        + label("D" + r, item.originalReference) + label("E" + r, item.reference));
    })];
  const decisions = result.history.filter(item => item.disposition).map(item => [item,
    item.disposition === "carry" ? "Carry as outstanding" : "Exclude from this reconciliation"]);
  const fundingRows = [row(1, columnHeaders(["Review", "Debit row", "Debit date", "Txn", "Funding row", "Funding date", "Amount"], "ABCDEFG")),
    ...result.fundingReview.map(({ debit, funding, type }, index) => {
      const r = index + 2;
      return row(r, label("A" + r, type) + number("B" + r, debit.row) + date("C" + r, debit.date) + label("D" + r, debit.code)
        + number("E" + r, funding.row) + date("F" + r, funding.date) + number("G" + r, funding.amount / 100));
    })];
  return createExcelWorkbook([
    ["Reconciliation", worksheet("A1:E30", [[1, 1, 38], [2, 4, 20], [5, 5, 18]], summary)],
    ["Paste cells", worksheet("A1:E2", [[1, 1, 22], [2, 2, 10], [3, 3, 25], [4, 4, 16], [5, 5, 45]], pasteRows.join(""))],
    ...completionPasteSheets(result.cover.auditTarget, audit),
    ...(result.cover.file ? [["Reviewed cover items", worksheet("A1:C" + lineRows.length, [[1, 1, 22], [2, 2, 45], [3, 3, 16]], lineRows.join(""))]] : []),
    ["Opening checks", worksheet("B1:F" + (openingEnd + 1), [[2, 5, 19], [6, 6, 50]], opening)],
    ["Checks Outstanding", worksheet("B1:F" + totalRow, [[2, 5, 19], [6, 6, 50]], outstanding)],
    ["CC Worksheet", worksheet("A1:I" + (includedHistory.length + 2), [[1, 9, 17]], history)],
    ["Matches", worksheet("A1:G" + (matched.length + 1), [[1, 6, 18], [7, 7, 45]], auditRows(matched))],
    ["Exceptions", worksheet("A1:G" + (result.exceptions.length + 1), [[1, 6, 18], [7, 7, 48]], auditRows(result.exceptions))],
    ...(result.fundingReview.length ? [["Journal funding review", worksheet("A1:G" + fundingRows.length,
      [[1, 1, 38], [2, 7, 16]], fundingRows.join(""))]] : []),
    ...(corrections.length ? [["Reference corrections", worksheet("A1:E" + correctionRows.length,
      [[1, 1, 24], [2, 3, 15], [4, 5, 22]], correctionRows.join(""))]] : []),
    ...(decisions.length ? [["Reviewed draft decisions", worksheet("A1:G" + (decisions.length + 1),
      [[1, 6, 18], [7, 7, 43]], auditRows(decisions))]] : []),
  ], audit);
}

for (const [kind, select] of [["draft", selectDraft], ["prior", selectPriorChecks],
  ["closing", file => selectCheckCover("closing", file)], ["cover", file => selectCheckCover("cover", file)]]) {
  const input = checkField(kind + "-file");
  const drop = checkField(kind + "-drop");
  input.addEventListener("change", () => select(input.files[0]));
  drop.addEventListener("dragover", event => { event.preventDefault(); drop.classList.add("active"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("active"));
  drop.addEventListener("drop", event => {
    event.preventDefault();
    drop.classList.remove("active");
    if (event.dataTransfer.files.length !== 1) return checkStatus("Drop one workbook at a time.", true);
    input.value = "";
    select(event.dataTransfer.files[0]);
  });
}
document.addEventListener("dragover", event => event.preventDefault());
document.addEventListener("drop", event => event.preventDefault());
for (const input of document.querySelectorAll("input:not([type=file]), select, textarea")) input.addEventListener("input", invalidateChecks);
checkField("draft-sheet").addEventListener("change", () => {
  suggestDraftCheckColumns().catch(error => { checkField("draft-status").textContent = error.message; });
});

checkField("run").addEventListener("click", async () => {
  invalidateChecks();
  const revision = checkRevision;
  try {
    const month = checkField("month").value;
    requireReport(month && draftBook, "Choose the month and load the draft history.");
    requireReport(Boolean(priorFile) !== checkField("empty-prior").checked, "Load the prior workbook or confirm that the prior outstanding list is empty.");
    const columns = ["date", "code", "reference", "amount"].map(name => checkField(name + "-column").value.trim().toUpperCase());
    requireReport(columns.every(column => /^[A-Z]{1,3}$/.test(column) && columnNumber(column) <= 16384)
      && new Set(columns).size === 4, "Enter four different Excel column letters.");
    const first = checkField("first-row").valueAsNumber;
    const last = checkField("last-row").value ? checkField("last-row").valueAsNumber : 1048576;
    requireReport(Number.isInteger(first) && Number.isInteger(last) && first >= 1 && last >= first && last <= 1048576,
      "Check the first and last data rows.");
    checkStatus("Comparing the draft history and prior checks…");
    const draftSheet = checkField("draft-sheet").value;
    const [draft, prior, closing, cover] = await Promise.all([
      readDraftChecks(draftBook, draftSheet, columns, first, last, month),
      readPriorChecks(priorFile, month),
      readCheckClosing(month, checkCoverFiles.closing),
      readCheckCover(month, checkCoverFiles.cover),
    ]);
    if (revision !== checkRevision) return;
    for (const { item, value } of draftReviewRows(draft, checkField("reference-corrections").value, "Correction")) {
      item.originalReference = item.reference;
      item.reference = value;
    }
    for (const { item, value } of draftReviewRows(draft, checkField("draft-decisions").value, "Decision")) {
      const action = value.toLowerCase();
      requireReport(action === "carry" || action === "exclude", "Use Carry or Exclude for draft row " + item.row + ".");
      requireReport(action !== "carry" || item.kind === "issued", "Only a positive SU issue can be carried from draft row " + item.row + ".");
      item.disposition = action;
      if (action === "carry") item.reason = "Reviewed carryover";
    }
    requireReport(draft.length || prior.entries.length || prior.outside.length, "There are no checks for the selected period.");
    requireReport(prior.entries.every(item => !item.date || item.date.slice(0, 7) < month), "The prior outstanding workbook contains checks dated in or after the selected month.");
    const comparison = reconcileChecks(prior.entries, draft);
    const fundingReview = suggestCheckFunding(draft, closing.accountHistory, month);
    comparison.exceptions.push(...prior.outside.map(item => [item, "Outside the prior workbook's outstanding total range"]));
    const total = comparison.outstanding.reduce((sum, item) => sum + item.amount, 0);
    const openingTotal = prior.entries.reduce((sum, item) => sum + item.amount, 0), priorSaved = prior.range?.savedTotal;
    const destination = cover.destination ?? prior.range;
    const firstPaste = destination?.first ?? 4, lastCheck = firstPaste + comparison.outstanding.length - 1;
    const end = Math.max(firstPaste, lastCheck, destination?.last ?? 0);
    const pasteRange = { first: firstPaste, end, totalRow: Math.max(end + 1, destination?.totalRow ?? 4),
      overflow: Boolean(destination && lastCheck > destination.last) };
    const { value: balance, source: closingSource } = closing, { adjustments } = cover;
    const adjusted = balance + adjustments, difference = total - adjusted;
    requireReport([total, openingTotal, adjustments, adjusted, difference].every(Number.isSafeInteger), "The totals exceed the supported amount range.");
    checkResult = { ...comparison, month, balance, closingSource, cover, adjustments, total, adjusted, difference, history: draft, fundingReview, opening: prior.entries,
      openingTotal, priorSaved, priorCount: prior.entries.length, pasteRange,
      draft: draftFile.name, draftSheet, prior: priorFile?.name ?? "None confirmed" };
    displayChecks(checkResult);
    checkStatus(priorSaved !== undefined && priorSaved !== openingTotal ? "The saved prior total differs from the imported opening rows. Review before pasting."
      : comparison.exceptions.length ? "Review the exceptions before using the proposed balances." : "Proposed balances are ready for review.");
  } catch (error) {
    if (revision === checkRevision) checkStatus(error.message, true);
  }
});

function fillChecksTemplate(result) {
  const audit = reconciliationAudit(), cover = result.cover;
  const updates = [["Reconciliation", "E5", result.balance / 100]];
  if (cover.auditTarget) {
    const { sheet, initials, date } = cover.auditTarget;
    updates.push([sheet, initials, audit.by], [sheet, date, { date: audit.on }]);
  }
  if (cover.outstanding && !result.pasteRange.overflow) {
    const { range, cells, formulas } = cover.outstanding;
    const entries = result.outstanding.map(item => [
      item.date ? { date: item.date } : item.written || null, item.account || null, item.amount / 100, item.reference || null, item.reason || null,
    ]);
    updates.push(...templateTableUpdates("Checks Outstanding", range, "BCDEF", entries, cells, formulas));
  }
  return fillExcelTemplate(cover.template, updates);
}

async function downloadChecks(button, build, name) {
  const result = checkResult;
  if (!result) return;
  button.disabled = true;
  try {
    const blob = await build(result);
    if (checkResult !== result) return;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = name(result);
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 30000);
  } catch (error) {
    if (checkResult === result) checkStatus(error.message, true);
  } finally {
    button.disabled = false;
  }
}

checkField("download").addEventListener("click", event =>
  downloadChecks(event.currentTarget, buildChecks, result => "Cashiers Checks " + result.month + ".xlsx"));
checkField("download-template").addEventListener("click", event =>
  downloadChecks(event.currentTarget, fillChecksTemplate,
    result => "Cashiers Checks " + result.month + " filled draft." + result.cover.file.split(".").at(-1)));
