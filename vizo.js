"use strict";

const vizoField = id => document.getElementById(id);
const vizoFiles = { provider: { revision: 0 }, ledger: { revision: 0 }, statement: { revision: 0 }, if67: { revision: 0 },
  loanstreet: { revision: 0 }, invoice: { revision: 0 }, transit: { revision: 0 }, "atm-opening": { revision: 0 }, boundary: { revision: 0 }, template: { revision: 0 } };
let vizoResult;
let vizoRevision = 0;
let relatedVizoFiles = [];
addTrialBalanceDetection("ledger", () => vizoFiles.ledger.book ? [vizoFiles.ledger] : []);

function vizoStatus(message, error = false) {
  const status = vizoField("status");
  status.textContent = message;
  status.classList.toggle("error", error);
}

function invalidateVizo() {
  vizoRevision++;
  vizoResult = undefined;
  vizoField("results").hidden = true;
  vizoField("confirm").checked = false;
  vizoStatus("");
}

async function selectVizoFile(kind, file) {
  const state = vizoFiles[kind];
  const current = ++state.revision;
  state.file = file;
  state.book = undefined;
  state.csv = undefined;
  state.nativeProvider = undefined;
  state.nativeWisdom = undefined;
  if (kind === "provider" || kind === "boundary") vizoField("deposit-allocations").value = "";
  if (kind === "provider") vizoField("overnight-empty").checked = false;
  invalidateVizo();
  const status = vizoField(kind + "-status");
  if (["invoice", "transit", "boundary", "template"].includes(kind)) {
    status.textContent = file?.name ?? "";
    if (kind === "transit") showVizoAccount();
    return;
  }
  vizoField(kind + "-map").hidden = true;
  status.textContent = file ? "Opening " + file.name + "…" : "";
  if (kind === "provider" && vizoField("statement-mode").value === "running") {
    vizoField("statement-mode").value = "manual";
    showStatementSource();
  }
  if (kind === "ledger" && vizoField("ledger-type").value === "wisdom") {
    vizoField("ledger-type").value = vizoAccounts[vizoField("mode").value].ledger;
    showVizoAccount();
  }
  if (!file) return;
  try {
    if (kind === "provider" && /\.csv$/i.test(file.name)) {
      requireReport(file.size <= 20 * 1024 * 1024, "The CSV exceeds the 20 MB import limit.");
      const csv = readCsvCells(await file.text());
      if (current !== state.revision) return;
      state.csv = csv;
      state.nativeProvider = vizoNativeProvider(csv);
      if (state.nativeProvider) {
        for (const [field, value] of Object.entries({ date: "C", desc: "E", amount: "G", first: state.nativeProvider.first, last: "" })) {
          vizoField("provider-" + field).value = value;
        }
        vizoField("statement-mode").value = "running";
        showStatementSource();
      }
    } else {
      const book = await openExcel(file, true);
      if (current !== state.revision) return;
      state.book = book;
      if (kind === "ledger" && book.sheets.size === 1) {
        const sheet = [...book.sheets.keys()][0];
        const cells = await readCells(book, sheet);
        if (current !== state.revision) return;
        const layout = wisdomHistoryLayout(cells);
        if (layout) {
          state.nativeWisdom = { ...layout, ...wisdomHistory(cells, book.date1904, layout) };
          vizoField("ledger-type").value = "wisdom";
          showVizoAccount();
        }
      }
    }
    const names = state.csv ? ["CSV"] : [...state.book.sheets.keys()];
    vizoField(kind + "-sheet").replaceChildren(...names.map(name => new Option(name, name)));
    vizoField(kind + "-map").hidden = false;
    status.textContent = file.name + (state.nativeProvider ? " · VIZO Transaction Summary detected"
      : state.nativeWisdom ? " · Wisdom Detailed Trial Balance detected" : "");
  } catch (error) {
    if (current === state.revision) {
      state.book = state.csv = undefined;
      status.textContent = error.message;
    }
  }
}

async function readVizoTransit(file, mode, month) {
  const book = await openExcel(file, true), formulas = new Map();
  const sheet = vizoAccounts[mode].paste.sheet ?? "Reconciliation";
  const cells = await readCells(book, sheet, formulas);
  const label = address => String(cells.get(address) ?? "").trim();
  const formula = address => String(formulas.get(address) ?? "").replace(/[\s$]/g, "").toUpperCase();
  const layout = mode === "misc" ? /^vizo misc$/i : /^vizo settlement account$/i;
  const expectedTotal = mode === "misc" ? formula("L46") === "SUM(L24:L45)"
    : formula("L47") === "SUM(L28:L45)+L24+L25";
  requireReport(layout.test(label("C11")) && /^in transit items$/i.test(label("C26")) && expectedTotal
    && ["Line Item", "Item Description", "Will Clear On", "Date", "Amount"].every((name, i) => label("DFIKL"[i] + "27") === name),
  "Use the existing in-transit layout for the selected VIZO account.");
  requireReport(mode === "misc" || !label("L25") || cents(cells.get("L25")) === 0,
    "The Settlement cover has another adjustment at L25. Review it before importing the open items.");
  const period = label("C13").replace(/^for the month of\s+/i, "").toLowerCase();
  const selected = [dateOffset(month + "-01", -1).slice(0, 7), month];
  const format = new Intl.DateTimeFormat("en-US", { month: "long", timeZone: "UTC" });
  requireReport(selected.some(value => {
    const name = format.format(new Date(value + "-01T00:00:00Z")).toLowerCase();
    return [name, name + " " + value.slice(0, 4), value].includes(period);
  }), "Check the current or prior month on the uploaded cover.");
  const items = [];
  for (let row = 28; row <= 45; row++) {
    const values = [..."DFIKL"].map(column => cells.get(column + row));
    if (values.every(value => !String(value ?? "").trim())) continue;
    requireReport([..."DFIKL"].every(column => !formulas.has(column + row))
      && String(values[1] ?? "").trim() && String(values[4] ?? "").trim(),
    "Check the description, amount and formulas on in-transit row " + row + ".");
    const fields = [..."DFIKL"].flatMap((column, i) => {
      const value = values[i];
      if (!String(value ?? "").trim()) return [];
      const isDate = ["I", "K"].includes(column) && typeof value === "number";
      return [{ cell: column + row, isDate, value: column === "L" ? cents(value) / 100
        : isDate ? sourceDate(value, book.date1904) : String(value) }];
    });
    items.push({ row, fields, amount: cents(values[4]) });
  }
  return { sheet, fileName: file.name, items, total: items.reduce((sum, item) => sum + item.amount, 0) };
}

async function readVizoTransferReview(primary, files, month) {
  requireReport(primary?.account, "Load the original VIZO Transaction Summary CSV for the selected account.");
  const reports = await Promise.all(files.map(async file => {
    requireReport(/\.csv$/i.test(file.name) && file.size <= 20 * 1024 * 1024, "Choose VIZO CSVs under 20 MB.");
    const report = vizoNativeProvider(readCsvCells(await file.text()));
    requireReport(report?.account && report.rows.some(item => item.date.startsWith(month)),
      file.name + " needs a VIZO Transaction Summary for the selected month.");
    return { ...report, file: file.name };
  }));
  const all = [primary, ...reports];
  requireReport(new Set(all.map(report => report.account)).size === all.length, "Choose one CSV per VIZO account.");
  const entries = all.flatMap(report => report.rows.filter(item => item.amount && item.date.startsWith(month)
    && /\b(?:transfer|overdraft)\b/i.test(item.description)).map(item => ({
      ...item, account: report.account, accountName: report.name, file: report.file,
      otherAccount: /\b(?:to|from)\s+(\d{4,})\b/i.exec(item.description)?.[1],
    })));
  const buckets = new Map();
  for (const item of entries) {
    const key = item.date + "|" + item.amount, bucket = buckets.get(key) ?? [];
    bucket.push(item);
    buckets.set(key, bucket);
  }
  const candidates = item => (buckets.get(item.date + "|" + -item.amount) ?? []).filter(other =>
    other.account !== item.account && (!item.otherAccount || item.otherAccount === other.account)
    && (!other.otherAccount || other.otherAccount === item.account));
  return entries.filter(item => item.account === primary.account).map(item => {
    const found = candidates(item);
    const unique = found.length === 1 && candidates(found[0]).length === 1;
    return { item, found, status: unique ? "Unique candidate · verify" : found.length ? "Competing candidates · review" : "Counterpart missing from supplied files" };
  });
}

function selectRelatedVizoFiles(files) {
  relatedVizoFiles = [...files];
  vizoField("related-status").textContent = relatedVizoFiles.map(file => file.name).join(", ");
  invalidateVizo();
}

vizoField("related-files").addEventListener("change", event => selectRelatedVizoFiles(event.target.files));
vizoField("related-drop").addEventListener("dragover", event => {
  event.preventDefault();
  vizoField("related-drop").classList.add("active");
});
vizoField("related-drop").addEventListener("dragleave", () => vizoField("related-drop").classList.remove("active"));
vizoField("related-drop").addEventListener("drop", event => {
  event.preventDefault();
  vizoField("related-drop").classList.remove("active");
  vizoField("related-files").value = "";
  selectRelatedVizoFiles(event.dataTransfer.files);
});

function vizoAccountMatches(mode, name) {
  const patterns = {
    atm: /\batm\b/, overnight: /\b(?:overnight|osl)\b/, psa: /\b(?:psa|premium share)\b/,
    pcc: /\b(?:pcc|premium club)\b/, misc: /\bmisc(?:ellaneous)?\b/,
    charge: /\bcharge cards?\b/, settlement: /^(?:vizo )?settlement(?: account)?$/,
  };
  return patterns[mode].test(name.trim().toLowerCase());
}

function vizoRange(kind, fields) {
  const columns = fields.map(field => vizoField(kind + "-" + field).value.trim().toUpperCase());
  requireReport(columns.every(column => /^[A-Z]{1,3}$/.test(column) && columnNumber(column) <= 16384)
    && new Set(columns).size === columns.length, "Use different Excel column letters for the " + kind + " fields.");
  const first = vizoField(kind + "-first").valueAsNumber;
  const lastField = vizoField(kind + "-last");
  const last = lastField.value ? lastField.valueAsNumber : 1048576;
  requireReport(Number.isInteger(first) && Number.isInteger(last) && first >= 1 && last >= first && last <= 1048576,
    "Check the first and last " + kind + " data rows.");
  return { columns, first, last, sheet: vizoField(kind + "-sheet").value };
}

async function vizoClosingCell(kind, month, address, dateColumn) {
  const { book, csv, file } = vizoFiles[kind];
  requireReport(book || csv, "Load the closing statement workbook first.");
  const sheet = vizoField(kind + "-sheet").value;
  const cell = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(address);
  requireReport(cell && columnNumber(cell[1]) <= 16384 && Number(cell[2]) <= 1048576,
    "Enter a valid closing-balance cell.");
  requireReport(!dateColumn || (/^[A-Z]{1,3}$/.test(dateColumn) && columnNumber(dateColumn) <= 16384 && dateColumn !== cell[1]),
    "Use a date column different from the closing-balance column.");
  const cells = csv ?? await readCells(book, sheet);
  if (dateColumn) {
    const date = sourceDate(cells.get(dateColumn + cell[2]), book?.date1904);
    requireReport(date.slice(0, 7) === month, "The closing balance in " + kind + " is outside the selected month.");
  }
  try {
    return { amount: cents(cells.get(address)), source: file.name + " · " + sheet + "!" + address };
  } catch {
    throw new Error("The closing balance at " + sheet + "!" + address + " is missing or invalid.");
  }
}

async function vizoRows(kind, fields, from, through) {
  const { columns, first, last, sheet } = vizoRange(kind, fields);
  const classification = kind === "provider" && vizoField("mode").value === "settlement" ? vizoField("provider-class").value.trim().toUpperCase() : "";
  requireReport(!classification || (/^[A-Z]{1,3}$/.test(classification) && columnNumber(classification) <= 16384 && !columns.includes(classification)),
    "Choose a separate reviewed deposit classification column.");
  const usedColumns = new Set([...columns, classification]);
  const state = vizoFiles[kind];
  const cells = state.csv ?? await readCells(state.book, sheet);
  const rows = new Set();
  for (const [address, value] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (!match || !usedColumns.has(match[1]) || !String(value ?? "").trim()) continue;
    const row = Number(match[2]);
    if (row >= first && row <= last) rows.add(row);
  }
  return Array.from(rows).sort((a, b) => a - b).flatMap(row => {
    const [dateValue, descValue, firstAmount, secondAmount] = columns.map(column => cells.get(column + row));
    try {
      const date = sourceDate(dateValue, state.book?.date1904 ?? false);
      if (date < from || date > through) return [];
      const description = String(descValue ?? "").trim().replace(/\s+/g, " ");
      requireReport(description, "The description is blank.");
      if (kind === "provider") {
        const reviewed = String(cells.get(classification + row) ?? "").trim().toLowerCase();
        const reviewedCategory = { "branch capture": "Branch capture", "mobile capture": "Mobile capture" }[reviewed];
        if (reviewed) {
          const detected = vizoAccountCategory(description);
          requireReport(reviewedCategory && (detected === "Unclassified forward collection" || detected === reviewedCategory),
            "Use Branch Capture or Mobile Capture for the corresponding forward collection row.");
        }
        return [{ date, description, amount: cents(firstAmount), row, reviewedCategory }];
      }
      requireReport(String(firstAmount ?? "").trim() || String(secondAmount ?? "").trim(), "Debit and credit are both blank.");
      const debit = String(firstAmount ?? "").trim() ? cents(firstAmount) : 0;
      const credit = String(secondAmount ?? "").trim() ? cents(secondAmount) : 0;
      requireReport(debit >= 0 && credit >= 0, "Use separate nonnegative debit and credit columns.");
      return [{ date, description, debit, credit, amount: debit - credit, row }];
    } catch (error) {
      throw new Error(kind + " " + sheet + " row " + row + ": " + error.message);
    }
  });
}

function vizoCategory(description) {
  const value = description.toUpperCase();
  if (/(?:^|[- ])EBT (?:FEES|SETTLE|SURCH)(?:\s+\d+)?$/.test(value)) return "ATM daily";
  if (/OVERDRAFT/.test(value)) return "Overdraft";
  if (/TRANSFER/.test(value)) return "Transfer";
  if (/DIVIDEND|INTEREST/.test(value)) return "Interest or dividend";
  if (/\bSPC\b/.test(value)) return "SPC";
  if (/\bAA\b/.test(value)) return "AA";
  if (/FEE|INTERCHANGE/.test(value)) return "Fees";
  return "Other";
}

function vizoGroups(provider) {
  const totals = new Map();
  for (const entry of provider) {
    const category = vizoRowCategory(entry), key = JSON.stringify([entry.description.toLowerCase(), category]);
    const group = totals.get(key) ?? { description: entry.description, category, count: 0, amount: 0 };
    group.count++;
    group.amount += entry.amount;
    requireReport(Number.isSafeInteger(group.amount), "A grouped total exceeds the supported range.");
    totals.set(key, group);
  }
  return [...totals.values()].sort((a, b) => a.description.localeCompare(b.description));
}

async function readChargeLoanStreet(from, end) {
  const { book, file } = vizoFiles.loanstreet;
  requireReport(book, "Load a readable LoanStreet G/L History workbook.");
  const normalize = text => text.toLowerCase().replace(/[\s-]+/g, "");
  const filter = normalize(vizoField("loanstreet-filter").value);
  const priorSign = Number(vizoField("loanstreet-prior-sign").value), currentSign = Number(vizoField("loanstreet-current-sign").value);
  requireReport(filter, "Enter the description text identifying LoanStreet entries.");
  const priorStart = dateOffset(from, -1).slice(0, 7) + "-01";
  const rows = await vizoRows("loanstreet", ["date", "desc", "debit", "credit"], priorStart, end);
  const included = rows.filter(item => normalize(item.description).includes(filter)).map(item => {
    const prior = item.date < from;
    const period = prior ? "Prior month" : "Current month", sign = prior ? priorSign : currentSign;
    requireReport(sign === 1 || sign === -1, "Choose the amount sign for the " + period.toLowerCase() + " LoanStreet rows.");
    return { ...item, period, contribution: item.amount * sign };
  });
  const total = period => {
    const entries = included.filter(item => item.period === period);
    const amount = entries.reduce((sum, item) => sum + item.contribution, 0);
    requireReport(Number.isSafeInteger(amount), "A LoanStreet total exceeds the supported range.");
    return entries.length ? amount : undefined;
  };
  return { prior: total("Prior month"), current: total("Current month"), rows: included, fileName: file.name,
    sheet: vizoField("loanstreet-sheet").value, filter: vizoField("loanstreet-filter").value.trim() };
}

function chargeLoanStreetFromWisdom(ledger, month, sign, file, sheet) {
  const period = Number(month.slice(5)) + "/" + month.slice(2, 4);
  const rows = (ledger ?? []).filter(item => item.date.startsWith(month)
    && /^LS\s+\S/i.test(item.batch ?? "") && item.batch.endsWith(" " + period))
    .map(item => ({ ...item, period: "Current month", contribution: -item.amount * sign }));
  if (!rows.length) return undefined;
  const current = rows.reduce((total, item) => total + item.contribution, 0);
  requireReport(Number.isSafeInteger(current), "The LoanStreet total exceeds the supported range.");
  return { current, rows, fileName: file.name, sheet, filter: "LS journal names for " + period };
}

async function readVizoLendKeyInvoice(file, month) {
  const text = await readLocalPdfText(file);
  requireReport(text.some(item => /^lendkey$/i.test(item.text)) && text.some(item => /^technologies,?$/i.test(item.text)),
    "The PDF does not identify a LendKey invoice.");
  const named = value => text.filter(item => item.text.toLowerCase() === value.toLowerCase());
  const sameRow = (a, b) => Math.abs(a.y - b.y) < 2;
  const pair = (first, second) => {
    const candidates = named(second).filter(b => named(first).some(a => a.x < b.x && sameRow(a, b)));
    requireReport(candidates.length === 1, "Check the " + first + " " + second + " invoice label.");
    return candidates[0];
  };
  const rightAmount = line => {
    const values = text.filter(item => item.x > line.x && sameRow(item, line) && Number.isFinite(readBalance(item.text)));
    requireReport(values.length === 1, "An invoice amount is missing or ambiguous.");
    return cents(values[0].text);
  };
  const dateLabel = pair("Invoice", "Date:");
  const dateMatch = text.find(item => item.x > dateLabel.x && sameRow(item, dateLabel)
    && /^\d{1,2}-[A-Za-z]{3}-\d{4}$/.test(item.text));
  requireReport(dateMatch, "The LendKey invoice date could not be verified.");
  const [day, abbreviation, year] = dateMatch.text.split("-");
  const months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const invoiceDate = sourceDate(year + "-" + String(months.indexOf(abbreviation.toLowerCase()) + 1).padStart(2, "0") + "-" + day.padStart(2, "0"));
  const priorMonth = dateOffset(month + "-01", -1).slice(0, 7);
  requireReport(invoiceDate.startsWith(priorMonth), "The manual calls for the " + priorMonth + " invoice for this reconciliation.");
  const subtotalLabel = pair("Invoice", "Total:"), subtotal = rightAmount(subtotalLabel);
  const tax = rightAmount(pair("Sales", "Tax:")), due = rightAmount(pair("Amount", "Due:"));
  const number = named("NO.")[0], quantity = named("QTY")[0], amount = named("AMOUNT").find(item => number && sameRow(item, number));
  requireReport(number && quantity && amount && sameRow(number, quantity) && number.x < quantity.x && quantity.x < amount.x,
    "The invoice item columns could not be verified.");
  const items = text.filter(item => Math.abs(item.x - number.x) < 12 && item.y < number.y
    && item.y > subtotalLabel.y && /^\d+$/.test(item.text)).sort((a, b) => b.y - a.y);
  requireReport(items.length && items.every((item, i) => Number(item.text) === i + 1), "The invoice item numbers could not be verified.");
  const lines = items.map((item, i) => {
    const description = text.filter(word => word.x > number.x + 12 && word.x < quantity.x - 5
      && word.y <= item.y + 2 && word.y > Math.max(item.y - 19, (items[i + 1]?.y ?? subtotalLabel.y) + 2))
      .sort((a, b) => Math.abs(a.y - b.y) < 2 ? a.x - b.x : b.y - a.y).map(word => word.text).join(" ");
    requireReport(description, "An invoice item description is missing.");
    return { number: Number(item.text), description, amount: rightAmount({ x: amount.x - 3, y: item.y }) };
  });
  requireReport(lines.reduce((sum, item) => sum + item.amount, 0) === subtotal && subtotal + tax === due && due > 0,
    "The line items, tax and invoice amount due do not agree.");
  return { file: file.name, invoiceDate, lines, subtotal, tax, due };
}

function compareVizoLendKeyInvoice(invoice, provider, ledger, sign) {
  const named = item => /\blend[ -]*key\b/i.test(item.description);
  const providerRows = provider.filter(item => named(item) && item.amount === -invoice.due);
  const ledgerRows = (ledger ?? []).filter(item => named(item) && item.amount * sign < 0);
  const ledgerDates = new Map();
  for (const item of ledgerRows) {
    const rows = ledgerDates.get(item.date) ?? [];
    rows.push(item);
    ledgerDates.set(item.date, rows);
  }
  const ledgerMatches = [...ledgerDates].flatMap(([date, rows]) => {
    const exact = rows.filter(item => item.amount * sign === -invoice.due);
    if (exact.length) return exact.map(item => ({ date, rows: [item] }));
    return rows.reduce((sum, item) => sum + item.amount * sign, 0) === -invoice.due ? [{ date, rows }] : [];
  });
  return { ...invoice, providerRows, ledgerMatches, ledgerLoaded: Boolean(ledger) };
}

async function readAtmIf67(from, through) {
  const { book, file } = vizoFiles.if67;
  requireReport(book, "Load a readable IF-67 daily totals workbook.");
  const sheet = vizoField("if67-sheet").value;
  const { columns: [dateColumn, settlementColumn, feeColumn], first, last } = vizoRange("if67", ["date", "settlement", "isa"]);
  const cells = await readCells(book, sheet), entries = new Map();
  const relevant = new RegExp("^(?:" + [dateColumn, settlementColumn, feeColumn].join("|") + ")([1-9]\\d*)$");
  const rows = new Set([...cells.keys()].map(address => Number(relevant.exec(address)?.[1])).filter(row => row >= first && row <= last));
  for (const row of [...rows].sort((a, b) => a - b)) {
    try {
      const date = sourceDate(cells.get(dateColumn + row), book.date1904);
      if (date < from || date > through) continue;
      requireReport(!entries.has(date), "Use one verified IF-67 total per ledger date.");
      entries.set(date, { date, settlement: cents(cells.get(settlementColumn + row)), fee: cents(cells.get(feeColumn + row)), row });
    } catch (error) {
      throw new Error(file.name + " · " + sheet + " row " + row + ": " + error.message);
    }
  }
  return { entries, fileName: file.name, sheet };
}

async function readAtmOpening(first, cutoff, filter) {
  const { book, file } = vizoFiles["atm-opening"];
  requireReport(book, "Load a readable opening ATM workbook.");
  const sheet = vizoField("atm-opening-sheet").value;
  const cells = await readCells(book, sheet);
  const layout = wisdomHistoryLayout(cells);
  const includesDate = date => date >= first && date < cutoff;
  if (layout) {
    requireReport(vizoAccountMatches("atm", layout.name), "Use the prior-month Wisdom history for the ATM account.");
    const history = wisdomHistory(cells, book.date1904, layout);
    return history.rows.filter(entry => includesDate(entry.date) && entry.description.toLowerCase().includes(filter))
      .map(entry => ({ ...entry, opening: { file: file.name, sheet, archived: false } }));
  }
  const heading = address => String(cells.get(address) ?? "").trim().toLowerCase();
  requireReport((book.sheets.has("Reconciliation ATM") || /\batm\b/i.test(sheet))
    && heading("A2") === "date" && /^(portico|cusa|wisdom)$/.test(heading("B2")) && /^(amount|amt)$/.test(heading("C2")),
    "Select the reviewed ATM Matchup worksheet with Date, Portico and Amount in A2:C2.");
  return [...cells].flatMap(([address, raw]) => {
    const row = Number(/^A([1-9]\d*)$/.exec(address)?.[1]);
    if (!(row > 2)) return [];
    const description = String(cells.get("B" + row) ?? "");
    if (!description.toLowerCase().includes(filter)) return [];
    const date = sourceDate(raw, book.date1904);
    if (!includesDate(date)) return [];
    return [{ date, row, description, amount: cents(cells.get("C" + row)), opening: { file: file.name, sheet, archived: true } }];
  });
}

function firstAtmLedgerDate(providerDate) {
  const monday = new Date(providerDate + "T00:00:00Z").getUTCDay() === 1;
  return dateOffset(providerDate, monday ? -3 : -1);
}

function atmReview(provider, ledger, filter, sign, if67) {
  const daily = new Map(), internal = new Map();
  for (const entry of provider) {
    if (vizoCategory(entry.description) !== "ATM daily") continue;
    const day = daily.get(entry.date) ?? { date: entry.date, amount: 0, count: 0 };
    day.amount += entry.amount;
    day.count++;
    daily.set(entry.date, day);
  }
  for (const entry of ledger) {
    if (!entry.description.toLowerCase().includes(filter)) continue;
    if (!internal.has(entry.date)) internal.set(entry.date, []);
    internal.get(entry.date).push(entry);
  }
  const days = [...daily.values()].sort((a, b) => a.date.localeCompare(b.date)).map((day, index, ordered) => {
    const previous = ordered[index - 1]?.date;
    const first = previous ?? firstAtmLedgerDate(day.date);
    const dates = [];
    for (let date = first; date < day.date; date = dateOffset(date, 1)) dates.push(date);
    return { ...day, dates };
  });
  const used = new Set(), if67Review = [];
  const comparisons = days.map(day => {
    const candidates = day.dates.flatMap(date => internal.get(date) ?? []);
    const archived = candidates.some(entry => entry.opening?.archived);
    const amount = candidates.reduce((sum, entry) => sum + entry.amount * sign, 0);
    const difference = day.amount - amount;
    requireReport([day.amount, amount, difference].every(Number.isSafeInteger), "A daily settlement exceeds the supported range.");
    let status;
    if (!candidates.length) status = "No ledger candidate";
    else if (day.dates.length > 1) status = "Multi-day batch · review IF-67";
    else if (difference !== 0) status = "Difference · verify ISA fee in IF-67";
    else status = "Amounts agree";
    if (if67) {
      const report = day.dates.map(date => {
        const entry = if67.entries.get(date), ledgerRows = internal.get(date) ?? [];
        const ledger = ledgerRows.length ? ledgerRows.reduce((sum, row) => sum + row.amount * sign, 0) : undefined;
        const variance = entry && ledger !== undefined ? entry.settlement - ledger - entry.fee : undefined;
        requireReport([ledger, variance].every(value => value === undefined || Number.isSafeInteger(value)), "An IF-67 daily total exceeds the supported range.");
        return { date, providerDate: day.date, entry, ledger, variance,
          review: !entry ? "IF-67 date missing" : ledger === undefined ? "Ledger date missing" : variance ? "IF-67 less ledger differs from ISA" : "Daily ISA difference agrees" };
      });
      if67Review.push(...report);
      if (candidates.length) {
        if (report.some(item => !item.entry || item.ledger === undefined)) status = "Incomplete IF-67 or ledger coverage";
        else {
          const settlements = report.reduce((sum, item) => sum + item.entry.settlement, 0);
          const fees = report.reduce((sum, item) => sum + item.entry.fee, 0);
          requireReport([settlements, fees].every(Number.isSafeInteger), "An IF-67 batch total exceeds the supported range.");
          status = settlements === day.amount && fees === difference && report.every(item => item.variance === 0)
            ? "IF-67 settlement and ISA agree · review" : "IF-67 totals differ · review";
        }
      }
    }
    if (status === "Amounts agree" || status === "IF-67 settlement and ISA agree · review") candidates.forEach(entry => used.add(entry));
    if (archived) status += " · includes archived opening rows";
    return { ...day, amountLedger: candidates.length ? amount : "", difference: candidates.length ? difference : "",
      ledgerDates: day.dates.join(", "), ledgerRows: candidates.map(entry => entry.opening
        ? entry.opening.sheet + (entry.opening.archived ? "!C" : " row ") + entry.row : entry.row).join(", "), status };
  });
  if (if67) {
    const comparedDates = new Set(if67Review.map(item => item.date));
    for (const entry of if67.entries.values()) {
      if (!comparedDates.has(entry.date)) if67Review.push({ date: entry.date, entry, review: "No selected provider batch" });
    }
  }
  return { comparisons, pending: [...internal.values()].flat().filter(entry => !used.has(entry)), if67Review };
}

function showVizo(result) {
  const money = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
  const table = document.createElement("table");
  const metrics = [
    ["Account template", result.profile.name],
    ["Provider transaction groups", result.groups.length],
    ["Statement ending balance", money.format(result.statement / 100)],
    ["Ledger balance after sign and adjustment", money.format(result.adjusted / 100)],
    ["Ending difference", money.format(result.difference / 100)],
  ];
  if (result.mode === "settlement") metrics.push(["Next-period provider activity",
    result.boundaryRows.length ? result.boundaryRows.length + " rows available for cutoff review" : "Awaiting the next-period export"]);
  if (result.noProviderActivity) metrics.push(["Overnight Settlement activity", "None confirmed by preparer"]);
  if (result.transit) metrics.push(["Reviewed in-transit items", result.transit.items.length]);
  if (result.transferReview) metrics.push(["Unique transfer candidates", result.transferReview.filter(row => row.status.startsWith("Unique")).length]);
  if (result.mode === "charge") metrics.push(
    ["LoanStreet transit", result.loanStreetTransit === undefined ? "Review needed" : money.format(result.loanStreetTransit / 100)]);
  if (result.invoice) metrics.push(["LendKey invoice amount due", money.format(result.invoice.due / 100)],
    ["Exact VIZO invoice candidates", result.invoice.providerRows.length], ["Wisdom invoice candidates", result.invoice.ledgerMatches.length]);
  if (result.atm) metrics.push(
    ["ATM dates requiring review", result.atm.comparisons.filter(day => day.status !== "Amounts agree").length],
    ["Ledger dailies without an agreeing batch", result.atm.pending.length],
  );
  const body = table.createTBody();
  for (const [name, amount] of metrics) {
    const row = body.insertRow();
    const heading = document.createElement("th");
    heading.scope = "row";
    heading.textContent = name;
    row.append(heading);
    row.insertCell().textContent = amount;
  }
  vizoField("summary").replaceChildren(table);
  vizoField("download-template").hidden = !result.template;
  vizoField("template-download-note").hidden = !result.template;
  vizoField("results").hidden = false;
}

function vizoCoverPaste(result, audit) {
  const target = result.profile.paste;
  const sheet = target.sheet ?? "Reconciliation";
  const rows = [
    [sheet, "L20", "VIZO statement ending balance", result.statement / 100, result.statementSource],
    [sheet, target.ledger, "Ledger ending balance", result.ending * result.sign / 100,
      result.endingSource + (result.sign === -1 ? " · sign reversed" : "")],
    [sheet, target.initials, "Completed by", audit.by, "Preparer"],
    [sheet, target.date, "Completed on", audit.on, "Preparer"],
  ];
  if (result.mode === "charge" && result.loanStreetTransit !== undefined) {
    rows.push([sheet, "L21", "LoanStreet in transit on statement side", -result.loanStreetTransit / 100,
      result.chargeLoanStreet?.current !== undefined ? result.chargeLoanStreet.fileName + " · LoanStreet G/L rows" : "Reviewed adjustment"]);
  }
  return rows;
}

async function fillVizoTemplate(result) {
  const audit = reconciliationAudit();
  const target = result.profile.paste;
  const sheet = target.sheet ?? "Reconciliation";
  const book = await openExcel(result.template, true);
  const formulas = new Set();
  const cells = await readCells(book, sheet, formulas);
  const label = address => String(cells.get(address) ?? "").trim();
  const period = label("C13").toLowerCase().replace(/^for the month of\s+/, "").replace(/^septmeber\b/, "september");
  const monthName = new Intl.DateTimeFormat("en-US", { month: "long", timeZone: "UTC" })
    .format(new Date(result.month + "-01T00:00:00Z")).toLowerCase();
  const year = result.month.slice(0, 4);
  requireReport(vizoAccountMatches(result.mode, label("C11")) && [monthName, monthName + " " + year, result.month].includes(period),
    "Load the prepared reconciliation for " + result.profile.name + " and " + result.month + ".");
  const ledgerLabel = "C" + target.ledger.slice(1);
  requireReport(/^balance per vizo statement$/i.test(label("C20"))
    && /^balance per (?:wisdom|portico|cusa) g\/l$/i.test(label(ledgerLabel)) && formulas.has(target.difference),
    "The workbook does not match the balance fields for this account.");
  const auditRows = [target.initials.slice(1), target.date.slice(1)];
  requireReport(auditRows.some(row => /^prepared by\s*:?$/i.test(label("H" + row)))
    && (!cells.has(target.initials) || typeof cells.get(target.initials) === "string")
    && (!cells.has(target.date) || typeof cells.get(target.date) === "number"),
    "Check the completion fields on the prepared cover.");
  const updates = vizoCoverPaste(result, audit).map(([tab, cell, , value]) =>
    [tab, cell, cell === target.date ? { date: value } : value]);
  requireReport(!updates.some(([, cell]) => formulas.has(cell))
    && (result.mode !== "charge" || result.loanStreetTransit === undefined || /loanstreet in transit/i.test(label("E21"))),
    "Check the mapped input cells on the prepared cover.");
  return fillExcelTemplate(result.template, updates);
}

async function buildVizo(result) {
  const audit = reconciliationAudit();
  const { label, number, date, row, worksheet } = excelOutputCells();
  function table(name, headers, entries, widths, dateCells = []) {
    const rows = [row(1, headers.map((header, i) => label(csvColumn(i + 1) + "1", header, 1)).join(""))];
    for (const [i, entry] of entries.entries()) {
      const r = i + 2;
      rows.push(row(r, entry.map((value, j) => {
        const address = csvColumn(j + 1) + r;
        if (dateCells.includes(address)) return date(address, value);
        return typeof value === "number" ? number(address, value) : label(address, value);
      }).join("")));
    }
    return [name, worksheet("A1:" + csvColumn(headers.length) + rows.length,
      widths.map((width, i) => [i + 1, i + 1, width]), rows.join(""))];
  }
  const provider = result.provider.slice().sort((a, b) => a.description.localeCompare(b.description) || a.date.localeCompare(b.date));
  const coverInputs = [
    ["Period", result.month], ["Account", result.profile.name],
    ["Statement balance", result.statement / 100], ["Ledger ending balance", result.ending / 100],
    ["Ledger sign multiplier", result.sign], ["Other reviewed ledger adjustments", result.adjustment / 100],
  ];
  if (result.mode === "charge") coverInputs.push(["LoanStreet in transit", result.loanStreetTransit === undefined ? "" : result.loanStreetTransit / 100]);
  const coverRows = coverInputs.map(([name, amount], i) => {
    const r = i + 2;
    return row(r, label("A" + r, name) + (typeof amount === "number" ? number("B" + r, amount) : label("B" + r, amount)));
  });
  const adjustedRow = coverRows.length + 2;
  const formula = (r, title, expression, value) => row(r, label("A" + r, title)
    + `<c r="B${r}" s="2"><f>${expression}</f><v>${value / 100}</v></c>`);
  const cover = row(1, label("A1", result.profile.name + " reconciliation", 1))
    + coverRows.join("")
    + formula(adjustedRow, "Adjusted ledger balance", "B5*B6+" + (result.mode === "charge" ? "SUM(B7:B8)" : "B7"), result.adjusted)
    + formula(adjustedRow + 1, "Draft difference", "B4-B" + adjustedRow, result.difference)
    + row(adjustedRow + 3, label("A" + (adjustedRow + 3), result.profile.note))
    + row(adjustedRow + 5, label("A" + (adjustedRow + 5), "Completed by") + label("B" + (adjustedRow + 5), audit.by))
    + row(adjustedRow + 6, label("A" + (adjustedRow + 6), "Completed on") + date("B" + (adjustedRow + 6), audit.on));
  const opposite = result.mode === "overnight";
  const target = result.profile.paste;
  const sheet = target.sheet ?? "Reconciliation";
  const pasteRows = [
    ...vizoCoverPaste(result, audit),
    [sheet, target.difference, "Difference formula", "", "Recalculate in the existing workbook after entering reviewed in-transit items."],
  ];
  const sheets = [
    [result.profile.cover, worksheet("A1:B" + (adjustedRow + 6), [[1, 1, 47], [2, 2, 25]], cover)],
    table("Cover paste cells", ["Destination worksheet", "Cell", "Field", "Paste value", "Source or action"],
      pasteRows, [33, 11, 34, 19, 76], ["D5"]),
    table(result.profile.groups, ["Category", "Description", "Rows", "Provider amount", ...(opposite ? ["Opposite sign for review"] : [])],
      result.groups.map(group => [group.category, group.description, group.count, group.amount / 100,
        ...(opposite ? [-group.amount / 100] : [])]), [23, 68, 12, 21, ...(opposite ? [26] : [])]),
    table("Provider detail", ["Date", "Description", "Category", "Signed amount", "Source row"],
      provider.map(item => [item.date, item.description, vizoRowCategory(item), item.amount / 100, item.row]),
      [15, 68, 23, 19, 15]),
    table("Sources", ["Field", "Value"], [
      ["Provider file", result.providerFile], ["Provider worksheet", result.providerSheet], ["Ledger file", result.ledgerFile],
      ...(result.noProviderActivity ? [["Provider activity", "No activity confirmed by preparer"]] : []),
      ["Ledger worksheet", result.ledgerSheet], ["Statement balance source", result.statementSource], ["Ending balance source", result.endingSource],
      ...(result.boundaryFile ? [["Next-period provider file", result.boundaryFile]] : []),
      ...(result.if67 ? [["IF-67 file", result.if67.fileName], ["IF-67 worksheet", result.if67.sheet]] : []),
      ...(result.mode === "settlement" ? [["Branch capture ledger offset (days)", result.branchOffset]] : []),
    ], [39, 65]),
  ];
  if (result.transit?.items.length) {
    const fields = result.transit.items.flatMap(item => item.fields);
    sheets.push(table("In-transit paste cells", ["Destination worksheet", "Cell", "Paste value", "Source workbook"],
      fields.map(field => [result.transit.sheet, field.cell, field.value, result.transit.fileName]),
      [34, 12, 60, 42], fields.flatMap((field, i) => field.isDate ? ["C" + (i + 2)] : [])));
  }
  if (result.transferReview?.length) sheets.push(table("Account transfer review",
    ["Date", "Source row", "Description", "Amount", "Other account", "Other source", "Other source row", "Review"],
    result.transferReview.flatMap(({ item, found, status }) => (found.length ? found : [undefined]).map(other =>
      [item.date, item.row, item.description, item.amount / 100, other?.accountName ?? "", other?.file ?? "", other?.row ?? "", status])),
    [16, 12, 54, 18, 28, 42, 16, 42]));
  if (result.mode === "misc") {
    const miscCover = [
      row(2, label("C2", "Preparation date") + date("F2", audit.on)),
      row(3, label("C3", "Report prepared by") + label("F3", audit.by)),
      row(11, label("C11", "VIZO MISC", 1)),
      row(13, label("C13", result.month)),
      row(20, label("C20", "Balance per VIZO statement") + number("L20", result.statement / 100)),
      row(24, label("C24", "Balance per Wisdom G/L") + number("L24", result.ending * result.sign / 100)),
      row(26, label("C26", "In transit items", 1)),
      row(27, label("D27", "Line Item", 1) + label("F27", "Item Description", 1)
        + label("I27", "Will Clear On", 1) + label("K27", "Date", 1) + label("L27", "Amount", 1)),
      ...(result.transit ? result.transit.items.map(item => row(item.row, item.fields.map(field =>
        field.isDate ? date(field.cell, field.value)
          : typeof field.value === "number" ? number(field.cell, field.value) : label(field.cell, field.value)).join("")))
        : result.adjustment ? [row(28, label("F28", "Reviewed in-transit total") + number("L28", result.adjustment / 100))] : []),
      row(46, label("D46", "Adjusted balance") + `<c r="L46" s="2"><f>SUM(L24:L45)</f><v>${result.adjusted / 100}</v></c>`),
      row(48, label("D48", "Difference") + `<c r="L48" s="2"><f>L20-L46</f><v>${result.difference / 100}</v></c>`),
      row(52, label("L52", audit.by)),
      row(53, label("H53", "Prepared by") + date("L53", audit.on)),
      row(57, label("H57", "Verified by")),
    ];
    sheets.unshift(["Reconciliation", worksheet("A1:L57", [[1, 2, 8], [3, 3, 21], [4, 5, 12], [6, 6, 28],
      [7, 7, 9], [8, 8, 18], [9, 10, 9], [11, 11, 14], [12, 12, 21]], miscCover.join(""),
    ["C2:E2", "C3:E3", "C11:L11", "C13:L13", "C20:K20", "C24:K24", "C26:L26", "F27:H27",
      ...(!result.transit ? ["F28:J28"] : [])])]);
  }
  if (result.boundaryRows.length) sheets.push(table("Provider boundary rows", ["Date", "Description", "Amount", "Source row", "File"],
    result.boundaryRows.map(item => [item.date, item.description, item.amount / 100, item.reviewRow ?? item.row, item.source ?? result.providerFile]),
    [16, 65, 18, 24, 40]));
  if (result.unclassifiedDeposits.length) {
    const deposits = vizoForwardDepositCandidates(result.unclassifiedDeposits, result.ledger ?? [], result.sign, result.branchOffset);
    sheets.push(table("Forward deposit review", ["Date", "Description", "Amount", "Source row", "Review"],
      deposits.map(({ item, candidates }) => [item.date, item.description, item.amount / 100, item.reviewRow ?? item.row,
        candidates.length ? candidates.map(match => match.category + " · " + match.date + " · Wisdom rows " + match.rows.join(", ")).join("; ")
          : "Verify branch or mobile capture"]), [16, 60, 18, 17, 64]));
    const candidates = deposits.flatMap(({ item, candidates }) => candidates.filter(match => match.branchAmount >= 0 && match.mobileAmount >= 0)
      .map(match => [item.reviewRow ?? item.row, match.date, match.branchAmount / 100, match.mobileAmount / 100, match.rows.join(", ")]));
    if (candidates.length) sheets.push(table("Deposit allocation candidates", ["Provider row", "Ledger date", "Branch amount", "Mobile amount", "Wisdom rows"],
      candidates, [16, 16, 18, 18, 40]));
  }
  if (result.allocations.length) sheets.push(table("Reviewed deposit allocations", ["Provider row", "Provider date", "Ledger date", "Branch amount", "Mobile amount", "Provider total"],
    result.allocations.map(({ item, ledgerDate, branchAmount, mobileAmount }) => [
      item.reviewRow ?? item.row, item.date, ledgerDate, branchAmount / 100, mobileAmount / 100, item.amount / 100,
    ]), [16, 16, 16, 18, 18, 18]));
  if (result.ledger) sheets.push(table("Ledger history", ["Date", "Description", "Debit", "Credit", "Debit less credit", "Source row", "Category"],
    result.ledger.map(item => [item.date, item.description, item.debit / 100, item.credit / 100, item.amount / 100, item.row,
      vizoAccountCategory(item.description)]), [15, 65, 19, 19, 19, 15, 23]));
  const unclassifiedDates = new Set(result.unclassifiedDeposits.map(item => item.date));
  for (const [name, category] of result.profile.tabs) {
    const branch = result.mode === "settlement" && category === "Branch capture";
    const deposit = branch || category === "Mobile capture";
    const offset = branch ? result.branchOffset : 0;
    const reviewed = deposit ? result.allocations.flatMap(({ item, ledgerDate, branchAmount, mobileAmount }) => {
      const amount = branch ? branchAmount : mobileAmount;
      return amount ? [{ ...item, amount, ledgerDate, reviewedCategory: category }] : [];
    }) : [];
    const providerItems = branch ? [...provider, ...result.branchProvider, ...reviewed] : deposit ? [...provider, ...reviewed] : provider;
    const data = vizoAccountComparison(providerItems, result.ledger ?? [], category, result.sign, offset)
      .filter(item => item.date >= result.month + "-01" && (item.date <= result.end
        || (branch && dateOffset(item.date, offset) <= result.end)))
      .map(item => {
        const hasProvider = item.providerRows.length > 0, hasLedger = item.ledgerRows.length > 0;
        const providerDates = item.providerDates.size ? [...item.providerDates].sort() : [item.date];
        const status = category === "Returned deposit" ? "Verify member posting with Ops"
          : deposit && providerDates.some(date => unclassifiedDates.has(date)) ? "Unclassified forward collection on this date"
          : !result.ledger ? "Load G/L History to compare"
          : item.date > result.end ? "Possible month-end branch carryover"
            : !hasProvider && deposit && item.date === result.end && !result.boundaryRows.length ? "Awaiting next-period provider activity"
              : !hasProvider ? "Ledger activity only" : !hasLedger ? "Provider activity only"
              : item.provider !== item.ledger ? "Amounts differ"
                : category === "Money order" ? "Date totals agree · verify batches" : "Date totals agree · verify source rows";
        return [providerDates.join(", "), dateOffset(item.date, offset), item.direction, hasProvider ? item.provider / 100 : "",
          hasLedger ? item.ledger / 100 : "", hasProvider && hasLedger ? (item.provider - item.ledger) / 100 : "",
          item.providerRows.join(", "), item.ledgerRows.join(", "), status];
      });
    sheets.push(table(name, ["Provider date", "Ledger date", "Direction", "Provider", "Ledger", "Difference", "Provider rows", "Ledger rows", "Review"],
      data, [16, 16, 12, 18, 18, 18, 18, 18, 43]));
  }
  if (result.profile.interest) {
    const entries = provider.filter(item => vizoAccountCategory(item.description) === "Interest or dividend");
    const incoming = entries.filter(item => item.amount > 0).reduce((sum, item) => sum + item.amount, 0);
    const outgoing = entries.filter(item => item.amount < 0).reduce((sum, item) => sum + item.amount, 0);
    sheets.push(table(result.profile.name + " interest", ["Field", "Value"], [
      ["Credits", incoming / 100], ["Withdrawals", outgoing / 100], ["Net", (incoming + outgoing) / 100],
      ["Review", incoming && outgoing && incoming + outgoing === 0 ? "Net zero · verify the interest transfer pair" : "Review interest and related transfers"],
    ], [28, 62]));
  }
  if (result.mode === "charge") {
    const loanStreet = provider.filter(item => vizoAccountCategory(item.description) === "LoanStreet");
    const amount = loanStreet.reduce((sum, item) => sum + item.amount, 0);
    sheets.push(table("LoanStreet carryover", ["Field", "Amount or review"], [
      ["Prior month reviewed total", result.priorLoanStreet === undefined ? "" : result.priorLoanStreet / 100],
      ["Provider rows labeled LoanStreet", loanStreet.length ? amount / 100 : ""],
      ["Difference to prior total", loanStreet.length && result.priorLoanStreet !== undefined ? (amount - result.priorLoanStreet) / 100 : ""],
      ["Current signed cover adjustment", result.loanStreetTransit === undefined ? "" : result.loanStreetTransit / 100],
      ...(result.chargeLoanStreet ? [
        ["LoanStreet history", result.chargeLoanStreet.fileName + " · " + result.chargeLoanStreet.sheet],
        ["Description filter", result.chargeLoanStreet.filter],
        ["Prior month basis", result.chargeLoanStreet.prior !== undefined ? "Matched Wisdom rows" : result.priorLoanStreet !== undefined ? "Manual fallback" : "No matched rows"],
        ["Current month basis", result.chargeLoanStreet.current !== undefined ? "Matched Wisdom rows" : result.loanStreetTransit !== undefined ? "Manual fallback" : "No matched rows"],
      ] : []),
      ["Review", result.loanStreetTransit === undefined ? "Verify the current LoanStreet in-transit amount" : "Verify the prior and current LoanStreet amounts"],
    ], [40, 62]));
    if (result.chargeLoanStreet) sheets.push(table("LoanStreet G-L rows",
      ["Date", "Period", "Journal name", "Description", "Debit", "Credit", "Signed contribution", "Source row"],
      result.chargeLoanStreet.rows.map(item => [item.date, item.period, item.batch ?? "", item.description, item.debit / 100, item.credit / 100,
        item.contribution / 100, item.row]), [15, 18, 25, 51, 17, 17, 23, 14]));
    if (result.invoice) {
      const { invoice } = result;
      const status = count => count === 1 ? "One candidate · verify" : count ? "Multiple candidates · review" : "No exact candidate";
      sheets.push(table("LendKey invoice review", ["Field", "Amount or review"], [
        ["Source PDF", invoice.file], ["Invoice date", invoice.invoiceDate], ["Invoice subtotal", invoice.subtotal / 100],
        ["Tax", invoice.tax / 100], ["Amount due", invoice.due / 100],
        ["VIZO billing debit", status(invoice.providerRows.length)],
        ["Wisdom G/L History", invoice.ledgerLoaded ? status(invoice.ledgerMatches.length) : "Load G/L History to check postings"],
      ], [29, 52]));
      sheets.push(table("LendKey invoice items", ["Item", "Description", "Amount"],
        invoice.lines.map(item => [item.number, item.description, item.amount / 100]), [12, 64, 18]));
      sheets.push(table("LendKey invoice candidates", ["Source", "Date", "Description", "Comparison amount", "Source rows"],
        [...invoice.providerRows.map(item => ["VIZO", item.date, item.description, item.amount / 100, item.row]),
          ...invoice.ledgerMatches.map(match => ["Wisdom", match.date, match.rows.map(item => item.description).join(" · "),
            -invoice.due / 100, match.rows.map(item => item.row).join(", ")])], [16, 16, 63, 22, 25]));
    }
  }
  if (result.atm) {
    sheets.push(
      table("ATM comparison", ["Provider date", "Rows", "Provider amount", "Candidate ledger dates", "Source ledger rows", "Candidate ledger amount", "Difference", "Review"],
        result.atm.comparisons.map(day => [day.date, day.count, day.amount / 100, day.ledgerDates, day.ledgerRows,
          day.amountLedger === "" ? "" : day.amountLedger / 100, day.difference === "" ? "" : day.difference / 100, day.status]),
        [16, 10, 20, 35, 21, 24, 18, 43]),
      table("Ledger dailies to review", ["Date", "Description", "Signed amount", "Source row", "Review"],
        result.atm.pending.map(item => [item.date, item.description, item.amount / 100, item.row,
          item.date === result.end ? "Possible month-end carryover" : "Review settlement and IF-67"]), [15, 65, 20, 15, 32]),
    );
    if (result.opening.length) sheets.push(table("ATM opening sources", ["Date", "Description", "Signed ledger amount", "Workbook", "Worksheet", "Source row", "Basis"],
      result.opening.map(item => [item.date, item.description, item.amount / 100, item.opening.file, item.opening.sheet,
        item.row, item.opening.archived ? "Archived matchup · verify with Wisdom" : "Prior Wisdom history"]),
      [16, 56, 21, 32, 32, 14, 38]));
  }
  if (result.if67) {
    const dollars = amount => amount === undefined ? "" : amount / 100;
    sheets.push(table("IF-67 daily review", ["Ledger date", "Provider date", "IF-67 source row", "IF-67 settlement", "Ledger total", "IF-67 ISA",
      "Unexplained difference", "Review"], result.atm.if67Review.map(item => [
      item.date, item.providerDate ?? "", item.entry?.row ?? "", dollars(item.entry?.settlement), dollars(item.ledger),
      dollars(item.entry?.fee), dollars(item.variance), item.review,
    ]), [16, 16, 17, 21, 19, 16, 24, 39]));
  }
  return createExcelWorkbook(sheets, audit);
}

for (const kind of Object.keys(vizoFiles)) {
  const input = vizoField(kind + "-file"), drop = vizoField(kind + "-drop");
  input.addEventListener("change", () => selectVizoFile(kind, input.files[0]));
  drop.addEventListener("dragover", event => { event.preventDefault(); drop.classList.add("active"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("active"));
  drop.addEventListener("drop", event => {
    event.preventDefault();
    drop.classList.remove("active");
    if (event.dataTransfer.files.length !== 1) return vizoStatus("Drop one workbook or CSV at a time.", true);
    input.value = "";
    selectVizoFile(kind, event.dataTransfer.files[0]);
  });
}
document.addEventListener("dragover", event => event.preventDefault());
document.addEventListener("drop", event => event.preventDefault());
for (const input of document.querySelectorAll("input:not([type=file]):not(#confirm), select, textarea")) {
  input.addEventListener("input", invalidateVizo);
}
function showStatementSource() {
  const mode = vizoField("statement-mode").value;
  vizoField("statement").disabled = mode !== "manual";
  vizoField("statement-upload").hidden = mode !== "statement";
  vizoField("statement-cell-fields").hidden = !["provider", "statement"].includes(mode);
}
vizoField("statement-mode").addEventListener("input", showStatementSource);
function showVizoAccount() {
  const mode = vizoField("mode").value, profile = vizoAccounts[mode];
  const atm = mode === "atm";
  const type = vizoField("ledger-type").value, history = type === "history", native = type === "wisdom";
  vizoField("ledger-type").disabled = false;
  vizoField("atm-map").hidden = !history;
  vizoField("trial-map").hidden = type !== "trial";
  vizoField("atm-filter").hidden = !atm;
  vizoField("if67-panel").hidden = !atm;
  vizoField("charge-fields").hidden = mode !== "charge";
  vizoField("loanstreet-panel").hidden = mode !== "charge";
  vizoField("invoice-panel").hidden = mode !== "charge";
  const canImportTransit = ["misc", "settlement"].includes(mode);
  vizoField("transit-panel").hidden = !canImportTransit;
  vizoField("adjustment").disabled = canImportTransit && Boolean(vizoFiles.transit.file);
  vizoField("settlement-fields").hidden = mode !== "settlement";
  vizoField("boundary-provider").hidden = mode !== "settlement";
  vizoField("overnight-empty-field").hidden = mode !== "overnight";
  vizoField("provider-class-field").hidden = mode !== "settlement";
  vizoField("provider-class-help").hidden = mode !== "settlement";
  vizoField("account-note").textContent = profile.note;
  vizoField("ledger-instructions").textContent = atm
    ? "Use G/L History for ATM dailies. Include the last three dates of the prior month to cover a possible weekend batch."
    : native ? "The tool reads signed activity and uses the last running balance through the selected month end."
      : history ? "Use the account's G/L History and enter its verified ending balance or select the ending-balance cell."
      : "Use the trial balance for the account and selected month. Select G/L History to compare activity.";
}
vizoField("mode").addEventListener("change", () => {
  vizoField("ledger-type").value = vizoFiles.ledger.nativeWisdom ? "wisdom" : vizoAccounts[vizoField("mode").value].ledger;
  vizoField("overnight-empty").checked = false;
  showVizoAccount();
});
vizoField("ledger-type").addEventListener("change", showVizoAccount);
vizoField("month").addEventListener("input", () => { vizoField("overnight-empty").checked = false; });
showVizoAccount();
vizoField("confirm").addEventListener("change", () => {
  if (!vizoField("confirm").checked) invalidateVizo();
});

vizoField("run").addEventListener("click", async () => {
  const current = ++vizoRevision;
  vizoResult = undefined;
  vizoField("results").hidden = true;
  try {
    const month = vizoField("month").value, mode = vizoField("mode").value, profile = vizoAccounts[mode];
    const sign = Number(vizoField("sign").value);
    requireReport(/^\d{4}-\d{2}$/.test(month) && (sign === 1 || sign === -1), "Choose the month and ledger comparison sign.");
    requireReport(vizoField("confirm").checked, "Verify the account, period, column selections, and balances first.");
    const providerLoaded = Boolean(vizoFiles.provider.book || vizoFiles.provider.csv);
    const noProviderActivity = mode === "overnight" && vizoField("overnight-empty").checked;
    requireReport(vizoFiles.ledger.book && (providerLoaded || noProviderActivity), mode === "overnight"
      ? "Load the provider and ledger exports or confirm no Overnight Settlement activity." : "Load the provider and ledger exports.");
    const from = month + "-01";
    const [year, m] = month.split("-").map(Number);
    const end = new Date(Date.UTC(year, m, 0)).toISOString().slice(0, 10);
    const statementMode = vizoField("statement-mode").value;
    const providerReport = vizoFiles.provider.nativeProvider;
    if (providerReport) requireReport(vizoAccountMatches(mode, providerReport.name), "The VIZO export account does not match the selected template.");
    const transferReview = relatedVizoFiles.length && providerLoaded ? await readVizoTransferReview(
      providerReport && { ...providerReport, file: vizoFiles.provider.file.name }, relatedVizoFiles, month) : undefined;
    if (current !== vizoRevision) return;
    const running = providerReport?.rows.findLast(row => row.date <= end && row.date >= from);
    if (statementMode === "running") requireReport(running, "Load the VIZO Transaction Summary with activity in the selected month.");
    const statementInput = statementMode === "manual" ? { amount: cents(vizoField("statement").value), source: "Entered manually" }
      : statementMode === "running" ? { amount: running.balance, source: vizoFiles.provider.file.name + " · CSV!H" + running.row + " · " + running.date }
        : await vizoClosingCell(statementMode, month, vizoField("statement-cell").value.trim().toUpperCase(), vizoField("statement-date").value.trim().toUpperCase());
    if (current !== vizoRevision) return;
    const { amount: statement, source: statementSource } = statementInput;
    const transitFile = ["misc", "settlement"].includes(mode) ? vizoFiles.transit.file : undefined;
    const transit = transitFile ? await readVizoTransit(transitFile, mode, month) : undefined;
    if (current !== vizoRevision) return;
    const adjustment = transit ? transit.total : cents(vizoField("adjustment").value);
    const optionalAmount = id => vizoField(id).value.trim() ? cents(vizoField(id).value) : undefined;
    const branchOffset = mode === "settlement" ? vizoField("branch-offset").valueAsNumber : 0;
    requireReport(Number.isInteger(branchOffset) && branchOffset >= -7 && branchOffset <= 7, "Choose a branch capture offset between −7 and 7 days.");
    let chargeLoanStreet = mode === "charge" && vizoFiles.loanstreet.file ? await readChargeLoanStreet(from, end) : undefined;
    if (current !== vizoRevision) return;
    const priorLoanStreet = mode === "charge" ? chargeLoanStreet?.prior ?? optionalAmount("prior-loanstreet") : undefined;
    const manualTransit = mode === "charge" ? optionalAmount("current-loanstreet") : undefined;
    const providerSheet = vizoField("provider-sheet").value, ledgerSheet = vizoField("ledger-sheet").value;
    const boundaryEnd = mode === "settlement" ? dateOffset(end, 7) : end;
    const providerRows = providerLoaded ? await vizoRows("provider", ["date", "desc", "amount"], from, boundaryEnd) : [];
    const boundaryFile = mode === "settlement" ? vizoFiles.boundary.file : undefined;
    if (boundaryFile) providerRows.push(...await readVizoBoundary(boundaryFile, providerReport, end, boundaryEnd));
    if (current !== vizoRevision) return;
    const provider = providerRows.filter(item => item.date <= end);
    const boundaryRows = providerRows.filter(item => item.date > end);
    const branchProvider = boundaryRows.filter(item => vizoRowCategory(item) === "Branch capture");
    const unclassifiedDeposits = mode === "settlement" ? providerRows.filter(item => vizoRowCategory(item) === "Unclassified forward collection") : [];
    requireReport(!noProviderActivity || !provider.length, "The provider export contains activity in the selected month. Clear the no-activity confirmation.");
    requireReport(provider.length || noProviderActivity, "The provider export has no rows for the selected month.");
    let ending, endingSource, atm, ledger, if67, opening = [];
    const ledgerType = vizoField("ledger-type").value;
    requireReport(mode !== "atm" || ledgerType !== "trial", "Use a detailed Wisdom report or G/L History for ATM.");
    if (["wisdom", "history"].includes(ledgerType)) {
      const native = ledgerType === "wisdom" ? vizoFiles.ledger.nativeWisdom : undefined;
      requireReport(ledgerType !== "wisdom" || (native && vizoAccountMatches(mode, native.name)),
        "Load the original detailed Wisdom report for the selected VIZO account.");
      const address = vizoField("ledger-ending-cell").value.trim().toUpperCase();
      const closing = native ? native.balanceOn(end) : address
        ? await vizoClosingCell("ledger", month, address, vizoField("ledger-ending-date").value.trim().toUpperCase())
        : { amount: cents(vizoField("ledger-ending").value), source: "Entered from closing balance" };
      if (current !== vizoRevision) return;
      ending = closing.amount;
      endingSource = native ? ledgerSheet + "!" + closing.address + " · " + end : closing.source;
      const first = dateOffset(from, mode === "atm" ? -3 : Math.min(0, branchOffset));
      const last = dateOffset(end, Math.max(0, branchOffset));
      ledger = native ? native.rows.filter(row => row.date >= first && row.date <= last)
        : await vizoRows("ledger", ["date", "desc", "debit", "credit"], first, last);
      requireReport(native ? native.from <= from : ledger.length, "The ledger history does not cover the selected period.");
      if (mode === "atm") {
        const filter = native ? "corporate settlements gl" : vizoField("ledger-filter").value.trim().toLowerCase();
        requireReport(filter, "Enter the description text that identifies ATM daily ledger entries.");
        const ledgerStart = native?.from ?? ledger.reduce((earliest, item) => item.date < earliest ? item.date : earliest, from);
        const cutoff = ledgerStart < from ? ledgerStart : from;
        if (vizoFiles["atm-opening"].file) {
          const firstProvider = provider.filter(item => vizoCategory(item.description) === "ATM daily").map(item => item.date).sort()[0];
          const openingStart = firstProvider && firstAtmLedgerDate(firstProvider);
          if (openingStart && openingStart < cutoff) opening = await readAtmOpening(openingStart, cutoff, filter);
        }
        if (current !== vizoRevision) return;
        if67 = vizoFiles.if67.file ? await readAtmIf67(first, last) : undefined;
        atm = atmReview(provider, [...opening, ...ledger], filter, sign, if67);
      }
    } else {
      const account = normalizeAccount(vizoField("ledger-gl").value);
      requireReport(account, "Enter the settlement G/L account from your monthly reconciliation.");
      const { columns, first, last } = vizoRange("ledger", ["account", "balance"]);
      const matches = await readWisdom(vizoFiles.ledger.book, ledgerSheet, ...columns, first, [account], last, true, month);
      const entry = matches.get(account);
      ending = cents(entry.amount);
      endingSource = ledgerSheet + "!" + entry.address;
    }
    if (current !== vizoRevision) return;
    const allocations = mode === "settlement" ? vizoReviewedDeposits(vizoField("deposit-allocations").value, providerRows, ledger ?? [], end) : [];
    const allocated = new Set(allocations.map(entry => entry.item.row));
    if (mode === "charge" && !chargeLoanStreet && manualTransit === undefined && ledgerType === "wisdom") {
      chargeLoanStreet = chargeLoanStreetFromWisdom(ledger, month, sign, vizoFiles.ledger.file, ledgerSheet);
    }
    const loanStreetTransit = chargeLoanStreet?.current ?? manualTransit;
    const invoiceFile = mode === "charge" ? vizoFiles.invoice.file : undefined;
    const invoice = invoiceFile ? compareVizoLendKeyInvoice(await readVizoLendKeyInvoice(invoiceFile, month), provider, ledger, sign) : undefined;
    if (current !== vizoRevision) return;
    const adjusted = ending * sign + adjustment + (loanStreetTransit ?? 0), difference = statement - adjusted;
    requireReport([adjusted, difference].every(Number.isSafeInteger), "The ending balance exceeds the supported range.");
    vizoResult = { mode, profile, month, end, sign, statement, statementSource, ending, adjustment, adjusted, difference, endingSource,
      branchProvider, boundaryRows, boundaryFile: boundaryFile?.name, transit, noProviderActivity,
      unclassifiedDeposits: unclassifiedDeposits.filter(item => !allocated.has(item.row)), transferReview, allocations,
      priorLoanStreet, loanStreetTransit, chargeLoanStreet, branchOffset, invoice,
      provider, groups: vizoGroups(provider), ledger, atm, if67, opening, providerSheet, ledgerSheet, template: vizoFiles.template.file,
      providerFile: providerLoaded ? vizoFiles.provider.file.name : "", ledgerFile: vizoFiles.ledger.file.name };
    showVizo(vizoResult);
    vizoStatus(atm && !atm.comparisons.length ? "No ATM daily descriptions were recognized. Review Provider detail."
      : mode === "charge" && loanStreetTransit === undefined ? "Verify the LoanStreet in-transit adjustment before using the draft difference."
        : vizoResult.unclassifiedDeposits.length ? "Review unclassified forward collection rows before using the deposit comparisons." : "");
  } catch (error) {
    if (current === vizoRevision) vizoStatus(error.message, true);
  }
});

async function downloadVizo(button, build, suffix) {
  const result = vizoResult;
  if (!result) return;
  button.disabled = true;
  try {
    const blob = await build(result);
    if (result !== vizoResult) return;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "VIZO " + result.profile.name + " " + result.month + suffix(result);
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 30000);
  } catch (error) {
    if (result === vizoResult) vizoStatus(error.message, true);
  } finally {
    button.disabled = false;
  }
}

vizoField("download").addEventListener("click", event =>
  downloadVizo(event.currentTarget, buildVizo, () => ".xlsx"));
vizoField("download-template").addEventListener("click", event =>
  downloadVizo(event.currentTarget, fillVizoTemplate, result => " filled template." + result.template.name.split(".").at(-1)));
