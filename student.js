"use strict";

const scField = id => document.getElementById(id);
const scFiles = { journal: [], fed: [], cover: [], statement: [], wisdom: [], invoice: [] };
let scResult, scRevision = 0;
addTrialBalanceDetection("wisdom", () => scFiles.wisdom.map(file => ({ file })));

const journalHeaders = {
  date: ["effective date"], code: ["transaction code", "tran code", "txn code", "transaction type", "tran type", "code"],
  amount: ["amount"], principal: ["prin"], capitalized: ["cint"], interest: ["int"], fees: ["cfees"], late: ["late"],
};
const fedHeaders = {
  date: ["date", "effective date", "settlement date"], direction: ["direction", "dr/cr", "transaction type"],
  amount: ["amount"], description: ["description", "transaction description"],
};

function scStatus(message, error = false) {
  scField("status").textContent = message;
  scField("status").classList.toggle("error", error);
}

function invalidateSc() {
  scRevision++;
  scResult = undefined;
  scField("confirm").checked = false;
  scField("results").hidden = true;
  scStatus("");
}

function selectScFiles(kind, files) {
  scFiles[kind] = Array.from(files);
  scField(kind + "-status").textContent = files.length === 1 ? files[0].name : files.length ? files.length + " files selected" : "";
  if (kind === "invoice") {
    const pdf = /\.pdf$/i.test(files[0]?.name ?? "");
    scField("invoice-amount").disabled = Boolean(files.length);
    for (const id of ["invoice-sheet", "invoice-cell"]) scField(id).closest("label").hidden = pdf;
  }
  invalidateSc();
}

async function scPdfInvoice(file) {
  const text = await readLocalPdfText(file);
  requireReport(text.some(item => /^CU Student Choice$/i.test(item.text)), "The PDF does not identify a CU Student Choice invoice.");
  const labels = name => text.filter(item => item.text.toUpperCase() === name);
  const [dateHeading] = labels("INVOICE DATE"), [amountHeading] = labels("INVOICE AMOUNT");
  const [customerHeading] = labels("CUSTOMER NO.");
  requireReport(dateHeading && amountHeading && customerHeading && dateHeading.x < amountHeading.x && amountHeading.x < customerHeading.x
    && Math.abs(dateHeading.y - amountHeading.y) < 2 && Math.abs(amountHeading.y - customerHeading.y) < 2,
  "The invoice remittance headings could not be verified.");
  const dates = text.filter(item => item.x >= dateHeading.x - 2 && item.x < amountHeading.x
    && item.y < dateHeading.y && dateHeading.y - item.y < 35 && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(item.text));
  requireReport(dates.length === 1, "The PDF invoice date is missing or ambiguous.");
  const date = sourceDate(dates[0].text);
  const totals = labels("TOTAL").filter(item => item.y > amountHeading.y);
  requireReport(totals.length === 1, "The printed invoice total is missing or ambiguous.");
  const printedAmounts = text.filter(item => item.x > totals[0].x && Math.abs(item.y - totals[0].y) < 2
    && Number.isFinite(readBalance(item.text)));
  const remittanceAmounts = text.filter(item => item.x >= amountHeading.x - 2 && item.x < customerHeading.x
    && Math.abs(item.y - dates[0].y) < 2 && Number.isFinite(readBalance(item.text)));
  requireReport(printedAmounts.length === 1 && remittanceAmounts.length === 1
    && cents(printedAmounts[0].text) === cents(remittanceAmounts[0].text), "The printed and remittance invoice totals could not be confirmed.");
  return { total: cents(printedAmounts[0].text), invoiceDate: date };
}

async function scInvoice(fed, month) {
  const file = scFiles.invoice[0], raw = scField("invoice-amount").value.trim();
  const date = scField("invoice-date").value;
  if (!file && !raw) {
    requireReport(!date, "Enter the invoice amount to check an expected Fed date.");
    return undefined;
  }
  requireReport(!date || date.slice(0, 7) === month, "The expected Fed debit date must be in the selected reconciliation month.");
  let total, invoiceDate, source = "Entered manually";
  if (/\.pdf$/i.test(file?.name ?? "")) {
    ({ total, invoiceDate } = await scPdfInvoice(file));
    source = file.name + " · page 1 · Total";
  } else if (file) {
    const book = await openExcel(file, true);
    const sheet = scField("invoice-sheet").value.trim() || (book.sheets.size === 1 ? [...book.sheets.keys()][0] : "");
    const address = scField("invoice-cell").value.trim().toUpperCase();
    const cell = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(address);
    requireReport(book.sheets.has(sheet) && cell && columnNumber(cell[1]) <= 16384 && Number(cell[2]) <= 1048576,
      "Choose the invoice worksheet and its verified total cell.");
    const cells = await readCells(book, sheet);
    total = Math.abs(cents(cells.get(address)));
    source = file.name + " · " + sheet + "!" + address;
  } else {
    total = Math.abs(cents(raw));
  }
  requireReport(total > 0, "Enter a nonzero invoice total.");
  const candidates = fed.filter(item => item.cash === -total && (!date || item.date === date));
  const match = candidates.length === 1 ? candidates[0] : undefined;
  const review = match ? "One exact Fed debit · verify" : candidates.length ? "Multiple exact Fed debits · review" : "No exact Fed debit";
  return { total, source, date, invoiceDate, candidates, match, review };
}

function scOffset(date, days) {
  const value = new Date(date + "T00:00:00Z");
  value.setUTCDate(value.getUTCDate() + days);
  return value.toISOString().slice(0, 10);
}

async function scJournalFiles(files) {
  const journals = [];
  for (const file of files) {
    if (!/\.zip$/i.test(file.name)) {
      journals.push({ file, name: file.name });
      continue;
    }
    requireReport(file.size <= 20 * 1024 * 1024, file.name + " exceeds the ZIP import limit.");
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const members = Object.values(zip.files).filter(entry => !entry.dir && /\.(?:xls[xm]|csv)$/i.test(entry.name)
      && !entry.name.startsWith("__MACOSX/") && !entry.name.split("/").at(-1).startsWith("~$"));
    requireReport(members.length, file.name + " contains no Excel or CSV journals.");
    for (const entry of members) {
      const name = entry.name.split("/").at(-1);
      const bytes = await entry.async("uint8array");
      journals.push({ file: new File([bytes], name), name: file.name + " · " + name });
    }
  }
  return journals;
}

function scFindHeader(cells, aliases) {
  const rows = new Map();
  for (const [address, raw] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (!match || Number(match[2]) > 50 || typeof raw !== "string") continue;
    const value = raw.trim().toLowerCase().replace(/\s+/g, " ");
    const key = Object.keys(aliases).find(field => aliases[field].includes(value));
    if (!key) continue;
    const row = Number(match[2]);
    const fields = rows.get(row) ?? new Map();
    fields.set(key, fields.has(key) ? "" : match[1]);
    rows.set(row, fields);
  }
  return [...rows].find(([, fields]) => fields.size === Object.keys(aliases).length && [...fields.values()].every(Boolean));
}

async function scReadRows(file, source, kind) {
  const aliases = kind === "journal" ? journalHeaders : fedHeaders;
  const candidates = [];
  let date1904 = false;
  const add = (sheet, cells) => {
    const header = scFindHeader(cells, aliases);
    if (header) candidates.push({ sheet, cells, header });
  };
  if (kind === "journal" && /\.csv$/i.test(file.name)) {
    requireReport(file.size <= 20 * 1024 * 1024, source + " exceeds the CSV import limit.");
    add("CSV", readCsvCells(await file.text()));
  } else {
    const book = await openExcel(file, true);
    date1904 = book.date1904;
    const names = [...book.sheets.keys()];
    const preferred = kind === "journal" ? names.filter(name => /transaction journal/i.test(name)) : [];
    for (const sheet of preferred.length ? preferred : names) {
      const cells = await readCells(book, sheet);
      if (kind === "fed" && names.length === 1) {
        const history = providerFedHistory(cells, date1904, "Student Choice");
        if (history) return history.map(item => ({ date: item.date, posted: item.posted,
          amount: Math.abs(item.amount), cash: -item.amount, description: item.description,
          invoice: /\binvoice\b/i.test(item.description), source, sheet, row: item.row, financeProvider: "Student Choice" }));
      }
      add(sheet, cells);
    }
  }
  requireReport(candidates.length === 1,
    source + (candidates.length ? " has several worksheets with matching headers." : " has no worksheet with the required headings."));
  const { sheet, cells, header: [first, fields] } = candidates[0];
  const financeColumn = kind === "fed" ? financeProviderColumn(cells, first) : undefined;
  const transactionColumns = kind === "journal" ? [...cells].flatMap(([address, value]) => {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    return Number(match?.[2]) === first && String(value).trim().toLowerCase() === "transaction id" ? [match[1]] : [];
  }) : [];
  requireReport(transactionColumns.length <= 1, "Use one Transaction ID column in the journal.");
  const [transactionColumn] = transactionColumns;
  const rowNumbers = new Set();
  const cols = new Set([...fields.values(), financeColumn].filter(Boolean));
  for (const [address, value] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (match && Number(match[2]) > first && cols.has(match[1]) && String(value ?? "").trim()) rowNumbers.add(Number(match[2]));
  }
  return Array.from(rowNumbers).sort((a, b) => a - b).flatMap(row => {
    if (financeColumn && !matchesFinanceProvider(cells.get(financeColumn + row), "Student Choice")) return [];
    const get = field => cells.get(fields.get(field) + row);
    const optionalMoney = field => String(get(field) ?? "").trim() ? cents(get(field)) : 0;
    try {
      const date = sourceDate(get("date"), date1904);
      const amount = cents(get("amount"));
      if (kind === "fed") {
        const direction = String(get("direction") ?? "").trim().toUpperCase();
        requireReport(["CR", "CREDIT", "DR", "DB", "DEBIT"].includes(direction), "Use a debit or credit direction.");
        const description = String(get("description") ?? "").trim();
        const cash = Math.abs(amount) * (direction === "CR" || direction === "CREDIT" ? 1 : -1);
        return { date, amount, cash, description, invoice: /\binvoice\b/i.test(description), source, sheet, row,
          financeProvider: financeColumn ? String(cells.get(financeColumn + row)).trim() : "" };
      }
      const code = String(get("code") ?? "").trim().toUpperCase();
      requireReport(code, "The transaction code is missing.");
      const principal = optionalMoney("principal"), capitalized = optionalMoney("capitalized"), interest = optionalMoney("interest");
      const collectedFees = optionalMoney("fees"), late = optionalMoney("late");
      const payment = code === "PMT" || code === "PMTR";
      const fees = collectedFees + (payment ? late : 0);
      const cash = payment ? -(principal + capitalized + interest + fees) : code === "DIS" ? -amount : 0;
      requireReport(Number.isSafeInteger(cash), "The amount exceeds the supported range.");
      const transactionId = transactionColumn ? String(cells.get(transactionColumn + row) ?? "").trim() : "";
      return { date, code, amount, principal, capitalized, interest, fees, late, cash, source, sheet, row, payment, transactionId };
    } catch (error) {
      throw new Error(source + " · " + sheet + " row " + row + ": " + error.message);
    }
  });
}

function scCompare(journal, fed, from, to, window, invoice) {
  const entries = journal.filter(item => item.date >= from && item.date <= to).sort((a, b) => a.date.localeCompare(b.date));
  const outside = journal.length - entries.length;
  const fingerprint = item => JSON.stringify([item.date, item.code, item.amount, item.principal, item.capitalized, item.interest, item.fees, item.late]);
  const review = new Map();
  const byId = new Map();
  for (const entry of journal) {
    if (!entry.transactionId) continue;
    const rows = byId.get(entry.transactionId) ?? [];
    rows.push(entry);
    byId.set(entry.transactionId, rows);
  }
  for (const rows of byId.values()) {
    if (rows.length < 2) continue;
    const identical = rows.every(entry => fingerprint(entry) === fingerprint(rows[0]));
    for (const entry of identical ? rows.slice(1) : rows) {
      review.set(entry, identical ? "Repeated transaction ID · counted once" : "Conflicting values for one transaction ID · review");
    }
  }
  const duplicates = new Map();
  for (const entry of entries) {
    if (review.has(entry)) continue;
    const key = fingerprint(entry), group = duplicates.get(key) ?? { count: 0, missingId: false };
    group.count++;
    group.missingId ||= !entry.transactionId;
    duplicates.set(key, group);
  }
  const usable = entries.filter(entry => {
    if (review.has(entry)) return false;
    const duplicatesAtAmount = duplicates.get(fingerprint(entry));
    if (duplicatesAtAmount.count > 1 && duplicatesAtAmount.missingId) review.set(entry, "Repeated journal values · review overlap");
    else if (!entry.cash) review.set(entry, "Review transaction code or zero settlement amount");
    return !review.has(entry);
  });
  const possible = new Map(), uses = new Map();
  const invoiceCandidates = new Set(invoice?.candidates ?? []);
  for (const item of fed) {
    if (item.invoice || invoiceCandidates.has(item) || !item.cash) continue;
    const earliest = scOffset(item.date, -window);
    const candidates = usable.filter(entry => entry.date >= earliest && entry.date <= item.date && Math.sign(entry.cash) === Math.sign(item.cash));
    const matches = [];
    const target = Math.abs(item.cash);
    let start = 0, total = 0;
    for (const [end, candidate] of candidates.entries()) {
      total += Math.abs(candidate.cash);
      requireReport(Number.isSafeInteger(total), "A candidate group exceeds the supported amount range.");
      while (total > target) total -= Math.abs(candidates[start++].cash);
      if (total === target) matches.push(candidates.slice(start, end + 1));
    }
    // A funding debit can have the same amount as another disbursement in the lookback window.
    // Use the effective date to distinguish them when an exact same-day disbursement group exists.
    const sameDayFunding = item.posted && item.cash < 0
      ? matches.filter(group => group.every(entry => entry.code === "DIS" && entry.date === item.date)) : [];
    const candidatesToReview = sameDayFunding.length ? sameDayFunding : matches;
    possible.set(item, candidatesToReview);
    for (const entry of new Set(candidatesToReview.flat())) uses.set(entry, (uses.get(entry) ?? 0) + 1);
  }
  const matched = new Map(), fedReview = new Map(), groups = [];
  for (const item of fed) {
    const candidates = possible.get(item) ?? [];
    const group = candidates[0];
    if (candidates.length === 1 && group.every(entry => uses.get(entry) === 1)) {
      const id = groups.length + 1;
      groups.push({ id, fed: item, journal: group });
      group.forEach(entry => matched.set(entry, id));
    } else {
      fedReview.set(item, invoiceCandidates.has(item) ? (item === invoice.match ? "Exact invoice debit · verify" : "Possible invoice debit · review")
        : item.invoice ? "Billing invoice · review separately"
        : !candidates.length ? "No exact chronological group" : "Multiple or overlapping groups");
      for (const entry of candidates.flat()) review.set(entry, "Possible Fed match · verify group");
    }
  }
  const carry = usable.filter(entry => !matched.has(entry) && !review.has(entry) && entry.payment);
  for (const entry of usable) {
    if (!matched.has(entry) && !review.has(entry) && !entry.payment) review.set(entry, "Unmatched disbursement · review");
  }
  const principal = carry.reduce((sum, entry) => sum + entry.principal + entry.capitalized, 0);
  const interest = carry.reduce((sum, entry) => sum + entry.interest, 0);
  const fees = carry.reduce((sum, entry) => sum + entry.fees, 0);
  requireReport([principal, interest, fees].every(Number.isSafeInteger), "Carryover totals exceed the supported range.");
  for (const entry of review.keys()) {
    if (entry.date < from || entry.date > to) review.delete(entry);
  }
  return { entries, outside, fed, groups, matched, review, fedReview, carry, principal, interest, fees };
}

async function scStatement(file, month) {
  const book = await openExcel(file, true);
  const sections = [
    ["Principal", "principal", { beginning: ["beginning balance", "starting balance"], disbursed: ["disbursed"], cancelled: ["cancelled", "canceled"], collected: ["collected"] }],
    ["Capitalized Interest", "capitalized", { beginning: ["beginning balance", "starting balance"], capitalized: ["capitalized"], collected: ["collected"] }],
    ["Interest", "interest", { beginning: ["beginning balance", "starting balance"], accrued: ["accrued"], capitalized: ["capitalized"], collected: ["collected"] }],
  ];
  return Object.fromEntries(await Promise.all(sections.map(async ([name, key, headers]) => {
    const sheet = [...book.sheets.keys()].find(item => item.trim().toLowerCase() === name.toLowerCase());
    requireReport(sheet, "The statement needs a " + name + " worksheet.");
    const cells = await readCells(book, sheet);
    const monthlyHeader = scFindHeader(cells, { ...headers, year: ["year"], month: ["month"] });
    const header = monthlyHeader ?? scFindHeader(cells, headers);
    const rowField = scField(key + "-total");
    let totals = rowField.valueAsNumber;
    if (monthlyHeader) {
      const [headerRow, columns] = monthlyHeader;
      const [year, periodMonth] = month.split("-").map(Number);
      const yearColumn = new RegExp("^" + columns.get("year") + "([1-9]\\d*)$");
      const matches = [];
      for (const [address, value] of cells) {
        const row = Number(yearColumn.exec(address)?.[1]);
        if (row > headerRow && Number(value) === year && Number(cells.get(columns.get("month") + row)) === periodMonth) matches.push(row);
      }
      requireReport(matches.length === 1, "The " + name + " worksheet needs one row for " + month + ". Found " + matches.length + ".");
      requireReport(!rowField.value || totals === matches[0], "The selected totals row on " + name + " does not match " + month + ".");
      totals = matches[0];
    }
    requireReport(header && Number.isInteger(totals) && totals > header[0] && totals <= 1048576,
      "Check the " + name + " column headings and choose a totals row below them.");
    const fields = Object.fromEntries(Object.keys(headers).map(field => {
      const column = header[1].get(field);
      const address = column + totals;
      try { return [field, { amount: cents(cells.get(address)), source: sheet + "!" + address }]; }
      catch { throw new Error("Use a verified numeric statement total at " + sheet + "!" + address + ". Enter zero in the source if it does not apply."); }
    }));
    const optional = key === "principal" ? { ending: ["ending balance"], refund: ["refunded", "refunds"] } : { ending: ["ending balance"] };
    for (const [field, labels] of Object.entries(optional)) {
      const columns = [...cells].flatMap(([address, value]) => {
        const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
        return Number(match?.[2]) === header[0] && labels.includes(String(value ?? "").trim().toLowerCase()) ? [match[1]] : [];
      });
      if (columns.length !== 1) continue;
      const address = columns[0] + totals, value = cells.get(address);
      if (!String(value ?? "").trim()) continue;
      try { fields[field] = { amount: cents(value), source: sheet + "!" + address }; }
      catch { throw new Error("Check the statement amount at " + sheet + "!" + address + "."); }
    }
    return [key, fields];
  })));
}

async function scCover(file, month, comparison) {
  const formulas = new Set();
  const cells = await readCells(await openExcel(file, true), "Cover Page", formulas);
  const period = new Intl.DateTimeFormat("en-US", { month: "long", year: "numeric", timeZone: "UTC" })
    .format(new Date(month + "-01T00:00:00Z"));
  requireReport(String(cells.get("B5") ?? "").toLowerCase().includes(period.toLowerCase()), "Check the period in Cover Page!B5.");
  for (const [row, expected] of [[10, "starting balance"], [17, "ending balance"], [19, "in transit"], [29, "starting interest"], [35, "in transit"]]) {
    requireReport(String(cells.get("C" + row) ?? "").trim().toLowerCase() === expected, "Check the cover layout at C" + row + ".");
  }
  const inputs = new Map();
  const setInput = (row, label, amount, source) => {
    const cell = "K" + row;
    requireReport(!formulas.has(cell), "The proposed input would replace a formula at Cover Page!" + cell + ".");
    const previous = cells.get(cell);
    const saved = typeof previous === "number" ? cents(previous) : undefined;
    const change = saved === undefined && previous !== undefined ? undefined : amount - (saved ?? 0);
    requireReport(Number.isSafeInteger(amount) && (change === undefined || Number.isSafeInteger(change)),
      "A proposed cover amount or change exceeds the supported range at Cover Page!" + cell + ".");
    inputs.set(row, { label, amount, source, saved, change });
  };
  const statement = scFiles.statement.length ? await scStatement(scFiles.statement[0], month) : undefined;
  if (statement) {
    const { principal, capitalized, interest } = statement;
    setInput(10, "Starting principal including capitalized interest", principal.beginning.amount + capitalized.beginning.amount,
      principal.beginning.source + " and " + capitalized.beginning.source);
    for (const [row, title, entry, reduction] of [
      [11, "New disbursals", principal.disbursed], [12, "Capitalized interest", capitalized.capitalized],
      [13, "Capitalized interest collected", capitalized.collected, true], [14, "Cancelled principal", principal.cancelled, true],
      [16, "Principal received", principal.collected, true], [29, "Starting interest", interest.beginning],
      [30, "New interest", interest.accrued], [31, "Interest capitalized", interest.capitalized, true], [32, "Interest received", interest.collected, true],
    ]) setInput(row, title, reduction ? -Math.abs(entry.amount) : entry.amount, entry.source);
    if (principal.refund && /^refunds?$/i.test(String(cells.get("C15") ?? "").trim()) && !formulas.has("K15")) {
      setInput(15, "Principal refunds", principal.refund.amount, principal.refund.source);
    }
  }
  if (scFiles.wisdom.length) {
    const book = await openExcel(scFiles.wisdom[0], true);
    const sheet = scField("wisdom-sheet").value.trim() || (book.sheets.size === 1 ? [...book.sheets.keys()][0] : "");
    requireReport(book.sheets.has(sheet), "Enter the exact Wisdom worksheet name.");
    const rows = [[18, "Principal Wisdom ending balance"], [34, "Interest Wisdom ending balance"]];
    const accounts = rows.map(([row]) => {
      const label = String(cells.get("B" + row) ?? "").trim();
      return normalizeAccount(/^balance per cusa gl\s+(\S+)$/i.exec(label)?.[1]);
    });
    requireReport(accounts.every(Boolean) && new Set(accounts).size === 2, "Check the two G/L labels at Cover Page!B18 and B34.");
    const last = scField("wisdom-last");
    const balances = await readWisdom(book, sheet, scField("wisdom-account").value.trim().toUpperCase(),
      scField("wisdom-balance").value.trim().toUpperCase(), scField("wisdom-first").valueAsNumber, accounts,
      last.value ? last.valueAsNumber : 1048576, true, month);
    rows.forEach(([row, title], index) => {
      const match = balances.get(accounts[index]);
      setInput(row, title, cents(match.amount), sheet + "!" + match.address);
    });
  }
  if (comparison) {
    setInput(19, "Principal in transit", comparison.principal, "Journal carryover candidates");
    setInput(35, "Interest in transit", comparison.interest, "Journal carryover candidates");
  }
  const get = (row, required = false) => {
    if (inputs.has(row)) return inputs.get(row).amount;
    const value = cells.get("K" + row);
    if (value === undefined && !required && !formulas.has("K" + row)) return 0;
    try { return cents(value); }
    catch { throw new Error("Enter a numeric cover input in K" + row + "."); }
  };
  const sum = (rows, required) => rows.reduce((total, row) => total + get(row, required.includes(row)), 0);
  const principalEnd = sum([10, 11, 12, 13, 14, 15, 16], [10]);
  const interestEnd = sum([29, 30, 31, 32], [29]);
  const principalAdjusted = sum([18, 19, 20, 21], [18]);
  const interestAdjusted = sum([34, 35, 36], [34]);
  requireReport([principalEnd, interestEnd, principalAdjusted, interestAdjusted, principalEnd - principalAdjusted,
    interestEnd - interestAdjusted].every(Number.isSafeInteger), "Cover totals exceed the supported range.");
  const providerChecks = [];
  if (statement?.principal.ending && statement.capitalized.ending) {
    const [principal, capitalized] = [statement.principal.ending, statement.capitalized.ending];
    providerChecks.push({ label: "Principal including capitalized interest", calculated: principalEnd,
      statement: principal.amount + capitalized.amount, source: principal.source + " and " + capitalized.source });
  }
  if (statement?.interest.ending) {
    providerChecks.push({ label: "Accrued interest", calculated: interestEnd,
      statement: statement.interest.ending.amount, source: statement.interest.ending.source });
  }
  requireReport(providerChecks.every(item => Number.isSafeInteger(item.statement) && Number.isSafeInteger(item.calculated - item.statement)),
    "A provider ending balance exceeds the supported range.");
  return { principalEnd, interestEnd, principalAdjusted, interestAdjusted, inputs, providerChecks,
    completion: completionTarget("Cover Page", cells, formulas, "G49", "G53", "K49", "K50") };
}

function scSummary(result) {
  const comparison = result.comparison;
  const currency = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
  const items = comparison ? [
    ["Proposed Fed matches", comparison.groups.length],
    ["Fed entries to review", comparison.fedReview.size],
    ["Journal entries to review", comparison.review.size],
    ["Proposed payment carryovers", comparison.carry.length],
    ["Carryover principal", currency.format(comparison.principal / 100)],
    ["Carryover interest", currency.format(comparison.interest / 100)],
    ["Carryover fees", currency.format(comparison.fees / 100)],
  ] : [];
  if (result.cover) items.push(
    ["Principal variance", currency.format((result.cover.principalEnd - result.cover.principalAdjusted) / 100)],
    ["Interest variance", currency.format((result.cover.interestEnd - result.cover.interestAdjusted) / 100)],
    ["Proposed amounts differing from the prepared cover", [...result.cover.inputs.values()].filter(input => input.change).length]);
  for (const check of result.cover?.providerChecks ?? []) {
    items.push([check.label + " less provider ending balance", currency.format((check.calculated - check.statement) / 100)]);
  }
  if (result.invoice) items.push(["Invoice debit", result.invoice.review]);
  const table = document.createElement("table"), body = table.createTBody();
  for (const [label, value] of items) {
    const row = body.insertRow(), heading = document.createElement("th");
    heading.scope = "row";
    heading.textContent = label;
    row.append(heading);
    row.insertCell().textContent = value;
  }
  scField("summary").replaceChildren(table);
  scField("results").hidden = false;
}

async function buildStudent(result) {
  const audit = reconciliationAudit();
  const { label, number, date, row, worksheet } = excelOutputCells();
  const comparison = result.comparison;
  function sheet(name, headers, data, widths) {
    const columns = "ABCDEFGHIJKLMNOP".slice(0, headers.length);
    const body = row(1, headers.map((title, i) => label(columns[i] + "1", title, 1)).join(""))
      + data.map((values, i) => row(i + 2, values.map((value, j) => {
        const address = columns[j] + (i + 2);
        return typeof value === "number" ? number(address, value) : label(address, value);
      }).join(""))).join("");
    return [name, worksheet("A1:" + columns.at(-1) + (data.length + 1),
      widths.map((width, i) => [i + 1, i + 1, width]), body)];
  }
  const journalRow = item => [item.date, item.amount / 100, item.code, (item.principal + item.capitalized) / 100,
    item.interest / 100, item.fees / 100, item.source, item.sheet, item.row];
  const headers = ["Date", "Amount", "Code", "Principal incl. CINT", "Interest", "Fees", "Source", "Worksheet", "Source row"];
  const widths = [14, 16, 12, 22, 16, 16, 48, 23, 15];
  const summary = [["Period", result.month]];
  if (comparison) {
    const matchedPayments = [...comparison.matched.keys()].filter(item => item.payment);
    summary.push(["Earliest journal date", result.from], ["Match lookback (days)", result.window],
      ["Proposed Fed matches", comparison.groups.length], ["Journal rows outside the period", comparison.outside],
      ["Principal carryover", comparison.principal / 100], ["Interest carryover", comparison.interest / 100], ["Fees carryover", comparison.fees / 100],
      ["Matched PMT principal (PRIN)", matchedPayments.filter(item => item.code === "PMT").reduce((sum, item) => sum + item.principal, 0) / 100],
      ["Matched payment interest", matchedPayments.reduce((sum, item) => sum + item.interest, 0) / 100],
      ["Matched payment fees", matchedPayments.reduce((sum, item) => sum + item.fees, 0) / 100]);
  }
  if (result.cover) {
    summary.push(["Prepared cover", result.coverFile]);
    if (result.statementFile) summary.push(["Month-end statement", result.statementFile]);
    if (result.wisdomFile) summary.push(["Wisdom trial balance", result.wisdomFile]);
    summary.push(["Principal statement ending", result.cover.principalEnd / 100],
      ["Principal adjusted book balance", result.cover.principalAdjusted / 100],
      ["Principal variance", (result.cover.principalEnd - result.cover.principalAdjusted) / 100],
      ["Interest statement ending", result.cover.interestEnd / 100], ["Interest adjusted book balance", result.cover.interestAdjusted / 100],
      ["Interest variance", (result.cover.interestEnd - result.cover.interestAdjusted) / 100]);
  }
  if (result.invoice) summary.push(["Invoice amount", result.invoice.total / 100], ["Invoice source", result.invoice.source],
    ...(result.invoice.invoiceDate ? [["Invoice date", result.invoice.invoiceDate]] : []),
    ["Invoice review", result.invoice.review]);
  const coverInputs = result.cover?.inputs ?? new Map(result.carryoverPending ? [] : [
    [19, { label: "Principal in transit", amount: comparison.principal, source: "Journal carryover candidates" }],
    [35, { label: "Interest in transit", amount: comparison.interest, source: "Journal carryover candidates" }],
  ]);
  const sheets = [sheet("Summary", ["Field", "Value"], summary, [37, 32]),
    sheet("Proposed cover inputs", ["Destination", "Cell", "Input", "Amount", "Source", "Saved cover amount", "Change from cover"],
      [...coverInputs].sort(([a], [b]) => a - b).map(([r, input]) => ["Cover Page", "K" + r, input.label, input.amount / 100, input.source,
        input.saved === undefined ? "" : input.saved / 100, input.change === undefined ? "" : input.change / 100]),
      [21, 14, 47, 18, 49, 20, 20])];
  if (result.cover?.providerChecks?.length) sheets.push(sheet("Provider roll-forward", ["Balance", "Calculated ending", "Provider ending", "Difference", "Provider source"],
    result.cover.providerChecks.map(item => [item.label, item.calculated / 100, item.statement / 100,
      (item.calculated - item.statement) / 100, item.source]), [40, 20, 20, 18, 54]));
  sheets.push(...completionPasteSheets(result.cover?.completion, audit));
  if (!comparison) return createExcelWorkbook(sheets, audit);
  const formula = (address, expression, amount) =>
    `<c r="${address}" s="2"><f>${xmlText(expression)}</f><v>${amount / 100}</v></c>`;
  const journalCells = (item, r) => journalRow(item).slice(0, 6).map((value, j) => {
    const address = "ABCDEF"[j] + r;
    if (j === 0) return date(address, value);
    return typeof value === "number" ? number(address, value) : label(address, value);
  }).join("");
  const matchup = [
    row(1, label("A1", "Student Choice MATCHUP · " + result.month, 1)),
    row(3, [...["Date", "Amount", "Description", "Principal", "Interest", "Fees"]
      .map((title, i) => label("ABCDEF"[i] + "3", title, 1)),
      label("G3", "Difference", 1), label("H3", "FedRec Money Out", 1), label("J3", "FedRec Money In", 1)].join("")),
  ];
  let matchupRow = 4;
  for (const group of comparison.groups) {
    const first = matchupRow, last = first + group.journal.length - 1;
    const code = "C" + first + ":C" + last;
    const sumCode = (criterion, column) => `SUMIF(${code},"${criterion}",${column}${first}:${column}${last})`;
    const journalAmount = group.journal.every(item => item.payment) ? "SUM(D" + first + ":F" + last + ")"
      : group.journal.every(item => item.code === "DIS") ? "SUM(B" + first + ":B" + last + ")"
        : [sumCode("DIS", "B"), ...["D", "E", "F"].map(column => sumCode("<>DIS", column))].join("+");
    const expression = "ROUND(" + journalAmount + "-SUM(H" + first + ":H" + last + ")-SUM(J" + first + ":J" + last + "),2)";
    for (const item of group.journal) {
      const r = matchupRow++;
      const fed = r === last ? number((group.fed.cash < 0 ? "H" : "J") + r, -group.fed.cash / 100) : "";
      const difference = r === last ? formula("G" + r, expression, 0) : "";
      matchup.push(row(r, journalCells(item, r) + difference + fed));
    }
    matchupRow++;
  }
  if (comparison.groups.length) {
    const end = matchupRow - 2, r = matchupRow++;
    const payments = [...comparison.matched.keys()].filter(item => item.payment);
    const paymentTotal = (column, amount) => formula(column + r, ["PMT", "PMTR"].map(code =>
      `SUMIF(C4:C${end},"${code}",${column}4:${column}${end})`).join("+"), amount);
    const fedTotal = (column, sign) => formula(column + r, `SUM(${column}4:${column}${end})`,
      comparison.groups.filter(group => Math.sign(group.fed.cash) === sign).reduce((total, group) => total - group.fed.cash, 0));
    matchup.push(row(r, label("A" + r, "TOTALS", 1)
      + paymentTotal("D", payments.reduce((total, item) => total + item.principal + item.capitalized, 0))
      + paymentTotal("E", payments.reduce((total, item) => total + item.interest, 0))
      + paymentTotal("F", payments.reduce((total, item) => total + item.fees, 0))
      + fedTotal("H", -1) + fedTotal("J", 1)));
  }
  if (result.invoice?.match) {
    const r = matchupRow++;
    matchup.push(row(r, label("G" + r, "Billing invoice") + number("H" + r, -result.invoice.match.cash / 100)));
  }
  if (comparison.carry.length && !result.carryoverPending) {
    matchupRow++;
    matchup.push(row(matchupRow, label("A" + matchupRow, "Transit", 1)));
    matchupRow++;
    const first = matchupRow;
    for (const item of comparison.carry) {
      const r = matchupRow++;
      matchup.push(row(r, journalCells(item, r)));
    }
    const r = matchupRow++;
    matchup.push(row(r, label("A" + r, "TOTALS", 1) + [["D", comparison.principal], ["E", comparison.interest], ["F", comparison.fees]]
      .map(([column, amount]) => formula(column + r, `SUM(${column}${first}:${column}${r - 1})`, amount)).join("")));
  }
  if (result.invoice) {
    const { candidates, review, date: expected } = result.invoice;
    sheets.push(sheet("Invoice review", ["Fed date", "Debit", "Description", "Source row", "Review"],
      candidates.length ? candidates.map(item => [item.date, -item.cash / 100, item.description, item.row, review])
        : [[expected, "", "", "", review]], [17, 19, 45, 13, 33]));
    if (result.invoice.match) {
      const match = result.invoice.match;
      sheets.push(["Invoice paste row", worksheet("G1:H2", [[7, 7, 27], [8, 8, 19]],
        row(1, label("G1", "Reviewed invoice label", 1) + label("H1", "Fed debit", 1))
        + row(2, label("G2", "Billing invoice") + number("H2", -match.cash / 100)))]);
    }
  }
  return createExcelWorkbook([
    ...sheets,
    ["MATCHUP", worksheet("A1:J" + Math.max(matchupRow - 1, 3), [[1, 10, 19]], matchup.join(""))],
    sheet("Journal review", [...headers, "Proposed group or review"], comparison.entries.map(item => [
      ...journalRow(item), comparison.matched.get(item) ?? comparison.review.get(item) ?? "Carryover candidate",
    ]), [...widths, 43]),
    sheet("Fed review", ["Effective or supplied date", "Posted on", "Signed Fed amount", "Description", "Proposed group or review", "Source row", "Finance provider"],
      comparison.fed.map(item => [item.date, item.posted ?? "", item.cash / 100, item.description,
        comparison.groups.find(group => group.fed === item)?.id ?? comparison.fedReview.get(item), item.row, item.financeProvider]), [24, 15, 23, 50, 40, 14, 22]),
    sheet("Journal exceptions", [...headers, "Review"], [...comparison.review].map(([item, reason]) => [...journalRow(item), reason]), [...widths, 45]),
    sheet("Carryover candidates", headers, comparison.carry.map(journalRow), widths),
  ], audit);
}

for (const kind of Object.keys(scFiles)) {
  const input = scField(kind + "-file"), drop = scField(kind + "-drop");
  input.addEventListener("change", () => selectScFiles(kind, input.files));
  drop.addEventListener("dragover", event => { event.preventDefault(); drop.classList.add("active"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("active"));
  drop.addEventListener("drop", event => {
    event.preventDefault();
    drop.classList.remove("active");
    const files = event.dataTransfer.files;
    if (!files.length || (kind !== "journal" && files.length !== 1)) return scStatus("Select one workbook for " + kind + ".", true);
    input.value = "";
    selectScFiles(kind, files);
  });
}
document.addEventListener("dragover", event => event.preventDefault());
document.addEventListener("drop", event => event.preventDefault());
scField("month").addEventListener("input", () => {
  const month = scField("month").value;
  if (month) scField("start").value = scOffset(month + "-01", -7);
});
for (const input of document.querySelectorAll("input:not([type=file]):not(#confirm)")) input.addEventListener("input", invalidateSc);
scField("confirm").addEventListener("change", () => { if (!scField("confirm").checked) invalidateSc(); });

scField("run").addEventListener("click", async () => {
  const current = ++scRevision;
  scResult = undefined;
  scField("results").hidden = true;
  try {
    const month = scField("month").value, from = scField("start").value, window = scField("window").valueAsNumber;
    requireReport(month && scField("confirm").checked, "Choose the month and verify the source inputs.");
    const wantsMatch = scFiles.journal.length || scFiles.fed.length, wantsCover = scFiles.cover.length;
    requireReport(wantsMatch || wantsCover, "Load the journal and Fed files or the monthly cover workbook.");
    requireReport(!(scFiles.statement.length || scFiles.wisdom.length) || wantsCover,
      "Load the monthly cover to map the statement or Wisdom inputs.");
    requireReport(wantsMatch || !(scFiles.invoice.length || scField("invoice-amount").value.trim() || scField("invoice-date").value),
      "Load the journal and Fed files to check the invoice.");
    let comparison, invoice, carryoverPending = false;
    if (wantsMatch) {
      requireReport(from && from <= month + "-01" && Number.isInteger(window) && window >= 0 && window <= 31,
        "Check the earliest journal date and match lookback.");
      requireReport(scFiles.journal.length && scFiles.fed.length === 1, "Load both the journal and Fed files to compare activity.");
      const [year, m] = month.split("-").map(Number);
      const to = new Date(Date.UTC(year, m, 0)).toISOString().slice(0, 10);
      const through = scField("fed-through").value || to;
      requireReport(through >= to, "The Fed activity ending date must include the last day of the reconciliation month.");
      const sources = await scJournalFiles(scFiles.journal);
      const journal = [];
      for (const source of sources) {
        if (current !== scRevision) return;
        journal.push(...await scReadRows(source.file, source.name, "journal"));
      }
      if (current !== scRevision) return;
      const fed = (await scReadRows(scFiles.fed[0], scFiles.fed[0].name, "fed"))
        .filter(item => item.date >= month + "-01" && item.date <= through);
      if (current !== scRevision) return;
      requireReport(fed.length && journal.some(item => item.date >= from && item.date <= to), "Check that both sources have rows for the selected period.");
      invoice = await scInvoice(fed, month);
      if (current !== scRevision) return;
      comparison = scCompare(journal, fed, from, to, window, invoice);
      carryoverPending = fed.some(item => item.posted) && [...comparison.fedReview.keys()].some(item => item !== invoice?.match);
    }
    const template = scFiles.cover[0];
    const cover = template ? await scCover(template, month, carryoverPending ? undefined : comparison) : undefined;
    if (current !== scRevision) return;
    scResult = { month, from, window, comparison, cover, invoice, carryoverPending, template, coverFile: template?.name,
      statementFile: scFiles.statement[0]?.name, wisdomFile: scFiles.wisdom[0]?.name };
    const canFill = Boolean(cover && (cover.inputs.size || cover.completion));
    scField("download-template").hidden = !canFill;
    scField("template-download-note").hidden = !canFill;
    scSummary(scResult);
    scStatus(cover?.providerChecks.some(item => item.calculated !== item.statement)
      ? "A calculated ending balance differs from the provider statement. Review Provider roll-forward."
      : carryoverPending ? "Review unmatched Fed entries before using the journal carryovers on the cover." : "");
  } catch (error) {
    if (current === scRevision) scStatus(error.message, true);
  }
});

function fillStudentTemplate(result) {
  const audit = reconciliationAudit();
  const updates = [...result.cover.inputs].map(([row, input]) => ["Cover Page", "K" + row, input.amount / 100]);
  if (result.cover.completion) {
    const { sheet, initials, date } = result.cover.completion;
    updates.push([sheet, initials, audit.by], [sheet, date, { date: audit.on }]);
  }
  return fillExcelTemplate(result.template, updates);
}

async function downloadStudent(button, build, name) {
  const result = scResult;
  if (!result) return;
  button.disabled = true;
  try {
    const blob = await build(result);
    if (result !== scResult) return;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = name(result);
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 30000);
  } catch (error) {
    if (result === scResult) scStatus(error.message, true);
  } finally {
    button.disabled = false;
  }
}

scField("download").addEventListener("click", event =>
  downloadStudent(event.currentTarget, buildStudent, result => "Student Choice " + result.month + ".xlsx"));
scField("download-template").addEventListener("click", event =>
  downloadStudent(event.currentTarget, fillStudentTemplate,
    result => "Student Choice " + result.month + " filled template." + result.template.name.split(".").at(-1)));
