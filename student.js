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

const scMonthStart = monthEnd => monthEnd.slice(0, 8) + "01";

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
  // Refunds (RFD) and monthly fee rows (MFA) have their own places in the monthly block rather than Fed groups.
  const refunds = [], fees = [];
  for (const entry of entries) {
    if (review.has(entry)) continue;
    if (entry.code === "RFD" && entry.date >= scMonthStart(to)) refunds.push(entry);
    if (entry.code === "MFA") fees.push(entry);
  }
  // A payment correction posts a reversal (PMTR) and a re-posted payment (PMT) for the same date and amount in one
  // journal. The pair nets to zero, so it is left out of the Fed matching and kept together with that journal's group.
  const pairs = [];
  const paired = new Set();
  for (const reversal of entries.filter(entry => entry.code === "PMTR" && !review.has(entry))) {
    const options = entries.filter(entry => entry.code === "PMT" && !paired.has(entry) && !review.has(entry)
      && entry.date === reversal.date && entry.amount === -reversal.amount);
    const repost = options.find(entry => entry.source === reversal.source) ?? (options.length === 1 ? options[0] : undefined);
    if (!repost) continue;
    paired.add(reversal).add(repost);
    pairs.push([reversal, repost]);
  }
  const usable = entries.filter(entry => {
    if (review.has(entry) || paired.has(entry) || refunds.includes(entry) || fees.includes(entry)) return false;
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
  // Each correction pair joins the first group holding another entry from the journal that posted it.
  for (const pair of pairs) {
    const group = groups.find(item => item.journal.some(entry => entry.source === pair[0].source));
    if (group) {
      group.journal.push(...pair);
      pair.forEach(entry => matched.set(entry, group.id));
    }
  }
  const carry = [...usable, ...pairs.flat()].filter(entry => !matched.has(entry) && !review.has(entry) && entry.payment)
    .sort((a, b) => a.date.localeCompare(b.date));
  const priorMonth = [];
  for (const entry of usable) {
    if (matched.has(entry) || review.has(entry) || entry.payment) continue;
    // A disbursement dated before the month that no Fed entry this month settles was matched in last month's block.
    if (entry.date < scMonthStart(to)) priorMonth.push(entry);
    else review.set(entry, "Unmatched disbursement · review");
  }
  // Fee rows dated before the last matched payment belong with the invoice line; later ones are in transit.
  const lastMatched = [...matched.keys()].filter(entry => entry.payment).map(entry => entry.date).sort().at(-1) ?? "";
  const invoiceFees = fees.filter(entry => entry.date <= lastMatched), transitFees = fees.filter(entry => entry.date > lastMatched);
  const principal = carry.reduce((sum, entry) => sum + entry.principal + entry.capitalized, 0);
  const interest = carry.reduce((sum, entry) => sum + entry.interest, 0);
  const carryFees = carry.reduce((sum, entry) => sum + entry.fees, 0);
  requireReport([principal, interest, carryFees].every(Number.isSafeInteger), "Carryover totals exceed the supported range.");
  for (const entry of review.keys()) {
    if (entry.date < from || entry.date > to) review.delete(entry);
  }
  return { entries, outside, fed, groups, matched, review, fedReview, carry, principal, interest, fees: carryFees,
    refunds, invoiceFees, transitFees, pairs, priorMonth };
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
  const periodOf = first => new Intl.DateTimeFormat("en-US", { month: "long", year: "numeric", timeZone: "UTC" })
    .format(new Date(first + "T00:00:00Z"));
  const period = periodOf(month + "-01"), previousPeriod = periodOf(scOffset(month + "-01", -1).slice(0, 8) + "01");
  // The cover may already show this month, or be last month's workbook (the filled copy then updates the title).
  const title = String(cells.get("B5") ?? "");
  requireReport([period, previousPeriod].some(item => title.toLowerCase().includes(item.toLowerCase())), "Check the period in Cover Page!B5.");
  const titleUpdate = title.toLowerCase().includes(period.toLowerCase()) ? undefined
    : title.replace(new RegExp(previousPeriod, "i"), period);
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
    // The monthly cover carries new disbursements (and any capitalized interest) on the Capitalized line, K12.
    setInput(12, "New disbursals and capitalized interest", principal.disbursed.amount + capitalized.capitalized.amount,
      principal.disbursed.source + " and " + capitalized.capitalized.source);
    for (const [row, title, entry, reduction] of [
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
  return { principalEnd, interestEnd, principalAdjusted, interestAdjusted, inputs, providerChecks, titleUpdate,
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

// Source rows in the latest monthly MATCHUP block of the prepared workbook, by kind of row. A block starts with its date in
// column G (merged down three rows) and the DATE/AMOUNT header three rows below.
function scBlockStyleRows({ cells, merges }) {
  const text = address => String(cells.get(address) ?? "").trim().toUpperCase();
  const starts = merges.filter(m => m.c1 === "G" && m.c2 === "G" && m.r2 > m.r1 && text("A" + (m.r1 + 3)) === "DATE")
    .map(m => m.r1).sort((a, b) => a - b);
  const start = starts.at(-1);
  if (!start) return undefined;
  const rows = sheetRows(cells).filter(r => r > start + 3);
  const transitLabel = rows.find(r => text("A" + r) === "TRANSIT");
  const end = rows.find(r => r > (transitLabel ?? start) && text("A" + r) === "TOTALS") ?? rows.at(-1);
  const body = rows.filter(r => r <= end);
  const code = r => text("C" + r);
  const all = test => { const found = body.filter(test); return found.length ? found : undefined; };
  const beforeTransit = r => !transitLabel || r < transitLabel;
  return {
    start, end, date: start, band2: start + 1, band3: start + 2, header: start + 3,
    payment: all(r => ["PMT", "PMTR"].includes(code(r)) && !cells.has("J" + r) && beforeTransit(r)),
    groupEnd: all(r => ["PMT", "PMTR"].includes(code(r)) && cells.has("J" + r) && beforeTransit(r)),
    dis: all(r => code(r) === "DIS"), rfd: all(r => code(r) === "RFD"),
    totals: body.find(r => !cells.has("A" + r) && cells.has("D" + r) && cells.has("E" + r) && beforeTransit(r)),
    invoice: body.find(r => code(r) === "MFA" && cells.has("G" + r)),
    fee: all(r => code(r) === "MFA" && !cells.has("G" + r) && beforeTransit(r)),
    transitLabel, transitTotals: transitLabel && end,
    transit: all(r => transitLabel && r > transitLabel && r < end),
  };
}

// The monthly MATCHUP block in the workbook's layout, built from row 1: date, header, payment groups with their Fed
// amounts, disbursements, refunds, totals, the invoice fee line, Transit, and the journal-entry table to the right.
function scMonthlyBlock(result) {
  const { label, number, date } = excelOutputCells();
  const comparison = result.comparison;
  const grid = new Map(), roles = new Map();
  const put = (r, column, xml) => { if (!grid.has(r)) grid.set(r, new Map()); grid.get(r).set(column, xml); };
  const formula = (address, expression, cents) => `<c r="${address}" s="2"><f>${xmlText(expression)}</f><v>${cents / 100}</v></c>`;
  const money = (r, column, cents) => { if (cents) put(r, column, number(column + r, cents / 100)); };
  const journal = (item, r, role) => {
    put(r, "A", date("A" + r, item.date));
    put(r, "B", number("B" + r, item.amount / 100));
    put(r, "C", label("C" + r, item.code));
    money(r, "D", item.principal + item.capitalized);
    money(r, "E", item.interest);
    money(r, "F", item.code === "MFA" ? item.amount : item.fees);
    roles.set(r, role);
  };
  const [year, monthNumber] = result.month.split("-").map(Number);
  put(1, "G", date("G1", new Date(Date.UTC(year, monthNumber, 0)).toISOString().slice(0, 10)));
  roles.set(1, "date").set(2, "band2").set(3, "band3");
  [["A", "DATE"], ["B", "AMOUNT"], ["C", "DESCRIPTION"], ["D", "PRINCIPAL"], ["E", "INTEREST"], ["F", "FEES"],
    ["H", "FEDREC MONEY OUT"], ["J", "FEDREC MONEY IN"]].forEach(([column, title]) => put(4, column, label(column + "4", title)));
  roles.set(4, "header");
  const byDate = items => [...items].sort((a, b) => a.date.localeCompare(b.date) || (a.row - b.row));
  let r = 5;
  const paymentRows = [];
  // Payment groups: each ends with the check (G) and the Fed money in (J, as a negative like the monthly workbook).
  for (const group of comparison.groups.filter(item => item.fed.cash > 0)) {
    const entries = [...byDate(group.journal.filter(item => !comparison.pairs.flat().includes(item))),
      ...group.journal.filter(item => comparison.pairs.flat().includes(item))];
    const first = r;
    entries.forEach((item, i) => { journal(item, r, i === entries.length - 1 ? "groupEnd" : "payment"); paymentRows.push(r); r++; });
    const last = r - 1, fees = entries.some(item => item.fees);
    const total = entries.reduce((sum, item) => sum + item.principal + item.capitalized + item.interest + item.fees, 0);
    put(last, "G", formula("G" + last, `SUM(D${first}:${fees ? "F" : "E"}${last})-J${last}`, total + group.fed.cash));
    put(last, "J", number("J" + last, -group.fed.cash / 100));
    r++;
  }
  // Disbursements with the Fed money out on the last row of each funding, then one check over all of them.
  const disRows = [];
  for (const group of comparison.groups.filter(item => item.fed.cash < 0)) {
    for (const [i, item] of byDate(group.journal).entries()) {
      journal(item, r, "dis");
      if (i === group.journal.length - 1) put(r, "H", number("H" + r, -group.fed.cash / 100));
      disRows.push(r++);
    }
  }
  if (disRows.length) {
    const [first, last] = [disRows[0], disRows.at(-1)];
    const disbursed = comparison.groups.filter(item => item.fed.cash < 0).flatMap(item => item.journal).reduce((sum, item) => sum + item.amount, 0);
    const fedOut = comparison.groups.filter(item => item.fed.cash < 0).reduce((sum, item) => sum - item.fed.cash, 0);
    put(last, "G", formula("G" + last, `SUM(B${first}:B${last})-SUM(H${first}:H${last})`, disbursed - fedOut));
    r++;
  }
  const refundRows = [];
  for (const item of byDate(comparison.refunds)) { journal(item, r, "rfd"); refundRows.push(r++); }
  // Totals of the matched payments, and the invoice fee line just below them.
  const totals = r++;
  roles.set(totals, "totals");
  const span = column => paymentRows.length ? `${column}${paymentRows[0]}:${column}${paymentRows.at(-1)}` : `${column}${totals}:${column}${totals}`;
  const payments = comparison.groups.filter(item => item.fed.cash > 0).flatMap(item => item.journal);
  const sum = key => payments.reduce((total, item) => total + key(item), 0);
  put(totals, "D", formula("D" + totals, `SUM(${span("D")})`, sum(item => item.principal + item.capitalized)));
  put(totals, "E", formula("E" + totals, `SUM(${span("E")})`, sum(item => item.interest)));
  put(totals, "F", formula("F" + totals, `-SUM(${span("F")})`, -sum(item => item.fees)));
  put(totals, "J", formula("J" + totals, `-SUM(${span("J")})`, comparison.groups.filter(item => item.fed.cash > 0).reduce((total, item) => total + item.fed.cash, 0)));
  const invoiceRow = r;
  const feeRows = [];
  byDate(comparison.invoiceFees).forEach((item, i) => { journal(item, r, i ? "fee" : "invoice"); feeRows.push(r++); });
  if (result.invoice?.match) {
    // The invoice paid in the month is labeled with that month, like the provider's invoice (INV-… August).
    const monthName = new Intl.DateTimeFormat("en-US", { month: "long", timeZone: "UTC" }).format(new Date(Date.UTC(year, monthNumber - 1, 1)));
    put(invoiceRow, "G", label("G" + invoiceRow, monthName + " Invoice"));
    put(invoiceRow, "H", number("H" + invoiceRow, -result.invoice.match.cash / 100));
    roles.set(invoiceRow, "invoice");
    if (!feeRows.length) r++;
  }
  // Transit: the fee total, then this month's entries not yet settled at the Fed, then their totals.
  r += 2;
  const transitLabel = r++;
  put(transitLabel, "A", label("A" + transitLabel, "Transit"));
  const feeTotal = comparison.invoiceFees.reduce((total, item) => total + item.amount, 0);
  put(transitLabel, "F", formula("F" + transitLabel, feeRows.length ? `SUM(F${feeRows[0]}:F${feeRows.at(-1)})` : "0", feeTotal));
  roles.set(transitLabel, "transitLabel");
  const transitRows = [];
  for (const item of byDate([...comparison.carry, ...comparison.transitFees])) { journal(item, r, "transit"); transitRows.push(r++); }
  const end = r;
  put(end, "A", label("A" + end, "TOTALS"));
  roles.set(end, "transitTotals");
  const transit = [...comparison.carry, ...comparison.transitFees];
  for (const [column, amount] of [["B", transit.reduce((t, i) => t + i.amount, 0)], ["D", comparison.principal], ["E", comparison.interest]]) {
    put(end, column, formula(column + end, transitRows.length ? `SUM(${column}${transitRows[0]}:${column}${transitRows.at(-1)})` : "0", amount));
  }
  return { grid, roles, end, rows: { totals, invoice: invoiceRow, dis: disRows, refunds: refundRows, payments: paymentRows } };
}

// The journal-entry table to the right of each block (descriptions and G/L accounts) is copied from the prior block at the
// same row offsets; its amounts point at this block's totals.
function scJournalTable(block, source, roleRows) {
  const cells = new Map();
  if (!source || !roleRows) return cells;
  const { totals, invoice, dis, refunds } = block.rows;
  const amounts = {
    "SC DISBURSEMENTS": dis.length ? `SUM(B${dis[0]}:B${dis.at(-1)})` : "0",
    "SC CANCELLED DIS": "0", "SC PRINCIPAL": "D" + totals, "SC INTEREST": "-E" + totals, "SC CAPITALIZED INT": "0",
    "SC REFUNDS": refunds.length ? `SUM(B${refunds[0]}:B${refunds.at(-1)})` : undefined,
    "LATE PAYMENT FEES": "F" + totals, "SC INVOICE": "H" + invoice,
  };
  for (let offset = 0; offset <= roleRows.end - roleRows.start; offset++) {
    const from = roleRows.start + offset, to = 1 + offset;
    const name = String(source.cells.get("N" + from) ?? "").trim().toUpperCase();
    for (const column of ["L", "M", "N", "O", "P", "Q", "R"]) {
      const style = source.styles.get(column + from), value = source.cells.get(column + from);
      const sourceFormula = source.formulas.get(column + from);
      const address = column + to, s = style === undefined ? "0" : "x" + style;
      if (column === "O" && name in amounts) {
        if (amounts[name] !== undefined) cells.set(address, `<c r="${address}" s="${s}"><f>${xmlText(amounts[name])}</f></c>`);
        else if (style !== undefined) cells.set(address, `<c r="${address}" s="${s}"/>`);
      } else if (["N", "P", "Q"].includes(column) && typeof sourceFormula === "string" && /^O\d+\+O\d+$/i.test(sourceFormula)) {
        // The check of disbursements less cancellations, moved with the table.
        const shifted = sourceFormula.replace(/O(\d+)/gi, (_, row) => "O" + (Number(row) - roleRows.start + 1));
        cells.set(address, `<c r="${address}" s="${s}"><f>${shifted}</f></c>`);
      } else if (["N", "P", "Q"].includes(column) && offset >= 4 && value !== undefined && value !== "" && !sourceFormula) {
        // Descriptions and G/L accounts from the table's header row down (the rows above hold the preparer's Fed
        // amount list, which is month-specific and not copied).
        cells.set(address, typeof value === "number" ? `<c r="${address}" s="${s}"><v>${value}</v></c>`
          : `<c r="${address}" s="${s}" t="inlineStr"><is><t>${xmlText(value)}</t></is></c>`);
      } else if (style !== undefined) {
        cells.set(address, `<c r="${address}" s="${s}"/>`);
      }
    }
  }
  return cells;
}

// The formatted monthly block: borrowed formats for the block rows and the copied journal-entry table.
function scStyledBlock(result) {
  const block = scMonthlyBlock(result);
  const source = result.blockStyles?.sheets.MATCHUP, roleRows = source && scBlockStyleRows(source);
  const { styled, merges, heights } = roleRows
    ? styledBlock(block, source, roleRows, ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"], [["date", 1]])
    : { styled: block.grid, merges: [], heights: new Map() };
  const table = scJournalTable(block, source, roleRows);
  for (const [address, xml] of table) {
    const r = Number(/\d+$/.exec(address)[0]), column = address.replace(/\d+$/, "");
    if (!styled.has(r)) styled.set(r, new Map());
    styled.get(r).set(column, xml);
  }
  const end = Math.max(block.end, ...[...table.keys()].map(address => Number(/\d+$/.exec(address)[0])));
  return { styled, merges, heights, end, cols: roleRows && source.cols.length ? source.cols : [[1, 10, 15]] };
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
  // The month's MATCHUP block in the workbook's layout and formatting, ready to paste below the previous month.
  const block = comparison ? scStyledBlock(result) : undefined;
  if (block) {
    summary.push(["MATCHUP paste block", "A1:R" + block.end]);
    if (comparison.priorMonth.length) summary.push(["Prior-month disbursements left out (matched in last month's block)", comparison.priorMonth.length]);
    if (comparison.pairs.length) summary.push(["Payment corrections kept with their journal's group", comparison.pairs.length]);
  }
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
    ["MATCHUP", worksheet("A1:R" + block.end, block.cols, gridRows(block.styled, block.heights), block.merges)],
    sheet("Journal review", [...headers, "Proposed group or review"], comparison.entries.map(item => [
      ...journalRow(item), comparison.matched.get(item) ?? comparison.review.get(item) ?? "Carryover candidate",
    ]), [...widths, 43]),
    sheet("Fed review", ["Effective or supplied date", "Posted on", "Signed Fed amount", "Description", "Proposed group or review", "Source row", "Finance provider"],
      comparison.fed.map(item => [item.date, item.posted ?? "", item.cash / 100, item.description,
        comparison.groups.find(group => group.fed === item)?.id ?? comparison.fedReview.get(item), item.row, item.financeProvider]), [24, 15, 23, 50, 40, 14, 22]),
    sheet("Journal exceptions", [...headers, "Review"], [...comparison.review].map(([item, reason]) => [...journalRow(item), reason]), [...widths, 45]),
    sheet("Carryover candidates", headers, comparison.carry.map(journalRow), widths),
  ], audit, result.blockStyles && { styles: result.blockStyles.stylesXml, theme: result.blockStyles.theme });
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
    // Formatting is a convenience: if it cannot be read, the block is still produced with plain formatting.
    const blockStyles = template ? await openExcel(template, true).then(book => readWorkbookFormatting(book, ["MATCHUP"])).catch(() => undefined) : undefined;
    if (current !== scRevision) return;
    scResult = { month, from, window, comparison, cover, invoice, carryoverPending, template, coverFile: template?.name, blockStyles,
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
  // The filled workbook must tie to the provider: refuse it when the cover's ending balances, from the values it will hold,
  // differ from the provider statement (for example a line left with last month's number).
  const failing = (result.cover.providerChecks ?? []).filter(item => item.calculated !== item.statement);
  requireReport(!failing.length, "The filled workbook would not tie to the provider statement: " + failing.map(item =>
    item.label + " ends at " + (item.calculated / 100).toFixed(2) + " but the statement shows " + (item.statement / 100).toFixed(2)).join("; ")
    + ". Review Provider roll-forward before downloading.");
  const updates = [...result.cover.inputs].map(([row, input]) => ["Cover Page", "K" + row, input.amount / 100]);
  if (result.cover.titleUpdate) updates.push(["Cover Page", "B5", result.cover.titleUpdate]);
  if (result.cover.completion) {
    const { sheet, initials, date } = result.cover.completion;
    updates.push([sheet, initials, audit.by], [sheet, date, { date: audit.on }]);
  }
  // Add the month's MATCHUP block below the last one, formatted like the previous month.
  const appends = [];
  if (result.comparison && result.blockStyles) {
    const source = result.blockStyles.sheets.MATCHUP;
    const [year, monthNumber] = result.month.split("-").map(Number);
    const serial = time => (time - Date.UTC(1899, 11, 30)) / 86400000;
    const [first, last] = [serial(Date.UTC(year, monthNumber - 1, 1)), serial(Date.UTC(year, monthNumber, 0))];
    const already = source.merges.some(m => m.c1 === "G" && m.c2 === "G" && m.r2 > m.r1
      && typeof source.cells.get("G" + m.r1) === "number" && source.cells.get("G" + m.r1) >= first && source.cells.get("G" + m.r1) <= last
      && String(source.cells.get("A" + (m.r1 + 3)) ?? "").trim().toUpperCase() === "DATE");
    requireReport(!already, "This workbook already has the " + result.month + " MATCHUP block. Use the previous month's workbook.");
    const block = scStyledBlock(result);
    appends.push({ sheet: "MATCHUP", grid: block.styled, merges: block.merges, heights: block.heights, end: block.end });
  }
  return fillExcelTemplate(result.template, updates, appends);
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
