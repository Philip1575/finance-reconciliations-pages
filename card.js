"use strict";

const cardElement = id => document.getElementById(id);
const cardFiles = {
  source: { workbook: undefined, file: undefined, revision: 0, mappingRevision: 0 },
  prior: { workbook: undefined, file: undefined, revision: 0, mappingRevision: 0 },
};
let cardRows = [];
let cardBalanceReview;
let cardCoverReview;
let cardCoverFile;
let cardSelection = new Set();
let cardRevision = 0;
let cardGroup = 0;
let cardDownloading = false;
const cardMoney = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const cardColumns = {
  date: ["effective date", "funding date", "date"], code: ["txn", "txn code", "tran code", "transaction code", "code"],
  amount: ["transaction amount", "signed amount", "amount"], description: ["statement description", "description"],
};

function cardStatus(message, error = false) {
  cardElement("status").textContent = message;
  cardElement("status").classList.toggle("error", error);
}

function cardReviewStatus(message, error = false) {
  cardElement("review-status").textContent = message;
  cardElement("review-status").classList.toggle("error", error);
}

function invalidateCard() {
  cardRevision++;
  cardRows = [];
  cardBalanceReview = undefined;
  cardCoverReview = undefined;
  cardSelection.clear();
  cardElement("review").hidden = true;
  cardElement("confirmed").checked = false;
  cardElement("download").disabled = true;
  cardStatus("");
  cardReviewStatus("");
}

async function openCardHistory(file) {
  if (!/\.xls$/i.test(file.name)) return openExcel(file);
  requireReport(file.size <= 20 * 1024 * 1024, "Choose a legacy account history workbook under 20 MB.");
  const legacy = XLSX.read(await file.arrayBuffer(), { cellFormula: false, cellHTML: false, bookVBA: false });
  const sheets = new Map(legacy.SheetNames.map(name => [name, legacy.Sheets[name]]));
  requireReport(sheets.size, "The account history does not contain a worksheet.");
  return { sheets, legacy: true, date1904: Boolean(legacy.Workbook?.WBProps?.date1904) };
}

function readCardCells(workbook, sheet, formulas) {
  if (!workbook.legacy) return readCells(workbook, sheet, formulas);
  const cells = workbook.sheets.get(sheet);
  requireReport(cells, "Choose an account history worksheet.");
  return new Map(Object.entries(cells).flatMap(([address, cell]) =>
    /^[A-Z]+[1-9]\d*$/.test(address) && ["s", "n"].includes(cell.t) ? [[address, cell.v]] : []));
}

async function suggestCardMapping(kind) {
  const state = cardFiles[kind], { workbook, revision } = state;
  const sheet = cardElement(kind + "-sheet").value, current = ++state.mappingRevision;
  const formulas = new Map(), headers = new Map();
  const cells = await readCardCells(workbook, sheet, formulas);
  if (revision !== state.revision || current !== state.mappingRevision || sheet !== cardElement(kind + "-sheet").value) return;
  for (const [address, value] of cells) {
    const cell = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (!cell || Number(cell[2]) > 50 || typeof value !== "string") continue;
    const text = value.trim().toLowerCase().replace(/\s+/g, " ");
    const key = Object.keys(cardColumns).find(field => cardColumns[field].includes(text));
    if (!key) continue;
    const row = Number(cell[2]), columns = headers.get(row) ?? new Map();
    columns.set(key, columns.has(key) ? "" : cell[1]);
    headers.set(row, columns);
  }
  const matches = [...headers].filter(([, columns]) => columns.size === 4 && [...columns.values()].every(Boolean));
  if (matches.length !== 1) return;
  const [[header, columns]] = matches;
  let last = "";
  if (kind === "prior" && sheet === "Pending Charges") {
    const amount = columns.get("amount"), date = columns.get("date");
    const range = new RegExp("^SUM\\(" + amount + "(\\d+):" + amount + "(\\d+)\\)$");
    const totals = [...formulas].flatMap(([address, formula]) => {
      const row = Number(/^[A-Z]+(\d+)$/.exec(address)?.[1]);
      if (row <= header || String(cells.get(date + row) ?? "").trim().toLowerCase() !== "total") return [];
      const match = range.exec(String(formula).replace(/[\s$]/g, "").toUpperCase());
      return match && Number(match[1]) === header + 1 && Number(match[2]) >= header + 1 && Number(match[2]) < row ? [match[2]] : [];
    });
    if (totals.length === 1) last = totals[0];
  }
  for (const [field, column] of columns) cardElement(kind + "-" + field).value = column;
  cardElement(kind + "-first").value = header + 1;
  cardElement(kind + "-last").value = last;
}

async function selectCardFile(kind, file) {
  const state = cardFiles[kind];
  const current = ++state.revision;
  state.file = file;
  state.workbook = undefined;
  invalidateCard();
  cardElement(kind + "-mapping").hidden = true;
  cardElement(kind + "-status").textContent = file ? "Opening " + file.name + "…" : "";
  if (kind === "prior") cardElement("prior-clear").hidden = !file;
  if (!file) return;
  try {
    const workbook = kind === "source" ? await openCardHistory(file) : await openExcel(file);
    if (state.revision !== current) return;
    state.workbook = workbook;
    const select = cardElement(kind + "-sheet");
    select.replaceChildren(...Array.from(workbook.sheets.keys(), name => new Option(name, name)));
    if (kind === "prior") {
      const sheet = ["Open items", "Pending Charges"].find(name => workbook.sheets.has(name));
      if (sheet) select.value = sheet;
    }
    await suggestCardMapping(kind);
    if (state.revision !== current) return;
    cardElement(kind + "-mapping").hidden = false;
    cardElement(kind + "-status").textContent = file.name;
  } catch (error) {
    if (state.revision === current) cardElement(kind + "-status").textContent = error.message;
  }
}

async function readCardRows(kind, from, through) {
  const { workbook, file } = cardFiles[kind];
  requireReport(workbook, "Choose a readable " + (kind === "source" ? "account history" : "prior open items") + " workbook.");
  const sheet = cardElement(kind + "-sheet").value;
  const columns = ["date", "code", "amount", "description"].map(field => cardElement(kind + "-" + field).value.trim().toUpperCase());
  requireReport(columns.every(column => /^[A-Z]{1,3}$/.test(column) && columnNumber(column) <= 16384) && new Set(columns).size === 4,
    "Choose four different columns for " + kind + ".");
  const first = cardElement(kind + "-first").valueAsNumber;
  const lastInput = cardElement(kind + "-last");
  const last = lastInput.value ? lastInput.valueAsNumber : 1048576;
  requireReport(Number.isInteger(first) && Number.isInteger(last) && first >= 1 && last >= first && last <= 1048576,
    "Check the first and last data rows for " + kind + ".");
  const cells = await readCardCells(workbook, sheet);
  const approvedCarryovers = kind === "prior" && sheet === "Open items" && cells.get("F1") === "Carryover decision";
  const rowNumbers = new Set();
  const relevant = new Set(columns);
  for (const [address, value] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (!match || !relevant.has(match[1]) || !String(value ?? "").trim()) continue;
    const row = Number(match[2]);
    if (row >= first && row <= last) rowNumbers.add(row);
  }
  const rows = [];
  for (const row of Array.from(rowNumbers).sort((a, b) => a - b)) {
    const [dateValue, codeValue, amountValue, descriptionValue] = columns.map(column => cells.get(column + row));
    let date, amount;
    try {
      date = sourceDate(dateValue, workbook.date1904);
      amount = cents(amountValue);
    } catch (error) {
      throw new Error(sheet + " row " + row + ": " + error.message);
    }
    const code = String(codeValue ?? "").trim();
    requireReport(code, sheet + " row " + row + " has no transaction code.");
    const outside = kind === "source" && (date < from || date > through);
    requireReport(kind !== "prior" || date <= through, "A prior open item is dated after the selected ending date.");
    rows.push({
      id: kind + "-" + row, date, code, amount, description: String(descriptionValue ?? "").trim(),
      source: kind, fileName: file.name, sheet, row,
      state: outside || amount === 0 ? "excluded" : "open",
      carryover: amount !== 0 && approvedCarryovers && cells.get("F" + row) === "Confirmed carryover",
      reason: outside ? "Outside selected dates" : amount === 0 ? "Zero amount" : "", group: undefined,
    });
  }
  return rows;
}

async function readCardBalanceReview(month) {
  const { workbook, file } = cardFiles.source;
  const sheet = cardElement("source-sheet").value;
  const cells = await readCardCells(workbook, sheet);
  const header = column => String(cells.get(column + "1") ?? "").trim().replace(/\s+/g, " ").toLowerCase();
  if (!["post date", "transaction amount", "after balance"].every((name, index) => header("CKL"[index]) === name)) return undefined;
  try {
    const history = [...cells].flatMap(([address, value]) => {
      const row = Number(/^C([1-9]\d*)$/.exec(address)?.[1]);
      if (!(row > 1) || !String(value ?? "").trim()) return [];
      return [{ row, posted: sourceDate(value, workbook.date1904), amount: cents(cells.get("K" + row)), balance: cents(cells.get("L" + row)) }];
    }).sort((a, b) => a.row - b.row);
    const chronological = history.every((item, i) => i === 0 || (item.posted >= history[i - 1].posted
      && history[i - 1].balance + item.amount === item.balance));
    const reversed = history.every((item, i) => i === 0 || (item.posted <= history[i - 1].posted
      && item.balance + history[i - 1].amount === history[i - 1].balance));
    requireReport(chronological || reversed, "The report's After Balance amounts do not tie to the posted transactions.");
    if (!chronological) history.reverse();
    const closing = history.findLast(item => item.posted.startsWith(month));
    requireReport(closing, "The account history has no posted balance for the selected month.");
    const opening = history.findLast(item => item.posted < month + "-01");
    const posted = history.filter(item => item.posted.startsWith(month));
    const latest = history.at(-1).posted;
    const [year, period] = month.split("-").map(Number);
    const monthEnd = new Date(Date.UTC(year, period, 0)).toISOString().slice(0, 10);
    return { file: file.name, sheet, closing, opening, count: posted.length,
      movement: posted.reduce((sum, item) => sum + item.amount, 0), verifyCoverage: latest < monthEnd };
  } catch (error) {
    return { file: file.name, sheet, issue: error.message };
  }
}

function cardFingerprint(row) {
  return JSON.stringify([row.date, row.code, row.amount, row.description]);
}

async function cardPriorBoundary(workbook, date) {
  if (!workbook?.sheets.has("Matchup")) return new Map();
  const matchup = await readCells(workbook, "Matchup");
  const previous = dateOffset(date, -1).slice(0, 7);
  const period = /^Corporate Card · (\d{4}-\d{2})$/.exec(String(matchup.get("A1") ?? ""))?.[1];
  const generated = period && workbook.sheets.has("Import review") && workbook.sheets.has("Completion record");
  requireReport(!generated || period === previous, "Use the " + previous + " Corporate Card download for opening items.");
  const audit = generated ? await readCells(workbook, "Import review") : undefined;
  const verified = audit && ["Date", "Code", "Amount", "Description", "Status"].every((heading, i) => audit.get("DEFGH"[i] + "1") === heading);
  const cells = verified ? audit : matchup;
  const evidence = new Map();
  for (const [address, value] of cells) {
    const hit = (verified ? /^(D)([1-9]\d*)$/ : /^([AF])([1-9]\d*)$/).exec(address);
    if (!hit) continue;
    let effective;
    try { effective = sourceDate(value, workbook.date1904); } catch { continue; }
    if (effective !== date) continue;
    const [, column, r] = hit, adjacent = verified ? "EFG" : column === "A" ? "BCD" : "GHI";
    const code = String(cells.get(adjacent[0] + r) ?? "").trim();
    const description = String(cells.get(adjacent[2] + r) ?? "").trim();
    if (!code) continue;
    let amount;
    try { amount = cents(cells.get(adjacent[1] + r)); } catch { continue; }
    const key = cardFingerprint({ date, code, description, amount });
    const status = verified ? String(cells.get("H" + r) ?? "") : "";
    const rows = evidence.get(key) ?? [];
    rows.push({ address: (verified ? "Import review" : "Matchup") + "!" + address, matched: /^Matched group \d+$/i.test(status) });
    evidence.set(key, rows);
  }
  return evidence;
}

function flagCardBoundary(source, prior, evidence) {
  const counts = new Map();
  for (const row of [...source, ...prior]) {
    if (row.state !== "open") continue;
    const key = cardFingerprint(row);
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  for (const row of source) {
    if (row.state !== "open") continue;
    const key = cardFingerprint(row), records = evidence.get(key);
    if (!records) continue;
    row.priorEvidence = records.map(item => item.address).join(", ");
    if (records.length === 1 && records[0].matched && counts.get(key) === 1) {
      row.state = "excluded";
      row.reason = "Already matched in prior reconciliation";
      row.carryover = false;
    } else {
      row.reason = "Check possible prior-month reuse";
    }
  }
}

function renderCard() {
  cardElement("confirmed").checked = false;
  cardElement("download").disabled = true;
  const active = cardRows.filter(row => row.state !== "excluded");
  const open = cardRows.filter(row => row.state === "open");
  const fingerprints = new Map();
  for (const row of active) {
    const key = cardFingerprint(row);
    fingerprints.set(key, (fingerprints.get(key) ?? 0) + 1);
  }
  const duplicates = Array.from(fingerprints.values()).reduce((total, count) => total + Math.max(count - 1, 0), 0);
  const carryovers = open.filter(row => row.carryover);
  const pending = carryovers.filter(row => row.amount > 0);
  const pendingTotal = pending.reduce((sum, row) => sum + row.amount, 0);
  const boundary = cardRows.filter(row => row.priorEvidence).length;
  cardElement("counts").textContent = active.length + " included · " + carryovers.length + " confirmed carryovers · "
    + (open.length - carryovers.length) + " unmatched to review · "
    + (cardRows.length - active.length) + " excluded · " + duplicates + " duplicate rows to review"
    + (boundary ? " · " + boundary + " opening-date rows found in the prior workbook" : "");
  cardElement("pending-preview").textContent = pending.length + " pending funding rows · " + cardMoney.format(pendingTotal / 100)
    + (pending.length > 10 ? " · Extend the destination total to include the new rows." : "");
  cardElement("balance-preview").textContent = cardBalanceReview?.issue ?? (cardBalanceReview?.closing
    ? "Last posted balance for the month: " + cardMoney.format(cardBalanceReview.closing.balance / 100) + " on " + cardBalanceReview.closing.posted
      + (cardBalanceReview.verifyCoverage ? " · Verify month-end coverage" : "") : "");
  const table = document.createElement("table");
  const heading = table.createTHead().insertRow();
  for (const title of ["Select", "Effective date", "Code", "Amount", "Description", "Source row", "Status"]) {
    const cell = document.createElement("th");
    cell.scope = "col";
    cell.textContent = title;
    heading.append(cell);
  }
  const body = table.createTBody();
  for (const item of cardRows) {
    const row = body.insertRow();
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = cardSelection.has(item.id);
    checkbox.setAttribute("aria-label", "Select " + item.date + " " + item.description + " " + cardMoney.format(item.amount / 100));
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) cardSelection.add(item.id);
      else cardSelection.delete(item.id);
      showCardSelection();
    });
    row.insertCell().append(checkbox);
    for (const text of [item.date, item.code, cardMoney.format(item.amount / 100), item.description,
      item.source + " · " + item.row, (item.state === "matched" ? "Matched group " + item.group
        : item.carryover ? "Confirmed carryover" : item.state + (item.reason ? " · " + item.reason : ""))
      + (item.priorEvidence ? " · " + item.priorEvidence : "")]) {
      const cell = row.insertCell();
      cell.textContent = text;
    }
    row.cells[4].classList.add("description");
  }
  cardElement("transactions").replaceChildren(table);
  renderCardPairs();
  cardElement("review").hidden = false;
  showCardSelection();
}

function showCardSelection() {
  const selected = cardRows.filter(row => cardSelection.has(row.id));
  const sum = selected.reduce((total, row) => total + row.amount, 0);
  cardElement("selection").textContent = selected.length + " selected · net " + cardMoney.format(sum / 100);
}

function matchCardGroup(rows) {
  requireReport(rows.every(row => row.state === "open"), "Reopen the selected rows before matching.");
  const sum = rows.reduce((total, row) => total + row.amount, 0);
  requireReport(Number.isSafeInteger(sum) && sum === 0 && rows.some(row => row.amount > 0) && rows.some(row => row.amount < 0),
    "Selected funding and charges must total zero.");
  const group = ++cardGroup;
  for (const row of rows) {
    row.state = "matched";
    row.group = group;
    row.carryover = false;
  }
}

function changeCardRows(action) {
  const selected = cardRows.filter(row => cardSelection.has(row.id));
  try {
    requireReport(selected.length, "Select transactions first.");
    if (action === "matched") {
      matchCardGroup(selected);
    } else if (action === "carryover") {
      requireReport(selected.every(row => row.state === "open"), "Reopen selected rows before confirming a carryover.");
      selected.forEach(row => { row.carryover = true; });
    } else {
      requireReport(action !== "open" || selected.every(row => row.reason !== "Outside selected dates" && row.amount !== 0),
        "Change the effective-date range to include rows outside it. Zero-amount rows stay excluded.");
      const groups = new Set(selected.map(row => row.group).filter(group => group !== undefined));
      for (const row of cardRows) {
        if (groups.has(row.group)) {
          row.state = "open";
          row.group = undefined;
        }
      }
      for (const row of selected) {
        row.state = action;
        row.reason = action === "excluded" ? "Excluded by preparer" : "";
        row.group = undefined;
        row.carryover = false;
      }
    }
    cardRevision++;
    cardSelection.clear();
    renderCard();
    cardReviewStatus("");
  } catch (error) {
    cardReviewStatus(error.message, true);
  }
}

function cardUniquePairs(rows) {
  const positive = new Map(), negative = new Map();
  for (const row of rows) {
    if (row.state !== "open") continue;
    const map = row.amount > 0 ? positive : negative;
    const amount = Math.abs(row.amount);
    if (!map.has(amount)) map.set(amount, []);
    map.get(amount).push(row);
  }
  return [...positive].flatMap(([amount, funding]) => {
    const charges = negative.get(amount);
    return funding.length === 1 && charges?.length === 1 ? [[funding[0], charges[0]]] : [];
  });
}

function selectCardPair(pair) {
  cardSelection = new Set(pair.map(item => item.id));
  renderCard();
  cardReviewStatus("Selected a unique amount pair. Check its descriptions, then match it if appropriate.");
}

function renderCardPairs() {
  const region = cardElement("pair-suggestions");
  const pairs = cardUniquePairs(cardRows);
  const summary = cardElement("pair-count");
  const apply = cardElement("match-reviewed-pairs");
  const checked = new Set();
  apply.disabled = true;
  apply.onclick = () => {
    for (const pair of checked) matchCardGroup(pair);
    cardRevision++;
    cardSelection.clear();
    renderCard();
    cardReviewStatus(checked.size + (checked.size === 1 ? " pair matched." : " pairs matched."));
  };
  summary.textContent = pairs.length + (pairs.length === 1 ? " suggested pair" : " suggested pairs");
  if (!pairs.length) {
    region.replaceChildren();
    return;
  }
  const table = document.createElement("table"), header = table.createTHead().insertRow(), body = table.createTBody();
  for (const name of ["Amount", "Funding", "Charge", "Review"]) {
    const cell = document.createElement("th");
    cell.scope = "col";
    cell.textContent = name;
    header.append(cell);
  }
  const describe = item => [item.date, item.code, item.description, item.source + " row " + item.row].filter(Boolean).join(" · ");
  for (const pair of pairs) {
    const [funding, charge] = pair, row = body.insertRow();
    const amount = cardMoney.format(funding.amount / 100);
    for (const value of [amount, describe(funding), describe(charge)]) row.insertCell().textContent = value;
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.setAttribute("aria-label", "Match the " + amount + " pair");
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) checked.add(pair);
      else checked.delete(pair);
      apply.disabled = !checked.size;
    });
    row.insertCell().append(checkbox);
  }
  region.replaceChildren(table);
}

function suggestCardPair() {
  const [pair] = cardUniquePairs(cardRows);
  if (pair) return selectCardPair(pair);
  cardReviewStatus("No unique amount pair remains. Select related funding and charges manually.");
}

function suggestCardSplit() {
  const selected = cardRows.filter(row => cardSelection.has(row.id));
  if (selected.length !== 1 || selected[0].state !== "open") {
    cardReviewStatus("Select one open funding or charge row to look for a split.", true);
    return;
  }
  const anchor = selected[0], target = Math.abs(anchor.amount);
  const opposite = cardRows.filter(row => row.state === "open" && Math.sign(row.amount) !== Math.sign(anchor.amount));
  if (opposite.some(row => Math.abs(row.amount) === target)) {
    cardReviewStatus("A single-row amount also matches. Review the entries before selecting a split.");
    return;
  }
  const byAmount = new Map();
  let match;
  for (const item of opposite) {
    const amount = Math.abs(item.amount);
    if (amount >= target) continue;
    const complementary = byAmount.get(target - amount);
    if (complementary) {
      if (match || complementary.length > 1) {
        cardReviewStatus("Several two-row splits match this amount. Select the related entries manually.");
        return;
      }
      match = [complementary[0], item];
    }
    const same = byAmount.get(amount) ?? [];
    same.push(item);
    byAmount.set(amount, same);
  }
  if (!match) {
    cardReviewStatus("No exact two-row split was found for the selected entry.");
    return;
  }
  cardSelection = new Set([anchor.id, ...match.map(item => item.id)]);
  renderCard();
  cardReviewStatus("Selected an exact two-row split. Verify the dates and descriptions before matching.");
}

async function readCardCover(file, priorWorkbook) {
  const workbook = file ? await openExcel(file) : priorWorkbook;
  const sheet = "Corp Card";
  const hasCover = workbook?.sheets.has(sheet);
  requireReport(!file || hasCover, "Choose the workbook with the Corp Card cover.");
  if (!hasCover) return {};
  const formulas = new Set(), cells = await readCells(workbook, sheet, formulas);
  const recognized = /^balance per account$/i.test(String(cells.get("C10") ?? "").trim())
    && /^difference$/i.test(String(cells.get("K47") ?? "").trim());
  requireReport(!file || recognized, "Check the balance and difference labels on the prepared Corp Card cover.");
  const auditTarget = completionTarget(sheet, cells, formulas, "H54", "H58", "L53", "L54");
  const balance = cells.get("L10");
  const balanceTarget = recognized && !formulas.has("L10") && (balance === undefined || typeof balance === "number") ? "L10" : "";
  return { sheet, auditTarget, balanceTarget };
}

async function buildCardWorkbook(rows, month, balances, cover) {
  const completion = reconciliationAudit();
  const { label, number, date, row, worksheet } = excelOutputCells();
  const active = rows.filter(item => item.state !== "excluded");
  const open = rows.filter(item => item.state === "open");
  const carryovers = open.filter(item => item.carryover);
  const unmatched = open.filter(item => !item.carryover);
  const grouped = new Map();
  for (const item of active) {
    const key = item.state === "matched" ? "group-" + item.group : item.id;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(item);
  }
  const lines = [row(1, label("A1", "Corporate Card · " + month, 1)), row(3, label("A3", "Account Deposits", 1) + label("F3", "Account Charges", 1))];
  let next = 4;
  for (const items of grouped.values()) {
    const funding = items.filter(item => item.amount > 0);
    const charges = items.filter(item => item.amount < 0);
    const start = next;
    const length = Math.max(funding.length, charges.length);
    for (let i = 0; i < length; i++) {
      let cells = "";
      for (const [item, columns] of [[funding[i], "ABCD"], [charges[i], "FGHI"]]) {
        if (item) cells += date(columns[0] + next, item.date) + label(columns[1] + next, item.code)
          + number(columns[2] + next, item.amount / 100) + label(columns[3] + next, item.description);
      }
      lines.push(row(next++, cells));
    }
    const total = items.reduce((sum, item) => sum + item.amount, 0) / 100;
    lines.push(row(next, `<c r="E${next}" s="2"><f>SUM(C${start}:C${next - 1})+SUM(H${start}:H${next - 1})</f><v>${total}</v></c>`));
    next += 2;
  }
  const list = (items, pending = false) => items.map((item, i) => {
    const r = i + (pending ? 3 : 2);
    const columns = pending ? "BCDE" : "ABCD";
    return row(r, date(columns[0] + r, item.date) + label(columns[1] + r, item.code)
      + number(columns[2] + r, item.amount / 100) + label(columns[3] + r, item.description)
      + (pending ? "" : label("E" + r, item.fileName + " · " + item.sheet + " row " + item.row)
        + (item.carryover ? label("F" + r, "Confirmed carryover") : "")));
  }).join("");
  const headers = (titles, columns) => titles.map((title, i) => label(columns[i] + "1", title, 1)).join("");
  const openHeader = row(1, headers(["Effective Date", "Txn", "Transaction Amount", "Statement Description", "Source"], "ABCDE"));
  const carryoverHeader = row(1, headers(["Effective Date", "Txn", "Transaction Amount", "Statement Description", "Source", "Carryover decision"], "ABCDEF"));
  const pending = carryovers.filter(item => item.amount > 0);
  const pendingAmount = pending.reduce((sum, item) => sum + item.amount, 0);
  requireReport(Number.isSafeInteger(pendingAmount), "The pending funding total exceeds the supported amount range.");
  const pendingEnd = Math.max(12, pending.length + 2), totalRow = pendingEnd + 4, transitRow = totalRow + 3;
  const formula = (address, expression) => `<c r="${address}" s="2"><f>${expression}</f><v>${pendingAmount / 100}</v></c>`;
  const pendingRows = [
    row(2, ["Funding Date", "Tran Code", "Amount", "Description"].map((title, i) => label("BCDE"[i] + "2", title, 1)).join("")),
    list(pending, true),
    row(totalRow, label("B" + totalRow, "Total") + formula("C" + totalRow, "SUM(D3:D" + pendingEnd + ")")),
    row(transitRow, label("B" + transitRow, "In Transit From Matchup Sheet") + formula("C" + transitRow, "C" + totalRow)),
  ];
  const excluded = rows.filter(item => item.state === "excluded");
  const audit = [row(1, headers(["Source", "Worksheet", "Row", "Date", "Code", "Amount", "Description", "Status", "Prior workbook evidence"], "ABCDEFGHI")),
    ...rows.map((item, i) => {
      const r = i + 2;
      return row(r, label("A" + r, item.fileName) + label("B" + r, item.sheet) + number("C" + r, item.row)
        + label("D" + r, item.date) + label("E" + r, item.code) + number("F" + r, item.amount / 100)
        + label("G" + r, item.description) + label("H" + r, item.state === "matched" ? "Matched group " + item.group
          : item.carryover ? "Confirmed carryover" : item.state + " " + item.reason)
        + label("I" + r, item.priorEvidence ?? ""));
    }),
  ].join("");
  const balanceRows = balances ? [
    ["Period", month], ["Source", balances.file], ["Worksheet", balances.sheet],
    ...(balances.issue ? [["Review", balances.issue]] : [
      ...(balances.opening ? [["Prior posted balance", balances.opening.balance / 100, balances.opening.posted, "L" + balances.opening.row]] : []),
      ["Posted transactions in the month", balances.count], ["Net posted activity", balances.movement / 100],
      ["Last posted balance for the month", balances.closing.balance / 100, balances.closing.posted, "L" + balances.closing.row],
      ...(balances.verifyCoverage ? [["Review", "Verify report coverage through the end of the month."]] : []),
    ]),
  ] : [];
  const balancePaste = cover.balanceTarget && balances?.closing && !balances.verifyCoverage
    ? [["Cover balance inputs", worksheet("A1:E2", [[1, 1, 20], [2, 2, 13], [3, 3, 20], [4, 4, 18], [5, 5, 45]],
      row(1, headers(["Destination sheet", "Cell", "Paste value", "Posting date", "Source"], "ABCDE"))
      + row(2, label("A2", cover.sheet) + label("B2", cover.balanceTarget) + number("C2", balances.closing.balance / 100)
        + date("D2", balances.closing.posted) + label("E2", balances.sheet + "!L" + balances.closing.row)))]] : [];
  return createExcelWorkbook([
    ["Matchup", worksheet("A1:I" + Math.max(next - 1, 3), [[1, 2, 14], [3, 3, 18], [4, 4, 48], [5, 5, 17], [6, 7, 14], [8, 8, 18], [9, 9, 48]], lines.join(""))],
    ["Pending Charges", worksheet("B2:E" + transitRow, [[2, 2, 33], [3, 3, 15], [4, 4, 18], [5, 5, 60]], pendingRows.join(""))],
    ["Open items", worksheet("A1:F" + (carryovers.length + 1), [[1, 2, 16], [3, 3, 18], [4, 5, 60], [6, 6, 24]], carryoverHeader + list(carryovers))],
    ["Unmatched activity", worksheet("A1:E" + (unmatched.length + 1), [[1, 2, 16], [3, 3, 18], [4, 5, 60]], openHeader + list(unmatched))],
    ["Excluded", worksheet("A1:E" + Math.max(excluded.length + 1, 1), [[1, 2, 16], [3, 3, 18], [4, 5, 60]], openHeader + list(excluded))],
    ["Import review", worksheet("A1:I" + (rows.length + 1), [[1, 2, 35], [3, 6, 16], [7, 9, 55]], audit)],
    ...(balances ? [["Account balance", worksheet("A1:D" + (balanceRows.length + 1), [[1, 1, 36], [2, 2, 55], [3, 4, 18]],
      row(1, headers(["Field", "Value", "Posting date", "Source cell"], "ABCD")) + balanceRows.map((entry, index) => {
        const r = index + 2;
        return row(r, entry.map((item, i) => typeof item === "number" ? number("ABCD"[i] + r, item)
          : label("ABCD"[i] + r, item)).join(""));
      }).join(""))]] : []),
    ...balancePaste,
    ...completionPasteSheets(cover.auditTarget, completion),
  ], completion);
}

function bindCardFile(kind, select) {
  const input = cardElement(kind + "-file"), drop = cardElement(kind + "-drop");
  input.addEventListener("change", () => select(input.files[0]));
  drop.addEventListener("dragover", event => { event.preventDefault(); drop.classList.add("active"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("active"));
  drop.addEventListener("drop", event => {
    event.preventDefault();
    drop.classList.remove("active");
    if (event.dataTransfer.files.length !== 1) {
      cardStatus("Drop one workbook at a time.", true);
      return;
    }
    input.value = "";
    select(event.dataTransfer.files[0]);
  });
}
for (const [kind, state] of Object.entries(cardFiles)) {
  cardElement(kind + "-mapping").addEventListener("input", () => { state.mappingRevision++; });
  cardElement(kind + "-sheet").addEventListener("change", async () => {
    try { await suggestCardMapping(kind); }
    catch (error) { cardStatus(error.message, true); }
  });
  bindCardFile(kind, file => selectCardFile(kind, file));
}
bindCardFile("cover", file => {
  cardCoverFile = file;
  cardElement("cover-status").textContent = file?.name ?? "";
  invalidateCard();
});
document.addEventListener("dragover", event => event.preventDefault());
document.addEventListener("drop", event => event.preventDefault());
cardElement("month").addEventListener("input", () => {
  const month = cardElement("month").value;
  if (month) {
    cardElement("from").value = month + "-01";
    const [year, m] = month.split("-").map(Number);
    cardElement("through").value = new Date(Date.UTC(year, m, 1)).toISOString().slice(0, 10);
  }
});
for (const input of document.querySelectorAll(".mapping input, .mapping select, #month, #from, #through")) {
  input.addEventListener("input", invalidateCard);
}
cardElement("prior-clear").addEventListener("click", () => {
  cardElement("prior-file").value = "";
  selectCardFile("prior", undefined);
});
cardElement("load").addEventListener("click", async () => {
  invalidateCard();
  const current = cardRevision;
  try {
    const month = cardElement("month").value;
    const from = cardElement("from").value, through = cardElement("through").value;
    requireReport(month && from && through && from <= through, "Choose a month and a valid effective-date range.");
    const [source, prior, boundary, balances, cover] = await Promise.all([
      readCardRows("source", from, through), cardFiles.prior.file ? readCardRows("prior", from, through) : [],
      cardPriorBoundary(cardFiles.prior.workbook, month + "-01"),
      readCardBalanceReview(month),
      readCardCover(cardCoverFile, cardFiles.prior.workbook),
    ]);
    requireReport(source.some(item => item.state !== "excluded"), "The account export has no transactions in the selected range.");
    if (current !== cardRevision) return;
    flagCardBoundary(source, prior, boundary);
    cardRows = [...source, ...prior];
    cardBalanceReview = balances;
    cardCoverReview = cover;
    cardGroup = 0;
    renderCard();
    cardStatus("Loaded transactions. Review the rows before matching.");
  } catch (error) {
    if (current === cardRevision) cardStatus(error.message, true);
  }
});
cardElement("suggest").addEventListener("click", suggestCardPair);
cardElement("suggest-split").addEventListener("click", suggestCardSplit);
cardElement("match").addEventListener("click", () => changeCardRows("matched"));
cardElement("carry").addEventListener("click", () => changeCardRows("carryover"));
cardElement("exclude").addEventListener("click", () => changeCardRows("excluded"));
cardElement("restore").addEventListener("click", () => changeCardRows("open"));
cardElement("confirmed").addEventListener("change", () => {
  cardElement("download").disabled = !cardRows.length || !cardElement("confirmed").checked || cardDownloading;
});
cardElement("download").addEventListener("click", async () => {
  if (cardElement("download").disabled) return;
  const current = cardRevision;
  const selected = cardRows.map(item => ({ ...item }));
  cardDownloading = true;
  cardElement("download").disabled = true;
  try {
    const blob = await buildCardWorkbook(selected, cardElement("month").value, cardBalanceReview, cardCoverReview);
    if (current !== cardRevision || !cardElement("confirmed").checked) return;
    const url = URL.createObjectURL(blob), link = document.createElement("a");
    link.href = url;
    link.download = "Corporate Card " + cardElement("month").value + ".xlsx";
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  } catch (error) {
    cardStatus(error.message, true);
  } finally {
    cardDownloading = false;
    cardElement("download").disabled = !cardRows.length || !cardElement("confirmed").checked;
  }
});
