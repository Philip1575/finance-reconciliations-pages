"use strict";

// VIZO batch: every VIZO account reconciled at once. Money moves between the accounts automatically (overdraft
// protection, transfers, interest sweeps), so each transfer is paired with the other account's side once, and every
// entry still needed in Wisdom is listed with its debit and credit G/Ls.

const vbField = id => document.getElementById(id);
const vbFiles = { vizo: [], wisdom: [], books: [], invoice: [] };
let vbResult, vbRevision = 0;

// The accounts are recognized by generic words in the VIZO export, the Wisdom report and the workbook titles.
// Order matters: an ATM or overnight settlement account is not the main Settlement account.
const vbKinds = [
  { kind: "ATM", short: "ATM", pattern: /\batm\b/i },
  { kind: "OSL", short: "OSL", pattern: /overnight|\bosl\b/i },
  { kind: "SETTLEMENT", short: "SETTLE", pattern: /settlement/i },
  { kind: "PSA", short: "PSA", pattern: /\bpsa\b|premium share/i },
  { kind: "PCC", short: "PCC", pattern: /\bpcc\b|perpetual/i },
  { kind: "MISC", short: "MISC", pattern: /misc/i },
  { kind: "CC", short: "CC", pattern: /charge card/i },
];
const vbKindOf = name => vbKinds.find(item => item.pattern.test(String(name ?? "")))?.kind;
const vbShort = kind => vbKinds.find(item => item.kind === kind)?.short ?? kind;
const vbMoney = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });

// G/L numbers as the workbooks write them (123.4 or 123400) in one form.
function vbGl(value) {
  if (typeof value === "string") value = value.replace(/^\s*(dr|cr)\s+/i, "");
  if (typeof value === "number") value = Number.isInteger(value) ? String(value) : String(Number(value.toFixed(3)));
  return normalizeAccount(value) || undefined;
}

function vbStatus(message, error = false) {
  vbField("status").textContent = message;
  vbField("status").classList.toggle("error", error);
}

async function vbSheetCells(file) {
  if (/\.csv$/i.test(file.name)) return readCsvCells(await file.text());
  const book = await openExcel(file, true);
  return readCells(book, [...book.sheets.keys()][0]);
}

// The VIZO exports (one for all accounts, or one per account): Account Number, Account Name, dates, description, amount
// and running balance. Rows after the month are kept to show how month-end items cleared; overlapping exports are merged.
async function vbReadVizo(files, month) {
  const accounts = new Map(), seen = new Set();
  for (const file of files) await vbReadVizoFile(file, month, accounts, seen);
  requireReport(accounts.size, "The VIZO exports have no activity for the month.");
  for (const entry of accounts.values()) {
    entry.rows.sort((a, b) => a.date.localeCompare(b.date) || a.order - b.order);
    entry.after.sort((a, b) => a.date.localeCompare(b.date) || a.order - b.order);
    const first = entry.rows[0], last = entry.rows.at(-1);
    entry.opening = first.balance - first.amount;
    entry.closing = last.balance;
    let running = entry.opening;
    for (const item of entry.rows) {
      running += item.amount;
      requireReport(running === item.balance, "The VIZO running balance does not agree for " + entry.name + " on " + item.date + ".");
    }
  }
  return accounts;
}

async function vbReadVizoFile(file, month, accounts, seen) {
  const cells = await vbSheetCells(file);
  const headers = new Map();
  let headerRow;
  for (const [address, value] of cells) {
    const [, column, row] = /^([A-Z]+)(\d+)$/.exec(address);
    if (Number(row) > 20) continue;
    const name = String(value).trim().toLowerCase();
    if (name === "account number") headerRow = Number(row);
    if (headerRow === Number(row) || ["account number", "account name", "transaction date", "transaction description", "amount", "balance"].includes(name)) {
      headers.set(name + "|" + row, column);
    }
  }
  requireReport(headerRow, "The VIZO export needs an Account Number heading.");
  const column = name => headers.get(name + "|" + headerRow);
  for (const name of ["account number", "account name", "transaction date", "transaction description", "amount", "balance"]) {
    requireReport(column(name), "The VIZO export needs a " + name + " column.");
  }
  const rows = [...new Set([...cells.keys()].map(address => Number(/\d+$/.exec(address)[0])))].filter(r => r > headerRow).sort((a, b) => a - b);
  for (const row of rows) {
    const get = name => cells.get(column(name) + row);
    const number = String(get("account number") ?? "").trim();
    if (!number) continue;
    const item = { date: sourceDate(get("transaction date")), description: String(get("transaction description") ?? "").trim(),
      amount: cents(get("amount")), balance: cents(get("balance")), row, file: file.name, order: seen.size, side: "VIZO", account: number };
    const key = [number, item.date, item.description, item.amount, item.balance].join("|");
    if (item.date < month + "-01" || seen.has(key)) continue;
    seen.add(key);
    const entry = accounts.get(number) ?? { number, name: String(get("account name") ?? "").trim(), rows: [], after: [] };
    entry.kind ??= vbKindOf(entry.name);
    (item.date.slice(0, 7) === month ? entry.rows : entry.after).push(item);
    accounts.set(number, entry);
  }
}

// The Wisdom Detailed Trial Balance for several G/Ls (one report or one per G/L), in either export layout (one date and a
// signed amount, or effective and posted dates with separate debit and credit columns). Only the month's rows are used.
async function vbReadWisdom(files, month) {
  const ledgers = new Map();
  for (const file of files) {
    for (const [gl, ledger] of await vbReadWisdomFile(file, month)) {
      requireReport(!ledgers.has(gl), "G/L " + gl + " is in more than one Wisdom report.");
      ledgers.set(gl, ledger);
    }
  }
  return ledgers;
}

async function vbReadWisdomFile(file, month) {
  const book = await openExcel(file, true);
  const cells = await readCells(book, [...book.sheets.keys()][0]);
  requireReport(/^detailed trial balance report$/i.test(String(cells.get("B1") ?? "").trim()), "Load the Wisdom Detailed Trial Balance.");
  const rows = [...new Set([...cells.keys()].map(address => Number(/\d+$/.exec(address)[0])))].sort((a, b) => a - b);
  const ledgers = new Map();
  let current;
  for (const row of rows) {
    const get = column => cells.get(column + row);
    const compactHeader = /^beg bal:$/i.test(String(get("D") ?? "").trim()) && vbGl(get("A"));
    const datedHeader = /^starting balance:/i.test(String(get("I") ?? "").trim()) && vbGl(get("D"));
    if (compactHeader || datedHeader) {
      const gl = compactHeader ? vbGl(get("A")) : vbGl(get("D"));
      const name = String(compactHeader ? get("C") : get("H")).trim();
      const beginning = cents(compactHeader ? get("E") : String(get("I")).replace(/^starting balance:\s*/i, ""));
      current = { gl, name, kind: vbKindOf(name), beginning, ending: beginning, running: beginning, rows: [], layout: compactHeader ? "compact" : "dated" };
      ledgers.set(gl, current);
      continue;
    }
    if (!current || typeof get("A") !== "number") continue;
    const dated = current.layout === "dated";
    const amount = dated ? cents(get("I") ?? 0) + cents(get("J") ?? 0) : cents(get("F"));
    const balance = cents(get(dated ? "K" : "G"));
    requireReport(current.running + amount === balance, file.name + ": check the Wisdom running balance for G/L " + current.gl + " at row " + row + ".");
    current.running = balance;
    const date = sourceDate(get("A"), book.date1904);
    if (date.slice(0, 7) < month) current.beginning = current.ending = balance;
    if (date.slice(0, 7) !== month) continue;
    current.ending = balance;
    current.rows.push({ date, reference: String(get(dated ? "C" : "B") ?? "").trim(),
      description: String(get(dated ? "G" : "D") ?? "").trim(), amount, balance, row, side: "Wisdom", gl: current.gl });
  }
  requireReport(ledgers.size, file.name + ": no G/L sections were found in the Wisdom report.");
  return ledgers;
}

const vbCellText = value => String(value ?? "").trim();
const vbIsVizoText = text => /^(external (deposit|withdrawal)|overdraft protection|deposit|withdrawal|credit interest|descriptive)/i.test(text);

// A short signature of a VIZO description that stays the same from month to month: the direction and the payee words
// or EBT code (deposits and withdrawals of the same kind are booked differently, e.g. SPC disputes and the SPC bill).
function vbSignature(description) {
  const full = vbCellText(description);
  const direction = /^(external|descriptive)?\s*deposit/i.test(full) ? "IN " : /^(external|descriptive)?\s*withdraw/i.test(full) ? "OUT " : "";
  const text = full.replace(/^(external|descriptive)\s+(deposit|withdrawal)\s*/i, "");
  const ebt = /-\s*EBT\s+([A-Z]+)/i.exec(text);
  if (ebt) return direction + "EBT " + ebt[1].toUpperCase();
  return direction + (text.toUpperCase().match(/[A-Z]{2,}/g) ?? []).slice(0, 3).join(" ");
}

// Last month's workbook: account, G/L, cover cells and in-transit items, and the entries in its latest block (memo,
// amount, debit and credit G/Ls, and the VIZO rows the entry covered).
async function vbReadBook(file, knownGls) {
  const book = await openExcel(file, true);
  const sheetNames = [...book.sheets.keys()];
  const coverName = sheetNames[0];
  const coverFormulas = new Map();
  const cover = await readCells(book, coverName, coverFormulas);
  const title = vbCellText(cover.get("C11"));
  const kind = vbKindOf(title);
  requireReport(kind, file.name + ": the cover title at C11 does not name a VIZO account.");
  const find = test => [...cover].find(([address, value]) => /^[A-D]\d+$/.test(address) && test(vbCellText(value).toUpperCase()));
  const rowOf = entry => entry && Number(/\d+$/.exec(entry[0])[0]);
  const statementRow = rowOf(find(text => text.startsWith("BALANCE PER VIZO STATEMENT")));
  const ledgerRow = rowOf(find(text => /^BALANCE PER .*G\/L/.test(text)));
  const transitRow = rowOf(find(text => text === "IN TRANSIT ITEMS"));
  const adjustedRow = rowOf(find(text => /ADJUSTED BOOK BALANCE/.test(text)));
  const titleRow = rowOf(find(text => text.startsWith("FOR THE MONTH OF")));
  const glRow = ledgerRow && [14, 15].find(r => vbGl(cover.get("D" + r)) && knownGls.has(vbGl(cover.get("D" + r))));
  const gl = glRow ? vbGl(cover.get("D" + glRow)) : undefined;
  requireReport(statementRow && ledgerRow, file.name + ": the cover needs BALANCE PER VIZO STATEMENT and BALANCE PER G/L rows.");
  const items = [];
  if (transitRow) {
    const last = adjustedRow ?? transitRow + 15;
    for (let r = transitRow + 2; r < last; r++) {
      const amount = cover.get("L" + r);
      if (typeof amount !== "number") continue;
      items.push({ row: r, description: vbCellText(cover.get("F" + r)), clears: cover.get("I" + r), date: cover.get("K" + r),
        note: vbCellText(cover.get("M" + r)), amount: cents(amount) });
    }
  }
  // Entries in the latest block of the worksheet (ATM keeps a separate "Matchup <Month>" sheet per month).
  const matchups = sheetNames.filter(name => /^matchup\s+[a-z]+\s+\d{2,4}/i.test(name.trim()));
  const sheet = matchups[0] ?? sheetNames[1];
  const templates = [];
  if (sheet) {
    const formulas = new Map();
    const cells = await readCells(book, sheet, formulas);
    const rows = [...new Set([...cells.keys()].map(address => Number(/\d+$/.exec(address)[0])))].sort((a, b) => a - b);
    const heads = rows.filter(r => ["A", "B"].some(column => /^(vizo|[a-z]+ match up)\b/i.test(vbCellText(cells.get(column + r)))));
    const start = matchups[0] ? 1 : heads.at(-1) ?? rows.at(-1) - 80;
    const block = rows.filter(r => r >= start);
    const columnsOf = r => [...cells.keys()].filter(address => Number(/\d+$/.exec(address)[0]) === r).map(address => address.replace(/\d+$/, ""))
      .sort((a, b) => columnNumber(a) - columnNumber(b));
    const vizoText = r => columnsOf(r).map(column => vbCellText(cells.get(column + r))).find(vbIsVizoText);
    const glPair = r => {
      const columns = columnsOf(r);
      for (let i = 0; i < columns.length - 1; i++) {
        const [a, b] = [vbGl(cells.get(columns[i] + r)), vbGl(cells.get(columns[i + 1] + r))];
        if (a && b && (knownGls.has(a) || knownGls.has(b)) && columnNumber(columns[i + 1]) - columnNumber(columns[i]) === 1) {
          return { dr: a, cr: b, columns: [columns[i], columns[i + 1]] };
        }
      }
      return undefined;
    };
    // Each memo with a debit/credit G/L pair: on its own row, on the rows its amount formula points to (=C11094 or
    // SUM(H233:H237)), or on the row below. Memos whose formula names the pair's row claim it first.
    const candidates = [];
    for (const r of block) {
      const memoColumn = columnsOf(r).find(column => {
        const text = vbCellText(cells.get(column + r));
        return typeof cells.get(column + r) === "string" && text && !vbIsVizoText(text) && !/^(date|description|amount|debit|credit|member acc|dr |cr )/i.test(text)
          && /[A-Z]{2,}/i.test(text);
      });
      if (!memoColumn) continue;
      const amountColumn = columnsOf(r).find(column => typeof cells.get(column + r) === "number" && !(glPair(r)?.columns ?? []).includes(column)
        && !(Number.isInteger(cells.get(column + r)) && cells.get(column + r) > 40000 && cells.get(column + r) < 60000)
        && columnNumber(column) > columnNumber(memoColumn));
      const formula = amountColumn && formulas.get(amountColumn + r);
      const text = typeof formula === "string" ? formula.replace(/[\s$]/g, "") : "";
      const range = /^SUM\([A-Z]+(\d+):[A-Z]+(\d+)\)$/i.exec(text), single = /^[A-Z]+(\d+)$/i.exec(text);
      const referenced = range ? Array.from({ length: Number(range[2]) - Number(range[1]) + 1 }, (_, k) => Number(range[1]) + k)
        : single ? [Number(single[1])] : [];
      candidates.push({ r, memoColumn, amountColumn, referenced, rows: referenced.length ? [r, ...referenced.slice().reverse()] : [r, r + 1] });
    }
    const memoRows = new Set(candidates.map(item => item.r)), claimed = new Set();
    for (const item of [...candidates].sort((a, b) => b.referenced.length - a.referenced.length || a.r - b.r)) {
      const pairRow = item.rows.find(row => !claimed.has(row) && glPair(row) && (row === item.r || item.referenced.includes(row) || !memoRows.has(row)));
      if (pairRow === undefined) continue;
      claimed.add(pairRow);
      const pair = glPair(pairRow);
      // The VIZO rows the entry covers: the rows its formula points to, or else the contiguous VIZO rows ending on the
      // pair's row.
      const vizoRows = item.referenced.map(vizoText).filter(Boolean);
      if (!item.referenced.length) for (let up = vizoText(pairRow) ? pairRow : item.r; up >= start && vizoText(up); up--) vizoRows.unshift(vizoText(up));
      templates.push({ memo: vbCellText(cells.get(item.memoColumn + item.r)), dr: pair.dr, cr: pair.cr,
        amount: item.amountColumn ? cents(cells.get(item.amountColumn + item.r)) : undefined, signatures: [...new Set(vizoRows.map(vbSignature))], row: item.r });
    }
    templates.sort((a, b) => a.row - b.row);
  }
  // Whether the cover adds its in-transit items to the G/L balance (Adjusted Book Balance = SUM from the G/L row).
  const adjustedFormula = adjustedRow && coverFormulas.get("L" + adjustedRow);
  const itemsToLedger = typeof adjustedFormula === "string" ? adjustedFormula.split(/[^A-Z0-9]+/).includes("L" + ledgerRow) : undefined;
  // The sign the cover applies to its items, from last month's own tie (statement = G/L ± items), else from the formula.
  const statementValue = cents(cover.get("L" + statementRow) ?? 0), ledgerValue = cents(cover.get("L" + ledgerRow) ?? 0);
  const itemsTotal = items.reduce((sum, item) => sum + item.amount, 0);
  const itemsSign = statementValue === ledgerValue + itemsTotal ? 1 : statementValue === ledgerValue - itemsTotal ? -1 : itemsToLedger === false ? -1 : 1;
  // A statement-side in-transit line with an amount (Charge Cards: "Loanstreet In Transit Clearing").
  const clearingRow = [...Array(Math.max(0, ledgerRow - statementRow - 1)).keys()].map(k => statementRow + 1 + k)
    .find(r => typeof cover.get("L" + r) === "number" && ["C", "D", "E"].some(column => /transit/i.test(vbCellText(cover.get(column + r)))));
  return { file, kind, title, gl, coverName, itemsToLedger, itemsSign, clearingRow, statementRow, ledgerRow, titleRow, items, templates, adjustedRow, transitRow,
    statement: cents(cover.get("L" + statementRow) ?? 0), ledgerBalance: cents(cover.get("L" + ledgerRow) ?? 0), sheet };
}

// The memo of a learned entry with this month's date in place of last month's (08/26 → 09/26, 8/26 → 9/26).
function vbMemo(memo, month) {
  const [year, number] = month.split("-");
  return memo.replace(/\b(\d{1,2})\/(\d{2})\b/g, (all, m, y) => (m.length === 2 ? number : String(Number(number))) + "/" + year.slice(2));
}

function vbClassify(item) {
  const text = item.description;
  let match;
  if (/overdraft protection withdraw/i.test(text)) return { type: "OD", direction: "out", other: /(\d{6})/.exec(text)?.[1] };
  if (/overdraft protection deposit/i.test(text)) return { type: "OD", direction: "in" };
  if ((match = /interest transfer (to|from)\s*(\d{6})/i.exec(text))) return { type: "INT", direction: match[1].toLowerCase() === "to" ? "out" : "in", other: match[2] };
  if (/^credit interest/i.test(text)) return { type: "INTCREDIT" };
  if ((match = /transfer from\s*(\d{6})/i.exec(text))) return { type: "XFER", direction: "in", other: match[1] };
  if ((match = /transfer to\s*(\d{6})/i.exec(text))) return { type: "XFER", direction: "out", other: match[1] };
  return { type: "EXT" };
}

const vbDays = (a, b) => Math.round((Date.parse(a) - Date.parse(b)) / 86400000);

// Pairs VIZO rows with Wisdom rows of the same G/L: exact amounts within a few days, then a Wisdom row equal to two or
// three VIZO rows (or the reverse) within a few days.
function vbMatch(vizoRows, wisdomRows) {
  const matched = new Map();
  const open = { vizo: [...vizoRows], wisdom: [...wisdomRows] };
  for (const item of [...open.vizo]) {
    const candidates = open.wisdom.filter(other => other.amount === item.amount && Math.abs(vbDays(item.date, other.date)) <= 5)
      .sort((a, b) => Math.abs(vbDays(item.date, a.date)) - Math.abs(vbDays(item.date, b.date)));
    if (!candidates.length) continue;
    matched.set(item, [candidates[0]]);
    matched.set(candidates[0], [item]);
    open.vizo.splice(open.vizo.indexOf(item), 1);
    open.wisdom.splice(open.wisdom.indexOf(candidates[0]), 1);
  }
  const combine = (one, many) => {
    for (const item of [...open[one]]) {
      const pool = open[many].filter(other => Math.sign(other.amount) === Math.sign(item.amount) && Math.abs(vbDays(item.date, other.date)) <= 4);
      if (pool.length > 25) continue;
      let found;
      const search = (start, remaining, picked) => {
        if (found === null) return;
        if (remaining === 0 && picked.length > 1) { found = found ? null : picked; return; }
        if (picked.length === 4) return;
        for (let i = start; i < pool.length; i++) {
          if (Math.abs(pool[i].amount) <= Math.abs(remaining)) search(i + 1, remaining - pool[i].amount, [...picked, pool[i]]);
        }
      };
      search(0, item.amount, []);
      if (!found) continue;
      matched.set(item, found);
      for (const other of found) { matched.set(other, [item]); open[many].splice(open[many].indexOf(other), 1); }
      open[one].splice(open[one].indexOf(item), 1);
    }
  };
  // Wisdom postings reversed the same day cancel out.
  for (const item of [...open.wisdom]) {
    if (!open.wisdom.includes(item)) continue;
    const reversal = open.wisdom.find(other => other !== item && other.date === item.date && other.amount === -item.amount);
    if (!reversal) continue;
    for (const [a, b] of [[item, reversal], [reversal, item]]) { matched.set(a, [b]); a.reversed = true; open.wisdom.splice(open.wisdom.indexOf(a), 1); }
  }
  combine("wisdom", "vizo");
  combine("vizo", "wisdom");
  const take = (group, others) => {
    for (const item of group) { matched.set(item, others); open[item.side === "VIZO" ? "vizo" : "wisdom"].splice(open[item.side === "VIZO" ? "vizo" : "wisdom"].indexOf(item), 1); }
  };
  const subsets = (pool, size, required) => {
    const out = [];
    const walk = (start, picked) => {
      if (picked.length && (!required || picked.includes(required))) out.push(picked);
      if (picked.length === size) return;
      for (let i = start; i < pool.length; i++) walk(i + 1, [...picked, pool[i]]);
    };
    walk(0, []);
    return out;
  };
  const total = rows => rows.reduce((sum, item) => sum + item.amount, 0);
  // A few deposits that together equal a few Wisdom postings within the same days (checks and RDC deposits posted
  // separately in Wisdom and combined by VIZO, or the reverse).
  for (const item of [...open.vizo]) {
    if (!open.vizo.includes(item)) continue;
    const near = (other, before, after) => vbDays(other.date, item.date) >= -before && vbDays(other.date, item.date) <= after;
    const vizoPool = open.vizo.filter(other => near(other, 0, 3)).slice(0, 10);
    const wisdomPool = open.wisdom.filter(other => near(other, 6, 3)).slice(0, 10);
    const sums = new Map();
    for (const group of subsets(wisdomPool, 3)) { const key = total(group); if (!sums.has(key)) sums.set(key, group); }
    const found = subsets(vizoPool, 3, item).find(group => group.length + (sums.get(total(group))?.length ?? 9) > 2 && sums.has(total(group)));
    if (!found) continue;
    const others = sums.get(total(found));
    take(found, others);
    take(others, found);
  }
  // Activity VIZO lists line by line that Wisdom books as a month total (investment interest and paydowns by CUSIP).
  const category = description => {
    const [payee, ...rest] = vbCellText(description).replace(/^(external|descriptive)\s+(deposit|withdrawal)\s*/i, "").split(" - ");
    const words = text => (text ?? "").toUpperCase().split(/\s+/).filter(word => /^[A-Z]+$/.test(word));
    return words(payee).slice(0, 2).join(" ") + "|" + words(rest.join(" - ")).slice(0, 1).join(" ");
  };
  const categories = new Map();
  for (const item of open.vizo) {
    const key = category(item.description) + "|" + Math.sign(item.amount);
    categories.set(key, [...(categories.get(key) ?? []), item]);
  }
  for (const group of categories.values()) {
    if (group.length < 2) continue;
    const pool = open.wisdom.slice(0, 30);
    const found = subsets(pool, 4).find(others => total(others) === total(group));
    if (!found) continue;
    take(group, found);
    take(found, group);
  }
  // Investment trades and returned items Wisdom books days later: the same amount anywhere in the month.
  for (const item of [...open.vizo]) {
    const other = open.wisdom.find(candidate => candidate.amount === item.amount);
    if (other) { take([item], [other]); take([other], [item]); }
  }
  return { matched, vizo: open.vizo, wisdom: open.wisdom };
}

// The last business day of the month through month-end: items there clear the next business day (in transit).
function vbTransitWindow(month) {
  const [year, number] = month.split("-").map(Number);
  let day = new Date(Date.UTC(year, number, 0));
  const end = day.toISOString().slice(0, 10);
  while ([0, 6].includes(day.getUTCDay())) day = new Date(day.getTime() - 86400000);
  return { from: day.toISOString().slice(0, 10), to: end };
}

function vbReconcile({ month, vizo, wisdom, books, invoice }) {
  const monthEnd = vbTransitWindow(month).to, window = vbTransitWindow(month);
  const byKind = new Map();
  for (const [gl, ledger] of wisdom) if (ledger.kind) byKind.set(ledger.kind, { kind: ledger.kind, gl, ledger });
  for (const account of vizo.values()) {
    if (!account.kind) continue;
    const entry = byKind.get(account.kind) ?? { kind: account.kind };
    entry.vizo = account;
    byKind.set(account.kind, entry);
  }
  for (const book of books) {
    const entry = byKind.get(book.kind) ?? { kind: book.kind };
    entry.book = book;
    entry.gl ??= book.gl;
    byKind.set(book.kind, entry);
  }
  const accounts = [...byKind.values()];
  const numberKind = new Map([...vizo.values()].map(account => [account.number, account.kind]));
  const glOf = kind => byKind.get(kind)?.gl;
  const entries = [], review = [], transit = [];
  const templatesFor = kind => byKind.get(kind)?.book?.templates ?? [];
  const allTemplates = accounts.flatMap(account => (account.book?.templates ?? []).map(item => ({ ...item, kind: account.kind })));
  const addReview = (account, item, reason) => review.push({ account, date: item?.date ?? monthEnd, description: item?.description ?? "", amount: item?.amount ?? 0, reason });

  // Internal activity between the VIZO accounts: overdraft sweeps, transfers and interest sweeps.
  const internal = [];
  for (const account of accounts) {
    for (const item of account.vizo?.rows ?? []) {
      item.class = vbClassify(item);
      if (item.class.type !== "EXT") internal.push({ account, item });
    }
  }
  const deposits = internal.filter(({ item }) => item.class.direction === "in");
  const groups = new Map();
  for (const { account, item } of internal) {
    const { type, direction, other } = item.class;
    if (type === "INTCREDIT" || direction !== "out") continue;
    const to = numberKind.get(other);
    const partnerType = type === "INT" ? "INT" : type;
    const partner = deposits.find(({ account: target, item: deposit }) => target.kind === to && !deposit.paired && deposit.date === item.date
      && deposit.amount === -item.amount && deposit.class.type === partnerType);
    if (partner) { partner.item.paired = item; item.paired = partner.item; }
    else if (!to) { addReview(account.kind, item, "Transfer to an account that is not in the VIZO export"); continue; }
    if (type === "INT") continue;
    const key = [type, account.kind, to].join("|");
    const group = groups.get(key) ?? { type, from: account.kind, to, rows: [], amount: 0 };
    group.rows.push(item);
    group.amount += -item.amount;
    groups.set(key, group);
  }
  // One deposit can be two accounts' overdraft sweeps on the same day (for example PSA and Settlement both covering ATM).
  for (const { account, item } of deposits) {
    if (item.paired || item.class.type !== "OD") continue;
    const sources = internal.filter(({ item: other }) => other.class.type === "OD" && other.class.direction === "out" && !other.paired
      && other.date === item.date && numberKind.get(other.class.other) === account.kind);
    const pair = sources.flatMap((a, i) => sources.slice(i + 1).map(b => [a, b])).find(([a, b]) => -(a.item.amount + b.item.amount) === item.amount);
    if (pair) { item.paired = pair.map(other => other.item); for (const other of pair) other.item.paired = item; }
  }
  for (const { account, item } of deposits) {
    if (item.paired || item.class.type === "INT") continue;
    // A transfer in with no matching transfer out in another account (for example from an account not in the export).
    const from = item.class.other ? numberKind.get(item.class.other) : undefined;
    if (from && item.class.type === "XFER") {
      const key = ["XFER", from, account.kind].join("|");
      const group = groups.get(key) ?? { type: "XFER", from, to: account.kind, rows: [], amount: 0 };
      group.rows.push(item);
      group.amount += item.amount;
      groups.set(key, group);
    } else addReview(account.kind, item, "Overdraft or transfer deposit without the other account's withdrawal");
  }
  const memoFor = (dr, cr, fallback) => {
    const template = allTemplates.find(item => item.dr === dr && item.cr === cr && /\b(od|overdraft|to)\b/i.test(item.memo));
    return template ? vbMemo(template.memo, month) : fallback;
  };
  for (const group of groups.values()) {
    const [dr, cr] = [glOf(group.to), glOf(group.from)];
    const fallback = (group.type === "OD" ? "OD " : "VIZO ") + vbShort(group.from) + " TO " + vbShort(group.to);
    entries.push({ account: group.from, date: monthEnd, memo: memoFor(dr, cr, fallback), amount: group.amount, dr, cr, rows: group.rows,
      basis: group.rows.length + " VIZO " + (group.type === "OD" ? "overdraft sweeps" : "transfers") });
  }
  // Interest: each account's interest is swept to Settlement; Settlement books the total.
  const settlement = byKind.get("SETTLEMENT");
  const interestRows = (settlement?.vizo?.rows ?? []).filter(item => item.class.type === "INTCREDIT" || (item.class.type === "INT" && item.class.direction === "in"));
  if (interestRows.length) {
    const template = templatesFor("SETTLEMENT").find(item => /\bint\b|interest/i.test(item.memo));
    entries.push({ account: "SETTLEMENT", date: monthEnd, memo: template ? vbMemo(template.memo, month) : "VIZO INT " + month.slice(5) + "/" + month.slice(2, 4),
      amount: interestRows.reduce((sum, item) => sum + item.amount, 0), dr: settlement.gl, cr: template?.cr,
      rows: interestRows, basis: interestRows.length + " interest credits and sweeps", needsGl: !template?.cr });
  }
  for (const account of accounts) {
    for (const item of account.vizo?.rows ?? []) {
      if (item.class.type === "INT" && item.class.direction === "out" && !item.paired) addReview(account.kind, item, "Interest sweep without Settlement's deposit");
    }
  }

  // External activity: matched to Wisdom, booked from last month's entries, in transit, or for review.
  for (const account of accounts) {
    const ledgerRows = (account.ledger?.rows ?? []).filter(item => !/^vizo\b/i.test(item.reference));
    const booked = (account.ledger?.rows ?? []).filter(item => /^vizo\b/i.test(item.reference));
    account.booked = booked;
    // VIZO activity booked by hand last month (fees, card payments, disputes) is an entry again; the rest is matched to Wisdom.
    const templatesOf = item => templatesFor(account.kind).filter(template => template.signatures.includes(vbSignature(item.description)));
    const external = (account.vizo?.rows ?? []).filter(item => item.class.type === "EXT");
    const match = vbMatch(external.filter(item => !templatesOf(item).length), ledgerRows);
    account.match = match;
    const unmatched = new Map();
    for (const item of external.filter(item => templatesOf(item).length)) {
      const signature = vbSignature(item.description);
      const group = unmatched.get(signature) ?? { signature, templates: templatesOf(item), rows: [] };
      group.rows.push(item);
      unmatched.set(signature, group);
    }
    // Last month's in-transit items clear early this month: one VIZO or Wisdom row of the same amount, or up to three rows
    // of one side in the first ten days that add up to it. Items that do not clear carry forward.
    const early = item => item.date < window.from && vbDays(item.date, month + "-01") <= 10;
    account.cleared = [];
    account.carried = [];
    // An item's effect on the tie (t) is cleared by a VIZO row of -t (Wisdom posted first) or a Wisdom row of t.
    const sign = account.book?.itemsSign ?? 1;
    const want = (pool, item) => (pool === match.vizo ? -1 : 1) * sign * item.amount;
    for (const item of account.book?.items ?? []) {
      const pools = [match.vizo, match.wisdom];
      let found;
      for (const pool of pools) {
        const single = pool.find(other => other.amount === want(pool, item) && other.date < window.from);
        if (single) { found = [pool, [single]]; break; }
      }
      for (const pool of found ? [] : pools) {
        const near = pool.filter(early).slice(0, 15);
        for (let a = 0; a < near.length && !found; a++) for (let b = a + 1; b < near.length && !found; b++) {
          if (near[a].amount + near[b].amount === want(pool, item)) found = [pool, [near[a], near[b]]];
          for (let c = b + 1; c < near.length && !found; c++) {
            if (near[a].amount + near[b].amount + near[c].amount === want(pool, item)) found = [pool, [near[a], near[b], near[c]]];
          }
        }
        if (found) break;
      }
      if (!found) { account.carried.push(item); continue; }
      for (const other of found[1]) found[0].splice(found[0].indexOf(other), 1);
      account.cleared.push({ item, rows: found[1] });
    }
    // Two or three of last month's items cleared by one row.
    for (const pool of [match.vizo, match.wisdom]) {
      for (const row of pool.filter(early)) {
        const open = account.carried, sizes = [];
        for (let a = 0; a < open.length; a++) for (let b = a + 1; b < open.length; b++) {
          sizes.push([open[a], open[b]]);
          for (let c = b + 1; c < open.length; c++) sizes.push([open[a], open[b], open[c]]);
        }
        const items = sizes.find(group => group.reduce((sum, item) => sum + want(pool, item), 0) === row.amount);
        if (!items) continue;
        pool.splice(pool.indexOf(row), 1);
        account.carried = open.filter(item => !items.includes(item));
        for (const item of items) account.cleared.push({ item, rows: [row] });
      }
    }
    for (const item of match.vizo) {
      if (item.date >= window.from && ["SETTLEMENT", "ATM"].includes(account.kind)) transit.push({ account: account.kind, item, amount: item.amount, description: item.description, side: "VIZO" });
      else addReview(account.kind, item, "VIZO activity not in Wisdom and no entry like it last month");
    }
    for (const group of unmatched.values()) {
      const amount = group.rows.reduce((sum, item) => sum + item.amount, 0);
      const { templates } = group;
      const each = templates.length > 1 && templates.every(item => item.memo === templates[0].memo);
      if (/lendkey/i.test(group.signature) && templatesFor(account.kind).filter(item => /\blk\b|lendkey/i.test(item.memo)).length > 1) {
        // The LendKey payment split by product from the invoice. Lines naming no product (the collection fee) went to Student
        // in last month's split; sales tax goes with Home Improvement.
        const lkTemplates = templatesFor(account.kind).filter(item => /\blk\b|lendkey/i.test(item.memo));
        if (!invoice || invoice.due !== -amount) {
          addReview(account.kind, group.rows[0], invoice ? "The LendKey payment is not the invoice amount due (" + vbMoney.format(invoice.due / 100) + ")" : "Load the LendKey invoice to split this payment");
          continue;
        }
        const products = [["HIL", /home improvement/i], ["CNL", /consolidation/i], ["STU", /student/i]];
        const productOf = line => products.find(([, pattern]) => pattern.test(line.description))?.[0];
        for (const line of invoice.lines.filter(line => !productOf(line))) {
          addReview(account.kind, { date: monthEnd, description: "LendKey invoice line " + line.number + ": " + line.description, amount: -line.amount },
            "Add to the LK entry of the product it belongs to (it is not in the split)");
        }
        const split = products.map(([code]) => [code, invoice.lines.filter(line => productOf(line) === code).reduce((sum, line) => sum + line.amount, 0) + (code === "HIL" ? invoice.tax : 0)]);
        for (const [code, value] of split) {
          if (!value) continue;
          const template = lkTemplates.find(item => item.memo.toUpperCase().split(/[^A-Z]+/).includes(code));
          if (!template) { addReview(account.kind, group.rows[0], "LendKey invoice product " + code + " has no entry last month"); continue; }
          entries.push({ account: account.kind, date: monthEnd, memo: vbMemo(template.memo, month), amount: value, dr: template.dr, cr: template.cr,
            rows: group.rows, basis: "LendKey payment split by invoice " + invoice.file });
        }
        continue;
      }
      const rowsList = each ? group.rows.map(item => [item]) : [group.rows];
      for (const rows of rowsList) {
        const template = templates[0];
        const value = rows.reduce((sum, item) => sum + item.amount, 0);
        const needsSplit = templates.length > 1 && !each;
        entries.push({ account: account.kind, date: monthEnd, memo: vbMemo(template.memo, month), amount: Math.abs(value), dr: template.dr, cr: template.cr,
          rows, basis: rows.length + " VIZO rows like last month's " + template.memo, needsGl: needsSplit,
          note: needsSplit ? "Last month this was split into " + templates.length + " entries; add the split" : "" });
      }
    }
    for (const item of match.wisdom) {
      if (item.date >= window.from) transit.push({ account: account.kind, item, amount: -item.amount, description: item.description, side: "Wisdom" });
      else {
        // Wisdom activity that never reached VIZO stays open on the cover (e.g. card funding suspense, odd physical checks).
        transit.push({ account: account.kind, item, amount: -item.amount, description: item.description, side: "Wisdom (open item)" });
        addReview(account.kind, item, "Wisdom activity not on the VIZO statement (listed as an open item on the cover)");
      }
    }
  }
  // Entries already in Wisdom (the export was pulled after they were made) are shown as booked, not to make.
  for (const entry of entries) {
    const ledger = byKind.get(entry.account)?.ledger;
    const pool = [...byKind.values()].flatMap(account => account.booked ?? []);
    // The same memo on the entry's G/L (net of reversals), or else one posting of the same amount that day.
    const words = text => vbCellText(text).toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim().replace(/^REV /, "");
    const sameMemo = pool.filter(item => (item.gl === entry.dr || item.gl === entry.cr) && words(item.description) === words(entry.memo));
    const net = Math.abs(sameMemo.filter(item => item.gl === sameMemo[0]?.gl).reduce((sum, item) => sum + item.amount, 0));
    entry.booked = sameMemo.length ? net === entry.amount : pool.some(item => Math.abs(item.amount) === Math.abs(entry.amount) && item.date === entry.date
      && (item.gl === entry.dr || item.gl === entry.cr));
    if (sameMemo.length && !entry.booked) entry.note = [entry.note, "In Wisdom as " + vbMoney.format(net / 100)].filter(Boolean).join(" · ");
    entry.status = entry.needsGl || (sameMemo.length && !entry.booked) ? "Review" : entry.booked ? "Already in Wisdom" : "To make";
    if (!ledger && !entry.booked && entry.status !== "Review") entry.status = "To make";
  }

  // Each account: statement against the Wisdom balance after the entries, the new in-transit items and carried items.
  const summary = accounts.filter(account => account.kind).map(account => {
    const gl = account.gl;
    // An account missing from an export had no activity there: last month's cover balance stands (flagged for review).
    const statement = account.vizo ? account.vizo.closing : account.book?.statement ?? 0;
    const ledger = account.ledger ? account.ledger.ending : account.book?.ledgerBalance ?? 0;
    if (!account.vizo) addReview(account.kind, undefined, "No VIZO activity in the exports; last month's statement balance is used. Confirm the account had no activity.");
    if (!account.ledger) addReview(account.kind, undefined, "G/L not in the Wisdom reports; last month's G/L balance is used. Confirm it had no activity.");
    const toMake = entries.filter(entry => !entry.booked && entry.status !== "Review");
    const effect = toMake.reduce((sum, entry) => sum + (entry.dr === gl ? entry.amount : 0) - (entry.cr === gl ? entry.amount : 0), 0);
    const newTransit = transit.filter(item => item.account === account.kind);
    // Open items carried from last month's cover, signed the way the cover adds them (to the G/L or to the statement).
    const carried = account.carried ?? account.book?.items ?? [];
    const carriedSign = account.book?.itemsSign ?? 1;
    const transitTotal = newTransit.reduce((sum, item) => sum + item.amount, 0) + carriedSign * carried.reduce((sum, item) => sum + item.amount, 0);
    const difference = statement - (ledger + effect + transitTotal);
    return { account, kind: account.kind, gl, statement, ledger, effect, projected: ledger + effect, newTransit, carried, cleared: account.cleared ?? [], transitTotal, difference };
  });
  return { month, accounts, entries, review, transit, summary };
}

function vbShow(result) {
  const table = document.createElement("table"), body = table.createTBody();
  const add = (name, value) => {
    const row = body.insertRow(), heading = document.createElement("th");
    heading.scope = "row";
    heading.textContent = name;
    row.append(heading);
    row.insertCell().textContent = value;
  };
  add("Entries to make in Wisdom", result.entries.filter(entry => entry.status === "To make").length);
  add("Entries already in Wisdom", result.entries.filter(entry => entry.status === "Already in Wisdom").length);
  add("Items to review", result.review.length + result.entries.filter(entry => entry.status === "Review").length);
  for (const item of result.summary) {
    add(vbShort(item.kind) + " difference after entries and in transit", vbMoney.format(item.difference / 100));
  }
  vbField("summary").replaceChildren(table);
  vbField("results").hidden = false;
}

// Each account's workbook with the cover's month, statement and Wisdom balances (after the entries) and the in-transit
// list: last month's open items that did not clear, then this month's new items. A cover without enough item rows keeps
// last month's list and is flagged.
async function vbCovers(result) {
  const monthName = new Intl.DateTimeFormat("en-US", { month: "long", timeZone: "UTC" }).format(new Date(result.month + "-01T00:00:00Z"));
  let clears = new Date(vbTransitWindow(result.month).to + "T00:00:00Z");
  do clears = new Date(clears.getTime() + 86400000); while ([0, 6].includes(clears.getUTCDay()));
  const nextBusinessDay = clears.toISOString().slice(0, 10);
  const covers = [], notes = [];
  for (const item of result.summary) {
    const book = item.account.book;
    if (!book) continue;
    const sheet = book.coverName;
    const updates = [[sheet, "L" + book.statementRow, item.statement / 100], [sheet, "L" + book.ledgerRow, item.projected / 100]];
    if (book.titleRow) updates.push([sheet, "C" + book.titleRow, "For the Month of " + monthName]);
    const sign = book.itemsSign ?? 1;
    // A cover that carries its month-end Wisdom postings on one clearing line gets their total there instead of a list.
    const clearing = book.clearingRow && !book.items.length && item.newTransit.length && item.newTransit.every(entry => entry.side === "Wisdom");
    if (clearing) updates.push([sheet, "L" + book.clearingRow, -item.newTransit.reduce((sum, entry) => sum + entry.amount, 0) / 100]);
    const lines = clearing ? [] : [
      ...item.carried.map(entry => [entry.description, entry.clears ?? null, entry.date ?? null, entry.amount / 100, entry.note || null]),
      ...item.newTransit.map(entry => [entry.description, /open/.test(entry.side) ? null : { date: nextBusinessDay }, { date: entry.item.date }, sign * entry.amount / 100, null]),
    ];
    const rows = book.adjustedRow && book.transitRow ? Array.from({ length: book.adjustedRow - book.transitRow - 2 }, (_, k) => book.transitRow + 2 + k) : [];
    const itemUpdates = rows.length >= lines.length ? rows.flatMap((row, k) => ["F", "I", "K", "L", "M"].map((column, c) =>
      [sheet, column + row, lines[k]?.[c] ?? null])) : [];
    const name = book.file.name.replace(/^\d{1,2}[.-]\d{2}/, String(Number(result.month.slice(5))) + "." + result.month.slice(2, 4));
    let blob;
    try {
      blob = await fillExcelTemplate(book.file, [...updates, ...itemUpdates]);
      if (!itemUpdates.length && lines.length + book.items.length) throw new Error("not enough item rows");
    } catch (error) {
      blob ??= await fillExcelTemplate(book.file, updates);
      notes.push([book.title, result.month + "-01", "Cover in-transit list", 0, "The cover keeps last month's in-transit list: add " + lines.length
        + " items (" + (rows.length < lines.length ? "the list has " + rows.length + " rows" : error.message) + ")."]);
    }
    covers.push([name, blob]);
  }
  return { covers, notes };
}

async function vbBuild(result) {
  const audit = reconciliationAudit();
  const { covers, notes } = await vbCovers(result);
  const { label, number, date, row, worksheet } = excelOutputCells();
  // Yellow (entry to make) and red (review) fills, borrowed like a workbook's formats: text, number and date for each.
  const fill = color => `<fill><patternFill patternType="solid"><fgColor rgb="${color}"/><bgColor indexed="64"/></patternFill></fill>`;
  const highlights = { styles: `<styleSheet xmlns="${spreadsheetNs}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
    <fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fill("FFFFFF00")}${fill("FFFF7C80")}</fills>
    <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count="6">
    ${[[0, 2], [0, 3], [4, 2], [4, 3], [14, 2], [14, 3]].map(([format, fillId]) => `<xf numFmtId="${format}" fontId="0" fillId="${fillId}" borderId="0" applyFill="1" applyNumberFormat="1"/>`).join("")}
    </cellXfs></styleSheet>` };
  const yellow = 0, red = 1;
  const cellsFor = (r, values, highlight) => values.map((value, i) => {
    const address = csvColumn(i + 1) + r, shade = kind => highlight === undefined ? undefined : "x" + (kind + highlight);
    if (typeof value === "number") return number(address, value, shade(2) ?? 2);
    if (/^\d{4}-\d{2}-\d{2}$/.test(value ?? "")) return date(address, value, shade(4) ?? 3);
    return label(address, value ?? "", shade(0) ?? 0);
  }).join("");
  const table = (name, headers, rows, widths, highlightOf) => [name, worksheet("A1:" + csvColumn(headers.length) + (rows.length + 1),
    widths.map((width, i) => [i + 1, i + 1, width]),
    row(1, headers.map((title, i) => label(csvColumn(i + 1) + "1", title, 1)).join(""))
    + rows.map((values, i) => row(i + 2, cellsFor(i + 2, values, highlightOf?.(i)))).join(""))];
  const name = kind => result.accounts.find(account => account.kind === kind)?.book?.title || vbShort(kind);
  const entryRows = [...result.entries].sort((a, b) => a.status.localeCompare(b.status) || a.account.localeCompare(b.account));
  const reviewRows = [
    ...result.entries.filter(entry => entry.status === "Review").map(entry => [name(entry.account), entry.date, entry.memo, entry.amount / 100, entry.note || "Add the credit G/L (not shown in last month's workbook)"]),
    ...result.review.map(item => [name(item.account), item.date, item.description, item.amount / 100, item.reason]),
    ...notes,
  ];
  const sheets = [
    table("Entries to make", ["Account", "Date", "Memo", "Amount", "Debit G/L", "Credit G/L", "Status", "Basis"],
      entryRows.map(entry => [name(entry.account), entry.date, entry.memo, entry.amount / 100, entry.dr ?? "", entry.cr ?? "", entry.status, entry.basis + (entry.note ? " · " + entry.note : "")]),
      [26, 13, 32, 16, 12, 12, 18, 60], i => entryRows[i].status === "To make" ? yellow : entryRows[i].status === "Review" ? red : undefined),
    table("Review", ["Account", "Date", "Description", "Amount", "Reason"], reviewRows, [26, 13, 60, 16, 60], () => red),
    table("Accounts", ["Account", "G/L", "VIZO statement", "Wisdom balance", "Entries to make", "Wisdom after entries", "In transit and open items", "Difference"],
      result.summary.map(item => [name(item.kind), item.gl ?? "", item.statement / 100, item.ledger / 100, item.effect / 100, item.projected / 100, item.transitTotal / 100, item.difference / 100]),
      [26, 10, 18, 18, 16, 20, 22, 14], i => result.summary[i].difference ? red : undefined),
    table("In transit", ["Account", "Date", "Description", "Amount", "Source"],
      result.summary.flatMap(item => [
        ...item.newTransit.map(entry => [name(item.kind), entry.item.date, entry.description, entry.amount / 100, entry.side + " · clears next business day"]),
        ...item.carried.map(entry => [name(item.kind), typeof entry.date === "number" ? sourceDate(entry.date) : vbCellText(entry.date), entry.description, entry.amount / 100, "Carried from last month's cover" + (entry.note ? " · " + entry.note : "")]),
      ]), [26, 13, 60, 16, 44]),
  ];
  const workbook = await createExcelWorkbook(sheets, audit, highlights);
  const zip = new JSZip();
  zip.file("VIZO entries " + result.month + ".xlsx", workbook);
  for (const [name, blob] of covers) zip.file(name, blob);
  return zip.generateAsync({ type: "blob", compression: "DEFLATE" });
}

for (const kind of Object.keys(vbFiles)) {
  const input = vbField(kind + "-file"), drop = vbField(kind + "-drop");
  const select = files => {
    vbFiles[kind] = [...files];
    vbField(kind + "-status").textContent = vbFiles[kind].map(file => file.name).join(", ");
    vbResult = undefined;
    vbField("results").hidden = true;
  };
  input.addEventListener("change", () => select(input.files));
  drop.addEventListener("dragover", event => { event.preventDefault(); drop.classList.add("active"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("active"));
  drop.addEventListener("drop", event => { event.preventDefault(); drop.classList.remove("active"); input.value = ""; select(event.dataTransfer.files); });
}
document.addEventListener("dragover", event => event.preventDefault());
document.addEventListener("drop", event => event.preventDefault());

vbField("run").addEventListener("click", async () => {
  const current = ++vbRevision;
  vbResult = undefined;
  vbField("results").hidden = true;
  try {
    const month = vbField("month").value;
    requireReport(month && vbField("confirm").checked, "Choose the month and confirm the source files.");
    requireReport(vbFiles.vizo.length && vbFiles.wisdom.length, "Load the VIZO exports and the Wisdom reports.");
    const [vizo, wisdom] = await Promise.all([vbReadVizo(vbFiles.vizo, month), vbReadWisdom(vbFiles.wisdom, month)]);
    const knownGls = new Set(wisdom.keys());
    const books = [];
    for (const file of vbFiles.books) books.push(await vbReadBook(file, knownGls));
    const invoice = vbFiles.invoice.length ? await readVizoLendKeyInvoice(vbFiles.invoice[0], month) : undefined;
    if (current !== vbRevision) return;
    vbResult = vbReconcile({ month, vizo, wisdom, books, invoice });
    vbShow(vbResult);
    vbStatus("");
  } catch (error) {
    if (current === vbRevision) vbStatus(error.message, true);
  }
});

vbField("download").addEventListener("click", async event => {
  const result = vbResult;
  if (!result) return;
  event.currentTarget.disabled = true;
  try {
    const blob = await vbBuild(result);
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "VIZO " + result.month + ".zip";
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 30000);
  } catch (error) {
    vbStatus(error.message, true);
  } finally {
    event.currentTarget.disabled = false;
  }
});
