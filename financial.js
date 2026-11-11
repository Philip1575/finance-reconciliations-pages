"use strict";

async function readLocalPdfPages(file, maxPages = 1) {
  requireReport(/\.pdf$/i.test(file.name) && file.size <= 10 * 1024 * 1024, "Choose a PDF under 10 MB.");
  const task = pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()), useWasm: false,
    disableFontFace: true, useSystemFonts: false, isEvalSupported: false, stopAtErrors: true, verbosity: 0 });
  try {
    const pdf = await task.promise;
    requireReport(pdf.numPages <= maxPages, maxPages === 1 ? "Choose a one-page PDF."
      : "Choose a PDF with no more than " + maxPages + " pages.");
    const pages = [];
    for (let number = 1; number <= pdf.numPages; number++) {
      const page = await pdf.getPage(number);
      const { items } = await page.getTextContent();
      pages.push(items.filter(item => item.str?.trim()).map(item => ({
        x: item.transform[4], y: item.transform[5], width: item.width,
        text: item.str.trim().replace(/\s+/g, " "),
      })));
    }
    return pages;
  } finally {
    await task.destroy();
  }
}

async function readLocalPdfText(file) {
  return (await readLocalPdfPages(file))[0];
}

function cents(value) {
  const amount = readBalance(value);
  const result = Math.round(amount * 100);
  requireReport(Number.isSafeInteger(result) && Math.abs(amount * 100 - result) < 0.000001,
    "Use a valid amount with at most two decimal places.");
  return result;
}

function dateOffset(date, offset) {
  const time = new Date(date + "T00:00:00Z");
  time.setUTCDate(time.getUTCDate() + offset);
  return time.toISOString().slice(0, 10);
}

function readReviewedDepositAllocations(text, items, ledgerDates) {
  const eligible = new Map(items.map(item => [String(item.reviewRow ?? item.row), item]));
  return text.trim().split(/\r?\n/).filter(line => line.trim()).flatMap((line, index) => {
    const [reference, rawDate, rawBranch, rawMobile, ...extra] = line.split("\t").map(value => value.trim());
    if (index === 0 && /^provider row$/i.test(reference)) return [];
    const item = eligible.get(reference);
    requireReport(item && rawDate && rawBranch && rawMobile && !extra.length,
      "Deposit allocation line " + (index + 1) + " needs an unused provider row, Wisdom date, branch amount and mobile amount.");
    const ledgerDate = sourceDate(rawDate), branchAmount = cents(rawBranch), mobileAmount = cents(rawMobile);
    requireReport(ledgerDates.has(ledgerDate) && branchAmount >= 0 && mobileAmount >= 0 && branchAmount + mobileAmount === item.amount,
      "Use a date in the Wisdom history and amounts that sum to the provider deposit on row " + reference + ".");
    eligible.delete(reference);
    return [{ item, ledgerDate, branchAmount, mobileAmount }];
  });
}

function financeProviderColumn(cells, first, last = first) {
  const matches = [];
  for (const [address, value] of cells) {
    if (typeof value !== "string" || value.trim().toLowerCase().replace(/\s+/g, " ") !== "finance provider") continue;
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    const row = Number(match?.[2]);
    if (row >= first && row <= last) matches.push(match[1]);
  }
  requireReport(matches.length <= 1, "Use one Finance provider heading in the Fed activity table.");
  return matches[0];
}

function matchesFinanceProvider(value, provider) {
  const normalize = text => String(text ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const name = normalize(value);
  return name === normalize(provider) || (provider === "Student Choice" && name === "custudentchoice");
}

function fedLoanProvider(description) {
  if (/\bwire\b/i.test(description)) return "";
  if (/\blend[\s-]*key\b/i.test(description)) return "LendKey";
  if (/\bstudent choice\b|\bcampus[\s-]*door\b|\buniversity accou\b[^,]*,[^,]*,\s*edi pymts\b/i.test(description)) return "Student Choice";
  return "";
}

function providerFedHistory(cells, date1904, provider) {
  if (!/^detailed trial balance report$/i.test(String(cells.get("B1") ?? "").trim())) return undefined;
  const layout = wisdomHistoryLayout(cells);
  requireReport(layout.separateDates && /\bfederal reserve\b/i.test(layout.name),
    "Load the Federal Reserve Detailed Trial Balance or a reviewed provider activity workbook.");
  return wisdomHistory(cells, date1904, layout).rows.filter(row => row.amount && fedLoanProvider(row.description) === provider);
}

function sourceDate(value, date1904) {
  if (typeof value === "number") return excelDate(Math.floor(value), date1904);
  const text = String(value ?? "").trim();
  let match = /^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/.exec(text);
  if (!match) {
    const american = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
    if (american) match = [text, american[3], american[1].padStart(2, "0"), american[2].padStart(2, "0")];
  }
  requireReport(match, "Use Excel dates, YYYY-MM-DD, or MM/DD/YYYY in the date column.");
  const iso = match.slice(1).join("-");
  const timestamp = new Date(iso + "T00:00:00Z");
  requireReport(Number.isFinite(timestamp.getTime()) && timestamp.toISOString().slice(0, 10) === iso,
    "The source contains an invalid date.");
  return iso;
}

function wisdomHistoryLayout(cells) {
  if (!/^detailed trial balance report$/i.test(String(cells.get("B1") ?? "").trim())) return undefined;
  const starting = /^starting balance:\s*(.+)$/i.exec(String(cells.get("I2") ?? "").trim());
  const separateDates = Boolean(starting);
  const account = normalizeAccount(cells.get(separateDates ? "D2" : "A2"));
  const dates = /^From:\s*(\S+)\s+To:\s*(\S+)$/i.exec(String(cells.get("E1") ?? "").trim());
  const totals = [...cells].filter(([address, value]) => new RegExp("^" + (separateDates ? "A" : "B") + "[1-9]\\d*$").test(address) && value === "Totals:");
  requireReport(account && dates && (separateDates || /^beg bal:$/i.test(String(cells.get("D2") ?? "").trim())) && totals.length === 1,
    "The detailed Wisdom report must contain one G/L account with a date range, beginning balance and totals.");
  const total = Number(totals[0][0].slice(1));
  requireReport(cells.get((separateDates ? "B" : "C") + total) === "Beginning Balance"
    && cells.get((separateDates ? "F" : "G") + total) === "Ending Balance",
    "The detailed Wisdom report totals have an unexpected layout.");
  return { account, name: String(cells.get(separateDates ? "H2" : "C2") ?? ""), from: sourceDate(dates[1]),
    through: sourceDate(dates[2]), total, separateDates, openingText: starting?.[1] };
}

function wisdomHistory(cells, date1904, layout) {
  const { total, from, through, separateDates } = layout;
  const opening = cents(separateDates ? layout.openingText : cells.get("E2"));
  const balanceColumn = separateDates ? "K" : "G";
  requireReport(opening === cents(cells.get("H" + total)), "The Wisdom beginning balances disagree.");
  let running = opening, lastDate = from;
  const rows = [];
  for (let row = 3; row < total; row++) {
    if (!String(cells.get("A" + row) ?? "").trim()) {
      requireReport(!String(cells.get("F" + row) ?? "").trim() && !String(cells.get("G" + row) ?? "").trim(),
        "A Wisdom amount is missing its date at row " + row + ".");
      continue;
    }
    const date = sourceDate(cells.get("A" + row), date1904);
    const description = String(cells.get((separateDates ? "G" : "D") + row) ?? "").trim();
    const posted = separateDates ? sourceDate(cells.get("B" + row), date1904) : undefined;
    const rawDebit = separateDates ? cents(cells.get("I" + row)) : undefined;
    const rawCredit = separateDates ? cents(cells.get("J" + row)) : undefined;
    const amount = separateDates ? rawDebit + rawCredit : cents(cells.get("F" + row));
    const balance = cents(cells.get(balanceColumn + row));
    requireReport(!separateDates || (rawDebit >= 0 && rawCredit <= 0), "Check the Wisdom debit and credit signs at row " + row + ".");
    requireReport(date >= lastDate && date <= through && description && running + amount === balance,
      "Check the Wisdom date, description and running balance at row " + row + ".");
    lastDate = date;
    running = balance;
    rows.push({ date, posted, description, amount, debit: rawDebit ?? Math.max(0, amount), credit: rawCredit === undefined ? Math.max(0, -amount) : -rawCredit,
      batch: String(cells.get((separateDates ? "C" : "B") + row) ?? "").trim(), balance, row });
  }
  requireReport(running === cents(cells.get("L" + total)) && running - opening === cents(cells.get("I" + total)),
    "The detailed Wisdom activity does not reproduce the report totals.");
  const balanceOn = date => {
    requireReport(date >= from && date <= through, "The detailed Wisdom report does not cover the selected closing date.");
    const last = rows.findLast(row => row.date <= date);
    return { amount: last?.balance ?? opening, address: last ? balanceColumn + last.row : separateDates ? "I2" : "E2" };
  };
  return { rows, balanceOn };
}

function readCsvCells(text) {
  const records = [];
  let record = [], field = "", quoted = false, closedQuote = false;
  text = text.replace(/^\uFEFF/, "");
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (quoted) {
      if (char === '"' && text[i + 1] === '"') {
        field += '"';
        i++;
      } else if (char === '"') {
        quoted = false;
        closedQuote = true;
      } else {
        field += char;
      }
      continue;
    }
    if (char === "," || char === "\n" || char === "\r") {
      record.push(field);
      field = "";
      closedQuote = false;
      if (char !== ",") {
        records.push(record);
        record = [];
        if (char === "\r" && text[i + 1] === "\n") i++;
      }
    } else if (char === '"' && !field && !closedQuote) {
      quoted = true;
    } else {
      requireReport(!closedQuote && char !== '"', "The CSV has malformed quoting at record " + (records.length + 1) + ".");
      field += char;
    }
  }
  requireReport(!quoted, "The CSV ends inside a quoted field.");
  if (record.length || field || closedQuote) records.push([...record, field]);
  requireReport(records.length, "The CSV contains no rows.");
  const cells = new Map();
  for (const [row, values] of records.entries()) {
    for (const [column, value] of values.entries()) {
      if (value.trim()) cells.set(csvColumn(column + 1) + (row + 1), value);
    }
  }
  return cells;
}

function csvColumn(index) {
  let letters = "";
  for (let value = index; value > 0; value = Math.floor((value - 1) / 26)) {
    letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
  }
  return letters;
}
