"use strict";

const fedField = id => document.getElementById(id);
const dailyFiles = {};
let dailyResult, dailyRevision = 0;
let dailyBatchRevision = 0, dailyBatch = [], dailyReferences = [];
let dailyBatchAssignments = new Map();
const dailyRoles = [
  ["statement", "Fed Statement of Account"], ["ledger", "Fed Wisdom G/L History"],
  ["template", "Prepared Fed Rec with Ops inputs"], ["vizo", "VIZO Settlement activity"],
  ["settlement", "Wisdom Settlement G/L History"], ["refresh-ledger", "Refreshed Fed Wisdom history"],
  ["sent-current", "Current day Sent Files Summary"], ["sent-prior", "Prior day Sent Files Summary"],
  ["returns-current", "Current day Report 174"], ["returns-prior", "Prior day Report 174"],
  ["reference", "Supporting file for manual review"], ["", "Leave out"],
];
const dailyCategories = ["Wire in", "Wire out", "ACH in", "ACH out", "ACH settlement", "FedNow in", "FedNow out"];
const settlementCategories = [
  ["Share drafts", /^share\s*drafts?$/i, 0],
  ["Branch deposits", /branch capture/i, 1],
  ["Mobile deposits", /mobile capture/i, 1],
  ["Returned deposits", /forward check collection\s*-\s*ret\b/i, 0],
  ["Share draft suspense", /^(?:sd|share\s*draft) suspense$/i, 0],
];

function dailyStatus(message, error = false) {
  fedField("status").textContent = message;
  fedField("status").classList.toggle("error", error);
}

function invalidateDaily() {
  dailyRevision++;
  dailyResult = undefined;
  fedField("results").hidden = true;
  fedField("confirm").checked = false;
  dailyStatus("");
}

function selectDaily(kind, file) {
  dailyFiles[kind] = file;
  fedField(kind + "-status").textContent = file?.name ?? "";
  if (kind === "template") fedField("adjustment").disabled = Boolean(file);
  if (kind.startsWith("sent-")) fedField("sent-approved").checked = false;
  if (kind === "vizo" || kind === "settlement") fedField("fed-deposit-allocations").value = "";
  if (kind.startsWith("returns-")) {
    for (const field of ["date", "debit", "credit"]) fedField(kind + "-" + field).value = "";
    fedField(kind + "-page").replaceChildren();
  }
  invalidateDaily();
}

async function suggestDailyRole(file, date) {
  if (/\.txt$/i.test(file.name) && file.size <= 10 * 1024 * 1024) {
    const beginning = await file.slice(0, 32000).text();
    if (/SUMMARY STATEMENT OF ACCOUNT/i.test(beginning)) return "statement";
    if (/Sent Files Summary Report/i.test(beginning)) {
      const reportDate = sentFilesDate(beginning);
      return date && reportDate < date ? "sent-prior" : !date || date === reportDate ? "sent-current" : "";
    }
  }
  if (/\.pdf$/i.test(file.name)) {
    const returns = /^(\d{2})\.(\d{2})\.(\d{2}|\d{4})\b.*\b174\b/i.exec(file.name);
    if (returns && date) {
      const reportDate = sourceDate(returns[3].padStart(4, "20") + "-" + returns[1] + "-" + returns[2]);
      return reportDate === date ? "returns-current" : reportDate < date ? "returns-prior" : "reference";
    }
    return /\bstatement of accounts?\b/i.test(file.name) ? "statement" : "reference";
  }
  if (/\.csv$/i.test(file.name) && file.size <= 20 * 1024 * 1024) {
    const beginning = await file.slice(0, 32000).text();
    if (/Account Name.+Transaction Date.+Transaction Description.+Balance/i.test(beginning)
      && /(?:^|,)"?Settlement Account"?,/mi.test(beginning)) return "vizo";
  }
  if (/\.(?:xlsx|xlsm)$/i.test(file.name)) {
    const book = await openExcel(file, true);
    if (["Cover", "MATCHUP", "OPS BALANCING"].every(sheet => book.sheets.has(sheet))) return "template";
    if (book.sheets.size === 1) {
      const cells = await readCells(book, [...book.sheets.keys()][0]);
      const name = wisdomHistoryLayout(cells)?.name.trim() ?? "";
      if (/^vizo settlement(?: account)?$/i.test(name)) return "settlement";
      if (/\bfederal reserve\b|^fed(?: account)?$/i.test(name)) return "ledger";
    }
  }
  return "";
}

async function selectDailyBatch(files) {
  const current = ++dailyBatchRevision;
  const message = fedField("batch-status");
  dailyBatch = [];
  fedField("batch-review").hidden = true;
  message.textContent = files.length ? "Checking file types…" : "";
  const date = fedField("date").value;
  const suggested = await Promise.all(files.map(async file => ({ file, role: await suggestDailyRole(file, date).catch(() => "") })));
  if (current !== dailyBatchRevision || !suggested.length) return;
  const counts = new Map();
  for (const { role } of suggested) counts.set(role, (counts.get(role) ?? 0) + 1);
  const fields = suggested.map(({ file, role }) => {
    const label = document.createElement("label"), input = document.createElement("select");
    label.textContent = file.name;
    input.replaceChildren(...dailyRoles.map(([key, title]) => new Option(title, key)));
    input.value = role === "reference" || counts.get(role) === 1 ? role : "";
    label.append(input);
    dailyBatch.push({ file, input });
    return label;
  });
  fedField("batch-roles").replaceChildren(...fields);
  fedField("batch-review").hidden = false;
  message.textContent = "Review the assignments for this date. Supporting files are listed for manual review.";
}

fedField("batch-apply").addEventListener("click", () => {
  const selected = dailyBatch.filter(({ input }) => input.value && input.value !== "reference");
  const assignments = new Map(selected.map(({ file, input }) => [input.value, file]));
  if (assignments.size !== selected.length) {
    fedField("batch-status").textContent = "Choose one file for each reconciliation role.";
    return;
  }
  for (const [role, file] of dailyBatchAssignments) {
    if (!assignments.has(role) && dailyFiles[role] === file) selectDaily(role, undefined);
  }
  for (const [role, file] of assignments) {
    fedField(role + "-file").value = "";
    if (dailyFiles[role] !== file) selectDaily(role, file);
  }
  dailyBatchAssignments = assignments;
  dailyReferences = dailyBatch.filter(({ input }) => input.value === "reference").map(({ file }) => file.name);
  invalidateDaily();
  fedField("batch-status").textContent = assignments.size + " input files assigned. "
    + dailyReferences.length + " supporting files for manual review.";
});

function sentFilesDate(contents) {
  const dates = [...contents.matchAll(/\bProcess Date:\s*(\d{2})\/(\d{2})\/(\d{2}|\d{4})\b/gi)];
  const unique = new Set(dates.map(([, month, day, year]) => sourceDate(year.padStart(4, "20") + "-" + month + "-" + day)));
  requireReport(unique.size === 1, "The Sent Files Summary must identify one process date.");
  return [...unique][0];
}

async function showReturnsSummary(role) {
  const file = dailyFiles[role], panel = fedField(role + "-page");
  if (!file) {
    panel.textContent = "Load the Report 174 PDF first.";
    return;
  }
  panel.textContent = "Opening the last page…";
  let task;
  try {
    requireReport(/\.pdf$/i.test(file.name) && file.size <= 10 * 1024 * 1024, "Choose a Report 174 PDF under 10 MB.");
    const data = new Uint8Array(await file.arrayBuffer());
    if (dailyFiles[role] !== file) return;
    task = pdfjsLib.getDocument({ data, useWasm: false, disableFontFace: true, isEvalSupported: false, stopAtErrors: true });
    const pdf = await task.promise;
    const page = await pdf.getPage(pdf.numPages);
    const viewport = page.getViewport({ scale: 1.8 });
    requireReport(viewport.width * viewport.height <= 10_000_000, "The report page is too large to preview.");
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    canvas.setAttribute("aria-label", "Report 174 last page");
    await page.render({ canvas, viewport }).promise;
    if (dailyFiles[role] === file) panel.replaceChildren(canvas);
  } catch (error) {
    if (dailyFiles[role] === file) panel.textContent = error.message;
  } finally {
    if (task) await task.destroy();
  }
}

function reviewedReturns(date) {
  return ["returns-current", "returns-prior"].flatMap(role => {
    const file = dailyFiles[role], reportDate = fedField(role + "-date").value;
    const debit = fedField(role + "-debit").value.trim(), credit = fedField(role + "-credit").value.trim();
    if (!reportDate && !debit && !credit) return [];
    const prior = role === "returns-prior";
    requireReport(file && /\.pdf$/i.test(file.name) && reportDate && (prior ? reportDate < date : reportDate === date) && debit && credit,
      "Load the Report 174 PDF and verify its " + (prior ? "prior" : "current") + " run date and both totals.");
    const dr = cents(debit), cr = cents(credit);
    requireReport(dr >= 0 && cr >= 0, "Enter positive Report 174 totals or zero.");
    return [{ name: file.name, date: reportDate, prior, debit: dr, credit: cr }];
  });
}

for (const role of ["returns-current", "returns-prior"]) {
  fedField(role + "-preview").addEventListener("click", () => showReturnsSummary(role));
}

async function readSentFiles(file, date, prior, includeApproved) {
  requireReport(file.size <= 10 * 1024 * 1024, "Choose a Sent Files Summary under 10 MB.");
  const contents = await file.text();
  const reportDate = sentFilesDate(contents);
  requireReport(/Sent Files Summary Report/i.test(contents) && (prior ? reportDate < date : reportDate === date),
    "Check the process date on the " + (prior ? "prior" : "current") + " Sent Files Summary.");
  const lines = contents.split(/\r?\n/), rows = [], included = new Set();
  const header = lines.findIndex(line => /^\s*File\s+Transmission\/.*\bDebits\s+Credits\s*$/i.test(line));
  requireReport(header >= 0, "The Sent Files Summary needs debit and credit columns.");
  let debit = 0, credit = 0;
  for (let i = header + 3; i < lines.length; i++) {
    const line = lines[i];
    if (!/^\S/.test(line)) continue;
    const status = line.slice(0, 20).trim();
    const process = /^\s*\d{9}\s+(\d{2})\/(\d{2})\/(\d{2}|\d{4})\b/.exec(line.slice(20));
    if (!process && !/^(?:accepted|approved|file not confirmed)$/i.test(status)) continue;
    requireReport(process && sourceDate(process[3].padStart(4, "20") + "-" + process[1] + "-" + process[2]) === reportDate,
      "Check the process date on Sent Files Summary line " + (i + 1) + ".");
    const amounts = /(\$[\d,]+\.\d{2})\s+(\$[\d,]+\.\d{2})\s*$/.exec(line);
    requireReport(amounts, "Check the amounts on Sent Files Summary line " + (i + 1) + ".");
    const count = /^accepted$/i.test(status) || (includeApproved && /^approved$/i.test(status));
    const dr = cents(amounts[1]), cr = cents(amounts[2]);
    if (count) {
      const metadata = lines[i + 1]?.trim().split(/\s+/);
      requireReport(metadata?.length === 6 && /^\d+$/.test(metadata[1]) && /^\d{2}:\d{2}$/.test(metadata[2]),
        "Check the file identification on Sent Files Summary line " + (i + 2) + ".");
      const key = JSON.stringify([metadata[0], metadata[1], metadata[2], metadata[5], dr, cr]);
      requireReport(!included.has(key), "Review duplicate accepted or approved files on the Sent Files Summary.");
      included.add(key);
      debit += dr;
      credit += cr;
    }
    rows.push({ line: i + 1, status, debit: dr, credit: cr, count });
  }
  requireReport(rows.length && Number.isSafeInteger(debit) && Number.isSafeInteger(credit),
    "The Sent Files Summary has no readable file rows.");
  return { name: file.name, date: reportDate, prior, rows, debit, credit };
}

fedField("batch-files").addEventListener("change", event => selectDailyBatch([...event.target.files]));
fedField("batch-drop").addEventListener("dragover", event => {
  event.preventDefault();
  fedField("batch-drop").classList.add("active");
});
fedField("batch-drop").addEventListener("dragleave", () => fedField("batch-drop").classList.remove("active"));
fedField("batch-drop").addEventListener("drop", event => {
  event.preventDefault();
  fedField("batch-drop").classList.remove("active");
  fedField("batch-files").value = "";
  selectDailyBatch([...event.dataTransfer.files]);
});

function dailyHeader(cells) {
  const aliases = {
    date: ["date", "effective date", "transaction date"],
    category: ["category", "description", "transaction description"],
    finance: ["finance provider"],
    amount: ["signed amount", "net amount", "amount"],
    debit: ["debit", "debits", "debit amount"],
    credit: ["credit", "credits", "credit amount"],
  };
  const rows = new Map();
  for (const [address, value] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (!match || Number(match[2]) > 50 || typeof value !== "string") continue;
    const key = Object.keys(aliases).find(field => aliases[field].includes(value.trim().toLowerCase()));
    if (!key) continue;
    const row = Number(match[2]), fields = rows.get(row) ?? new Map();
    if (key === "date" && value.trim().toLowerCase() === "effective date"
      && cells.get("C" + row) === "Transaction Date" && cells.get("B" + row) === "Account Name") continue;
    fields.set(key, fields.has(key) ? "" : match[1]);
    rows.set(row, fields);
  }
  return [...rows].find(([, fields]) => fields.get("date") && fields.get("category")
    && (fields.get("amount") || (fields.get("debit") && fields.get("credit"))));
}

function dailyCategory(text, amount, source) {
  const description = text.trim().replace(/^\(\d+\)\s*/, "").toLowerCase();
  if (/^(closing balance|ending balance)$/.test(description)) return "Closing balance";
  const explicit = /^(wire|ach|fednow)\s+(in|out)$/.exec(description);
  if (explicit) {
    const inbound = explicit[2] === "in";
    if (amount !== 0 && (amount > 0) !== inbound) return "Direction conflicts with label";
    const prefix = { wire: "Wire", ach: "ACH", fednow: "FedNow" }[explicit[1]];
    return prefix + " " + explicit[2];
  }
  const direction = amount < 0 ? "out" : "in";
  if (/^(?:10\s+)?transfer of funds\b|^wire\b|^scr mem\b|\b(?:dom |intl? )?wire (?:in|out)\b/.test(description)) return "Wire " + direction;
  if (/^(?:\d{2}\s+)?fednow\b/.test(description)) return "FedNow " + direction;
  if (/^(?:57\s+)?settlement for ach\b|^ach (?:settlement|advice|ori?gination|suspense clearing)\b/.test(description)) {
    return source === "ledger" ? "ACH settlement" : "ACH " + direction;
  }
  return "Other · review";
}

function financeProvider(description) {
  const text = description.replace(/[-_]/g, " ").replace(/\s+/g, " ").trim();
  if (/\bwire\b|^scr mem\b|^fednow\b/i.test(text)) return "";
  const names = [
    ["LendKey", /\blend\s*key\b/], ["Student Choice", /\bstudent choice\b|\bcampus\s*door\b/],
    ["TruStage", /\btru\s*stage\b/], ["BankACount", /\bbank\s*a\s*count\b/],
    ["LoanStreet", /\bloan\s*street\b/], ["RA Rodgers", /\br\s*a\s+rodg?ers\b/],
    ["KS State", /\bks\s+state\b/],
  ];
  return names.find(([, pattern]) => pattern.test(text.toLowerCase()))?.[0] ?? "";
}

function nextFedPostingDate(date) {
  const next = new Date(date + "T00:00:00Z");
  do { next.setUTCDate(next.getUTCDate() + 1); } while ([0, 6].includes(next.getUTCDay()));
  return next.toISOString().slice(0, 10);
}

function settlementCategory(text) {
  const description = text.replace(/[-_]/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
  if (/\b(?:sd|share\s*draft) suspense\b/.test(description)) return "Share draft suspense";
  if (/\b(?:sd|share\s*draft) discrep(?:ancy|ency)\b/.test(description)) return "Share draft discrepancy · review";
  if (/\b(?:forward (?:check )?collection.*ret(?:urn)?|returned deposits?|rdi)\b/.test(description)) return "Returned deposits";
  if (/\b(?:share\s*draft(?:s| settlement| activity)?|cash letter|sd clearing)\b/.test(description)) return "Share drafts";
  if (/\b(?:mobile capt(?:ure)?|mobile deposits?|rdc(?: deposits?)?)\b/.test(description)) return "Mobile deposits";
  if (/\b(?:branch capture|branch deposits?|daily deposits?|physical checks?)\b/.test(description)) return "Branch deposits";
  if (/\bforward (?:check )?collection\b/.test(description)) return "Unclassified forward collection";
  return /^no activity$/.test(description) ? "No activity" : "Other · review";
}

async function readFedStatement(file, date) {
  try {
    let pages;
    if (/\.txt$/i.test(file.name)) {
      requireReport(file.size <= 10 * 1024 * 1024, "Choose a statement text file under 10 MB.");
      const contents = await file.text();
      pages = contents.split("\f").map(page => page.split(/\r?\n/).flatMap((line, i) =>
        [...line.matchAll(/\S+(?: \S+)*/g)].map(match => ({
          x: match.index, y: -3 * i, width: match[0].length, text: match[0],
        }))));
    } else {
      pages = await readLocalPdfPages(file, 100);
    }
    const lines = pages.flatMap((items, page) => {
      const rows = [];
      for (const item of items.sort((a, b) => b.y - a.y || a.x - b.x)) {
        if (!rows.length || rows.at(-1).y - item.y > 2) rows.push({ page: page + 1, y: item.y, items: [] });
        rows.at(-1).items.push(item);
      }
      return rows.map((row, i) => {
        row.items.sort((a, b) => a.x - b.x);
        return { ...row, number: i + 1, text: row.items.map(item => item.text).join(" ") };
      });
    });
    const titles = lines.filter(row => /^\**\s*summary (?:statement of|of statement of) account\s*\**$/i.test(row.text));
    requireReport(titles.length === 1, "Find one readable Summary Statement of Account in the PDF.");
    const title = titles[0], start = lines.indexOf(title);
    const afterTitle = lines.slice(start + 1);
    const detail = afterTitle.findIndex(row => /^(?:\**\s*)?(?:detail (?:statement|of)|future items summary)/i.test(row.text));
    const summary = detail < 0 ? afterTitle : afterTitle.slice(0, detail);
    const closing = summary.filter(row => /^closing balance as of\b/i.test(row.text));
    requireReport(closing.length === 1, "Find one dated closing balance in the statement summary.");
    const balanceLine = closing[0];
    const current = summary.slice(0, summary.indexOf(balanceLine));
    const datePart = /^closing balance as of\s+(\d{1,2})\/(\d{1,2})(?:\/(\d{4}|\d{2})(?!\d))?\s/i.exec(balanceLine.text);
    const [year, month, day] = date.split("-").map(Number);
    const yearMatches = value => Number(value) === year || (value.length === 2 && Number(value) === year % 100);
    const headerDate = lines.slice(0, start).filter(row => row.page === title.page)
      .flatMap(row => [...row.text.matchAll(/\b(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})\b/g)]);
    requireReport(datePart && Number(datePart[1]) === month && Number(datePart[2]) === day
      && (datePart[3] ? yearMatches(datePart[3]) : headerDate.some(([, m, d, y]) =>
        Number(m) === month && Number(d) === day && yearMatches(y))),
    "The PDF statement date does not match the selected date.");

    const money = /(?<![\d,.(])(?:-?\$?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}|\(\$?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}\))(?![\d,.)])/g;
    const scalar = row => {
      const amounts = [...row.text.matchAll(money)];
      requireReport(amounts.length === 1 && amounts[0].index + amounts[0][0].length === row.text.length,
        "Check the PDF amount on page " + row.page + ", line " + row.number + ".");
      return cents(amounts[0][0]);
    };
    const one = pattern => {
      const matching = current.filter(row => pattern.test(row.text));
      requireReport(matching.length === 1, "Check the opening balance and total debits and credits in the PDF summary.");
      return matching[0];
    };
    const opening = scalar(one(/^opening balance\s+(?!as of\b)/i));
    const totalDebits = scalar(one(/^total debits\b/i)), totalCredits = scalar(one(/^total credits\b/i));
    const balance = scalar(balanceLine);
    requireReport(totalDebits >= 0 && totalCredits >= 0 && opening - totalDebits + totalCredits === balance,
      "The opening balance and activity totals do not reproduce the closing balance.");

    const columns = current.find(row => /^debits?\s+credits?$/i.test(row.text));
    const sections = current.flatMap((row, index) => {
      const code = /^(\d{2})\s+[a-z]/i.exec(row.text)?.[1];
      return code ? [{ row, index, code, description: row.text.split(money, 1)[0].trim() }] : [];
    });
    const recap = current.find(row => /^own activity recap\b/i.test(row.text));
    const activityEnd = current.findIndex(row => /^(?:own activity recap|subaccount activity|osrtn activity|respondent activity|total\b|opening balance\s+(?!as of\b))/i.test(row.text));
    const activity = sections.filter(section => section.index < activityEnd);
    requireReport(new Set(activity.map(section => section.code)).size === activity.length,
      "A summary activity code appears more than once.");

    const columnAmounts = rows => {
      const amounts = [0, 0], seen = new Set();
      for (const row of rows) {
        const headings = current.filter(line => line.page === row.page && /^debits?\s+credits?$/i.test(line.text));
        const header = headings.at(-1) ?? columns;
        const [debit, credit] = header?.items ?? [];
        requireReport(debit && credit && /^debits?$/i.test(debit.text) && /^credits?$/i.test(credit.text) && credit.x - debit.x >= 12,
          "The debit and credit columns could not be distinguished in the statement.");
        for (const item of row.items) {
          for (const match of item.text.matchAll(money)) {
            const right = item.x + item.width * (match.index + match[0].length) / item.text.length;
            requireReport(right > debit.x - 20 && right < credit.x + credit.width + 60,
              "A PDF amount is outside the debit and credit columns.");
            const col = right < (debit.x + debit.width + credit.x + credit.width) / 2 ? 0 : 1;
            requireReport(!seen.has(col), "A PDF activity total has more than one amount in a column.");
            amounts[col] = cents(match[0]);
            requireReport(amounts[col] >= 0, "PDF summary debit and credit totals must be positive or zero.");
            seen.add(col);
          }
        }
      }
      requireReport(seen.size, "An activity code in the PDF summary has no readable totals.");
      return amounts;
    };

    const entry = (row, description, amount, category) => ({
      source: "statement", name: "PDF page " + row.page, row: row.number, date, description, raw: amount, amount, category,
    });
    const result = [entry(balanceLine, "Closing balance", balance, "Closing balance")];
    const sums = [0, 0];
    for (const [i, section] of activity.entries()) {
      const end = activity[i + 1]?.index ?? activityEnd;
      const amounts = columnAmounts(current.slice(section.index, end));
      for (const [col, amount] of amounts.entries()) {
        sums[col] += amount;
        if (!amount) continue;
        const signed = col === 0 ? -amount : amount;
        result.push(entry(section.row, section.description, signed, dailyCategory(section.description, signed, "statement")));
      }
    }
    const own = recap ? columnAmounts(/^own activity recap$/i.test(recap.text)
      ? current.slice(current.indexOf(recap), current.indexOf(recap) + 2) : [recap]) : [totalDebits, totalCredits];
    requireReport(sums.every((amount, i) => Number.isSafeInteger(amount) && amount === own[i]),
      "The activity codes do not add up to the PDF summary totals.");
    for (const [col, amount] of [totalDebits - own[0], totalCredits - own[1]].entries()) {
      requireReport(amount >= 0, "Own activity exceeds the statement total.");
      if (amount) result.push(entry(recap, "Additional account activity", col === 0 ? -amount : amount, "Other · review"));
    }
    return result;
  } catch (error) {
    throw new Error(file.name + " · " + error.message);
  }
}

async function readDailySource(file, dates, source, sign) {
  const choices = [];
  const addTable = (name, cells, date1904) => {
    const header = dailyHeader(cells);
    if (header) choices.push({ name, cells, header, date1904 });
  };
  if (/\.csv$/i.test(file.name)) {
    requireReport(file.size <= 20 * 1024 * 1024, "The CSV exceeds the 20 MB import limit.");
    addTable("CSV", readCsvCells(await file.text()), false);
  } else {
    const book = await openExcel(file, true);
    for (const name of book.sheets.keys()) {
      const cells = await readCells(book, name);
      const layout = ["ledger", "settlement"].includes(source) && book.sheets.size === 1 ? wisdomHistoryLayout(cells) : undefined;
      if (!layout) {
        addTable(name, cells, book.date1904);
        continue;
      }
      const history = wisdomHistory(cells, book.date1904, layout);
      const rows = history.rows.filter(row => dates.has(row.date)).map(row => {
        const finance = source !== "ledger" ? "" : financeProvider(row.description)
          || (layout.separateDates ? fedLoanProvider(row.description) : "");
        const ambiguousFinance = source === "ledger" && !finance && /\bgu\s+univ|\bun(?:iv|vi)ersity\s+accou/i.test(row.description);
        // The Fed matchup shows provider debit and credit from the provider's side of the transfer.
        const amount = row.amount * sign * (finance && layout.separateDates ? -1 : 1);
        const category = source === "settlement" ? settlementCategory(row.description)
          : finance ? "Finance ACH" : dailyCategory(row.description, amount, source);
        return { ...row, source, name, raw: row.amount, amount, category, finance, ambiguousFinance };
      });
      if (source === "ledger") for (const date of dates) {
        if (date < layout.from || date > layout.through) continue;
        const closing = history.balanceOn(date);
        rows.push({ source, name, row: Number(closing.address.slice(1)), date, description: "Closing balance",
          amount: closing.amount, raw: closing.amount, category: "Closing balance", separateDates: layout.separateDates });
      }
      return rows;
    }
  }
  requireReport(choices.length === 1, file.name + " must contain one table with the expected date, description and amount headings.");
  const { name, cells, date1904, header: [first, fields] } = choices[0];
  const cols = new Set(fields.values()), rows = new Set();
  for (const [address, value] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (match && cols.has(match[1]) && Number(match[2]) > first && String(value ?? "").trim()) rows.add(Number(match[2]));
  }
  const result = [];
  for (const row of [...rows].sort((a, b) => a - b)) {
    const get = key => cells.get(fields.get(key) + row);
    try {
      const rowDate = sourceDate(get("date"), date1904);
      if (!dates.has(rowDate)) continue;
      const description = String(get("category") ?? "").trim();
      requireReport(description, "The description is blank.");
      const closing = /^(closing balance|ending balance)$/i.test(description);
      const finance = source === "ledger" && !closing ? String(get("finance") ?? "").trim() || financeProvider(description) : "";
      const ambiguousFinance = source === "ledger" && !finance
        && /\bgu\s+univ|\bun(?:iv|vi)ersity\s+accou/i.test(description);
      const multiplier = closing ? 1 : sign;
      const optional = key => String(get(key) ?? "").trim() ? cents(get(key)) : 0;
      let amounts;
      if (fields.get("amount")) {
        amounts = [cents(get("amount"))];
      } else {
        requireReport(String(get("debit") ?? "").trim() || String(get("credit") ?? "").trim(), "Debit and credit are both blank.");
        const debit = Math.abs(optional("debit")), credit = Math.abs(optional("credit"));
        amounts = closing || (!debit && !credit) ? [credit - debit] : [-debit, credit].filter(Boolean);
      }
      for (const raw of amounts) {
        const amount = raw * multiplier;
        const category = finance ? "Finance ACH" : ["vizo", "settlement"].includes(source)
          ? settlementCategory(description) : dailyCategory(description, amount, source);
        requireReport(category !== "No activity" || amount === 0, "A No activity row must be zero.");
        result.push({ source, name, row, date: rowDate, description, raw, amount, category, finance, ambiguousFinance });
      }
    } catch (error) {
      throw new Error(file.name + " · " + name + " row " + row + ": " + error.message);
    }
  }
  return result;
}

function dailyClosing(rows, field) {
  const entered = fedField(field).value.trim();
  const closings = rows.filter(item => item.category === "Closing balance");
  if (entered) {
    const amount = cents(entered);
    requireReport(!closings.length || (closings.length === 1 && closings[0].amount === amount), "The entered closing balance differs from the source.");
    return amount;
  }
  requireReport(closings.length === 1, "Enter the closing balance or include one Closing Balance row in the source.");
  return closings[0].amount;
}

async function readDailyTemplate(file, date, sentReports = [], returnReports = []) {
  const book = await openExcel(file, true);
  const coverFormulas = new Map(), matchupFormulas = new Set(), opsFormulas = new Map();
  const [cover, matchup, ops] = await Promise.all([
    readCells(book, "Cover", coverFormulas), readCells(book, "MATCHUP", matchupFormulas), readCells(book, "OPS BALANCING", opsFormulas),
  ]);
  requireReport(/federal reserve statement/i.test(String(cover.get("B10"))) && /general ledger/i.test(String(cover.get("B13"))),
    "Check the statement and ledger labels on the prepared Cover sheet.");
  const formulaAt = (formulas, address) => String(formulas.get(address) ?? "").replace(/[\s$]/g, "").toUpperCase();
  const coverDifferenceRows = [...cover].flatMap(([address, value]) => {
    const row = /^C([1-9]\d*)$/.exec(address)?.[1];
    return row && /^adjusted book balance$/i.test(String(value ?? "").trim()) ? [Number(row)] : [];
  });
  requireReport(coverDifferenceRows.length === 1, "Check the Adjusted Book Balance label on the prepared Cover sheet.");
  const totalCell = /^\+?K10-(K[1-9]\d*)$/.exec(formulaAt(coverFormulas, "K" + coverDifferenceRows[0]))?.[1];
  const range = /^SUM\(K(\d+):K(\d+)\)$/.exec(formulaAt(coverFormulas, totalCell));
  const first = Number(range?.[1]), last = Number(range?.[2]);
  requireReport(first > 10 && first <= 13 && last >= 13 && last < Number(totalCell?.slice(1)),
    "Check the line-item total and balance difference formulas on the prepared Cover sheet.");
  const headerDate = matchup.get("D6");
  if (String(headerDate ?? "").trim()) requireReport(sourceDate(headerDate, book.date1904) === date,
    "Set MATCHUP!D6 to the selected daily reconciliation date.");
  const matches = new Map();
  for (const [address, value] of matchup) {
    const match = /^A([1-9]\d*)$/.exec(address);
    if (!match) continue;
    let rowDate;
    try { rowDate = sourceDate(value, book.date1904); } catch { continue; }
    if (rowDate !== date) continue;
    const row = Number(match[1]), title = String(matchup.get("B" + row) ?? "").trim().toLowerCase();
    if (!dailyCategories.some(category => category.toLowerCase() === title)) continue;
    requireReport(!matches.has(title), "Several MATCHUP rows use the selected date and category " + title + ".");
    matches.set(title, row);
  }
  for (const category of ["ach in", "ach out", "wire in", "wire out"]) {
    requireReport(matches.has(category), "The selected date has no MATCHUP row for " + category + ".");
  }
  const zero = (cells, address, formulas) => {
    const value = cells.get(address);
    if (String(value ?? "").trim()) return cents(value);
    requireReport(!formulas?.has(address), "Recalculate and save the missing formula result at " + address + ".");
    return 0;
  };
  const achRow = matches.get("ach out");
  const achItems = ["J", "K", "L", "M", "N"].map(column => zero(matchup, column + achRow));
  const financeCells = ["M", "N"].map(column => column + achRow).filter(address => !matchupFormulas.has(address));
  const opsValues = new Map([4, 5, 6, 7, 11, 12, 13, 14, 18, 19].map(row => [row, zero(ops, "G" + row, opsFormulas)]));
  requireReport(formulaAt(opsFormulas, "E22") === "-(G11-G12)-(G4-G5)-G18",
    "Check the reviewed OPS BALANCING!E22 formula.");
  const g = row => opsValues.get(row);
  const opsPaste = [];
  if (sentReports.length) {
    requireReport(formulaAt(opsFormulas, "G8") === "-(G6-G7)+(G4-G5)" && !matchupFormulas.has("K" + achRow),
      "Check the Quickpay formula in OPS BALANCING and the destination input in MATCHUP.");
    for (const report of sentReports) {
      const start = report.prior ? 6 : 4;
      for (const [index, amount] of [report.debit, report.credit].entries()) {
        const row = start + index, cell = "G" + row;
        requireReport(!opsFormulas.has(cell) || /^[+-]?\d+(?:\.\d+)?(?:[+-]\d+(?:\.\d+)?)*$/.test(formulaAt(opsFormulas, cell)),
          "Check the originated item input at OPS BALANCING!" + cell + ".");
        opsValues.set(row, amount);
        opsPaste.push(["OPS BALANCING", cell, report.name, amount / 100]);
      }
    }
    achItems[1] = -(g(6) - g(7)) + (g(4) - g(5));
    opsPaste.push(["MATCHUP", "K" + achRow, "Quickpay from originated items", achItems[1] / 100]);
  }
  if (returnReports.length) {
    requireReport(formulaAt(opsFormulas, "G15") === "-(G13-G14)+(G11-G12)" && !matchupFormulas.has("J" + achRow),
      "Check the Returns formula in OPS BALANCING and the destination input in MATCHUP.");
    for (const report of returnReports) {
      const start = report.prior ? 13 : 11;
      for (const [index, amount] of [report.debit, report.credit].entries()) {
        const row = start + index, cell = "G" + row;
        requireReport(!opsFormulas.has(cell) || /^[+-]?\d+(?:\.\d+)?(?:[+-]\d+(?:\.\d+)?)*$/.test(formulaAt(opsFormulas, cell)),
          "Check the returns input at OPS BALANCING!" + cell + ".");
        opsValues.set(row, amount);
        opsPaste.push(["OPS BALANCING", cell, report.name + " · verified total", amount / 100]);
      }
    }
    achItems[0] = -(g(13) - g(14)) + (g(11) - g(12));
    opsPaste.push(["MATCHUP", "J" + achRow, "Returns from reviewed Report 174 totals", achItems[0] / 100]);
  }
  const opsCoverAdjustment = -(g(11) - g(12)) - (g(4) - g(5)) - g(18);
  const lineItems = [], financeSlots = new Map();
  for (let row = first; row <= last; row++) {
    if (row === 13) continue;
    const address = "K" + row;
    const description = String(cover.get("E" + row) ?? cover.get("B" + row) ?? "").trim();
    const slot = /^(SC|LK|STUDENT CHOICE|LENDKEY|TRUSTAGE|BANKACOUNT) (DEBIT|CREDIT)S?$/i.exec(description);
    if (slot && !coverFormulas.has(address)) {
      const provider = { SC: "Student Choice", LK: "LendKey" }[slot[1].toUpperCase()] ?? financeProvider(slot[1]);
      const key = provider.toLowerCase() + "|" + slot[2].toLowerCase();
      financeSlots.set(key, financeSlots.has(key) ? undefined : { address, current: cover.get(address) });
    }
    if (typeof cover.get(address) !== "number" && !coverFormulas.has(address)) continue;
    const linkedOps = formulaAt(coverFormulas, address) === "'OPSBALANCING'!E22";
    const amount = linkedOps ? opsCoverAdjustment : zero(cover, address, coverFormulas);
    lineItems.push([address, description, amount, linkedOps ? "Recalculated from OPS BALANCING inputs" : "Prepared cover"]);
  }
  const balanceWritable = coverFormulas.has("K13") ? /^[+-]?\d+(?:\.\d+)?$/.test(formulaAt(coverFormulas, "K13"))
    : cover.get("K13") === undefined || typeof cover.get("K13") === "number";
  const auditTarget = completionTarget("Cover", cover, coverFormulas, "G131", "G134", "K131", "H131");
  return { book, matches, lineItems, achItems, financeCells, financeSlots, ops: opsValues, opsCoverAdjustment, balanceWritable, opsPaste, auditTarget,
    adjustments: lineItems.reduce((sum, [, , value]) => sum + value, 0) };
}

function fedDepositCandidates(items, wisdom) {
  const byDate = new Map();
  for (const item of wisdom) {
    if (!["Branch deposits", "Mobile deposits"].includes(item.category)) continue;
    const totals = byDate.get(item.date) ?? { branch: 0, mobile: 0, branchRows: [], mobileRows: [] };
    const key = item.category === "Branch deposits" ? "branch" : "mobile";
    totals[key] += item.amount;
    totals[key + "Rows"].push(item.row);
    byDate.set(item.date, totals);
  }
  return items.flatMap(item => [...byDate].flatMap(([ledgerDate, { branch, mobile, branchRows, mobileRows }]) => {
    if (ledgerDate > item.date || ledgerDate < dateOffset(item.date, -4)) return [];
    const candidates = [];
    if (branch === item.amount) candidates.push([branch, 0, branchRows]);
    if (mobile === item.amount) candidates.push([0, mobile, mobileRows]);
    if (branch > 0 && mobile > 0 && branch + mobile === item.amount) {
      candidates.push([branch, mobile, [...branchRows, ...mobileRows]]);
    }
    return candidates.map(([branchAmount, mobileAmount, rows]) => ({ item, ledgerDate, branchAmount, mobileAmount, rows }));
  }));
}

async function reconcileVizo(date, template, sign, reviewedDeposits) {
  const formulas = new Map(), sheet = "VIZO MATCHUP";
  const cells = await readCells(template.book, sheet, formulas);
  const sections = [...cells].flatMap(([address, value]) => {
    const row = /^D([1-9]\d*)$/.exec(address)?.[1];
    if (!row || typeof value !== "string" || String(cells.get("A" + row)).trim().toLowerCase() !== "date") return [];
    return [{ row: Number(row), heading: value.trim() }];
  }).sort((a, b) => a.row - b.row);
  const references = [];
  const dateAt = address => {
    const value = cells.get(address);
    return value == null || value === "" ? "" : sourceDate(value, template.book.date1904);
  };
  for (const [category, heading, offset] of settlementCategories) {
    const matches = sections.filter(section => heading.test(section.heading));
    if (!matches.length && category === "Share draft suspense") continue;
    requireReport(matches.length === 1, "The VIZO MATCHUP tab needs one section for " + category + ".");
    const start = matches[0].row + 1, end = sections.find(section => section.row >= start)?.row ?? start + 40;
    const selected = new Set();
    for (let i = start; i < end; i++) {
      const vizoRow = i + offset;
      const wisdomDate = dateAt("E" + i), vizoDate = vizoRow < end ? dateAt("A" + vizoRow) : "";
      if (wisdomDate !== date && vizoDate !== date) continue;
      const stored = formulas.get("H" + i);
      const base = typeof stored === "string" ? i : Number(stored?.sharedFrom?.slice(1));
      const formula = typeof stored === "string" ? stored : stored?.expression;
      requireReport(String(formula ?? "").replace(/\$/g, "").toUpperCase() === "G" + base + "-C" + (base + offset),
        "Review the comparison formula at " + sheet + "!H" + i + ".");
      references.push({ category, wisdomDate, vizoDate, wisdomCell: "G" + i, vizoCell: "C" + vizoRow, comparison: "H" + i });
      for (const [side, current] of [["Wisdom", wisdomDate], ["VIZO", vizoDate]]) {
        if (current !== date) continue;
        requireReport(!selected.has(side), "The selected date appears twice in " + category + " for " + side + ".");
        selected.add(side);
      }
    }
    if (!selected.size) references.push({ category, wisdomDate: "", vizoDate: "", wisdomCell: "", vizoCell: "", comparison: "" });
  }
  const dates = new Set([date, ...references.flatMap(item => [item.wisdomDate, item.vizoDate]).filter(Boolean)]);
  const wisdomDates = new Set([...dates, ...[-1, -2, -3, -4].map(offset => dateOffset(date, offset))]);
  const [vizo, wisdom] = await Promise.all([
    readDailySource(dailyFiles.vizo, dates, "vizo", 1),
    readDailySource(dailyFiles.settlement, wisdomDates, "settlement", sign),
  ]);
  const unclassified = vizo.filter(item => item.date === date && item.amount > 0 && item.category === "Unclassified forward collection");
  const allocations = readReviewedDepositAllocations(reviewedDeposits, unclassified,
    new Set(wisdom.filter(item => item.date <= date).map(item => item.date)));
  const allocated = new Set(allocations.map(entry => entry.item));
  const depositCandidates = fedDepositCandidates(unclassified.filter(item => !allocated.has(item)), wisdom);
  const groups = new Map();
  for (const item of [...vizo, ...wisdom]) {
    const key = item.source + item.date, rows = groups.get(key) ?? [];
    rows.push(item);
    groups.set(key, rows);
  }
  if (!references.some(item => item.category === "Share draft suspense") &&
    [...vizo, ...wisdom].some(item => item.date === date && item.category === "Share draft suspense")) {
    references.push({ category: "Share draft suspense", wisdomDate: "", vizoDate: "", wisdomCell: "", vizoCell: "", comparison: "" });
  }
  for (const rows of groups.values()) {
    requireReport(!rows.some(item => item.category === "No activity") || rows.every(item => item.amount === 0),
      "A No activity row conflicts with activity on " + rows[0].date + ".");
  }
  const paste = [];
  const input = (source, targetDate, category, cell, pairedWisdomDate) => {
    if (!targetDate) return { note: "Date missing from template" };
    const rows = groups.get(source + targetDate);
    let reviewedAmount = 0;
    if (source === "vizo" && /^(Branch|Mobile) deposits$/.test(category)) {
      if (rows?.some(item => item.category === "Unclassified forward collection" && !allocated.has(item))) {
        return { note: "Classify forward collection" };
      }
      const component = category === "Branch deposits" ? "branchAmount" : "mobileAmount";
      const deposits = allocations.filter(entry => entry.item.date === targetDate && entry[component]);
      const conflicts = [...new Set(deposits.filter(entry => entry.ledgerDate !== pairedWisdomDate).map(entry => entry.ledgerDate))];
      if (conflicts.length) return { note: "Reviewed Wisdom date " + conflicts.join(", ") + " differs from template" };
      reviewedAmount = deposits.reduce((sum, entry) => sum + entry[component], 0);
    } else if (rows && /^(Branch|Mobile) deposits$/.test(category) && rows.some(item => item.category === "Unclassified forward collection")) {
      return { note: "Classify forward collection" };
    }
    if (rows && ["Share drafts", "Share draft suspense"].includes(category)
      && rows.some(item => item.category === "Share draft discrepancy · review")) {
      return { note: "Classify share draft discrepancy" };
    }
    if (rows) {
      const amount = rows.filter(item => item.category === category).reduce((sum, item) => sum + item.amount, reviewedAmount);
      requireReport(Number.isSafeInteger(amount), "A VIZO settlement total exceeds the supported range.");
      if (targetDate === date && !formulas.has(cell)) paste.push([sheet, cell, category + " · " + (source === "vizo" ? "VIZO" : "Wisdom"), amount / 100]);
      return { amount, note: targetDate === date && formulas.has(cell) ? "Template formula · review" : "Export" };
    }
    if (targetDate < date && typeof cells.get(cell) === "number") return { amount: cents(cells.get(cell)), note: "Earlier template value" };
    return { note: "Date missing from export" };
  };
  const comparisons = references.map(item => {
    const provider = input("vizo", item.vizoDate, item.category, item.vizoCell, item.wisdomDate);
    const ledger = input("settlement", item.wisdomDate, item.category, item.wisdomCell);
    const difference = provider.amount == null || ledger.amount == null ? undefined : ledger.amount - provider.amount;
    requireReport(difference == null || Number.isSafeInteger(difference), "A VIZO difference exceeds the supported range.");
    const review = item.category === "Returned deposits" && (provider.amount || ledger.amount) ? "Verify member posting with Ops"
      : [provider, ledger].some(side => side.note === "Template formula · review") ? "Review template formula"
        : difference == null ? "Pending source or template review" : difference ? "Difference to review" : "Amounts agree";
    return { ...item, provider, ledger, difference, review };
  });
  return { comparisons, paste, allocations, depositCandidates, sources: [...vizo, ...wisdom],
    vizoFile: dailyFiles.vizo.name, wisdomFile: dailyFiles.settlement.name };
}

function compareDaily(statement, ledger, template, finance, effectiveDateHistory) {
  const sum = (rows, label) => rows.filter(row => row.category === label).reduce((total, row) => total + row.amount, 0);
  const comparisons = ["Wire in", "Wire out", "FedNow in", "FedNow out"].map(category => {
    const a = Math.abs(sum(statement, category)), b = Math.abs(sum(ledger, category));
    return { category, statement: a, ledger: b, difference: category === "Wire out" ? b - a : a - b };
  });
  const achIn = sum(statement, "ACH in"), achOut = sum(statement, "ACH out");
  const timing = finance.debits - finance.credits;
  const net = sum(ledger, "ACH settlement") + (effectiveDateHistory ? timing : -timing);
  const achDifference = achIn + achOut - net;
  const [returns, dao, other, reviewedDebits, reviewedCredits] = template?.achItems ?? [0, 0, 0, 0, 0];
  const debits = finance.included.length ? finance.debits : reviewedDebits;
  const credits = finance.included.length ? finance.credits : reviewedCredits;
  const adjustedAch = achDifference + returns + dao + other + debits - credits;
  requireReport([achIn, achOut, net, achDifference, adjustedAch, ...comparisons.flatMap(row =>
    [row.statement, row.ledger, row.difference])].every(Number.isSafeInteger), "A daily total exceeds the supported range.");
  return { comparisons, achIn, achOut, net, achDifference, adjustedAch };
}

function fedFinanceTiming(rows, date, postingDate) {
  const covered = rows.some(item => item.date === postingDate);
  const review = rows.filter(item => item.finance || item.ambiguousFinance);
  const included = review.filter(item => item.date === postingDate && item.finance);
  const debits = -included.filter(item => item.amount < 0).reduce((total, item) => total + item.amount, 0);
  const credits = included.filter(item => item.amount > 0).reduce((total, item) => total + item.amount, 0);
  requireReport(Number.isSafeInteger(debits) && Number.isSafeInteger(credits), "Finance timing amounts exceed the supported range.");
  return { date, postingDate, covered, review, included, debits, credits };
}

function fedClearingCheck(template, selected, original, refreshed, statement, source) {
  const entries = [];
  const issues = [];
  for (const provider of selected) {
    let slots = 0;
    for (const direction of ["debit", "credit"]) {
      const slot = template.financeSlots.get(provider.toLowerCase() + "|" + direction);
      if (!slot) continue;
      slots++;
      let amount;
      try { amount = slot.current === undefined || slot.current === "" ? 0 : cents(slot.current); }
      catch { issues.push("Check the amount at Cover!" + slot.address + "."); continue; }
      if (!amount) continue;
      if ((direction === "debit") !== (amount < 0)) {
        issues.push(provider + " " + direction + " has the opposite sign at Cover!" + slot.address + ".");
        continue;
      }
      entries.push({ provider, direction, address: slot.address, amount });
    }
    if (slots < 2) issues.push("Check the prepared " + provider + " debit and credit fields.");
  }
  const cleared = entries.reduce((sum, item) => sum + item.amount, 0);
  const movement = refreshed - original;
  const remaining = template.adjustments - cleared;
  const difference = statement - (refreshed + remaining);
  requireReport([cleared, movement, remaining, difference].every(Number.isSafeInteger), "The refreshed cover check exceeds the supported range.");
  if (!entries.length) issues.push("No current provider timing items were found to clear.");
  if (!template.balanceWritable) issues.push("Check the Wisdom balance input at Cover!K13.");
  if (movement !== cleared) issues.push("The Wisdom balance changed by a different amount than the selected timing items.");
  if (difference) issues.push("A difference remains after the proposed clearing.");
  return { source, refreshed, movement, cleared, remaining, difference, entries, issues };
}

function dailySummary(result) {
  const cash = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
  const table = document.createElement("table"), body = table.createTBody();
  const details = [
    ["Cover difference", cash.format(result.difference / 100)],
    ["ACH difference before reviewed items", cash.format(result.achDifference / 100)],
    ["ACH difference after reviewed items", cash.format(result.adjustedAch / 100)],
    ["Wire and FedNow differences", result.comparisons.filter(row => row.difference !== 0).length],
    ["Uncategorized rows to review", result.exceptions.length],
    ["Finance timing date checked", result.finance.postingDate + (result.finance.covered ? "" : " · no export rows for this date")],
    ["Next-day finance rows included", result.finance.included.length],
  ];
  if (result.clearing) details.push(["Cover difference after proposed clearing", cash.format(result.clearing.difference / 100)],
    ["Provider clearing check", result.clearing.issues.join(" ") || "Balances agree · see Cover clearing review"]);
  if (result.vizo) details.push(["VIZO comparisons requiring review", result.vizo.comparisons.filter(item => item.review !== "Amounts agree").length]);
  for (const [name, value] of details) {
    const row = body.insertRow(), heading = document.createElement("th");
    heading.scope = "row";
    heading.textContent = name;
    row.append(heading);
    row.insertCell().textContent = value;
  }
  fedField("summary").replaceChildren(table);
  fedField("results").hidden = false;
}

function dailyPasteInputs(result) {
  const paste = [
    ["Cover", "K10", "Fed statement balance", result.statementBalance / 100],
    ...((!result.template || result.template.balanceWritable) && (!result.clearing || result.clearing.issues.length)
      ? [["Cover", "K13", "Wisdom ending balance", result.ledgerBalance / 100]] : []),
  ];
  const match = result.template?.matches;
  if (match) {
    paste.push(["MATCHUP", "C" + match.get("ach in"), "ACH In", result.achIn / 100],
      ["MATCHUP", "C" + match.get("ach out"), "ACH Out", result.achOut / 100],
      ["MATCHUP", "G" + match.get("ach out"), "ACH Settlement", result.net / 100]);
    for (const entry of result.comparisons) {
      const r = match.get(entry.category.toLowerCase());
      if (!r) continue;
      paste.push(["MATCHUP", "C" + r, entry.category + " statement", entry.statement / 100],
        ["MATCHUP", "G" + r, entry.category + " Wisdom", entry.ledger / 100]);
    }
    if (result.finance.included.length) {
      const r = match.get("ach out");
      if (result.template.financeCells.includes("M" + r)) paste.push(["MATCHUP", "M" + r, "Finance debits", result.finance.debits / 100]);
      if (result.template.financeCells.includes("N" + r)) paste.push(["MATCHUP", "N" + r, "Finance credits", result.finance.credits / 100]);
    }
  }
  if (result.vizo) paste.push(...result.vizo.paste);
  if (result.template?.opsPaste.length) paste.push(...result.template.opsPaste);
  return paste;
}

async function buildDaily(result) {
  const audit = reconciliationAudit();
  const { label, number, row, worksheet } = excelOutputCells();
  function table(name, headers, data, widths) {
    const cols = "ABCDEFGHIJKLMN".slice(0, headers.length);
    const body = row(1, headers.map((title, i) => label(cols[i] + "1", title, 1)).join(""))
      + data.map((entry, i) => row(i + 2, entry.map((value, j) => {
        const cell = cols[j] + (i + 2);
        return typeof value === "number" ? number(cell, value) : label(cell, value);
      }).join(""))).join("");
    return [name, worksheet("A1:" + cols.at(-1) + (data.length + 1),
      widths.map((width, i) => [i + 1, i + 1, width]), body)];
  }
  const sourceRow = entry => [entry.source, entry.name, entry.row, entry.date, entry.description,
    entry.raw / 100, entry.amount / 100, entry.category];
  const sourceHeaders = ["Source", "Worksheet", "Row", "Date", "Description", "Imported amount", "Comparison amount", "Category"];
  const sourceWidths = [14, 24, 10, 15, 44, 18, 20, 29];
  const sheets = [
    table("Daily summary", ["Field", "Value"], [
      ["Date", result.date], ["Statement source", result.statementFile], ["Wisdom source", result.ledgerFile],
      ...result.references.map(name => ["Supporting file for manual review", name]),
      ["Prepared workbook", result.templateFile], ["Fed closing balance", result.statementBalance / 100],
      ["Wisdom source closing balance", result.rawLedgerBalance / 100], ["Wisdom balance for cover", result.ledgerBalance / 100],
      ["Reviewed cover items", result.adjustments / 100],
      ["Adjusted book balance", result.adjusted / 100], ["Cover difference", result.difference / 100],
      ["ACH in", result.achIn / 100], ["ACH out", result.achOut / 100], ["Wisdom ACH settlement", result.net / 100],
      ["ACH difference before items", result.achDifference / 100], ["ACH difference after items", result.adjustedAch / 100],
      ["Finance timing date checked", result.finance.postingDate],
      ["Next-day finance debits", result.finance.debits / 100], ["Next-day finance credits", result.finance.credits / 100],
      ["Finance timing coverage", result.finance.covered ? "Posting date found in Wisdom export" : "Posting date missing from Wisdom export · verify"],
      ...(result.vizo ? [["VIZO settlement source", result.vizo.vizoFile], ["Wisdom settlement source", result.vizo.wisdomFile]] : []),
    ], [40, 60]),
    table("Wire and FedNow", ["Category", "Statement", "Wisdom", "Workbook difference"],
      result.comparisons.map(entry => [entry.category, entry.statement / 100, entry.ledger / 100, entry.difference / 100]), [18, 19, 19, 23]),
    table("Paste cells", ["Sheet", "Cell", "Input", "Value"], dailyPasteInputs(result), [18, 15, 31, 20]),
    ...completionPasteSheets(result.template?.auditTarget, audit),
    table("Activity review", sourceHeaders, result.activity.map(sourceRow), sourceWidths),
    table("Exceptions", sourceHeaders, result.exceptions.map(sourceRow), sourceWidths),
  ];
  const finance = result.finance;
  if (result.sentReports.length) sheets.push(table("FedACH sent-file review",
    ["Source", "Process date", "Line", "Status", "Debits", "Credits", "Counted"],
    result.sentReports.flatMap(report => report.rows.map(item => [
      report.name, report.date, item.line, item.status, item.debit / 100, item.credit / 100, item.count ? "Yes" : "Review",
    ])), [37, 16, 10, 24, 18, 18, 15]));
  if (result.returnReports.length) sheets.push(table("Report 174 review", ["Source PDF", "Run date", "Use", "Debit returns", "Credit returns"],
    result.returnReports.map(report => [report.name, report.date, report.prior ? "Prior business day" : "Current day",
      report.debit / 100, report.credit / 100]), [38, 17, 22, 18, 18]));
  sheets.push(table("Finance timing", ["Provider", "Wisdom effective or supplied date", "Wisdom posted on", "Fed Rec date", "Description", "Signed amount", "Wisdom sheet", "Source row", "Review"],
    finance.review.map(item => [item.finance, item.date, item.posted ?? "", item.date === finance.postingDate && item.finance ? finance.date : "",
      item.description, item.amount / 100, item.name, item.row, item.ambiguousFinance ? "Confirm provider or wire"
        : item.date === finance.postingDate ? "Included in ACH and Finance totals · review cover line item" : "Posted on statement date · review prior Fed Rec"]),
    [23, 25, 18, 18, 52, 18, 21, 14, 61]));
  if (finance.included.length) {
    const grouped = new Map();
    for (const item of finance.included) {
      const direction = item.amount < 0 ? "debit" : "credit", key = item.finance.toLowerCase() + "|" + direction;
      const group = grouped.get(key) ?? { provider: item.finance, direction, amount: 0 };
      group.amount += item.amount;
      grouped.set(key, group);
    }
    sheets.push(table("Finance cover review", ["Provider", "Direction", "Signed timing item", "Cover cell", "Current cover value", "Review"],
      [...grouped].map(([key, group]) => {
        const slot = result.template?.financeSlots.get(key);
        return [group.provider, group.direction, group.amount / 100, slot?.address ?? "", slot?.current ?? "",
          slot ? "Compare with the existing item before copying." : "Choose the appropriate cover line item."];
      }), [25, 15, 21, 14, 21, 49]));
  }
  if (result.clearing) {
    const { clearing } = result;
    sheets.push(table("Cover clearing review", ["Field", "Amount or review"], [
      ["Refreshed Wisdom source", clearing.source], ["Original Wisdom balance", result.ledgerBalance / 100],
      ["Refreshed Wisdom balance", clearing.refreshed / 100], ["Wisdom balance movement", clearing.movement / 100],
      ["Selected timing items", clearing.cleared / 100], ["Remaining cover line items", clearing.remaining / 100],
      ["Difference after proposed clearing", clearing.difference / 100],
      ...clearing.issues.map(issue => ["Review", issue]),
      ...(!clearing.issues.length ? [["Review", "Balances agree. Paste the updated balance and clear only the listed cells."]] : []),
    ], [41, 78]));
    if (clearing.entries.length) sheets.push(table("Cover timing items", ["Provider", "Direction", "Cover cell", "Current value", "Action"],
      clearing.entries.map(item => [item.provider, item.direction, item.address, item.amount / 100,
        clearing.issues.length ? "Investigate before clearing" : "Clear the cell"]), [23, 16, 17, 18, 42]));
    if (!clearing.issues.length) sheets.push(table("Refreshed cover paste cells", ["Sheet", "Cell", "Field", "Value or action"], [
      ["Cover", "K13", "Refreshed Wisdom closing balance", clearing.refreshed / 100],
      ...clearing.entries.map(item => ["Cover", item.address, item.provider + " " + item.direction, "Clear contents"]),
    ], [17, 12, 35, 23]));
  }
  if (result.template) {
    const { lineItems, ops, achItems, opsCoverAdjustment } = result.template;
    const g = row => ops.get(row);
    sheets.push(table("Cover line items", ["Source cell", "Description", "Amount", "Basis"],
      lineItems.map(([cell, description, amount, basis]) => [cell, description, amount / 100, basis]), [17, 52, 19, 38]));
    sheets.push(table("OPS comparison", ["Cell or source", "Calculation", "Amount"], [
      ["G8", "Current originated net less prior originated net", (-(g(6) - g(7)) + (g(4) - g(5))) / 100],
      ["G15", "Current returns net less prior returns net", (-(g(13) - g(14)) + (g(11) - g(12))) / 100],
      ["G20", "DAO ACH difference", (g(18) - g(19)) / 100],
      ["E22", "Returns, originated items and DAO ACH", opsCoverAdjustment / 100],
      ...achItems.map((amount, i) => ["MATCHUP " + "JKLMN"[i] + result.template.matches.get("ach out"), "Reviewed ACH item", amount / 100]),
    ], [23, 52, 19]));
  }
  if (result.vizo) {
    const value = amount => amount == null ? "" : amount / 100;
    sheets.push(table("VIZO settlement review", ["Category", "VIZO date", "VIZO cell", "VIZO amount", "VIZO source",
      "Wisdom date", "Wisdom cell", "Wisdom amount", "Wisdom source", "Comparison cell", "Wisdom less VIZO", "Review"],
      result.vizo.comparisons.map(item => [
        item.category, item.vizoDate, item.vizoDate ? item.vizoCell : "", value(item.provider.amount), item.provider.note,
        item.wisdomDate, item.wisdomDate ? item.wisdomCell : "", value(item.ledger.amount), item.ledger.note,
        item.comparison, value(item.difference), item.review,
      ]), [21, 15, 14, 16, 30, 15, 15, 17, 30, 18, 21, 36]));
    sheets.push(table("VIZO source review", sourceHeaders, result.vizo.sources.map(sourceRow), sourceWidths));
    const allocationColumns = ["Provider row", "Wisdom date", "Branch amount", "Mobile amount"];
    const allocationValues = entry => [entry.item.row, entry.ledgerDate, entry.branchAmount / 100, entry.mobileAmount / 100];
    if (result.vizo.depositCandidates.length) sheets.push(table("Deposit allocation candidates",
      [...allocationColumns, "Wisdom rows"],
      result.vizo.depositCandidates.map(entry => [...allocationValues(entry), entry.rows.join(", ")]), [16, 17, 17, 17, 28]));
    if (result.vizo.allocations.length) sheets.push(table("Reviewed deposit allocations",
      [...allocationColumns, "Provider total"],
      result.vizo.allocations.map(entry => [...allocationValues(entry), entry.item.amount / 100]), [16, 17, 17, 17, 17]));
    const issues = result.vizo.comparisons.filter(item => item.review !== "Amounts agree" &&
      ["Share drafts", "Share draft suspense", "Returned deposits"].includes(item.category));
    sheets.push(table("Clearing review", ["Category", "VIZO date", "Wisdom date", "VIZO amount", "Wisdom amount",
      "Difference", "Template comparison", "Review"],
    issues.map(item => [item.category, item.vizoDate, item.wisdomDate, value(item.provider.amount), value(item.ledger.amount),
      value(item.difference), item.comparison, item.review]), [24, 16, 16, 18, 18, 17, 23, 39]));
  }
  return createExcelWorkbook(sheets, audit);
}

for (const kind of ["statement", "ledger", "template", "vizo", "settlement", "refresh-ledger", "sent-current", "sent-prior", "returns-current", "returns-prior"]) {
  const input = fedField(kind + "-file"), drop = fedField(kind + "-drop");
  input.addEventListener("change", () => selectDaily(kind, input.files[0]));
  drop.addEventListener("dragover", event => { event.preventDefault(); drop.classList.add("active"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("active"));
  drop.addEventListener("drop", event => {
    event.preventDefault();
    drop.classList.remove("active");
    if (event.dataTransfer.files.length !== 1) return dailyStatus("Drop one workbook at a time.", true);
    input.value = "";
    selectDaily(kind, event.dataTransfer.files[0]);
  });
}
document.addEventListener("dragover", event => event.preventDefault());
document.addEventListener("drop", event => event.preventDefault());
for (const input of document.querySelectorAll("input:not([type=file]):not(#confirm), select, textarea")) input.addEventListener("input", invalidateDaily);
fedField("date").addEventListener("input", () => {
  const date = fedField("date").value;
  fedField("finance-posting-date").value = date ? nextFedPostingDate(date) : "";
  fedField("fed-deposit-allocations").value = "";
});
fedField("confirm").addEventListener("change", () => { if (!fedField("confirm").checked) invalidateDaily(); });

fedField("run").addEventListener("click", async () => {
  const current = ++dailyRevision;
  dailyResult = undefined;
  fedField("results").hidden = true;
  try {
    const date = fedField("date").value, sign = Number(fedField("sign").value);
    requireReport(date && [1, -1].includes(sign) && fedField("confirm").checked, "Verify the date, signs and source inputs.");
    const postingDate = fedField("finance-posting-date").value || nextFedPostingDate(date);
    requireReport(postingDate > date, "Choose the next Finance effective date after the statement date.");
    requireReport(dailyFiles.statement && dailyFiles.ledger, "Load the statement and Wisdom G/L History.");
    const wantsVizo = dailyFiles.vizo || dailyFiles.settlement;
    requireReport(!wantsVizo || (dailyFiles.vizo && dailyFiles.settlement && dailyFiles.template),
      "Load both settlement exports and the prepared daily workbook for VIZO.");
    const sentRoles = ["sent-current", "sent-prior"].filter(role => dailyFiles[role]);
    const returnReports = reviewedReturns(date);
    requireReport(!(sentRoles.length || returnReports.length) || dailyFiles.template, "Load the prepared daily workbook to use reviewed Ops report inputs.");
    const includeApproved = fedField("sent-approved").checked;
    const sentReports = await Promise.all(sentRoles.map(role => readSentFiles(dailyFiles[role], date, role === "sent-prior", includeApproved)));
    if (current !== dailyRevision) return;
    const dates = new Set([date]);
    const [statement, ledgerRows, template] = await Promise.all([
      /\.(?:pdf|txt)$/i.test(dailyFiles.statement.name) ? readFedStatement(dailyFiles.statement, date)
        : readDailySource(dailyFiles.statement, dates, "statement", 1),
      readDailySource(dailyFiles.ledger, new Set([date, postingDate]), "ledger", sign),
      dailyFiles.template ? readDailyTemplate(dailyFiles.template, date, sentReports, returnReports) : undefined,
    ]);
    if (current !== dailyRevision) return;
    const ledger = ledgerRows.filter(item => item.date === date);
    const finance = fedFinanceTiming(ledgerRows, date, postingDate);
    requireReport(statement.length && ledger.length, "Check that both exports contain rows for the selected date.");
    const statementBalance = dailyClosing(statement, "statement-balance");
    const rawLedgerBalance = dailyClosing(ledger, "ledger-balance");
    const effectiveDateHistory = ledger.some(item => item.separateDates);
    const ledgerBalance = rawLedgerBalance + (effectiveDateHistory ? finance.debits - finance.credits : 0);
    const adjustments = template?.adjustments ?? cents(fedField("adjustment").value);
    const adjusted = ledgerBalance + adjustments, difference = statementBalance - adjusted;
    requireReport([adjustments, adjusted, difference].every(Number.isSafeInteger), "A cover amount exceeds the supported range.");
    const activity = [...statement, ...ledger];
    const exceptions = activity.filter(item => !dailyCategories.includes(item.category) && !["Closing balance", "Finance ACH"].includes(item.category)
      || (item.source === "statement" && item.category === "ACH settlement")
      || (item.source === "ledger" && ["ACH in", "ACH out"].includes(item.category)));
    const vizo = wantsVizo ? await reconcileVizo(date, template, Number(fedField("settlement-sign").value),
      fedField("fed-deposit-allocations").value) : undefined;
    if (current !== dailyRevision) return;
    const selected = [["posted-lendkey", "LendKey"], ["posted-student", "Student Choice"]]
      .filter(([id]) => fedField(id).checked).map(([, provider]) => provider);
    const refreshFile = dailyFiles["refresh-ledger"];
    let clearing;
    if (selected.length || refreshFile || fedField("refreshed-balance").value.trim()) {
      requireReport(template && selected.length, "Load the prepared daily workbook and confirm which provider postings have cleared.");
      const refreshedRows = refreshFile ? await readDailySource(refreshFile, dates, "ledger", sign) : [];
      if (current !== dailyRevision) return;
      const refreshed = dailyClosing(refreshedRows, "refreshed-balance");
      clearing = fedClearingCheck(template, selected, ledgerBalance, refreshed, statementBalance,
        refreshFile ? refreshFile.name : "Entered from refreshed Wisdom balance");
    }
    const references = [...dailyReferences, ...["returns-current", "returns-prior"].map(role => dailyFiles[role]?.name)
      .filter(name => name && !returnReports.some(report => report.name === name))];
    const templateSource = dailyFiles.template;
    const result = { date, statementBalance, rawLedgerBalance, ledgerBalance, adjustments, adjusted, difference, template, activity, exceptions, vizo, finance, clearing, sentReports, returnReports, references, templateSource,
      statementFile: dailyFiles.statement.name, ledgerFile: dailyFiles.ledger.name, templateFile: templateSource?.name ?? "",
      ...compareDaily(statement, ledger, template, finance, effectiveDateHistory) };
    const note = fedField("template-download-note");
    const skipped = [];
    if (template) {
      const paste = dailyPasteInputs(result);
      const formulas = new Map(await Promise.all([...new Set(paste.map(([sheet]) => sheet))].map(async sheet => {
        const cells = new Set();
        await readCells(template.book, sheet, cells);
        return [sheet, cells];
      })));
      result.templateUpdates = [];
      for (const [sheet, address, , value] of paste) {
        if (formulas.get(sheet).has(address)) skipped.push(sheet + "!" + address);
        else result.templateUpdates.push([sheet, address, value]);
      }
    }
    if (current !== dailyRevision) return;
    const canFill = Boolean(template && (result.templateUpdates.length || template.auditTarget));
    fedField("download-template").hidden = !canFill;
    note.hidden = !canFill;
    note.textContent = [
      "The filled copy includes inputs and completion details in your prepared workbook.",
      skipped.length ? "Use the review Excel to check proposed values for formula cells " + skipped.join(", ") + "." : "",
      clearing ? "Use the review Excel for any proposed provider clearing changes." : "",
    ].filter(Boolean).join(" ");
    dailyResult = result;
    dailySummary(dailyResult);
    dailyStatus("");
  } catch (error) {
    if (current === dailyRevision) dailyStatus(error.message, true);
  }
});

function fillDailyTemplate(result) {
  const audit = reconciliationAudit();
  const updates = [...result.templateUpdates];
  if (result.template.auditTarget) {
    const { sheet, initials, date } = result.template.auditTarget;
    updates.push([sheet, initials, audit.by], [sheet, date, { date: audit.on }]);
  }
  return fillExcelTemplate(result.templateSource, updates);
}

async function downloadDaily(button, build, name) {
  const result = dailyResult;
  if (!result) return;
  button.disabled = true;
  try {
    const blob = await build(result);
    if (dailyResult !== result) return;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = name(result);
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 30000);
  } catch (error) {
    if (dailyResult === result) dailyStatus(error.message, true);
  } finally {
    button.disabled = false;
  }
}

fedField("download").addEventListener("click", event =>
  downloadDaily(event.currentTarget, buildDaily, result => "Federal Reserve " + result.date + ".xlsx"));
fedField("download-template").addEventListener("click", event =>
  downloadDaily(event.currentTarget, fillDailyTemplate,
    result => "Federal Reserve " + result.date + " filled template." + result.templateSource.name.split(".").at(-1)));
