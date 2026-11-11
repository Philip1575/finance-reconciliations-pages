"use strict";

function coverBalances(result) {
  const { covers, payable } = result.profile;
  const balances = covers.flatMap(([sheet, principal, interest], index) => [
    { sheet, metric: "Principal", destination: principal, amount: result.values[7][index] },
    { sheet, metric: "Accrued interest", destination: interest, amount: result.values[2][index] },
  ]);
  balances.push({
    sheet: payable[0],
    metric: "Ending accounts payable",
    destination: payable.slice(1),
    amount: result.payable,
  });
  return balances.map(({ sheet, metric, destination, amount }) => {
    const [account, providerCell, ledgerCell] = destination;
    return { sheet, metric, account, providerCell, ledgerCell, amount };
  });
}

function normalizeAccount(value) {
  const code = String(value ?? "").trim();
  if (/^\d{6}$/.test(code)) return code;
  const parts = /^(\d{3})(?:\.(\d{1,3})|-(\d{3}))$/.exec(code);
  if (!parts) return "";
  return parts[1] + (parts[2] || parts[3]).padEnd(3, "0");
}

function readBalance(value) {
  if (typeof value === "number") return value;
  const text = String(value ?? "").trim();
  const accounting = /^\(.*\)$/.test(text);
  const digits = (accounting ? text.slice(1, -1) : text).replace(/^(-?)\$/, "$1");
  const pattern = accounting ? /^(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/ : /^-?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/;
  if (!pattern.test(digits)) return NaN;
  return Number(digits.replace(/,/g, "")) * (accounting ? -1 : 1);
}

async function detectTrialBalanceColumns(workbook, sheet) {
  const cells = await readCells(workbook, sheet);
  if (/^consolidated trial balance report$/i.test(String(cells.get("B1") ?? "").trim())
    && normalizeAccount(cells.get("A2")) && Number.isFinite(readBalance(cells.get("H2")))
    && [...cells].some(([address, value]) => /^D\d+$/.test(address) && value === "Ending Balance"
      && cells.get("A" + address.slice(1)) === "Totals:")) {
    return { account: "A", balance: "H", first: 2 };
  }
  const headers = new Map();
  for (const [address, value] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (!match || Number(match[2]) > 30 || typeof value !== "string") continue;
    const label = value.trim().toLowerCase().replace(/\s+/g, " ");
    const field = ["g/l account", "gl account", "account", "account number"].includes(label) ? "account"
      : ["ending balance", "closing balance"].includes(label) ? "balance" : "";
    if (!field) continue;
    const row = Number(match[2]), columns = headers.get(row) ?? {};
    columns[field] = Object.hasOwn(columns, field) ? "" : match[1];
    headers.set(row, columns);
  }
  const candidates = [...headers].filter(([, columns]) => columns.account && columns.balance);
  if (candidates.length !== 1) return undefined;
  const [[row, columns]] = candidates;
  return { ...columns, first: row + 1 };
}

function addTrialBalanceDetection(prefix, sources, ids = ["account", "balance", "first", "sheet"].map(name => prefix + "-" + name)) {
  const [accountField, balanceField, firstField, sheetField] = ids.map(id => document.getElementById(id));
  const selectedSheet = () => sheetField.tagName === "SELECT" ? sheetField.value : sheetField.value.trim();
  const control = document.createElement("div");
  const button = document.createElement("button"), status = document.createElement("span");
  control.className = "column-detection";
  button.type = "button";
  button.textContent = "Detect trial balance columns";
  status.setAttribute("role", "status");
  control.append(button, status);
  accountField.closest("label").parentElement.append(control);
  button.addEventListener("click", async () => {
    const files = sources().map(({ book, file }) => ({ book, file }));
    const sheetName = selectedSheet();
    const fields = [accountField, balanceField, firstField];
    const current = fields.map(field => field.value);
    button.disabled = true;
    status.textContent = "Checking worksheet headings…";
    try {
      requireReport(files.length, "Load the trial balance workbook first.");
      const detected = [];
      for (const source of files) {
        const book = source.book ?? await openExcel(source.file, true);
        requireReport(sheetName || book.sheets.size === 1, "Choose the trial balance worksheet first.");
        detected.push(await detectTrialBalanceColumns(book, sheetName || [...book.sheets.keys()][0]));
      }
      const latest = sources();
      if (selectedSheet() !== sheetName || latest.length !== files.length
        || files.some((source, index) => source.book !== latest[index].book || source.file !== latest[index].file)
        || fields.some((field, index) => field.value !== current[index])) {
        status.textContent = "The selection changed. Detect the columns again if needed.";
        return;
      }
      const [mapping] = detected;
      requireReport(mapping && detected.every(item => item?.account === mapping.account && item.balance === mapping.balance && item.first === mapping.first),
        "A shared account and ending balance layout could not be identified. Enter the columns and first row manually.");
      accountField.value = mapping.account;
      balanceField.value = mapping.balance;
      firstField.value = mapping.first;
      accountField.dispatchEvent(new Event("input", { bubbles: true }));
      status.textContent = "Columns suggested. Verify the worksheet and data range.";
    } catch (error) {
      status.textContent = error.message;
    } finally {
      button.disabled = false;
    }
  });
  return control;
}

function validateWisdomMonth(cells, month) {
  if (!month || !/^consolidated trial balance report$/i.test(String(cells.get("B1") ?? "").trim())) return;
  const through = /\bto:\s*(\d{1,2})\/(\d{1,2})\/(\d{4})\s*$/i.exec(String(cells.get("E1") ?? ""));
  const [year, period] = month.split("-").map(Number);
  const lastDay = new Date(Date.UTC(year, period, 0)).getUTCDate();
  requireReport(through && Number(through[1]) === period && Number(through[2]) === lastDay && Number(through[3]) === year,
    "The Wisdom report must end on the last day of the selected month.");
}

async function readWisdom(workbook, sheet, accountColumn, balanceColumn, firstRow, accounts, lastRow = 1048576, requireAll = true, month) {
  requireReport([accountColumn, balanceColumn].every(column => /^[A-Z]{1,3}$/.test(column) && columnNumber(column) <= 16384)
    && accountColumn !== balanceColumn, "Enter two different Excel column letters for the account and ending balance.");
  requireReport(Number.isInteger(firstRow) && firstRow >= 1 && firstRow <= 1048576, "Enter the first row of ledger data.");
  requireReport(Number.isInteger(lastRow) && lastRow >= firstRow && lastRow <= 1048576, "Check the last row of ledger data.");
  const cells = await readCells(workbook, sheet);
  validateWisdomMonth(cells, month);
  const expected = new Set(accounts);
  const matches = new Map();
  const accountAddress = new RegExp("^" + accountColumn + "([1-9]\\d*)$");
  for (const [address, value] of cells) {
    const row = Number(accountAddress.exec(address)?.[1]);
    if (!(row >= firstRow && row <= lastRow)) continue;
    const account = normalizeAccount(value);
    if (!expected.has(account)) continue;
    requireReport(!matches.has(account), "Wisdom lists account " + account + " more than once. Use a trial balance with one ending balance per account.");
    const balanceAddress = balanceColumn + row;
    const amount = readBalance(cells.get(balanceAddress));
    requireReport(Number.isFinite(amount), "Account " + account + " has a missing or invalid balance at " + sheet + "!" + balanceAddress + ".");
    matches.set(account, { amount, address: balanceAddress });
  }
  const missing = accounts.filter(account => !matches.has(account));
  requireReport(!requireAll || !missing.length, "The selected columns are missing accounts " + missing.join(", ") + ".");
  return matches;
}

function coverWorksheets(result, wisdom, cells) {
  const { label, number, row, worksheet } = cells;
  const balances = coverBalances(result);
  const index = [
    row(1, label("A1", "LoanStreet cover inputs · " + result.month, 1)),
    row(2, label("A2", wisdom ? "Wisdom file confirmed by preparer" : "Provider figures only. Wisdom amounts pending.")),
    row(3, label("A3", result.fileName)),
    ...(wisdom ? [row(4, label("A4", wisdom.fileName + " · " + wisdom.sheet))] : []),
    row(6, ["Destination sheet", "Balance", "G/L account", "Provider cell", "Provider amount", "Ledger cell", "Wisdom amount", "Wisdom source"]
      .map((heading, index) => label("ABCDEFGH"[index] + "6", heading, 1)).join("")),
  ];
  const copies = new Map();
  for (const [i, entry] of balances.entries()) {
    const { sheet, metric, account, providerCell, ledgerCell, amount } = entry;
    const posted = wisdom?.matches.get(account);
    const line = i + 7;
    index.push(row(line, label("A" + line, sheet) + label("B" + line, metric) + label("C" + line, account)
      + label("D" + line, providerCell) + number("E" + line, amount) + label("F" + line, ledgerCell)
      + (posted ? number("G" + line, posted.amount) + label("H" + line, posted.source ?? posted.address) : "")));
    if (!copies.has(sheet)) copies.set(sheet, []);
    copies.get(sheet).push([providerCell, amount]);
    if (posted) copies.get(sheet).push([ledgerCell, posted.amount]);
  }
  return [
    ["Cover inputs", worksheet("A1:H23", [[1, 2, 29], [3, 7, 18], [8, 8, 45]], index.join(""))],
    ...Array.from(copies, ([sheet, amounts]) => {
      amounts.sort(([first], [second]) => Number(first.slice(1)) - Number(second.slice(1)));
      const rows = [
        row(1, label("A1", "Paste only the listed input cells into the matching cover sheet.", 1)),
        ...amounts.map(([address, value]) => row(Number(address.slice(1)), number(address, value))),
      ];
      return [sheet, worksheet("A1:J50", [[1, 1, 44], [2, 10, 17]], rows.join(""))];
    }),
  ];
}
