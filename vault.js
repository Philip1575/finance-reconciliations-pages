"use strict";

const byId = id => document.getElementById(id);
const vaultSources = {
  eod: { input: byId("eod-file"), workbook: undefined, file: undefined, revision: 0 },
  ledger: { input: byId("ledger-file"), workbook: undefined, file: undefined, revision: 0 },
  cover: { input: byId("cover-file"), workbook: undefined, file: undefined, revision: 0 },
};
let vaultResult;
let vaultRevision = 0;
let vaultDownloading = false;
const vaultTrialDetection = addTrialBalanceDetection("ledger", () =>
  vaultSources.ledger.workbook ? [{ book: vaultSources.ledger.workbook }] : []);

async function readVaultLog(workbook, sheet, dateColumn, balanceColumn, firstRow, month) {
  requireReport([dateColumn, balanceColumn].every(column => /^[A-Z]{1,3}$/.test(column) && columnNumber(column) <= 16384)
    && dateColumn !== balanceColumn, "Select different date and closing-balance columns.");
  requireReport(Number.isInteger(firstRow) && firstRow >= 1 && firstRow <= 1048576, "Enter the first data row.");
  const cells = await readCells(workbook, sheet);
  const pattern = new RegExp("^" + dateColumn + "([1-9]\\d*)$");
  let latest = "";
  let matches = [];
  for (const [address, value] of cells) {
    const row = Number(pattern.exec(address)?.[1]);
    if (!(row >= firstRow) || !String(value).trim()) continue;
    let date;
    try {
      date = sourceDate(value, workbook.date1904);
    } catch (error) {
      throw new Error(sheet + "!" + address + ": " + error.message);
    }
    if (date.slice(0, 7) !== month || date < latest) continue;
    if (date > latest) {
      latest = date;
      matches = [];
    }
    matches.push({ date, dateAddress: address, address: balanceColumn + row });
  }
  requireReport(matches.length, "The Daily EOD log has no dated entries for " + month + ".");
  requireReport(matches.length === 1, "The final EOD date appears more than once. Select a log with one final vault balance per day.");
  const result = matches[0];
  try {
    result.cents = cents(cells.get(result.address));
  } catch (error) {
    throw new Error(sheet + "!" + result.address + ": " + error.message);
  }
  return result;
}

async function readVaultEod(month, manual) {
  if (manual) {
    const date = byId("eod-manual-date").value;
    requireReport(date.slice(0, 7) === month, "Enter the verified EOD closing date in the selected month.");
    const balance = byId("eod-manual-balance").value.trim();
    requireReport(balance, "Enter the verified EOD closing total.");
    return { date, cents: cents(balance), source: byId("eod-manual-source").value.trim(), manual: true };
  }
  const { workbook, file } = vaultSources.eod;
  const sheet = byId("eod-sheet").value;
  const eod = await readVaultLog(workbook, sheet, byId("eod-date").value.trim().toUpperCase(),
    byId("eod-balance").value.trim().toUpperCase(), byId("eod-first").valueAsNumber, month);
  return { ...eod, file: file.name, sheet };
}

async function selectVaultFile(kind, file) {
  const state = vaultSources[kind];
  const current = ++state.revision;
  state.workbook = undefined;
  state.file = file;
  byId("confirm").checked = false;
  const mapping = byId(kind + "-mapping");
  if (mapping) mapping.hidden = true;
  if (kind === "cover") {
    for (const id of ["vault-account", "adjustment", "adjustment-description"]) byId(id).disabled = Boolean(file);
  }
  byId(kind + "-status").textContent = file ? "Opening " + file.name + "…" : "";
  validateVault();
  if (!file) return;
  try {
    const workbook = await openExcel(file, kind !== "eod");
    if (current !== state.revision) return;
    state.workbook = workbook;
    if (mapping) {
      byId(kind + "-sheet").replaceChildren(...Array.from(workbook.sheets.keys(), name => new Option(name, name)));
      mapping.hidden = kind === "eod" && byId("eod-mode").value === "manual";
    }
    byId(kind + "-status").textContent = file.name;
    validateVault();
  } catch (error) {
    if (current === state.revision) byId(kind + "-status").textContent = error.message;
  }
}

async function readVaultCover(month) {
  const { workbook, file } = vaultSources.cover;
  if (!file) return {
    account: normalizeAccount(byId("vault-account").value), adjustment: cents(byId("adjustment").value),
    description: byId("adjustment-description").value.trim(),
  };
  requireReport(workbook, "Load a readable prepared MV workbook.");
  const formulas = new Map(), cells = await readCells(workbook, "MV", formulas);
  const text = address => String(cells.get(address) ?? "").trim();
  const formula = address => String(formulas.get(address) ?? "").replace(/[\s$]/g, "").toUpperCase();
  const period = new Intl.DateTimeFormat("en-US", { month: "long", timeZone: "UTC" }).format(new Date(month + "-01T00:00:00Z"));
  const title = text("B13").toLowerCase();
  const expected = "for the month of " + period.toLowerCase();
  requireReport(title === expected || title === expected + " " + month.slice(0, 4), "Set MV!B13 to the selected month in the prepared workbook.");
  requireReport(/^main vault g\/l:?$/i.test(text("B15")) && /daily eod/i.test(text("B20")) && /wisdom/i.test(text("B22"))
    && !["B13", "K20", "K22", "E28", "K28"].some(address => formulas.has(address))
    && formula("K30") === "K22+K28" && formula("K32") === "K20-K30", "The prepared MV worksheet differs from the supported cover.");
  return { account: normalizeAccount(cells.get("D15")), adjustment: text("K28") ? cents(cells.get("K28")) : 0,
    description: text("E28"), file: file.name, template: file,
    completion: completionTarget("MV", cells, formulas, "G42", "G47", "K42", "K43") };
}

async function readVaultWisdom(account, month) {
  const book = vaultSources.ledger.workbook, sheet = byId("ledger-sheet").value;
  if (byId("ledger-mode").value === "trial") {
    const ledger = await readWisdom(book, sheet, byId("ledger-account").value.trim().toUpperCase(),
      byId("ledger-balance").value.trim().toUpperCase(), byId("ledger-first").valueAsNumber, [account], undefined, true, month);
    const match = ledger.get(account);
    return { address: match.address, cents: cents(match.amount) };
  }
  const address = byId("ledger-cell").value.trim().toUpperCase();
  const match = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(address);
  requireReport(match && columnNumber(match[1]) <= 16384 && Number(match[2]) <= 1048576, "Enter the verified Wisdom ending-balance cell.");
  const cells = await readCells(book, sheet);
  validateWisdomMonth(cells, month);
  return { address, cents: cents(cells.get(address)) };
}

function displayVault(result) {
  const table = document.createElement("table");
  const body = table.createTBody();
  const money = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
  for (const [name, value] of [
    ["Daily EOD · " + result.eod.date, result.eod.cents],
    ["Wisdom · " + result.account, result.ledgerCents],
    ["In-transit adjustment", result.adjustment],
    ["Adjusted Wisdom balance", result.adjusted],
    ["Difference", result.difference],
  ]) {
    const line = body.insertRow();
    const heading = document.createElement("th");
    heading.scope = "row";
    heading.textContent = name;
    line.append(heading);
    line.insertCell().textContent = money.format(value / 100);
  }
  byId("table").replaceChildren(table);
}

async function validateVault() {
  const current = ++vaultRevision;
  vaultResult = undefined;
  byId("download").disabled = true;
  byId("download-template").disabled = true;
  byId("preview").hidden = true;
  byId("status").textContent = "";
  byId("status").classList.remove("error");
  const month = byId("month").value;
  const manual = byId("eod-mode").value === "manual";
  if (!month || (!manual && !vaultSources.eod.workbook) || !vaultSources.ledger.workbook) return;
  if (!byId("confirm").checked) {
    byId("status").textContent = "Verify the column selections and confirm the closing inputs.";
    return;
  }
  try {
    const cover = await readVaultCover(month);
    if (current !== vaultRevision) return;
    const { account, adjustment, description } = cover;
    requireReport(account, "Enter the main vault G/L account from your reconciliation workbook.");
    const [eod, ledger] = await Promise.all([
      readVaultEod(month, manual),
      readVaultWisdom(account, month),
    ]);
    const ledgerCents = ledger.cents;
    requireReport(!adjustment || description, "Describe the nonzero in-transit adjustment.");
    const adjusted = ledgerCents + adjustment;
    const difference = eod.cents - adjusted;
    requireReport(Number.isSafeInteger(adjusted) && Number.isSafeInteger(difference), "The balances exceed the supported amount range.");
    if (current !== vaultRevision) return;
    vaultResult = { month, account, eod, ledgerCents, adjustment, description, adjusted, difference,
      ledgerFile: vaultSources.ledger.file.name, ledgerSheet: byId("ledger-sheet").value, ledgerAddress: ledger.address, coverFile: cover.file,
      completion: cover.completion, template: cover.template };
    displayVault(vaultResult);
    byId("preview").hidden = false;
    byId("download").disabled = vaultDownloading;
    byId("download-template").hidden = !cover.template;
    byId("download-template").disabled = vaultDownloading;
    byId("status").textContent = difference === 0 ? "Difference is $0.00." : "A difference remains. Review the inputs and reconciling items.";
  } catch (error) {
    if (current !== vaultRevision) return;
    byId("status").textContent = error.message;
    byId("status").classList.add("error");
  }
}

async function buildVault(result) {
  const audit = reconciliationAudit();
  const { label, number, date, row, worksheet } = excelOutputCells();
  const formula = (address, expression, value) =>
    `<c r="${address}" s="2"><f>${expression}</f><v>${value / 100}</v></c>`;
  const month = new Intl.DateTimeFormat("en-US", { month: "long", timeZone: "UTC" }).format(new Date(result.month + "-01T00:00:00Z"));
  const body = [
    row(12, label("B12", "Teller Vault and Cash Monthly Reconciliation", 1)),
    row(13, label("B13", "For the Month of " + month)),
    row(15, label("B15", "Main Vault G/L") + label("D15", result.account.slice(0, 3) + "-" + result.account.slice(3))),
    row(20, label("B20", "BALANCE PER DAILY EOD TOTAL") + number("K20", result.eod.cents / 100)),
    row(22, label("B22", "BALANCE PER WISDOM G/L") + number("K22", result.ledgerCents / 100)),
    row(26, label("B26", "IN TRANSIT ITEMS")),
    row(28, label("E28", result.description) + number("K28", result.adjustment / 100)),
    row(30, label("C30", "ADJUSTED WISDOM BALANCE") + formula("K30", "K22+K28", result.adjusted)),
    row(32, label("C32", "Difference") + formula("K32", "K20-K30", result.difference)),
    row(42, label("G42", "Prepared by") + label("K42", audit.by)),
    row(43, label("G43", "Completed on") + date("K43", audit.on)),
  ].join("");
  const sources = [
    ["Reconciliation month", result.month],
    ...(result.eod.manual ? [["Daily EOD source", "Entered manually"], ["Source note", result.eod.source]]
      : [["Daily EOD file", result.eod.file], ["Daily EOD worksheet", result.eod.sheet],
        ["Date cell", result.eod.dateAddress], ["Closing balance cell", result.eod.address]]),
    ["Closing date", result.eod.date],
    ["Wisdom file", result.ledgerFile], ["Wisdom worksheet", result.ledgerSheet], ["Balance cell", result.ledgerAddress],
    ["G/L account", result.account], ["In-transit description", result.description],
    ...(result.coverFile ? [["Prepared MV workbook", result.coverFile]] : []),
    ["Copy as values", "MV!B13, K20, K22, E28, K28"],
  ];
  const sourceRows = sources.map(([name, value], index) => row(index + 1, label("A" + (index + 1), name) + label("B" + (index + 1), value))).join("");
  return createExcelWorkbook([
    ["MV", worksheet("B12:K43", [[1, 1, 3], [2, 2, 37], [3, 10, 10], [11, 11, 20]], body)],
    ["Sources", worksheet("A1:B" + sources.length, [[1, 1, 26], [2, 2, 70]], sourceRows)],
    ...completionPasteSheets(result.completion, audit),
  ], audit);
}

for (const [kind, state] of Object.entries(vaultSources)) {
  state.input.addEventListener("change", () => selectVaultFile(kind, state.input.files[0]));
  const drop = byId(kind + "-drop");
  drop.addEventListener("dragover", event => {
    event.preventDefault();
    drop.classList.add("active");
  });
  drop.addEventListener("dragleave", () => drop.classList.remove("active"));
  drop.addEventListener("drop", event => {
    event.preventDefault();
    drop.classList.remove("active");
    if (event.dataTransfer.files.length !== 1) {
      byId(kind + "-status").textContent = "Drop one workbook at a time.";
      return;
    }
    state.input.value = "";
    selectVaultFile(kind, event.dataTransfer.files[0]);
  });
}
document.addEventListener("dragover", event => event.preventDefault());
document.addEventListener("drop", event => event.preventDefault());
byId("eod-mode").addEventListener("input", () => {
  const manual = byId("eod-mode").value === "manual";
  byId("eod-drop").hidden = manual;
  byId("eod-status").hidden = manual;
  byId("eod-mapping").hidden = manual || !vaultSources.eod.workbook;
  byId("eod-manual").hidden = !manual;
});
byId("ledger-mode").addEventListener("input", () => {
  const cell = byId("ledger-mode").value === "cell";
  vaultTrialDetection.hidden = cell;
  byId("ledger-cell-field").hidden = !cell;
  for (const id of ["ledger-account", "ledger-balance", "ledger-first"]) byId(id).disabled = cell;
});
for (const input of document.querySelectorAll("input:not([type=file]), select")) {
  input.addEventListener("input", () => {
    if (input.id !== "confirm") byId("confirm").checked = false;
    validateVault();
  });
}
function fillVaultTemplate(result) {
  const audit = reconciliationAudit();
  const updates = [["MV", "K20", result.eod.cents / 100], ["MV", "K22", result.ledgerCents / 100]];
  if (result.completion) {
    const { sheet, initials, date } = result.completion;
    updates.push([sheet, initials, audit.by], [sheet, date, { date: audit.on }]);
  }
  return fillExcelTemplate(result.template, updates);
}

async function downloadVault(fill) {
  if (!vaultResult || vaultDownloading) return;
  const selected = vaultResult;
  vaultDownloading = true;
  byId("download").disabled = true;
  byId("download-template").disabled = true;
  try {
    const blob = fill ? await fillVaultTemplate(selected) : await buildVault(selected);
    if (selected !== vaultResult) return;
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "Teller Vault " + selected.month + (fill ? " filled." + selected.template.name.split(".").at(-1) : ".xlsx");
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  } catch (error) {
    if (selected === vaultResult) byId("status").textContent = error.message;
  } finally {
    vaultDownloading = false;
    byId("download").disabled = !vaultResult;
    byId("download-template").disabled = !vaultResult;
  }
}

byId("download").addEventListener("click", () => downloadVault(false));
byId("download-template").addEventListener("click", () => downloadVault(true));
