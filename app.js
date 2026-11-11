"use strict";

const monthInput = document.querySelector("#month");
const reportInput = document.querySelector("#report");
const drop = document.querySelector("#drop");
const status = document.querySelector("#status");
const preview = document.querySelector("#preview");
const period = document.querySelector("#period");
const tableContainer = document.querySelector("#table");
const download = document.querySelector("#download");
const downloadTemplate = document.querySelector("#download-template");
const templateDownloadNote = document.querySelector("#template-download-note");
const wisdomInput = document.querySelector("#wisdom-file");
const wisdomDrop = document.querySelector("#wisdom-drop");
const wisdomMapping = document.querySelector("#wisdom-mapping");
const wisdomSheet = document.querySelector("#wisdom-sheet");
const wisdomAccount = document.querySelector("#wisdom-account");
const wisdomBalance = document.querySelector("#wisdom-balance");
const wisdomFirst = document.querySelector("#wisdom-first");
const wisdomLast = document.querySelector("#wisdom-last");
const wisdomConfirm = document.querySelector("#wisdom-confirm");
const wisdomStatus = document.querySelector("#wisdom-status");
const wisdomClear = document.querySelector("#wisdom-clear");
const detailInput = document.querySelector("#detail-file");
const detailStatus = document.querySelector("#detail-status");
const detailClear = document.querySelector("#detail-clear");
const profileInput = document.querySelector("#profile");
const profileDrop = document.querySelector("#profile-drop");
const profileStatus = document.querySelector("#profile-status");
const money = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });

let file;
let result;
let revision = 0;
let ledgerFiles = [];
let ledgerWorkbooks = [];
let ledgerResult;
let ledgerRevision = 0;
let detailFile;
let detailTarget;
let detailRevision = 0;
let downloading = false;
let loanStreetProfile;
let profileRevision = 0;

function refreshInputs() {
  wisdomConfirm.checked = false;
  if (ledgerWorkbooks.length) validateWisdom();
  validateDetail();
  validate();
}

async function selectProfile(file) {
  const current = ++profileRevision;
  loanStreetProfile = undefined;
  refreshInputs();
  profileStatus.classList.remove("error");
  if (!file) {
    profileStatus.textContent = "Choose a completed workbook or local profile.";
    return;
  }
  try {
    const profile = await readLoanStreetProfile(file);
    if (current !== profileRevision) return;
    loanStreetProfile = profile;
    profileStatus.textContent = /\.xlsx$/i.test(file.name) ? "Workbook mappings loaded for this session." : "Profile loaded for this session.";
    refreshInputs();
  } catch (error) {
    if (current !== profileRevision) return;
    profileStatus.classList.add("error");
    profileStatus.textContent = error.message;
  }
}
profileInput.addEventListener("change", () => selectProfile(profileInput.files[0]));
profileDrop.addEventListener("dragover", () => profileDrop.classList.add("active"));
profileDrop.addEventListener("dragleave", () => profileDrop.classList.remove("active"));
profileDrop.addEventListener("drop", event => {
  profileDrop.classList.remove("active");
  const files = event.dataTransfer.files;
  if (files.length !== 1) {
    profileStatus.textContent = "Drop one completed workbook or local profile.";
    profileStatus.classList.add("error");
    return;
  }
  profileInput.value = "";
  selectProfile(files[0]);
});

function syncDownload() {
  download.disabled = downloading || !result || Boolean(ledgerFiles.length && !ledgerResult) || Boolean(detailFile && !detailTarget);
  const canFill = Boolean(detailTarget && !detailTarget.coverIssue);
  downloadTemplate.hidden = !canFill;
  downloadTemplate.disabled = download.disabled || !canFill;
  templateDownloadNote.hidden = !canFill;
  templateDownloadNote.textContent = ledgerResult
    ? "The filled copy updates the consolidated covers and details. Copy KEY INFO from the review Excel."
    : "The filled copy is marked as a draft. Check the saved ledger amounts and use the review Excel for KEY INFO.";
}

async function validateDetail() {
  const current = ++detailRevision;
  detailTarget = undefined;
  detailClear.hidden = !detailFile;
  detailStatus.classList.remove("error");
  syncDownload();
  if (!detailFile) {
    detailStatus.textContent = "";
    return;
  }
  if (!monthInput.value) {
    detailStatus.textContent = "Choose the reconciliation month to check the detail sheets.";
    return;
  }
  if (!loanStreetProfile) {
    detailStatus.textContent = "Load a completed LoanStreet workbook or local profile first.";
    return;
  }
  detailStatus.textContent = "Checking the prepared month blocks…";
  try {
    const target = await readDetailDestinations(detailFile, monthInput.value, loanStreetProfile);
    if (current !== detailRevision) return;
    detailTarget = target;
    detailStatus.textContent = "Found 40 paste cells across " + target.addresses.size + " detail sheets."
      + (target.monthOnly ? " Verify the year for " + target.monthOnly + " month-only headings." : "")
      + " Found preparer fields on " + target.coverAudit.filter(([, fields]) => fields).length + " covers."
      + (target.coverIssue ? " Filled copy unavailable. " + target.coverIssue : "");
    syncDownload();
  } catch (error) {
    if (current !== detailRevision) return;
    detailStatus.textContent = error.message;
    detailStatus.classList.add("error");
  }
}

function showWisdomStatus(message, error = false) {
  wisdomStatus.textContent = message;
  wisdomStatus.classList.toggle("error", error);
}

async function validateWisdom() {
  const current = ++ledgerRevision;
  ledgerResult = undefined;
  syncDownload();
  if (!ledgerWorkbooks.length) return;
  if (!loanStreetProfile) {
    showWisdomStatus("Load a completed LoanStreet workbook or local profile first.");
    return;
  }
  if (!monthInput.value || !wisdomConfirm.checked) {
    showWisdomStatus("Select the columns and confirm the closing period and balances.");
    return;
  }
  showWisdomStatus("Checking the Wisdom accounts…");
  try {
    const sheet = wisdomSheet.value;
    const { accounts } = loanStreetProfile;
    const matches = new Map(), combined = ledgerWorkbooks.length > 1;
    for (const { workbook, file } of ledgerWorkbooks) {
      const worksheet = sheet || [...workbook.sheets.keys()][0];
      let entries;
      try {
        entries = await readWisdom(workbook, worksheet, wisdomAccount.value.trim().toUpperCase(), wisdomBalance.value.trim().toUpperCase(),
          wisdomFirst.valueAsNumber, accounts, wisdomLast.value ? wisdomLast.valueAsNumber : 1048576, !combined, monthInput.value);
      } catch (error) {
        throw new Error(file.name + ": " + error.message);
      }
      if (current !== ledgerRevision) return;
      for (const [account, entry] of entries) {
        requireReport(!matches.has(account), "Wisdom account " + account + " appears in more than one selected workbook.");
        matches.set(account, { ...entry, source: combined ? file.name + " · " + worksheet + "!" + entry.address : undefined });
      }
    }
    const missing = accounts.filter(account => !matches.has(account));
    requireReport(!missing.length, "The selected workbooks are missing accounts " + missing.join(", ") + ".");
    ledgerResult = { fileName: combined ? ledgerWorkbooks.length + " Wisdom workbooks" : ledgerWorkbooks[0].file.name,
      sheet: sheet || "Individual worksheets", matches };
    showWisdomStatus("Found all " + matches.size + " required G/L balances.");
    syncDownload();
  } catch (error) {
    if (current === ledgerRevision) showWisdomStatus(error.message, true);
  }
}

async function selectWisdom(files) {
  const current = ++ledgerRevision;
  const selected = Array.from(files);
  ledgerFiles = selected;
  ledgerWorkbooks = [];
  ledgerResult = undefined;
  wisdomConfirm.checked = false;
  wisdomMapping.hidden = true;
  wisdomClear.hidden = !selected.length;
  syncDownload();
  if (!selected.length) {
    showWisdomStatus("");
    return;
  }
  showWisdomStatus("Opening " + (selected.length === 1 ? selected[0].name : selected.length + " Wisdom workbooks") + "…");
  try {
    const books = [];
    for (const file of selected) {
      try { books.push({ file, workbook: await openExcel(file, true) }); }
      catch (error) { throw new Error(file.name + ": " + error.message); }
      if (current !== ledgerRevision) return;
    }
    if (current !== ledgerRevision) return;
    const common = [...books[0].workbook.sheets.keys()].filter(name => books.every(({ workbook }) => workbook.sheets.has(name)));
    const singleSheets = books.length > 1 && books.every(({ workbook }) => workbook.sheets.size === 1);
    requireReport(common.length || singleSheets, "Use workbooks with a common worksheet name or one worksheet per file.");
    wisdomSheet.replaceChildren(...common.map(name => new Option(name, name)),
      ...(singleSheets && !common.length ? [new Option("Each workbook's only worksheet", "")] : []));
    ledgerWorkbooks = books;
    wisdomMapping.hidden = false;
    validateWisdom();
  } catch (error) {
    if (current === ledgerRevision) showWisdomStatus(error.message, true);
  }
}

function showStatus(message, error = false) {
  status.textContent = message;
  status.classList.toggle("error", error);
}

function showTable(data) {
  const table = document.createElement("table");
  const heading = table.createTHead().insertRow();
  for (const label of ["Provider figures", ...data.profile.pools]) {
    const cell = document.createElement("th");
    cell.scope = "col";
    cell.textContent = label;
    heading.append(cell);
  }
  const body = table.createTBody();
  for (const [index, [name]] of metrics.entries()) {
    const row = body.insertRow();
    const label = document.createElement("th");
    label.scope = "row";
    label.textContent = name;
    row.append(label);
    for (const amount of data.values[index]) row.insertCell().textContent = money.format(amount);
  }
  const payable = document.createElement("p");
  payable.textContent = "Ending accounts payable (full portfolio) · " + money.format(data.payable);
  tableContainer.replaceChildren(table, payable);
}

async function validate() {
  const current = ++revision;
  result = undefined;
  syncDownload();
  preview.hidden = true;
  if (!file) {
    showStatus("");
    return;
  }
  if (!loanStreetProfile) {
    showStatus("Load a completed LoanStreet workbook or local profile to read the report.");
    return;
  }
  if (!monthInput.value) {
    showStatus("Selected " + file.name + ". Choose the reconciliation month to continue.");
    return;
  }
  showStatus("Reading the report…");
  try {
    const parsed = await readLoanStreet(file, monthInput.value, loanStreetProfile);
    if (current !== revision) return;
    showTable(parsed);
    period.textContent = new Intl.DateTimeFormat("en-US", { month: "long", year: "numeric", timeZone: "UTC" })
      .format(new Date(parsed.month + "-01T00:00:00Z"));
    result = parsed;
    syncDownload();
    preview.hidden = false;
    showStatus("Loaded " + parsed.fileName + ".");
  } catch (error) {
    if (current === revision) showStatus(error.message, true);
  }
}

reportInput.addEventListener("change", () => {
  file = reportInput.files[0];
  validate();
});
monthInput.addEventListener("change", refreshInputs);
document.addEventListener("dragover", event => event.preventDefault());
document.addEventListener("drop", event => event.preventDefault());
drop.addEventListener("dragover", () => drop.classList.add("active"));
drop.addEventListener("dragleave", () => drop.classList.remove("active"));
drop.addEventListener("drop", event => {
  drop.classList.remove("active");
  const files = event.dataTransfer.files;
  if (!files.length) return;
  reportInput.value = "";
  file = files.length === 1 ? files[0] : undefined;
  validate();
  if (!file) showStatus("Drop one Buy Side Combined Report at a time.", true);
});
wisdomInput.addEventListener("change", () => selectWisdom(wisdomInput.files));
wisdomClear.addEventListener("click", () => {
  wisdomInput.value = "";
  selectWisdom([]);
});
wisdomDrop.addEventListener("dragover", () => wisdomDrop.classList.add("active"));
wisdomDrop.addEventListener("dragleave", () => wisdomDrop.classList.remove("active"));
wisdomDrop.addEventListener("drop", event => {
  wisdomDrop.classList.remove("active");
  const files = event.dataTransfer.files;
  if (!files.length) return;
  wisdomInput.value = "";
  selectWisdom(files);
});
for (const field of [wisdomSheet, wisdomAccount, wisdomBalance, wisdomFirst, wisdomLast]) {
  field.addEventListener("input", () => {
    wisdomConfirm.checked = false;
    validateWisdom();
  });
}
wisdomConfirm.addEventListener("change", validateWisdom);
addTrialBalanceDetection("wisdom", () => ledgerWorkbooks.map(({ workbook, file }) => ({ book: workbook, file })));
detailInput.addEventListener("change", () => {
  detailFile = detailInput.files[0];
  validateDetail();
});
detailClear.addEventListener("click", () => {
  detailInput.value = "";
  detailFile = undefined;
  validateDetail();
});
function fillLoanStreetTemplate(file, result, wisdom, target) {
  const audit = reconciliationAudit();
  const updates = result.profile.details.flatMap(([sheet], index) =>
    target.addresses.get(sheet).map((address, metric) => [sheet, address, result.values[detailMetrics[metric][1]][index]]));
  for (const { sheet, account, providerCell, ledgerCell, amount } of coverBalances(result)) {
    updates.push([sheet, providerCell, amount]);
    const balance = wisdom?.matches.get(account);
    if (balance) updates.push([sheet, ledgerCell, balance.amount]);
  }
  for (const [sheet, address, , title] of target.coverPeriods) updates.push([sheet, address, title]);
  for (const [sheet, fields] of target.coverAudit) {
    if (fields) updates.push([sheet, fields.initials, audit.by], [sheet, fields.date, { date: audit.on }]);
  }
  return fillExcelTemplate(file, updates);
}

async function downloadLoanStreet(button, fill) {
  const selected = result;
  if (!selected || button.disabled) return;
  const selectedLedger = ledgerResult;
  const selectedLedgerRevision = ledgerRevision;
  const selectedDetail = detailTarget;
  const template = detailFile;
  const selectedDetailRevision = detailRevision;
  downloading = true;
  syncDownload();
  try {
    const workbook = fill ? await fillLoanStreetTemplate(template, selected, selectedLedger, selectedDetail)
      : await buildKeyInfo(selected, selectedLedger, selectedDetail);
    if (result !== selected || ledgerRevision !== selectedLedgerRevision || detailRevision !== selectedDetailRevision) return;
    const url = URL.createObjectURL(workbook);
    const link = document.createElement("a");
    link.href = url;
    link.download = fill ? "LoanStreet " + selected.month + " filled" + (selectedLedger ? "" : " draft") + ".xlsx"
      : "LoanStreet inputs " + selected.month + ".xlsx";
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  } catch (error) {
    if (result === selected) showStatus("Could not create the Excel download. " + error.message, true);
  } finally {
    downloading = false;
    syncDownload();
  }
}

download.addEventListener("click", event => downloadLoanStreet(event.currentTarget, false));
downloadTemplate.addEventListener("click", event => downloadLoanStreet(event.currentTarget, true));
