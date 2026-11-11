"use strict";

const detailMetrics = [
  ["Principal payment", 0],
  ["Net interest payment", 1],
  ["Ending accrued but unpaid interest", 2],
  ["Change in accounts payable", 3],
  ["Charge-offs", 4],
];
const monthNames = ["january", "february", "march", "april", "may", "june",
  "july", "august", "september", "october", "november", "december"];

function monthHeading(value) {
  const text = String(value ?? "").trim().toLowerCase();
  const match = /^([a-z]+)(?:\s+(\d{4})|\s+(\d{1,2})\/\d{1,2}(?:\/(\d{2}|\d{4}))?)?$/.exec(text);
  if (!match) return undefined;
  const index = monthNames.findIndex(month => month === match[1] || month.slice(0, 3) === match[1]);
  if (index < 0) return undefined;
  const dateYear = Number(match[3]) === index + 1 ? match[4] : "";
  const year = match[2] || (dateYear?.length === 2 ? "20" + dateYear : dateYear);
  return { month: monthNames[index], year };
}

function findLoanStreetCoverAudit(cells, formulas) {
  const prepared = [...cells].filter(([, value]) => /^\s*prepared by\s*:?\s*$/i.test(String(value ?? "")));
  if (prepared.length !== 1) return undefined;
  const [, labelColumn, rawRow] = /^([BF])([1-9]\d*)$/.exec(prepared[0][0]) ?? [];
  if (!labelColumn) return undefined;
  const row = Number(rawRow);
  if (!/^\s*verified by\s*:?\s*$/i.test(String(cells.get(labelColumn + (row + 4)) ?? ""))) return undefined;
  const inputColumn = labelColumn === "B" ? "F" : "J";
  const initials = inputColumn + row, date = inputColumn + (row + 1);
  const currentInitials = cells.get(initials), currentDate = cells.get(date);
  if (formulas.has(initials) || formulas.has(date)
    || (currentInitials !== undefined && typeof currentInitials !== "string")
    || (currentDate !== undefined && typeof currentDate !== "number")) return undefined;
  return { initials, date };
}

async function readDetailDestinations(file, month, profile) {
  const workbook = await openExcel(file);
  const expectedMonth = monthNames[Number(month.slice(5)) - 1];
  requireReport(expectedMonth, "Choose the reconciliation month first.");
  const addresses = new Map();
  const existingValues = new Map();
  let monthOnly = 0;
  for (const [sheet, principalCol, interestCol, apCol, chargeCol] of profile.details) {
    const formulas = new Set();
    const cells = await readCells(workbook, sheet, formulas);
    function lastBlock(column) {
      const addressPattern = new RegExp("^" + column + "([1-9]\\d*)$");
      let lastRow = 0;
      let lastHeading;
      for (const [address, value] of cells) {
        const row = Number(addressPattern.exec(address)?.[1]);
        if (!(row > lastRow)) continue;
        const heading = monthHeading(value);
        if (heading) {
          lastRow = row;
          lastHeading = heading;
        }
      }
      requireReport(lastHeading?.month === expectedMonth,
        sheet + " column " + column + ": the last month heading must be " + expectedMonth + ".");
      requireReport(!lastHeading.year || lastHeading.year === month.slice(0, 4),
        sheet + " column " + column + ": the last month heading has a different year.");
      if (!lastHeading.year) monthOnly++;
      return lastRow;
    }
    function checkLabel(column, row, pattern) {
      const address = column + row;
      requireReport(pattern.test(String(cells.get(address) ?? "").trim()),
        sheet + "!" + address + " has an unexpected label. Check the prepared month block.");
    }
    function checkFormula(address) {
      requireReport(formulas.has(address), sheet + "!" + address + " is missing an expected roll-forward formula.");
    }
    function pasteAddress(column, row) {
      const address = nextDetailColumn(column) + row;
      requireReport(!formulas.has(address), sheet + "!" + address + " is a formula. Check the prepared month block.");
      const currentValue = cells.get(address);
      requireReport(currentValue === undefined || (typeof currentValue === "number" && Number.isFinite(currentValue)),
        sheet + "!" + address + " is neither a number nor an empty input cell.");
      return address;
    }
    const principal = lastBlock(principalCol);
    const interest = lastBlock(interestCol);
    const ap = lastBlock(apCol);
    const charge = ap + (chargeCol === apCol ? 2 : 1);
    checkLabel(principalCol, principal + 4, /^principal payment\b/i);
    checkLabel(interestCol, interest + 4, /^net interest payment\b/i);
    checkLabel(interestCol, interest + 8, /^ending accrued but unpaid interest\b/i);
    checkLabel(apCol, ap + 1, /^change in a\/?p$/i);
    checkLabel(chargeCol, charge, /^charge[ -]offs$/i);
    for (const [column, rows] of [[principalCol, [principal + 2, principal + 6]], [interestCol, [interest + 2, interest + 6]]]) {
      for (const row of rows) checkFormula(nextDetailColumn(column) + row);
    }
    const inputs = [
      pasteAddress(principalCol, principal + 4),
      pasteAddress(interestCol, interest + 4),
      pasteAddress(interestCol, interest + 8),
      pasteAddress(apCol, ap + 1),
      pasteAddress(chargeCol, charge),
    ];
    addresses.set(sheet, inputs);
    existingValues.set(sheet, inputs.map(address => cells.get(address)));
  }
  const coverPeriods = [], coverAudit = [];
  const title = "For the Month of " + new Intl.DateTimeFormat("en-US", { month: "long", year: "numeric", timeZone: "UTC" })
    .format(new Date(month + "-01T00:00:00Z"));
  for (const sheet of new Set([...profile.covers.map(([name]) => name), profile.payable[0]])) {
    if (!workbook.sheets.has(sheet)) {
      coverAudit.push([sheet, undefined]);
      continue;
    }
    const formulas = new Set(), cells = await readCells(workbook, sheet, formulas);
    coverAudit.push([sheet, findLoanStreetCoverAudit(cells, formulas)]);
    const current = String(cells.get("A3") ?? "").trim();
    const heading = /^for the month of (.+)$/i.exec(current)?.[1];
    if (sheet !== profile.payable[0] && heading && monthHeading(heading) && !formulas.has("A3")) coverPeriods.push([sheet, "A3", current, title]);
  }
  return { fileName: file.name, addresses, existingValues, monthOnly, coverPeriods, coverAudit };
}

function nextDetailColumn(column) {
  return String.fromCharCode(column.charCodeAt(0) + 1);
}

function detailWorksheets(result, destination, cells, audit) {
  const { label, number, date, row, worksheet } = cells;
  const review = [
    row(1, label("A1", "LoanStreet detail inputs · " + result.month, 1)),
    row(2, label("A2", destination ? "Destination: " + destination.fileName : "Paste into the current month's prepared blocks.")),
    row(3, label("A3", "Review the sign and posting labels for changes in accounts payable.")),
    ...(destination?.monthOnly ? [row(4, label("A4", "Verify the year for " + destination.monthOnly + " month-only headings."))] : []),
    row(5, ["Destination sheet", "Input", "Amount", "Paste cell", "Provider source",
      ...(destination ? ["Current destination value", "Change on paste"] : [])].map((title, index) =>
      label("ABCDEFG"[index] + "5", title, 1)).join("")),
  ];
  const sheets = [];
  const { pools, details } = result.profile;
  for (const [index, [sheet]] of details.entries()) {
    const addresses = destination?.addresses.get(sheet);
    const existing = destination?.existingValues.get(sheet);
    const column = result.reportedPools.get(pools[index]);
    const amounts = detailMetrics.map(([metric, source], i) => {
      const line = 6 + index * detailMetrics.length + i;
      const amount = result.values[source][index];
      const sourceCell = "Summary!" + column + metrics[source][1];
      const old = existing?.[i];
      const change = old === undefined ? undefined : (Math.round(amount * 100) - Math.round(old * 100)) / 100;
      review.push(row(line, label("A" + line, sheet) + label("B" + line, metric)
        + number("C" + line, amount) + (addresses ? label("D" + line, addresses[i]) : "") + label("E" + line, sourceCell)
        + (old === undefined ? "" : number("F" + line, old) + number("G" + line, change))));
      return amount;
    });
    if (addresses) {
      const entries = addresses.map((address, i) => [Number(address.slice(1)), address, amounts[i]]);
      const rows = new Map();
      for (const [rowNumber, address, amount] of entries.sort((a, b) => a[0] - b[0] || a[1].localeCompare(b[1]))) {
        rows.set(rowNumber, (rows.get(rowNumber) ?? "") + number(address, amount));
      }
      const body = row(1, label("A1", "Paste only the cells listed in Detail inputs.", 1))
        + Array.from(rows, ([rowNumber, content]) => row(rowNumber, content)).join("");
      sheets.push([sheet, worksheet("A1:V" + Math.max(...rows.keys()), [[1, 22, 16]], body)]);
    }
  }
  if (destination?.coverPeriods.length) {
    const periodRows = [row(1, ["Destination sheet", "Cell", "Current period title", "Proposed period title"]
      .map((title, i) => label("ABCD"[i] + "1", title, 1)).join("")),
    ...destination.coverPeriods.map((entry, i) => row(i + 2, entry.map((value, j) => label("ABCD"[j] + (i + 2), value)).join("")))];
    sheets.push(["Cover period inputs", worksheet("A1:D" + periodRows.length, [[1, 1, 27], [2, 2, 12], [3, 4, 37]], periodRows.join(""))]);
  }
  if (destination?.coverAudit?.length) {
    const auditRows = [row(1, ["Destination sheet", "Initials cell", "Completed by", "Date cell", "Completed on", "Review"]
      .map((title, i) => label("ABCDEF"[i] + "1", title, 1)).join(""))];
    for (const [i, [sheet, target]] of destination.coverAudit.entries()) {
      const r = i + 2;
      const values = target ? label("B" + r, target.initials) + label("C" + r, audit.by)
        + label("D" + r, target.date) + date("E" + r, audit.on) : label("F" + r, "Check the preparer area in the destination cover.");
      auditRows.push(row(r, label("A" + r, sheet) + values));
    }
    sheets.push(["Cover audit inputs", worksheet("A1:F" + auditRows.length,
      [[1, 1, 27], [2, 5, 17], [6, 6, 54]], auditRows.join(""))]);
  }
  return [["Detail inputs", worksheet("A1:" + (destination ? "G" : "E") + "45",
    [[1, 1, 25], [2, 2, 39], [3, 5, 17], ...(destination ? [[6, 7, 25]] : [])], review.join(""))], ...sheets];
}
