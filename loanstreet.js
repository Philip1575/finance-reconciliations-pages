"use strict";

const spreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const officeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const packageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
const excelMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const metrics = [
  ["Principal (net cash flows)", 21, "Principal (net cash flows)"],
  ["Net interest", 37, "Net Interest"],
  ["Accrued but unpaid interest", 30, "End of Period Balance"],
  ["Change in accounts payable", 46, "Change in Accounts Payable"],
  ["Principal charged off", 23, "Principal Charged Off This Period"],
  ["Recoveries", 36, "Recoveries (cash)"],
  ["Fees and other", 40, "Fees"],
  ["Principal owned", 24, "End of Period Balance"],
];

function requireReport(condition, message) {
  if (!condition) throw new Error(message);
}

async function readXml(zip, path) {
  const entry = zip.file(path);
  requireReport(entry, "The report is missing " + path + ".");
  const xml = new DOMParser().parseFromString(await entry.async("string"), "application/xml");
  requireReport(!xml.querySelector("parsererror"), "The Excel report contains invalid XML.");
  return xml;
}

function excelDate(value, date1904) {
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}(?:T|$)/.test(value)) return value.slice(0, 10);
  requireReport(Number.isInteger(value) && value >= 0 && value < 2957000, "The report contains an invalid reporting date.");
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return new Date(epoch + value * 86400000).toISOString().slice(0, 10);
}

async function openExcel(file, allowMacroWorkbook = false) {
  const validExtension = allowMacroWorkbook ? /\.xls[xm]$/i : /\.xlsx$/i;
  requireReport(validExtension.test(file.name), allowMacroWorkbook ? "Choose an Excel .xlsx or .xlsm workbook." : "Choose an Excel .xlsx export.");
  requireReport(file.size <= 20 * 1024 * 1024, "The workbook exceeds the 20 MB import limit.");
  let zip;
  try {
    zip = await JSZip.loadAsync(await file.arrayBuffer());
  } catch {
    throw new Error("Excel could not be read. Choose an unencrypted .xlsx export.");
  }
  const workbook = await readXml(zip, "xl/workbook.xml");
  const relations = await readXml(zip, "xl/_rels/workbook.xml.rels");
  const sheets = new Map(Array.from(workbook.getElementsByTagNameNS(spreadsheetNs, "sheet"),
    sheet => [sheet.getAttribute("name"), sheet.getAttributeNS(officeRelNs, "id")]));
  const date1904 = ["1", "true"].includes(workbook.getElementsByTagNameNS(spreadsheetNs, "workbookPr")[0]?.getAttribute("date1904"));
  return { zip, relations, sheets, date1904 };
}

function excelWorksheetPath({ relations, sheets }, sheetName) {
  requireReport(sheets.has(sheetName), "The workbook has no " + sheetName + " worksheet.");
  const relationship = Array.from(relations.getElementsByTagNameNS(packageRelNs, "Relationship"))
    .find(rel => rel.getAttribute("Id") === sheets.get(sheetName));
  const target = relationship?.getAttribute("Target")?.replace(/^\/?(?:xl\/)?/, "xl/");
  requireReport(relationship?.getAttribute("Type") === officeRelNs + "/worksheet" &&
    relationship.getAttribute("TargetMode") !== "External" &&
    /^xl\/worksheets\/[\w.-]+\.xml$/.test(target), "The selected worksheet could not be located.");
  return target;
}

async function readCells(workbook, sheetName, formulas) {
  const { zip } = workbook;
  const worksheet = await readXml(zip, excelWorksheetPath(workbook, sheetName));
  let strings = [];
  if (zip.file("xl/sharedStrings.xml")) {
    const shared = await readXml(zip, "xl/sharedStrings.xml");
    strings = Array.from(shared.getElementsByTagNameNS(spreadsheetNs, "si"), readText);
  }
  const cells = new Map();
  const sharedBases = new Map(), sharedFollowers = [];
  for (const cell of worksheet.getElementsByTagNameNS(spreadsheetNs, "c")) {
    const address = cell.getAttribute("r");
    const expression = cell.getElementsByTagNameNS(spreadsheetNs, "f")[0];
    if (formulas && expression) {
      if (formulas instanceof Map) {
        formulas.set(address, expression.textContent);
        if (expression.getAttribute("t") === "shared") {
          const index = expression.getAttribute("si");
          if (expression.textContent) sharedBases.set(index, { address, expression: expression.textContent, range: expression.getAttribute("ref") });
          else sharedFollowers.push({ index, address });
        }
      } else formulas.add(address);
    }
    const type = cell.getAttribute("t");
    const value = cell.getElementsByTagNameNS(spreadsheetNs, "v")[0]?.textContent;
    if (type === "inlineStr") cells.set(address, readText(cell));
    else if (type === "s") cells.set(address, strings[Number(value)]);
    else if (value?.trim()) cells.set(address, type === "n" || !type ? Number(value) : value);
  }
  for (const { index, address } of sharedFollowers) {
    const base = sharedBases.get(index);
    const range = /^([A-Z]+)(\d+):\1(\d+)$/.exec(base?.range ?? "");
    const cell = /^([A-Z]+)(\d+)$/.exec(address);
    if (range && cell && cell[1] === range[1] && base.address === range[1] + range[2]
      && Number(cell[2]) >= Number(range[2]) && Number(cell[2]) <= Number(range[3])) {
      formulas.set(address, { sharedFrom: base.address, expression: base.expression });
    }
  }
  return cells;
}

// Index of a plain cell format with the given number format in the workbook's styles, adding one if needed.
function templateNumberStyle(stylesDoc, numFmtId) {
  const cellXfs = stylesDoc.getElementsByTagNameNS(spreadsheetNs, "cellXfs")[0];
  const xfs = Array.from(cellXfs.children).filter(node => node.localName === "xf");
  const plain = xfs.findIndex(xf => xf.getAttribute("numFmtId") === String(numFmtId)
    && ["fontId", "fillId", "borderId"].every(key => (xf.getAttribute(key) ?? "0") === "0"));
  if (plain >= 0) return plain;
  const xf = stylesDoc.createElementNS(spreadsheetNs, cellXfs.prefix ? cellXfs.prefix + ":xf" : "xf");
  for (const [key, value] of [["numFmtId", numFmtId], ["fontId", 0], ["fillId", 0], ["borderId", 0], ["xfId", 0], ["applyNumberFormat", 1]]) xf.setAttribute(key, String(value));
  cellXfs.append(xf);
  cellXfs.setAttribute("count", String(xfs.length + 1));
  return xfs.length;
}

// Appends a block of rows (built from row 1) below the last used row of a worksheet, shifting the block's own cell
// references, formulas and merges. Cells written with s="x<index>" use that format index of this same workbook.
function appendTemplateBlock(doc, block, styleFor) {
  const sheetData = doc.getElementsByTagNameNS(spreadsheetNs, "sheetData")[0];
  const rowNodes = Array.from(sheetData.children).filter(node => node.localName === "row");
  const used = rowNodes.filter(row => Array.from(row.children).some(cell =>
    Array.from(cell.children).some(child => ["v", "is", "f"].includes(child.localName) && child.textContent.trim() !== "")))
    .map(row => Number(row.getAttribute("r")));
  const mergeCells = doc.getElementsByTagNameNS(spreadsheetNs, "mergeCells")[0];
  const mergeEnds = Array.from(mergeCells?.children ?? []).map(node => Number(/(\d+)$/.exec(node.getAttribute("ref") ?? "")?.[1] ?? 0));
  const offset = Math.max(0, ...used, ...mergeEnds);
  const shiftRef = text => text.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])/g,
    (all, colAbs, column, rowAbs, row) => rowAbs ? all : colAbs + column + (Number(row) + offset));
  const shiftCell = xml => xml
    .replace(/\br="([A-Z]{1,3})(\d+)"/, (_, column, row) => `r="${column}${Number(row) + offset}"`)
    .replace(/\ss="([^"]*)"/, (_, style) => ` s="${style.startsWith("x") ? style.slice(1) : styleFor(style)}"`)
    .replace(/<f>([^<]*)<\/f>/, (_, expression) => `<f>${shiftRef(expression)}</f>`);
  const parser = new DOMParser();
  for (const [r, cells] of [...block.grid].filter(([r]) => r <= block.end).sort((a, b) => a[0] - b[0])) {
    const target = r + offset, height = block.heights?.get(r);
    const xml = `<row xmlns="${spreadsheetNs}" r="${target}"${height ? ` ht="${height}" customHeight="1"` : ""}>`
      + [...cells].sort((a, b) => columnNumber(a[0]) - columnNumber(b[0])).map(([, cell]) => shiftCell(cell)).join("") + "</row>";
    const parsed = parser.parseFromString(xml, "application/xml");
    requireReport(!parsed.querySelector("parsererror"), "A generated block row could not be written.");
    const incoming = doc.importNode(parsed.documentElement, true);
    const existing = rowNodes.find(row => Number(row.getAttribute("r")) === target);
    if (!existing) {
      sheetData.insertBefore(incoming, rowNodes.find(row => Number(row.getAttribute("r")) > target) ?? null);
      continue;
    }
    if (height) { existing.setAttribute("ht", String(height)); existing.setAttribute("customHeight", "1"); }
    for (const cell of Array.from(incoming.children)) {
      const address = cell.getAttribute("r"), column = columnNumber(address.replace(/\d+$/, ""));
      const current = Array.from(existing.children).find(node => node.getAttribute("r") === address);
      if (current) existing.replaceChild(cell, current);
      else existing.insertBefore(cell, Array.from(existing.children).find(node =>
        columnNumber((node.getAttribute("r") ?? "").replace(/\d+$/, "")) > column) ?? null);
    }
  }
  if (block.merges.length) {
    let merges = mergeCells;
    if (!merges) {
      merges = doc.createElementNS(spreadsheetNs, sheetData.prefix ? sheetData.prefix + ":mergeCells" : "mergeCells");
      const after = ["phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
        "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
        "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];
      doc.documentElement.insertBefore(merges, Array.from(doc.documentElement.children).find(node => after.includes(node.localName)) ?? null);
    }
    for (const ref of block.merges) {
      const merge = doc.createElementNS(spreadsheetNs, merges.prefix ? merges.prefix + ":mergeCell" : "mergeCell");
      merge.setAttribute("ref", shiftRef(ref));
      merges.append(merge);
    }
    merges.setAttribute("count", String(merges.children.length));
  }
  const dimension = doc.getElementsByTagNameNS(spreadsheetNs, "dimension")[0];
  const last = block.end + offset, ref = /^([A-Z]+\d+)(?::([A-Z]+)(\d+))?$/.exec(dimension?.getAttribute("ref") ?? "");
  if (ref && (!ref[3] || Number(ref[3]) < last)) dimension.setAttribute("ref", ref[1] + ":" + (ref[2] ?? "Z") + last);
  return offset + 1;
}

async function fillExcelTemplate(file, updates, appends = []) {
  const workbook = await openExcel(file, true);
  const serializer = new XMLSerializer();
  const element = (doc, parent, name) => doc.createElementNS(spreadsheetNs, parent.prefix ? parent.prefix + ":" + name : name);
  const groups = new Map();
  for (const [sheet, address, value] of updates) {
    const group = groups.get(sheet) ?? new Map();
    requireReport(!group.has(address), "Multiple values target " + sheet + "!" + address + ".");
    group.set(address, value);
    groups.set(sheet, group);
  }
  for (const [sheet, group] of groups) {
    const target = excelWorksheetPath(workbook, sheet);
    const doc = await readXml(workbook.zip, target);
    const rows = new Map(Array.from(doc.getElementsByTagNameNS(spreadsheetNs, "row"), row => [row.getAttribute("r"), row]));
    const cells = new Map(Array.from(doc.getElementsByTagNameNS(spreadsheetNs, "c"), cell => [cell.getAttribute("r"), cell]));
    for (const [address, value] of group) {
      const location = /^([A-Z]{1,3})([1-9]\d*)$/.exec(address);
      requireReport(location && rows.has(location[2]), "Missing template row at " + sheet + "!" + address + ".");
      let cell = cells.get(address);
      if (!cell && value === null) continue;
      if (!cell) {
        const parent = rows.get(location[2]);
        cell = element(doc, parent, "c");
        cell.setAttribute("r", address);
        const next = Array.from(parent.children).find(child => child.localName !== "c"
          || columnNumber(child.getAttribute("r").replace(/\d+$/, "")) > columnNumber(location[1]));
        parent.insertBefore(cell, next ?? null);
      }
      const existingFormula = cell.getElementsByTagNameNS(spreadsheetNs, "f")[0];
      const formulaText = value?.formula;
      requireReport(!existingFormula || formulaText === existingFormula.textContent,
        "The template contains a formula at " + sheet + "!" + address + ".");
      for (const child of Array.from(cell.children)) {
        if (["v", "is"].includes(child.localName)) child.remove();
      }
      if (value === null) {
        cell.removeAttribute("t");
        continue;
      }
      const text = typeof value === "string";
      cell.setAttribute("t", text ? "inlineStr" : "n");
      const output = element(doc, cell, text ? "is" : "v");
      if (text) {
        const node = element(doc, output, "t");
        node.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
        node.textContent = value;
        output.append(node);
      } else {
        const epoch = workbook.date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
        const number = typeof value === "number" ? value : formulaText ? value.cached
          : (Date.parse(value.date + "T00:00:00Z") - epoch) / 86400000;
        requireReport(Number.isFinite(number), "Invalid value for " + sheet + "!" + address + ".");
        output.textContent = String(number);
      }
      cell.insertBefore(output, existingFormula ? existingFormula.nextSibling : cell.firstChild);
      if (formulaText && !existingFormula) {
        const formula = element(doc, cell, "f");
        formula.textContent = formulaText;
        cell.insertBefore(formula, output);
      }
    }
    workbook.zip.file(target, serializer.serializeToString(doc));
  }
  if (appends.length) {
    const stylesDoc = await readXml(workbook.zip, "xl/styles.xml");
    const plain = new Map();
    // The app's own plain formats: 2 = number, 3 = date; anything else becomes the workbook's default format.
    const styleFor = style => {
      const numFmtId = { 2: 4, 3: 14 }[style];
      if (!numFmtId) return "0";
      if (!plain.has(numFmtId)) plain.set(numFmtId, templateNumberStyle(stylesDoc, numFmtId));
      return String(plain.get(numFmtId));
    };
    for (const block of appends) {
      const target = excelWorksheetPath(workbook, block.sheet);
      const doc = await readXml(workbook.zip, target);
      block.startRow = appendTemplateBlock(doc, block, styleFor);
      workbook.zip.file(target, serializer.serializeToString(doc));
    }
    workbook.zip.file("xl/styles.xml", serializer.serializeToString(stylesDoc));
  }
  const doc = await readXml(workbook.zip, "xl/workbook.xml");
  const root = doc.documentElement;
  let calc = doc.getElementsByTagNameNS(spreadsheetNs, "calcPr")[0];
  if (!calc) {
    calc = element(doc, root, "calcPr");
    const after = ["oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing",
      "fileRecoveryPr", "webPublishObjects", "extLst"];
    root.insertBefore(calc, Array.from(root.children).find(child => after.includes(child.localName)) ?? null);
  }
  calc.setAttribute("calcMode", "auto");
  calc.setAttribute("fullCalcOnLoad", "1");
  workbook.zip.file("xl/workbook.xml", serializer.serializeToString(doc));
  return workbook.zip.generateAsync({ type: "blob", compression: "DEFLATE" });
}

function templateTableUpdates(sheet, range, columns, entries, cells, formulas) {
  requireReport(entries.length <= range.last - range.first + 1, "Extend the total range on " + sheet + " to include all reviewed rows.");
  const address = /^([A-Z]+)([1-9]\d*)$/;
  const allowed = new Set(columns);
  const inRange = cell => {
    const match = address.exec(cell);
    return match && allowed.has(match[1]) && Number(match[2]) >= range.first && Number(match[2]) <= range.last;
  };
  requireReport(![...formulas.keys()].some(inRange), "The input range on " + sheet + " contains formulas. Review the paste list.");
  const updates = new Map([...cells.keys()].filter(inRange).map(cell => [cell, null]));
  entries.forEach((entry, i) => entry.forEach((value, j) => updates.set(columns[j] + (range.first + i), value)));
  return [...updates].map(([cell, value]) => [sheet, cell, value]);
}

async function readLoanStreet(file, month, profile) {
  requireReport(/^\d{4}-\d{2}$/.test(month), "Choose the reconciliation month first.");
  const { pools } = profile;
  const workbook = await openExcel(file);
  const cells = await readCells(workbook, "Summary");
  requireReport(cells.get("A1") === "Buy Side Combined Report" && cells.get("A3") === "Reporting Period:",
    "Choose a LoanStreet Buy Side Combined Report.");
  const start = excelDate(cells.get("B3"), workbook.date1904);
  const end = excelDate(cells.get("C3"), workbook.date1904);
  const [year, monthNumber] = month.split("-").map(Number);
  const lastDay = new Date(Date.UTC(year, monthNumber, 0)).getUTCDate();
  requireReport(start === month + "-01" && end === month + "-" + lastDay,
    "The report covers " + start + " through " + end + ". Choose its month or load the correct report.");

  requireReport(cells.get("A4") === "Sellers Complete" && cells.get("A5") === "Sellers Pending" &&
    cells.get("D7") === "Full Portfolio" && cells.get("D8") === "Ticker",
    "The report layout differs from the supported LoanStreet format.");
  requireReport(!Array.from(cells).some(([address, value]) => /^[B-Z][A-Z]*5$/.test(address) && String(value).trim()),
    "The report lists pending sellers. Download a complete report before continuing.");
  const poolStart = Array.from(cells).find(([address, value]) => /^[A-Z]+7$/.test(address) && value === "By Pool");
  requireReport(poolStart, "The Summary worksheet has no By Pool section.");
  const firstPoolColumn = columnNumber(poolStart[0].slice(0, -1));
  const reportedPools = new Map();
  for (const [address, value] of cells) {
    if (!/^[A-Z]+8$/.test(address) || !String(value).trim()) continue;
    const column = address.slice(0, -1);
    if (columnNumber(column) < firstPoolColumn) continue;
    requireReport(pools.includes(value) && !reportedPools.has(value),
      "The report includes an unexpected or duplicate pool. Have finance confirm the pool mapping.");
    reportedPools.set(value, column);
  }
  requireReport(pools.every(pool => reportedPools.has(pool)), "The report is missing one or more expected pools.");
  const completeSellers = new Set(Array.from(cells)
    .filter(([address, value]) => /^[A-Z]+4$/.test(address) && address !== "A4" && String(value).trim())
    .map(([, value]) => value));
  requireReport(Array.from(reportedPools.values()).every(column => completeSellers.has(cells.get(column + "9"))),
    "One or more pool sellers are missing from Sellers Complete.");

  function amount(address) {
    const value = cells.get(address);
    requireReport(typeof value === "number" && Number.isFinite(value), "Missing or invalid amount at Summary!" + address + ".");
    return value;
  }
  const requiredLabels = [...metrics.map(([, row, label]) => [row, label]),
    [19, "Owned Portfolio"], [26, "Accrued But Unpaid Interest (Less Servicing Spread)"],
    [41, "Other"], [47, "Accounts Payable at End of Period"]];
  for (const [row, label] of requiredLabels) {
    requireReport(cells.get("A" + row) === label, "Expected " + label + " at Summary!A" + row + ". Check the report layout.");
  }
  const values = metrics.map(([, row]) => pools.map(pool => {
    const column = reportedPools.get(pool);
    const value = amount(column + row);
    return row === 40 ? value + amount(column + 41) : value;
  }));
  return { month, fileName: file.name, values, payable: amount("D47"), reportedPools, profile };
}

function readText(element) {
  return Array.from(element.getElementsByTagNameNS(spreadsheetNs, "t"), node => node.textContent).join("");
}

function columnNumber(name) {
  return Array.from(name).reduce((number, letter) => number * 26 + letter.charCodeAt(0) - 64, 0);
}

function xmlText(value) {
  return String(value).replace(/[&<>"']/g, symbol => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;",
  })[symbol]).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");
}

function excelOutputCells() {
  const label = (cell, value, style = 0) =>
    `<c r="${cell}" s="${style}" t="inlineStr"><is><t>${xmlText(value)}</t></is></c>`;
  const number = (cell, value, style = 2) => `<c r="${cell}" s="${style}"><v>${value}</v></c>`;
  const date = (cell, value, style = 3) => {
    if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value) || value < "1900-03-01") return label(cell, value);
    const timestamp = Date.parse(value + "T00:00:00Z");
    if (!Number.isFinite(timestamp) || new Date(timestamp).toISOString().slice(0, 10) !== value) return label(cell, value);
    const serial = (timestamp - Date.UTC(1899, 11, 30)) / 86400000;
    return `<c r="${cell}" s="${style}"><v>${serial}</v></c>`;
  };
  const row = (index, content) => `<row r="${index}">${content}</row>`;
  const worksheet = (range, columns, rows, merges = []) => `<worksheet xmlns="${spreadsheetNs}">
    <dimension ref="${range}"/><cols>${columns.map(([min, max, width]) =>
      `<col min="${min}" max="${max}" width="${width}" customWidth="1"/>`).join("")}</cols>
    <sheetData>${rows}</sheetData>${merges.length ? `<mergeCells count="${merges.length}">${merges.map(ref => `<mergeCell ref="${ref}"/>`).join("")}</mergeCells>` : ""}
  </worksheet>`;
  return { label, number, date, row, worksheet };
}

document.addEventListener("DOMContentLoaded", () => {
  const section = document.createElement("section");
  section.className = "panel";
  section.setAttribute("aria-labelledby", "completion-heading");
  section.innerHTML = `<h2 id="completion-heading">Completion details</h2>
    <div class="mapping">
      <label>Completed by (initials)<input id="completed-by" type="text" maxlength="12" autocomplete="off" required></label>
      <label>Completed on<input id="completed-on" type="date" min="1900-03-01" max="9999-12-31" required></label>
    </div>`;
  document.querySelector("main > header").after(section);
});

function reconciliationAudit() {
  const initials = document.getElementById("completed-by");
  const completedOn = document.getElementById("completed-on");
  const by = initials.value.trim();
  if (!by) initials.focus();
  requireReport(by, "Enter the preparer's initials in Completed by.");
  if (!completedOn.validity.valid) completedOn.focus();
  requireReport(completedOn.validity.valid, "Enter a valid date in Completed on.");
  return { by, on: completedOn.value };
}

function completionTarget(sheet, cells, formulas, prepared, verified, initials, date) {
  const read = address => String(cells.get(address) ?? "").trim();
  if (!/^(?:reconciliation )?(?:prepared|completed) by\s*:?$/i.test(read(prepared))
    || !/^(?:reconciliation )?verified by\s*:?$/i.test(read(verified))) return undefined;
  if (formulas.has(initials) || formulas.has(date)) return undefined;
  if (cells.has(initials) && typeof cells.get(initials) !== "string") return undefined;
  if (cells.has(date) && typeof cells.get(date) !== "number") return undefined;
  return { sheet, initials, date };
}

async function readCompletionTarget(workbook, sheet, prepared, verified, initials, date) {
  if (!workbook?.sheets.has(sheet)) return undefined;
  const formulas = new Set();
  const cells = await readCells(workbook, sheet, formulas);
  return completionTarget(sheet, cells, formulas, prepared, verified, initials, date);
}

function completionPasteSheets(target, audit) {
  if (!target) return [];
  const { label, date, row, worksheet } = excelOutputCells();
  return [["Cover audit inputs", worksheet("A1:D3", [[1, 1, 27], [2, 2, 13], [3, 4, 20]],
    row(1, ["Destination sheet", "Cell", "Field", "Paste value"].map((title, i) => label("ABCD"[i] + "1", title, 1)).join(""))
    + row(2, label("A2", target.sheet) + label("B2", target.initials) + label("C2", "Completed by") + label("D2", audit.by))
    + row(3, label("A3", target.sheet) + label("B3", target.date) + label("C3", "Completed on") + date("D3", audit.on))) ]];
}

// Copies the cell formats a download borrows from an uploaded workbook (cells written with s="x<index>") into the
// download's style sheet, renumbering fonts, fills, borders and formats. Returns the new style sheet parts and the
// mapping from the uploaded workbook's format index to the download's.
function borrowedExcelStyles(sourceXml, indices, base) {
  const doc = new DOMParser().parseFromString(sourceXml, "application/xml");
  requireReport(!doc.querySelector("parsererror"), "The workbook formatting could not be read.");
  const serializer = new XMLSerializer();
  const list = (parent, child) => Array.from(doc.getElementsByTagNameNS(spreadsheetNs, parent)[0]?.children ?? [])
    .filter(node => node.localName === child);
  const sources = { fontId: list("fonts", "font"), fillId: list("fills", "fill"), borderId: list("borders", "border") };
  const formats = new Map(list("numFmts", "numFmt").map(node => [node.getAttribute("numFmtId"), node.getAttribute("formatCode")]));
  const xfs = list("cellXfs", "xf");
  const added = { fontId: [], fillId: [], borderId: [] }, renumbered = { fontId: new Map(), fillId: new Map(), borderId: new Map() };
  const numFmts = new Map(), cellXfs = [], map = new Map();
  for (const index of indices) {
    const xf = xfs[index];
    if (!xf) continue;
    const copy = xf.cloneNode(true);
    for (const key of ["fontId", "fillId", "borderId"]) {
      const from = xf.getAttribute(key) ?? "0";
      if (!renumbered[key].has(from)) {
        renumbered[key].set(from, base[key] + added[key].length);
        added[key].push(serializer.serializeToString(sources[key][Number(from)] ?? sources[key][0]));
      }
      copy.setAttribute(key, String(renumbered[key].get(from)));
    }
    const format = xf.getAttribute("numFmtId") ?? "0";
    if (Number(format) >= 164 && formats.has(format)) numFmts.set(format, formats.get(format));
    copy.setAttribute("xfId", "0");
    cellXfs.push(serializer.serializeToString(copy));
    map.set(index, base.xf + cellXfs.length - 1);
  }
  return { fonts: added.fontId, fills: added.fillId, borders: added.borderId, numFmts, cellXfs, map };
}

async function createExcelWorkbook(sheets, audit = reconciliationAudit(), styleSource) {
  const { label, date, row, worksheet } = excelOutputCells();
  sheets = [...sheets, ["Completion record", worksheet("A1:B2", [[1, 2, 23]],
    row(1, label("A1", "Completed by") + label("B1", audit.by))
    + row(2, label("A2", "Completed on") + date("B2", audit.on)))]];
  const zip = new JSZip();
  const types = "http://schemas.openxmlformats.org/package/2006/content-types";
  const docType = "application/vnd.openxmlformats-officedocument.spreadsheetml";
  zip.file("_rels/.rels", `<Relationships xmlns="${packageRelNs}">
    <Relationship Id="rId1" Type="${officeRelNs}/officeDocument" Target="xl/workbook.xml"/>
  </Relationships>`);
  const borrowedIndices = [...new Set(sheets.flatMap(([, xml]) => [...xml.matchAll(/\ss="x(\d+)"/g)].map(match => Number(match[1]))))];
  const borrowed = styleSource && borrowedIndices.length
    ? borrowedExcelStyles(styleSource.styles, borrowedIndices, { fontId: 2, fillId: 3, borderId: 1, xf: 4 }) : undefined;
  sheets = sheets.map(([name, xml]) => [name, xml.replace(/\ss="x(\d+)"/g, (_, index) => ` s="${borrowed?.map.get(Number(index)) ?? 0}"`)]);
  const fonts = [`<font><sz val="11"/><name val="Calibri"/></font>`, `<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>`,
    ...(borrowed?.fonts ?? [])];
  const fills = [`<fill><patternFill patternType="none"/></fill>`, `<fill><patternFill patternType="gray125"/></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FF136B63"/><bgColor indexed="64"/></patternFill></fill>`, ...(borrowed?.fills ?? [])];
  const borders = [`<border><left/><right/><top/><bottom/><diagonal/></border>`, ...(borrowed?.borders ?? [])];
  const cellXfs = [`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`,
    `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>`,
    `<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`,
    `<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`, ...(borrowed?.cellXfs ?? [])];
  const numFmts = [...(borrowed?.numFmts ?? [])];
  zip.file("xl/styles.xml", `<styleSheet xmlns="${spreadsheetNs}">
    ${numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${xmlText(code)}"/>`).join("")}</numFmts>` : ""}
    <fonts count="${fonts.length}">${fonts.join("")}</fonts>
    <fills count="${fills.length}">${fills.join("")}</fills>
    <borders count="${borders.length}">${borders.join("")}</borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="${cellXfs.length}">${cellXfs.join("")}</cellXfs>
    <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  </styleSheet>`);
  const theme = borrowed && styleSource.theme;
  if (theme) zip.file("xl/theme/theme1.xml", theme);
  const parts = [];
  const names = [];
  const relations = [];
  for (const [index, [name, xml]] of sheets.entries()) {
    const id = index + 1;
    const path = "worksheets/sheet" + id + ".xml";
    zip.file("xl/" + path, xml);
    parts.push(`<Override PartName="/xl/${path}" ContentType="${docType}.worksheet+xml"/>`);
    names.push(`<sheet name="${xmlText(name)}" sheetId="${id}" r:id="rId${id}"/>`);
    relations.push(`<Relationship Id="rId${id}" Type="${officeRelNs}/worksheet" Target="${path}"/>`);
  }
  zip.file("[Content_Types].xml", `<Types xmlns="${types}">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="${docType}.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="${docType}.styles+xml"/>
    ${theme ? `<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` : ""}
    ${parts.join("")}
  </Types>`);
  zip.file("xl/workbook.xml", `<workbook xmlns="${spreadsheetNs}" xmlns:r="${officeRelNs}">
    <sheets>${names.join("")}</sheets>
  </workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<Relationships xmlns="${packageRelNs}">
    ${relations.join("")}
    <Relationship Id="styles" Type="${officeRelNs}/styles" Target="styles.xml"/>
    ${theme ? `<Relationship Id="theme" Type="${officeRelNs}/theme" Target="theme/theme1.xml"/>` : ""}
  </Relationships>`);
  return zip.generateAsync({ type: "blob", mimeType: excelMime, compression: "DEFLATE" });
}

async function buildKeyInfo(result, wisdom, detailTarget) {
  const audit = reconciliationAudit();
  const outputCells = excelOutputCells();
  const { label, number, row, worksheet } = outputCells;
  const { pools } = result.profile;
  const sum = (row, values) => `<c r="O${row}" s="2"><f>SUM(G${row}:N${row})</f><v>${values.reduce((total, value) => total + value, 0)}</v></c>`;
  const columns = "GHIJKLMN";
  const header = row(2, label("F2", "LoanStreet KEY INFO · " + result.month, 1)
    + pools.map((pool, index) => label(columns[index] + "2", pool, 1)).join("") + label("O2", "Pool total", 1));
  const detail = metrics.map(([name], index) => {
    const rowNumber = index + 3;
    const amounts = result.values[index];
    const total = [8, 9].includes(rowNumber) ? sum(rowNumber, amounts) : "";
    return row(rowNumber, label("F" + rowNumber, name) +
      amounts.map((value, poolIndex) => number(columns[poolIndex] + rowNumber, value)).join("") + total);
  }).join("");
  const payable = row(12, label("F12", "Ending accounts payable · full portfolio") + number("G12", result.payable));
  const source = row(15, label("F15", "Source") + label("G15", result.fileName))
    + row(16, label("F16", "Provider worksheet") + label("G16", "Summary"))
    + row(17, label("F17", "Source pool columns")
      + pools.map((pool, index) => label(columns[index] + "17", result.reportedPools.get(pool))).join(""))
    + row(18, label("F18", "Source rows in output order") + label("G18", "21, 37, 30, 46, 23, 36, 40 + 41, 24. Full portfolio accounts payable from D47."));
  const keyInfo = worksheet("F2:O18", [[1, 5, 3], [6, 6, 43], [7, 15, 17]], header + detail + payable + source);
  return createExcelWorkbook([
    ["KEY INFO " + result.month, keyInfo],
    ...coverWorksheets(result, wisdom, outputCells),
    ...detailWorksheets(result, detailTarget, outputCells, audit),
  ], audit);
}

