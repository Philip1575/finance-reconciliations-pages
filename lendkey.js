"use strict";

const lendField = id => document.getElementById(id);
const lendSources = { ach: { revision: 0 }, fed: { revision: 0 }, portfolio: { revision: 0 }, performance: { revision: 0 }, wisdom: { revision: 0 } };
let lendResult, lendRevision = 0;
let lendStatementFiles = [];
addTrialBalanceDetection("wisdom", () => lendSources.wisdom.book ? [lendSources.wisdom] : []);

const lendProducts = [
  ["HIL", "Home Improvement", ["HIL", "HOME IMPROVEMENT", "HOME IMPROVEMENT LOAN"]],
  ["CNL", "Consolidation", ["CNL", "CONSOLIDATION", "CONSOLIDATION LOAN"]],
  ["STU", "Student", ["STU", "STUDENT", "STUDENT LOAN"]],
];
const lendPurposeFields = [
  ["prior-ar", "Prior A/R"], ["prior-ap", "Prior A/P"],
  ["new-ar", "New A/R"], ["new-ap", "New A/P"], ["late-fees", "Late fees"],
];
const lendProviderSheet = "LenderACHDetail_Monthly+Daily";
const lendProviderHeaders = {
  date: "Transaction_Date", direction: "ACH_Direction", product: "Product_Name", reason: "Transaction_Type",
  amount: "Transaction_Amount", settlement: "ACH_Date", "status-col": "AR_AP", reference: "ACH_Reference_ID",
};

const lendSchedules = [
  {
    name: "Loan Principal Recs",
    fields: [[5, "Beg Principal Balance"], [6, "Prior Period Adjustment"], [9, "Disbursed"], [10, "Payment reversals"],
      [11, "Origination Fee"], [12, "Refunds Applied to Principal"], [13, "Interest Capitalized"], [15, "Principal Purchased"],
      [19, "Borrower Repayment, Orig"], [21, "Borrower Repayment, Part"], [23, "Principal Cancelled/Defaulted"],
      [25, "Adjustments"], [27, "Principal Sold"], [29, "Principal Recovery"], [31, "Principal Adjustment"], [37, "Statement ending balance"]],
    plus: [9, 10, 11, 12, 13, 15], less: [19, 21, 23, 25, 27], other: [29, 31],
    totals: [16, 28], end: 33, change: 35, statement: 37, difference: 38,
  },
  {
    name: "Interest Recs",
    fields: [[5, "Beg Accrued Interest"], [6, "Prior Period Adjustment"], [9, "Interest Earned, Orig"], [11, "Interest Earned, Part"],
      [13, "Loan Accrued Purchased"], [18, "Interest Capitalized"], [20, "Interest Paid"], [22, "Interest Reversal"],
      [24, "Interest Loan Cancel/Default"], [26, "Loan Accrued Interest Sold"], [33, "Statement ending balance"]],
    plus: [9, 11, 13], less: [18, 20, 22, 24, 26], other: [],
    totals: [14, 27], end: 29, change: 31, statement: 33, difference: 34,
  },
];
const lendPortfolioSheet = "Recon Adj Wkbook";
const lendPortfolioFields = [
  [
    [5, "H9", "PR Ending Prin Balance"], [9, "H15", "DISB"], [10, "H17", "PMT Reversals (PRIN) (3)"],
    [12, "H18", "Refunds Applied to PRIN (4)"], [13, "H16", "Capitalized INT"],
    [21, "H24", "PMT Applied to PRIN (1)"], [27, "H27", "PRIN Claim Payment (2)"], [37, "H35", "Ending PRIN Balance"],
  ],
  [
    [5, "M9", "Accr INT Ending Balance"], [6, "M10", "Prior Period Adjustments"],
    [11, "M15", "INT ACCR During Period"], [18, "M23", "Capitalized INT"],
    [20, "M24", "PMT Applied to INT (1)"], [22, "M16", "PMT Reversals (INT) (3)", -1],
    [26, "M27", "INT Claim Payment (2)"], [33, "M35", "Ending INT Balance"],
  ],
];

function lendStatus(message, error = false) {
  lendField("status").textContent = message;
  lendField("status").classList.toggle("error", error);
}

function invalidateLend() {
  lendRevision++;
  lendResult = undefined;
  lendField("confirm").checked = false;
  lendField("results").hidden = true;
  lendStatus("");
}

async function selectLendFile(kind, file) {
  const state = lendSources[kind], current = ++state.revision;
  state.file = file;
  state.book = undefined;
  if (kind === "fed") {
    lendField("fed-exclusions").value = "";
    lendField("fed-confirmed").value = "";
  }
  if (state.providerLayout) {
    for (const [field, value] of Object.entries(state.providerLayout.previous)) lendField("ach-" + field).value = value;
    state.providerLayout = undefined;
  }
  invalidateLend();
  const status = lendField(kind + "-status"), mapping = lendField(kind + "-map");
  if (mapping) mapping.hidden = true;
  status.textContent = file ? "Opening " + file.name + "…" : "";
  if (!file) return;
  try {
    const book = await openExcel(file, true);
    if (current !== state.revision) return;
    const providerLayout = kind === "ach" && book.sheets.has(lendProviderSheet) ? await readLendProviderLayout(book) : undefined;
    if (current !== state.revision) return;
    state.book = book;
    if (mapping) {
      const sheet = lendField(kind + "-sheet");
      sheet.replaceChildren(...Array.from(book.sheets.keys(), name => new Option(name, name)));
      if (kind === "ach") {
        const ach = [...book.sheets.keys()].find(name => name.trim().toLowerCase() === "ach");
        if (ach) sheet.value = ach;
        if (providerLayout) {
          sheet.value = lendProviderSheet;
          const values = { ...providerLayout.columns, first: providerLayout.first, last: "", exclude: "" };
          const previous = Object.fromEntries(Object.keys(values).map(field => [field, lendField("ach-" + field).value]));
          for (const [field, value] of Object.entries(values)) lendField("ach-" + field).value = value;
          state.providerLayout = { ...providerLayout, previous };
        }
      }
      mapping.hidden = false;
    }
    status.textContent = file.name + (providerLayout ? " · Provider ACH columns detected." : "");
  } catch (error) {
    if (current === state.revision) status.textContent = error.message;
  }
}

async function readLendProviderLayout(book) {
  const cells = await readCells(book, lendProviderSheet);
  const rows = new Map(), fields = new Map(Object.entries(lendProviderHeaders).map(([field, heading]) => [heading.toLowerCase(), field]));
  for (const [address, value] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (!match || Number(match[2]) > 50 || typeof value !== "string") continue;
    const field = fields.get(value.trim().toLowerCase());
    if (!field) continue;
    const row = Number(match[2]), columns = rows.get(row) ?? new Map();
    columns.set(field, columns.has(field) ? "" : match[1]);
    rows.set(row, columns);
  }
  const matches = [...rows].filter(([, columns]) => columns.size === fields.size && [...columns.values()].every(Boolean));
  if (matches.length !== 1) return undefined;
  const [[row, columns]] = matches;
  return { first: row + 1, columns: Object.fromEntries(columns) };
}

function lendDirection(value) {
  const text = String(value ?? "").trim().toUpperCase();
  if (["DR", "DB", "DEBIT"].includes(text)) return "Debit";
  if (["CR", "CD", "CREDIT"].includes(text)) return "Credit";
  throw new Error("Expected a debit or credit transaction type.");
}

function readLendPostingDates() {
  const overrides = new Map();
  for (const [index, line] of lendField("posting-dates").value.trim().split(/\r?\n/).entries()) {
    if (!line.trim()) continue;
    const [achDate, direction, fedDate, ...extra] = line.split("\t").map(value => value.trim());
    if (index === 0 && achDate.toLowerCase() === "ach date") continue;
    try {
      requireReport(achDate && direction && fedDate && !extra.length, "Use three columns: ACH date, ACH direction, and Fed date.");
      const date = sourceDate(achDate), posting = sourceDate(fedDate);
      const type = lendDirection(direction), key = date + type;
      requireReport(!overrides.has(key), "An override already exists for this ACH date and direction.");
      overrides.set(key, { date, type, posting, used: false });
    } catch (error) {
      throw new Error("Posting date row " + (index + 1) + " · " + error.message);
    }
  }
  return overrides;
}

async function lendAchRows(kind, from, through, reverse, overrides) {
  const state = lendSources[kind];
  const required = kind === "ach" ? ["date", "direction", "product", "reason", "amount"] : ["date", "direction", "amount"];
  const optional = kind === "ach" ? ["status-col", "settlement", "reference", "exclude"] : ["reference", "exclude", "finance-provider"];
  const fields = new Map([...required, ...optional].map(key => [key, lendField(kind + "-" + key).value.trim().toUpperCase()]));
  const first = lendField(kind + "-first").valueAsNumber, lastField = lendField(kind + "-last");
  const last = lastField.value ? lastField.valueAsNumber : 1048576;
  requireReport(Number.isInteger(first) && Number.isInteger(last) && first >= 1 && first <= last && last <= 1048576,
    "Check the first and last " + kind.toUpperCase() + " data rows.");
  const sheet = lendField(kind + "-sheet").value;
  const provider = state.providerLayout?.columns;
  const originalProviderColumns = sheet === lendProviderSheet && provider
    && ["direction", "settlement"].every(field => fields.get(field) === provider[field]);
  const cells = await readCells(state.book, sheet);
  if (kind === "fed" && state.book.sheets.size === 1) {
    const history = providerFedHistory(cells, state.book.date1904, "LendKey");
    if (history) return history.map(item => {
      const direction = item.amount > 0 ? "Debit" : "Credit";
      const inRange = item.date >= from && item.date <= through;
      return { kind, date: item.date, posted: item.posted, settlement: "", periodDate: item.date, direction, comparison: direction,
        sourceAmount: -item.amount, amount: Math.abs(item.amount), row: item.row, reference: item.description, financeProvider: "LendKey",
        product: "", reason: "", status: "", override: "", exclusion: "", review: "", inRange, included: inRange };
    });
  }
  const financeColumn = kind === "fed" ? fields.get("finance-provider") || financeProviderColumn(cells, 1, Math.min(first - 1, 50)) : undefined;
  if (financeColumn) fields.set("finance-provider", financeColumn);
  const columns = [...fields.values()].filter(Boolean);
  requireReport(required.every(key => fields.get(key)) && new Set(columns).size === columns.length
    && columns.every(column => /^[A-Z]{1,3}$/.test(column) && columnNumber(column) <= 16384),
    "Choose valid, distinct " + kind.toUpperCase() + " columns.");
  const rows = new Set();
  for (const [address, value] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (!match || !columns.includes(match[1]) || !String(value ?? "").trim()) continue;
    const row = Number(match[2]);
    if (row >= first && row <= last) rows.add(row);
  }
  return Array.from(rows).sort((a, b) => a - b).flatMap(row => {
    if (financeColumn && !matchesFinanceProvider(cells.get(financeColumn + row), "LendKey")) return [];
    const get = field => cells.get(fields.get(field) + row);
    const text = field => String(get(field) ?? "").trim();
    try {
      const settlement = text("settlement") ? sourceDate(get("settlement"), state.book.date1904) : "";
      const date = text("date") || !settlement ? sourceDate(get("date"), state.book.date1904) : "";
      const missingSettlement = originalProviderColumns && (!settlement || !text("direction"));
      const direction = missingSettlement && !text("direction") ? "" : lendDirection(get("direction"));
      const comparison = missingSettlement ? "" : kind === "ach" && reverse ? (direction === "Debit" ? "Credit" : "Debit") : direction;
      const override = settlement && direction ? overrides?.get(settlement + direction) : undefined;
      if (override) override.used = true;
      const periodDate = override?.posting ?? (settlement || date);
      const sourceAmount = cents(get("amount"));
      const review = missingSettlement ? "Provider ACH date or direction missing" : "";
      const exclusion = text("exclude"), inRange = periodDate >= from && periodDate <= through;
      // The standard provider report lists positive amounts with a separate direction, so a negative row is an
      // adjustment within its batch and must reduce the batch total rather than add to it.
      const amount = originalProviderColumns ? sourceAmount : Math.abs(sourceAmount);
      return {
        kind, date, settlement, periodDate, direction, comparison, sourceAmount, amount, row,
        override: override?.posting ?? "", providerDetail: Boolean(originalProviderColumns),
        product: text("product"), reason: text("reason"), status: text("status-col"), reference: text("reference"),
        financeProvider: text("finance-provider"),
        inRange, exclusion, review, included: inRange && !exclusion && !review,
      };
    } catch (error) {
      throw new Error(sheet + " row " + row + ": " + error.message);
    }
  });
}

function compareLendAch(ach, fed) {
  const sums = { Debit: { fed: 0, ach: 0 }, Credit: { fed: 0, ach: 0 } };
  const dates = new Map();
  const all = [...fed, ...ach];
  for (const item of all) {
    if (!item.included) continue;
    const key = item.periodDate + item.comparison;
    const day = dates.get(key) ?? { date: item.periodDate, direction: item.comparison, ach: 0, fed: 0,
      achCount: 0, fedCount: 0, achRows: [], fedRows: [] };
    day[item.kind] += item.amount;
    day[item.kind + "Count"]++;
    day[item.kind + "Rows"].push(item);
    sums[item.comparison][item.kind] += item.amount;
    requireReport(Number.isSafeInteger(day[item.kind]) && Number.isSafeInteger(sums[item.comparison][item.kind]),
      "An ACH total exceeds the supported amount range.");
    dates.set(key, day);
  }
  const days = [...dates.values()].sort((a, b) => a.date.localeCompare(b.date) || a.direction.localeCompare(b.direction));
  const matches = lendBatchMatches(days);
  return {
    ach: ach.filter(item => item.included), fed: fed.filter(item => item.included),
    outside: all.filter(item => !item.inRange), excluded: all.filter(item => item.inRange && item.exclusion),
    review: all.filter(item => item.inRange && !item.exclusion && item.review), sums,
    dates: days, matches,
  };
}

function lendReviewedFedRows(fed, text, allowProduct = false) {
  const rows = new Map(fed.map(item => [String(item.row), item]));
  const decisions = [];
  for (const [index, line] of text.trim().split(/\r?\n/).entries()) {
    if (!line.trim()) continue;
    const [number, reason, productText = "", ...extra] = line.split("\t").map(value => value.trim());
    if (index === 0 && /^source row$/i.test(number) && /^reason$/i.test(reason ?? "")) continue;
    const item = rows.get(number);
    requireReport(item && reason && !extra.length && (allowProduct || !productText),
      "Review line " + (index + 1) + " needs a distinct eligible Fed source row and a reason"
        + (allowProduct ? ", with an optional product." : "."));
    const product = productText ? lendProducts.find(([, , aliases]) => aliases.includes(productText.toUpperCase()))?.[1] : "";
    requireReport(product !== undefined, "Review line " + (index + 1) + " product must be Consolidation, Home Improvement or Student.");
    decisions.push({ item, reason, product });
    rows.delete(number);
  }
  return decisions;
}

function lendReviewMissingProvider(result, text) {
  const matches = new Set(result.matches.map(match => match.fed));
  const unmatched = result.dates.flatMap(day => day.sameDate ? [] : day.fedRows
    .filter(item => item.amount && !matches.has(item)).map(item => ({ day, item })));
  const reviewed = lendReviewedFedRows(unmatched.map(entry => entry.item), text, true);
  result.unmatchedDays = new Map(unmatched.map(entry => [entry.item, entry.day]));
  for (const { item, reason, product } of reviewed) lendApplyProviderReview(result, item, reason, product);
  result.unmatchedFed = unmatched.map(entry => entry.item);
}

function lendApplyProviderReview(result, item, reason, product) {
  item.providerReview = reason;
  item.providerProduct = product;
  const day = result.unmatchedDays.get(item);
  day.reviewedProvider = (day.reviewedProvider ?? 0) + item.amount;
  const sums = result.sums[item.comparison];
  sums.reviewedProvider = (sums.reviewedProvider ?? 0) + item.amount;
  requireReport(Number.isSafeInteger(sums.reviewedProvider + sums.ach), "The reviewed provider total exceeds the supported range.");
  day.open--;
}

// The provider ACH export covers Home Improvement only, so Consolidation and Student payments appear only on the Fed.
// Each statement's "Remittance to Lender" total identifies them: assign the single combination of unmatched Fed
// credits (one to four) that equals a product's remittance. Rows the preparer already reviewed are left as entered.
function lendAutoReviewProvider(result, statementCash) {
  const assigned = [];
  for (const [code, product] of [["CNL", "Consolidation"], ["STU", "Student"]]) {
    const target = statementCash.get(code)?.get("remittance");
    if (!target) continue;
    const remittance = Math.abs(target);
    const entered = result.unmatchedFed.filter(item => item.providerProduct === product).reduce((sum, item) => sum + item.amount, 0);
    if (entered) continue;
    const candidates = result.unmatchedFed.filter(item => !item.providerReview && item.comparison === "Credit" && item.amount > 0);
    if (candidates.length > 40) continue;
    let found, ambiguous = false;
    (function search(start, remaining, selected) {
      if (ambiguous) return;
      if (remaining === 0 && selected.length) {
        if (found) ambiguous = true;
        else found = selected;
        return;
      }
      if (selected.length === 4) return;
      for (let i = start; i < candidates.length; i++) {
        if (candidates[i].amount <= remaining) search(i + 1, remaining - candidates[i].amount, [...selected, candidates[i]]);
      }
    })(0, remittance, []);
    if (!found || ambiguous) continue;
    for (const item of found) lendApplyProviderReview(result, item, code + " statement remittance to lender · assigned automatically", product);
    assigned.push(...found);
  }
  result.autoReviewed = assigned.length;
}

function lendLineupGroups(items) {
  const groups = new Map();
  for (const item of items) {
    // Remittances, reversals and refunds on the same date form one payment batch. The monthly workbook labels Home
    // Improvement batches Remittance or Remittance - Reversal and other products PMTS. Each funding is its own row.
    const payment = item.providerDetail && ["REMITTANCE", "REMITTANCE - REVERSAL", "REMITTANCE - REFUND"].includes(item.reason.toUpperCase());
    const reason = !payment ? item.reason
      : lendProducts[0][2].includes(item.product.toUpperCase()) ? (item.direction === "Credit" ? "Remittance" : "Remittance - Reversal") : "PMTS";
    const funding = ["FUNDING", "DISBURSEMENT"].includes(item.reason.toUpperCase());
    const key = item.providerDetail && !funding ? JSON.stringify([item.date, item.settlement, item.direction, item.product, payment ? "PMTS" : reason]) : item.row;
    const group = groups.get(key) ?? { ...item, reason, amount: 0, sourceRows: [] };
    if (group.status !== item.status) group.status = "Mixed · review";
    group.amount += item.amount;
    group.sourceRows.push(item.row);
    requireReport(Number.isSafeInteger(group.amount), "An Amounts Lineup subtotal exceeds the supported range.");
    groups.set(key, group);
  }
  return [...groups.values()];
}

function lendBatchMatches(days) {
  const open = { ach: [], fed: [] }, matches = [];
  for (const day of days) {
    day.achBatches = lendLineupGroups(day.achRows);
    day.sameDate = day.fedCount > 0 && day.achCount > 0 && day.fed === day.ach;
    day.open = day.sameDate ? 0 : day.fedCount + day.achBatches.length;
    if (day.sameDate) continue;
    for (const item of day.achBatches) if (item.amount) open.ach.push({ item, day });
    for (const item of day.fedRows) if (item.amount) open.fed.push({ item, day });
  }
  for (const sameDate of [true, false]) {
    const candidates = { ach: new Map(), fed: new Map() };
    for (const side of ["ach", "fed"]) {
      for (const entry of open[side]) {
        if (entry.match) continue;
        const key = JSON.stringify([sameDate ? entry.day.date : "", entry.day.direction, entry.item.amount]);
        const list = candidates[side].get(key) ?? [];
        list.push(entry);
        candidates[side].set(key, list);
      }
    }
    for (const [key, provider] of candidates.ach) {
      const fed = candidates.fed.get(key);
      if (provider.length !== 1 || fed?.length !== 1) continue;
      const [left] = provider, [right] = fed;
      const match = { provider: left.item, fed: right.item, direction: left.day.direction };
      left.match = right.match = match;
      left.day.open--;
      right.day.open--;
      matches.push(match);
    }
  }
  return matches;
}

const lendDayGap = (a, b) => Math.abs(Date.parse(a) - Date.parse(b)) / 86400000;

// Unique combination of two to four provider batches that totals the Fed amount, or undefined.
function lendUniqueSubset(candidates, target) {
  if (candidates.length > 40) return undefined;
  let found, ambiguous = false;
  (function search(start, remaining, selected) {
    if (ambiguous) return;
    if (remaining === 0 && selected.length >= 2) {
      if (found) ambiguous = true;
      else found = selected;
      return;
    }
    if (selected.length === 4) return;
    for (let i = start; i < candidates.length; i++) {
      if (candidates[i].amount <= remaining) search(i + 1, remaining - candidates[i].amount, [...selected, candidates[i]]);
    }
  })(0, target, []);
  return ambiguous ? undefined : found;
}

// Lines for the historical Amounts Lineup block: each line pairs Fed amounts with the LendKey batches they cover.
function lendLineupLines({ dates, matches }, reverse) {
  const byFed = new Map(matches.map(match => [match.fed, match.provider]));
  const used = new Set(matches.map(match => match.provider));
  const lines = [], pending = [];
  for (const day of dates) {
    for (const item of day.fedRows) {
      if (byFed.has(item)) lines.push({ fed: [item], provider: [byFed.get(item)] });
      else if (item.providerReview) {
        // LendKey remits a business day before the Fed settles; the monthly workbook dates these rows that way.
        let remitted = dateOffset(item.date, -1);
        while ([0, 6].includes(new Date(remitted + "T00:00:00Z").getUTCDay())) remitted = dateOffset(remitted, -1);
        lines.push({ fed: [item], provider: [{
          date: remitted, direction: reverse ? (item.direction === "Debit" ? "Credit" : "Debit") : item.direction,
          comparison: item.comparison, product: item.providerProduct, reason: item.providerProduct ? "PMTS" : "",
          note: item.providerReview, amount: item.amount,
        }] });
      } else if (item.amount) pending.push(item);
    }
  }
  let open = dates.flatMap(day => day.achBatches.filter(item => item.amount && !used.has(item)));
  const take = items => { open = open.filter(item => !items.includes(item)); };
  // Remaining Fed and LendKey rows on the same comparison date and direction with equal totals.
  for (const day of dates) {
    const fed = pending.filter(item => item.periodDate === day.date && item.comparison === day.direction);
    const provider = open.filter(item => item.periodDate === day.date && item.comparison === day.direction);
    const total = items => items.reduce((sum, item) => sum + item.amount, 0);
    if (!fed.length || !provider.length || total(fed) !== total(provider)) continue;
    lines.push({ fed, provider });
    fed.forEach(item => pending.splice(pending.indexOf(item), 1));
    take(provider);
  }
  // A Fed amount settles LendKey activity from the same or up to five earlier days. Try, in order: the closest single
  // batch with the same amount, then batches from one LendKey date that total the Fed amount (one wire funds that day's
  // loans), then a unique combination across those days.
  pending.sort((a, b) => a.periodDate.localeCompare(b.periodDate));
  for (const item of [...pending]) {
    const window = open.filter(batch => batch.comparison === item.comparison && batch.periodDate <= item.periodDate
      && lendDayGap(batch.periodDate, item.periodDate) <= 5);
    const gap = batch => lendDayGap(batch.periodDate, item.periodDate);
    const singles = window.filter(batch => batch.amount === item.amount).sort((a, b) => gap(a) - gap(b));
    let subset = singles.length && (singles.length === 1 || gap(singles[0]) < gap(singles[1])) ? [singles[0]] : undefined;
    if (!subset) {
      const solutions = [...new Set(window.map(batch => batch.periodDate))].flatMap(day => {
        const sameDay = window.filter(batch => batch.periodDate === day);
        if (sameDay.reduce((sum, batch) => sum + batch.amount, 0) === item.amount && sameDay.length > 1) return [sameDay];
        const found = sameDay.length <= 12 ? lendUniqueSubset(sameDay, item.amount) : undefined;
        return found ? [found] : [];
      });
      if (solutions.length === 1) [subset] = solutions;
    }
    subset ??= lendUniqueSubset(window, item.amount);
    if (!subset) continue;
    lines.push({ fed: [item], provider: subset });
    pending.splice(pending.indexOf(item), 1);
    take(subset);
  }
  for (const batch of open) lines.push({ fed: [], provider: [batch], note: "No Fed match · review" });
  for (const line of lines) {
    line.fed.sort((a, b) => a.date.localeCompare(b.date));
    line.provider.sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  }
  return { lines, unmatchedFed: pending };
}

// Historical order: debits (funding, then other activity, then rows without Fed), a blank row, then credits by product.
function lendLineupSections(lines) {
  const productRank = name => {
    const index = lendProducts.findIndex(([, , aliases]) => aliases.includes(String(name ?? "").toUpperCase()));
    return index < 0 ? lendProducts.length : index;
  };
  const lineDate = line => line.provider[0]?.date || line.fed[0]?.date || "";
  const direction = line => (line.fed[0] ?? line.provider[0]).comparison;
  // Ties on the same date follow the provider report's row order.
  const sourceRow = line => line.provider[0]?.row ?? Infinity;
  const keys = {
    Debit: line => [line.fed.length ? 0 : 1, String(line.provider[0]?.reason).toUpperCase() === "FUNDING" ? 0 : 1, lineDate(line), sourceRow(line)],
    Credit: line => [line.fed.length ? 0 : 1, productRank(line.provider[0]?.product), lineDate(line), sourceRow(line)],
  };
  const compare = key => (a, b) => {
    const left = key(a), right = key(b);
    for (let i = 0; i < left.length; i++) {
      if (left[i] !== right[i]) return typeof left[i] === "number" ? left[i] - right[i] : String(left[i]).localeCompare(String(right[i]));
    }
    return 0;
  };
  return ["Debit", "Credit"].map(name => lines.filter(line => direction(line) === name).sort(compare(keys[name])));
}

// Home Improvement late fees, calculated the same way as the provider workbook's Recon Adj Wkbook!AI21
// (Servicing Transaction Net, Late Fee). The provider saves that formula without a calculated value.
async function readLendServicingLateFees(book) {
  const sheet = "ServicingTransactionDetail";
  if (!book?.sheets.has(sheet)) return undefined;
  const cells = await readCells(book, sheet);
  const wanted = ["Late_Fee_Component", "Transaction_Type", "Transaction_Amount"];
  const headers = new Map();
  for (const [address, value] of cells) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (!match || Number(match[2]) > 50 || !wanted.includes(String(value ?? "").trim())) continue;
    const row = Number(match[2]), columns = headers.get(row) ?? {};
    columns[String(value).trim()] = match[1];
    headers.set(row, columns);
  }
  const found = [...headers].find(([, columns]) => wanted.every(name => columns[name]));
  if (!found) return undefined;
  const [headerRow, columns] = found;
  const rows = new Set();
  for (const address of cells.keys()) {
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
    if (match && match[1] === columns.Transaction_Type && Number(match[2]) > headerRow) rows.add(Number(match[2]));
  }
  let credits = 0, debits = 0;
  for (const row of rows) {
    const type = String(cells.get(columns.Transaction_Type + row) ?? "").trim().toLowerCase();
    const fee = cells.get(columns.Late_Fee_Component + row);
    if (fee === undefined || fee === "" || !type) continue;
    const amount = cents(fee), sign = Math.sign(readBalance(cells.get(columns.Transaction_Amount + row) ?? 0));
    if ((type === "regular payment" && sign < 0) || type === "claim payment") credits -= amount;
    else if ((type === "regular payment" && sign > 0) || type === "interest refund" || type === "principal claim refund") debits += amount;
  }
  return { amount: credits - debits, source: sheet + " late fees (Recon Adj Wkbook!AI21 calculation)" };
}

// Static summary text (journal entries and account notes) from the last Purpose Matchup block in the prepared workbook.
function lendPriorPurposeLabels(cells) {
  let start;
  for (const [address, value] of cells) {
    const match = /^C(\d+)$/.exec(address);
    if (match && String(value ?? "").trim().toUpperCase() === "A/R CLEARED") start = Math.max(start ?? 0, Number(match[1]));
  }
  if (!start) return undefined;
  return [0, 1, 2, 3, 4].map(offset => new Map(["E", "G", "H", "I", "K", "M"].flatMap(column => {
    const value = cells.get(column + (start + offset));
    return value === undefined || value === "" ? [] : [[column, value]];
  })));
}

// The monthly Purpose Matchup block: month-end date, then Consolidation, Home Improvement and Student sections, then the summary.
function lendPurposeBlock(result, lines, { label, number, date, formula }) {
  const grid = new Map();
  const put = (r, column, xml) => {
    if (!grid.has(r)) grid.set(r, new Map());
    grid.get(r).set(column, xml);
  };
  const text = (r, column, value) => {
    if (value !== undefined && value !== null && value !== "") put(r, column, typeof value === "number" ? number(column + r, value) : label(column + r, value));
  };
  const money = (r, column, amount) => { if (amount != null) put(r, column, number(column + r, amount / 100)); };
  const total = items => items.reduce((sum, item) => sum + item.amount, 0);
  const upper = value => String(value ?? "").toUpperCase();
  const codeOf = name => lendProducts.find(([, , aliases]) => aliases.includes(upper(name)))?.[0];
  const range = (column, rows) => rows.length ? `${column}${rows[0]}:${column}${rows.at(-1)}` : "";
  const [year, monthNumber] = result.month.split("-").map(Number);
  put(1, "C", date("C1", new Date(Date.UTC(year, monthNumber, 0)).toISOString().slice(0, 10)));
  const items = lines.flatMap(line => line.provider);
  const groups = new Map(result.purpose.groups.map(group => [group.code, group]));
  const cleared = { ar: 0, ap: 0 }, inTransit = { ar: [], ap: [] };
  let r = 2;
  for (const code of ["CNL", "HIL", "STU"]) {
    const group = groups.get(code), cash = result.statementCash?.get(code) ?? new Map();
    const rows = items.filter(item => codeOf(item.product) === code)
      .sort((a, b) => (a.date || "").localeCompare(b.date || "") || (a.row ?? 1e9) - (b.row ?? 1e9));
    // The prior-period settlement on the statement identifies the batch that clears last month's A/R or A/P. The
    // settlement can differ by cents from the batch (August 2026: 19,017.75 vs 19,017.38), so take the single
    // closest payment batch within one dollar.
    const clearing = (target, direction) => {
      if (!target) return undefined;
      const gap = item => Math.abs(item.amount - Math.abs(target));
      const found = rows.filter(item => item.comparison === direction && upper(item.reason) !== "FUNDING" && gap(item) <= 100)
        .sort((a, b) => gap(a) - gap(b));
      return found.length && (found.length === 1 || gap(found[0]) < gap(found[1])) ? found[0] : undefined;
    };
    const arItem = clearing(cash.get("priorAR") ?? group?.inputs["prior-ar"], "Credit");
    const apItem = clearing(cash.get("priorAP") ?? group?.inputs["prior-ap"], "Debit");
    const rest = rows.filter(item => item !== arItem && item !== apItem);
    const funding = rest.filter(item => item.comparison === "Debit" && upper(item.reason) === "FUNDING");
    const otherDebits = rest.filter(item => item.comparison === "Debit" && !funding.includes(item));
    const credits = rest.filter(item => item.comparison !== "Debit");
    const write = (item, note) => {
      if (item.date) put(r, "C", date("C" + r, item.date));
      text(r, "D", item.direction);
      text(r, "E", item.product);
      text(r, "G", item.reason);
      money(r, "J", item.amount);
      text(r, "K", note);
      return r++;
    };
    if (arItem) { write(arItem, "Cleared"); cleared.ar += arItem.amount; }
    const fundingRows = funding.map(item => write(item));
    const apRow = apItem ? write(apItem, "Cleared") : undefined;
    if (apItem) cleared.ap += apItem.amount;
    const otherRows = otherDebits.map(item => write(item));
    const creditRows = credits.map(item => write(item));
    const newAR = cash.get("newAR") ?? (group?.inputs["new-ar"] == null ? undefined : -group.inputs["new-ar"]);
    const newAP = cash.get("newAP") ?? group?.inputs["new-ap"] ?? undefined;
    const arRow = r++, apTransitRow = r++;
    text(arRow, "H", "A/R in Transit");
    money(arRow, "J", newAR);
    text(arRow, "K", "NEW A/R");
    text(apTransitRow, "H", "A/P in Transit");
    money(apTransitRow, "J", newAP);
    text(apTransitRow, "K", "NEW A/P");
    inTransit.ar.push([arRow, newAR ?? 0]);
    inTransit.ap.push([apTransitRow, newAP ?? 0]);
    // Remittance tie-out: transactions in the section against the statement's remittance line.
    const remittance = cash.has("remittance") ? Math.abs(cash.get("remittance")) : undefined;
    const remitted = total(credits);
    text(arRow, "L", code === "HIL" ? "PMT REMITTANCE" : code === "CNL" ? "remittance to lender" : "");
    text(arRow, "N", "PDF");
    text(apTransitRow, "L", "TOTAL");
    put(apTransitRow, "M", creditRows.length ? formula("M" + apTransitRow, `SUM(${range("J", creditRows)})`, remitted) : number("M" + apTransitRow, 0));
    money(apTransitRow, "N", remittance);
    if (code !== "STU" && remittance !== undefined) {
      put(r, "M", formula("M" + r, `ROUND(M${apTransitRow}-N${apTransitRow},2)`, remitted - remittance));
      if (code === "CNL") r++;
    }
    if (code !== "HIL") continue;
    // Funding tie-out against statement disbursements.
    if (fundingRows.length >= 5) {
      const [, , third, fourth, fifth] = fundingRows;
      const disbursed = lendStatementInput(result.statements, result.portfolio, "HIL", "Loan Principal Recs", 9)?.amount ?? group?.statementFunding;
      const funded = total(funding);
      put(third, "M", formula("M" + third, `K${fifth}-L${fifth}`, disbursed == null ? funded : funded - disbursed));
      text(fourth, "K", "Funding:");
      text(fourth, "L", "Disbursed:");
      put(fifth, "K", formula("K" + fifth, `SUM(${range("J", fundingRows)})`, funded));
      money(fifth, "L", disbursed);
    }
    // Reversal tie-out: payment reversals after the cleared A/P against the statement's remittance reversal line.
    const reversalStart = apRow ?? otherRows[0];
    if (fundingRows.length && reversalStart) {
      const reversal = cash.has("reversal") ? Math.abs(cash.get("reversal")) : undefined;
      const reversed = total(otherDebits);
      if (fundingRows.length !== 5) {
        text(fundingRows.at(-1), "L", "PMT REMITTANCE REVERSAL");
        text(fundingRows.at(-1), "N", "PDF");
      }
      text(reversalStart, "L", "TOTAL");
      put(reversalStart, "M", otherRows.length ? formula("M" + reversalStart, `SUM(${range("J", otherRows)})`, reversed) : number("M" + reversalStart, 0));
      money(reversalStart, "N", reversal);
      if (reversal !== undefined) {
        put(reversalStart + 1, "M", formula("M" + (reversalStart + 1), `ROUND(M${reversalStart}-N${reversalStart},2)`, reversed - reversal));
      }
    }
  }
  const s = r, prior = result.priorPurpose;
  const defaults = [
    new Map([["H", "DEBIT"], ["I", "CREDIT"], ["K", "LATE FEES+REVERSALS"]]), new Map([["G", "A/R Cleared"], ["K", "HIL"]]),
    new Map([["G", "A/P Cleared"], ["K", "CNL"]]), new Map([["G", "Late Fees"], ["K", "PSL"]]), new Map([["G", "New A/R"]]),
  ];
  ["A/R CLEARED", "A/P CLEARED", "LATE FEES", "NEW A/R", "NEW A/P"].forEach((title, i) => {
    text(s + i, "C", title);
    for (const [column, value] of prior?.[i] ?? defaults[i]) text(s + i, column, value);
  });
  money(s, "D", cleared.ar);
  money(s + 1, "D", cleared.ap);
  const lateFees = ["HIL", "CNL", "STU"].map(code => groups.get(code)?.inputs["late-fees"] ?? 0);
  lateFees.forEach((amount, i) => money(s + 1 + i, "L", amount));
  put(s + 2, "D", formula("D" + (s + 2), `SUM(L${s + 1}:L${s + 3})`, lateFees.reduce((sum, amount) => sum + (amount ?? 0), 0)));
  for (const [offset, entries] of [[3, inTransit.ar], [4, inTransit.ap]]) {
    put(s + offset, "D", formula("D" + (s + offset), `SUM(${entries.map(([row]) => "J" + row).join(",")})`,
      entries.reduce((sum, [, amount]) => sum + amount, 0)));
  }
  return { grid, last: s + 4 };
}

function lendClearingCandidate(items, target) {
  if (!target) return { status: target === 0 ? "No opening balance" : "Prior amount unavailable", items: [] };
  const eligible = items.filter(item => item.amount > 0 && item.amount <= target);
  if (eligible.length > 120) return { status: "More than 120 eligible rows · review manually", items: [] };
  let found;
  function search(start, remaining, selected) {
    if (remaining === 0) {
      if (found) return false;
      found = selected;
      return true;
    }
    if (selected.length === 3) return true;
    for (let i = start; i < eligible.length; i++) {
      const item = eligible[i];
      if (item.amount <= remaining && !search(i + 1, remaining - item.amount, [...selected, item])) return false;
    }
    return true;
  }
  if (!search(0, target, [])) return { status: "Multiple exact candidates", items: [] };
  if (!found) return { status: "No exact candidate within three rows", items: [] };
  return { status: "Exact candidate · verify", items: found };
}

function buildLendPurpose(ach, performance) {
  const classify = (reason, direction) => {
    const key = reason.toUpperCase();
    if (["PMTS", "PAYMENT", "PAYMENTS", "REMITTANCE"].includes(key)) return direction === "Debit" ? "Payment reversal" : "Payment";
    if (key === "REMITTANCE - REVERSAL") return "Payment reversal";
    if (key === "REMITTANCE - REFUND") return "Payment refund";
    if (["FUNDING", "DISBURSEMENT", "DISBURSEMENTS"].includes(key)) return "Funding";
    if (["CANCELLATION", "CANCELLATIONS"].includes(key)) return "Cancellation";
    if (key === "INVOICE") return "Invoice";
    if (key === "CASH ADJUSTMENT") return "Cash adjustment";
    return "Review purpose";
  };
  const lines = ach.ach.map(item => {
    const product = lendProducts.find(([, , aliases]) => aliases.includes(item.product.toUpperCase()));
    return { ...item, code: product?.[0], name: product?.[1] ?? item.product, purpose: classify(item.reason, item.direction) };
  });
  const groups = lendProducts.map(([code, name]) => {
    const rows = lines.filter(item => item.code === code);
    const inputs = Object.fromEntries(lendPurposeFields.map(([key, label]) => {
      const raw = lendField(code + "-" + key).value.trim();
      const amount = raw ? cents(raw) : null;
      requireReport(amount === null || amount >= 0, name + " " + label + " needs a positive amount or zero.");
      return [key, amount];
    }));
    const candidates = [["A/R", "Credit", "Payment", inputs["prior-ar"]], ["A/P", "Debit", "Payment reversal", inputs["prior-ap"]]]
      .map(([type, direction, purpose, target]) => {
        const result = lendClearingCandidate(rows.filter(item => item.direction === direction && item.purpose === purpose), target);
        for (const item of result.items) item.clearing = type + " clearing candidate";
        return { type, target, ...result };
      });
    const funding = rows.filter(item => item.direction === "Debit" && item.purpose === "Funding")
      .reduce((sum, item) => sum + item.amount, 0);
    const statementFunding = code === "HIL" ? performance?.[0].entries.find(item => item.product === code)?.values.get(9) : undefined;
    return { code, name, rows, inputs, candidates, funding, statementFunding };
  });
  return { groups, unknown: lines.filter(item => !item.code) };
}

async function readLendPortfolio(workbook, month) {
  requireReport(workbook.sheets.has(lendPortfolioSheet), "Choose the LendKey workbook with its Recon Adj Wkbook worksheet.");
  const cells = await readCells(workbook, lendPortfolioSheet);
  const title = String(cells.get("C3") ?? "");
  requireReport(/\bHome Improvement Loan Portfolio Summary\b/i.test(title), "The provider roll-forward currently supports the Home Improvement portfolio summary.");
  const [, reportMonth, day, year] = /\bperiod\s+ending\s+([a-z]+)\s+(\d{1,2}),?\s+(\d{4})\b/i.exec(title) ?? [];
  const [selectedYear, selectedMonth] = month.split("-").map(Number);
  const end = new Date(Date.UTC(selectedYear, selectedMonth, 0));
  const monthName = new Intl.DateTimeFormat("en-US", { month: "long", timeZone: "UTC" }).format(end);
  requireReport(reportMonth?.toLowerCase() === monthName.toLowerCase() && Number(year) === selectedYear && Number(day) === end.getUTCDate(),
    "The provider portfolio summary does not show the selected month-end.");
  return new Map(lendSchedules.map((schedule, index) => [schedule.name,
    new Map(lendPortfolioFields[index].map(([row, labelCell, label, sign = 1]) => {
      requireReport(String(cells.get(labelCell) ?? "").trim().toLowerCase() === label.toLowerCase(),
        "Check the provider statement label at " + lendPortfolioSheet + "!" + labelCell + ".");
      const address = (index === 0 ? "I" : "N") + labelCell.slice(1);
      let amount;
      try { amount = cents(cells.get(address)) * sign; }
      catch { throw new Error("A numeric provider total is required at " + lendPortfolioSheet + "!" + address + "."); }
      return [row, { amount, source: lendPortfolioSheet + "!" + address + (sign === -1 ? " · sign reversed" : "") }];
    })),
  ]));
}

function lendStatementInput(statements, portfolio, product, schedule, row) {
  return statements?.get(product)?.get(schedule)?.get(row) ?? (product === "HIL" ? portfolio?.get(schedule)?.get(row) : undefined);
}

async function readLendPerformance(workbook, portfolio, statements) {
  return Promise.all(lendSchedules.map(async schedule => {
    const cells = workbook ? await readCells(workbook, schedule.name) : undefined;
    for (const [row, label] of schedule.fields) {
      requireReport(!cells || String(cells.get("B" + row) ?? "").trim().toLowerCase() === label.toLowerCase(),
        schedule.name + "!B" + row + " does not match the expected statement line.");
    }
    const entries = ["D", "F", "H"].flatMap((column, index) => {
      const product = ["CNL", "HIL", "STU"][index];
      requireReport(!cells || String(cells.get(column + "4") ?? "").trim().toUpperCase() === product,
        "Check the product heading at " + schedule.name + "!" + column + "4.");
      const sourceInput = row => lendStatementInput(statements, portfolio, product, schedule.name, row);
      if (!cells && schedule.fields.some(([row]) => !sourceInput(row))) return [];
      const values = new Map(schedule.fields.map(([row]) => {
        const provider = sourceInput(row);
        if (provider) return [row, provider.amount];
        const address = column + row;
        try { return [row, cents(cells.get(address))]; }
        catch { throw new Error("Enter a numeric statement amount at " + schedule.name + "!" + address + ". Use zero when it does not apply."); }
      }));
      const sum = rows => rows.reduce((total, row) => total + values.get(row), 0);
      const plus = sum(schedule.plus), less = sum(schedule.less);
      const end = values.get(5) + values.get(6) + plus - less + sum(schedule.other);
      const change = end - values.get(5), difference = end - values.get(schedule.statement);
      requireReport([plus, less, end, change, difference].every(Number.isSafeInteger), "A roll-forward exceeds the supported amount range.");
      return { column, product, values, plus, less, end, change, difference };
    });
    return { ...schedule, entries };
  }));
}

async function readLendCover(workbook, performance, month) {
  const coverFormulas = new Map(), entryFormulas = new Set();
  const [, entries] = await Promise.all([
    readCells(workbook, "Cover Page", coverFormulas),
    readCells(workbook, "CUSA Entries", entryFormulas),
  ]);
  const adjustmentFormula = /^SUMIF\('CUSA ENTRIES'!C5:C32,(\d{6}),'CUSA ENTRIES'!K5:K32\)\+SUMIF\('CUSA ENTRIES'!E5:E32,\1,'CUSA ENTRIES'!J5:J32\)$/;
  const cover = performance.flatMap((schedule, index) => schedule.entries.map(entry => {
    const coverColumn = { D: "K", F: "V", H: "AG" }[entry.column];
    const [statementRow, ledgerRow, adjustmentRow, adjustedRow, differenceRow] = index === 0 ? [7, 10, 11, 12, 14] : [17, 19, 20, 21, 23];
    const statementCell = coverColumn + statementRow, ledgerCell = coverColumn + ledgerRow;
    const adjustmentCell = coverColumn + adjustmentRow, adjustedCell = coverColumn + adjustedRow, differenceCell = coverColumn + differenceRow;
    const formulaAt = address => String(coverFormulas.get(address) ?? "").replace(/\$/g, "").toUpperCase();
    const account = adjustmentFormula.exec(formulaAt(adjustmentCell))?.[1];
    const addition = [ledgerCell + "+" + adjustmentCell, ledgerCell + "+SUM(" + adjustmentCell + ":" + adjustmentCell + ")"];
    requireReport(account && !coverFormulas.has(ledgerCell)
      && formulaAt(statementCell) === ("'" + schedule.name + "'!" + entry.column + schedule.end).toUpperCase()
      && addition.includes(formulaAt(adjustedCell)) && formulaAt(differenceCell) === statementCell + "-" + adjustedCell,
    "The Cover Page layout differs at " + adjustmentCell + ". Use the supported monthly template.");
    let adjustment = 0;
    for (let row = 5; row <= 32; row++) {
      for (const [accountColumn, amountColumn] of [["C", "K"], ["E", "J"]]) {
        if (String(entries.get(accountColumn + row) ?? "").trim() !== account) continue;
        const address = amountColumn + row, amount = entries.get(address);
        requireReport(amount !== undefined || !entryFormulas.has(address), "Recalculate and save the reviewed CUSA Entries before importing. Missing result at " + address + ".");
        requireReport(amount === undefined || typeof amount === "number", "Use a numeric CUSA adjustment at " + address + ".");
        adjustment += amount === undefined ? 0 : cents(amount);
      }
    }
    return { product: entry.product, metric: index === 0 ? "Principal" : "Accrued interest",
      account, ledgerCell, statement: entry.end, adjustment };
  }));
  const accounts = cover.map(item => item.account);
  requireReport(new Set(accounts).size === accounts.length, "The Cover Page must identify six distinct G/L accounts.");
  const sheet = lendField("wisdom-sheet").value, last = lendField("wisdom-last");
  const balances = await readWisdom(lendSources.wisdom.book, sheet, lendField("wisdom-account").value.trim().toUpperCase(),
    lendField("wisdom-balance").value.trim().toUpperCase(), lendField("wisdom-first").valueAsNumber, accounts,
    last.value ? last.valueAsNumber : 1048576, true, month);
  return cover.map(item => {
    const match = balances.get(item.account);
    const ledger = cents(match.amount), adjusted = ledger + item.adjustment, difference = item.statement - adjusted;
    requireReport([item.adjustment, adjusted, difference].every(Number.isSafeInteger), "A cover balance exceeds the supported amount range.");
    return { ...item, ledger, adjusted, difference, source: sheet + "!" + match.address };
  });
}

function showLend(result) {
  const money = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
  const summary = [];
  if (result.ach) {
    for (const [direction, values] of Object.entries(result.ach.sums)) {
      summary.push([direction + " difference after review", money.format((values.fed - values.ach - (values.reviewedProvider ?? 0)) / 100)]);
    }
    summary.push(["Date and direction groups to review", result.ach.dates.filter(day => day.open).length]);
    summary.push(["Exact batch matches across dates", result.ach.matches.filter(match => match.provider.periodDate !== match.fed.periodDate).length]);
    summary.push(["Rows outside selected dates", result.ach.outside.length]);
    summary.push(["Fed rows excluded for another month's LendKey date", result.ach.excluded.filter(item => item.autoExcluded).length]);
    summary.push(["Fed credits assigned to Consolidation or Student from statement remittances", result.ach.autoReviewed ?? 0]);
    summary.push(["Rows excluded by reviewer", result.ach.excluded.filter(item => !item.autoExcluded).length]);
    if (result.ach.review.length) summary.push(["Rows with missing ACH settlement details", result.ach.review.length]);
    summary.push(["Prior carryovers with an exact candidate", result.purpose.groups.flatMap(group => group.candidates).filter(item => item.items.length).length]);
  }
  if (result.performance) {
    summary.push(["Statement roll-forwards", result.performance.flatMap(schedule => schedule.entries).length]);
    summary.push(["Statement variances", result.performance.flatMap(schedule => schedule.entries).filter(item => item.difference !== 0).length]);
  }
  if (result.portfolio) summary.push(["Home Improvement inputs read from provider", [...result.portfolio.values()].reduce((total, rows) => total + rows.size, 0)]);
  if (result.statements.size) summary.push(["Statement PDF portfolios", [...result.statements.keys()].join(", ")]);
  if (result.cover) summary.push(["Cover balances with a difference", result.cover.filter(item => item.difference !== 0).length]);
  const table = document.createElement("table"), body = table.createTBody();
  for (const [name, value] of summary) {
    const row = body.insertRow(), heading = document.createElement("th");
    heading.scope = "row";
    heading.textContent = name;
    row.append(heading);
    row.insertCell().textContent = value;
  }
  lendField("summary").replaceChildren(table);
  lendField("results").hidden = false;
}

async function buildLend(result) {
  const audit = reconciliationAudit();
  const { label, number, date, row, worksheet } = excelOutputCells();
  const formula = (address, expression, value) => `<c r="${address}" s="2"><f>${xmlText(expression)}</f><v>${value / 100}</v></c>`;
  const value = (address, item) => typeof item === "number" ? number(address, item) : label(address, item);
  function table(name, headers, entries, widths) {
    const columns = headers.map((_, i) => csvColumn(i + 1));
    const body = row(1, headers.map((header, i) => label(columns[i] + "1", header, 1)).join(""))
      + entries.map((entry, i) => row(i + 2, entry.map((item, j) => value(columns[j] + (i + 2), item)).join(""))).join("");
    return [name, worksheet("A1:" + columns.at(-1) + (entries.length + 1), widths.map((width, i) => [i + 1, i + 1, width]), body)];
  }
  const sheets = [];
  const summary = [["Period", result.month]];
  if (result.ach) {
    const { ach, fed, outside, excluded, review, sums, dates } = result.ach;
    summary.push(["Fed date range", result.from + " to " + result.through],
      ["Provider comparison direction", result.reverse ? "Reversed" : "As reported"],
      ["LendKey export", result.achFile], ["Fed export", result.fedFile]);
    for (const [direction, amounts] of Object.entries(sums)) {
      summary.push(["Fed " + direction.toLowerCase(), amounts.fed / 100],
        ["LendKey " + direction.toLowerCase() + " on Fed basis", amounts.ach / 100],
        ["Reviewed Fed items added to " + direction.toLowerCase(), (amounts.reviewedProvider ?? 0) / 100],
        [direction + " difference after review", (amounts.fed - amounts.ach - (amounts.reviewedProvider ?? 0)) / 100]);
    }
    // Same block layout as the monthly workbook's Amounts Lineup, so rows 1 through the month-end date paste directly.
    const lineup = [
      row(1, label("C1", "ACH Transactions") + label("L1", "Lendkey Transactions")),
      row(2, [["E", "DR/CR"], ["H", "Amount"], ["L", "Date"], ["M", "DR/CR"], ["P", "Description"], ["R", "Amount"]]
        .map(([col, title]) => label(col + "2", title)).join("")),
    ];
    const text = (address, item) => item ? label(address, item) : "";
    const directionCode = direction => direction === "Debit" ? "DR" : "CR";
    const { lines, unmatchedFed } = lendLineupLines(result.ach, result.reverse);
    let r = 3;
    lendLineupSections(lines).forEach((section, index) => {
      if (index && section.length) r++;
      for (const line of section) {
        const first = r, count = Math.max(line.fed.length, line.provider.length);
        const difference = line.fed.reduce((sum, item) => sum + item.amount, 0)
          - line.provider.reduce((sum, item) => sum + item.amount, 0);
        for (let i = 0; i < count; i++) {
          const left = line.fed[i], right = line.provider[i];
          const fedCells = left ? label("E" + r, directionCode(left.comparison)) + number("H" + r, left.amount / 100) : "";
          const check = i === 0 && line.fed.length
            ? formula("K" + r, `ROUND(SUM(H${first}:H${first + count - 1})-SUM(S${first}:S${first + count - 1}),2)`, difference) : "";
          const achCells = right ? (right.date ? date("L" + r, right.date) : "") + text("M" + r, right.direction)
            + text("N" + r, right.product) + text("P" + r, right.reason) + number("S" + r, right.amount / 100)
            + text("T" + r, right.note ?? (i === 0 ? line.note : "")) : "";
          lineup.push(row(r++, fedCells + check + achCells));
        }
      }
    });
    const [year, monthNumber] = result.month.split("-").map(Number);
    lineup.push(row(r, date("C" + r, new Date(Date.UTC(year, monthNumber, 0)).toISOString().slice(0, 10))));
    const blockEnd = r;
    if (unmatchedFed.length) {
      r += 4;
      lineup.push(row(r, label("C" + r, "Unmatched Fed items · review before pasting. Not part of the block above.", 1)));
      r++;
      for (const item of unmatchedFed) {
        lineup.push(row(r, date("C" + r, item.date) + label("E" + r, directionCode(item.comparison)) + number("H" + r, item.amount / 100)
          + text("J" + r, item.reference) + label("T" + r, "Fed row " + item.row)));
        r++;
      }
    }
    summary.push(["Amounts Lineup paste block", "C1:T" + blockEnd]);
    sheets.push(["Amounts Lineup", worksheet("C1:T" + (unmatchedFed.length ? r - 1 : blockEnd), [[3, 20, 14]], lineup.join(""))]);
    sheets.push(table("ACH date review", ["Comparison date", "Direction", "Fed rows", "LendKey rows", "Fed amount", "LendKey amount", "Reviewed Fed items", "Difference after review", "Review"],
      dates.map(day => [day.date, day.direction, day.fedCount, day.achCount, day.fed / 100, day.ach / 100, (day.reviewedProvider ?? 0) / 100,
        (day.fed - day.ach - (day.reviewedProvider ?? 0)) / 100, day.sameDate ? "Date totals agree"
          : day.open ? day.open + " unmatched batches or transactions" : day.reviewedProvider ? "Includes reviewer decisions" : "Matched by amount and direction"]),
      [19, 16, 14, 15, 18, 18, 21, 24, 42]));
    if (result.ach.unmatchedFed.length) sheets.push(table("Unmatched Fed items",
      ["Fed row", "Effective date", "Posted on", "Direction", "Amount", "Reference", "Reviewer reason"],
      result.ach.unmatchedFed.map(item => [item.row, item.date, item.posted ?? "", item.comparison, item.amount / 100,
        item.reference, item.providerReview ?? ""]), [12, 19, 16, 15, 17, 45, 52]));
    sheets.push(table("ACH source review", ["Source", "Row", "Transaction date", "Provider ACH date", "Type", "Compared as", "Product", "Reason", "Status", "Reference", "Source amount", "Selection", "Review reason", "Fed comparison date", "Confirmed Fed date", "Finance provider", "Wisdom posted on"],
      [...fed, ...ach, ...outside, ...excluded, ...review].map(item => [item.kind === "fed" ? "Fed" : "LendKey", item.row, item.date, item.settlement,
        item.direction, item.comparison, item.product, item.reason, item.status, item.reference, item.sourceAmount / 100,
        item.included ? "Included" : !item.inRange ? "Outside selected dates"
          : item.autoExcluded ? "Excluded automatically" : item.exclusion ? "Excluded by reviewer" : "Review settlement",
        [item.exclusion, item.review].filter(Boolean).join(" · "), item.periodDate, item.override, item.financeProvider, item.posted ?? ""]),
      [15, 10, 18, 18, 13, 15, 25, 24, 25, 19, 18, 24, 34, 20, 20, 24, 20]));
    if (result.ach.matches.length) sheets.push(table("ACH batch matches",
      ["Direction", "Amount", "Provider ACH date", "Provider comparison date", "Fed date", "Provider rows", "Fed row", "Fed reference"],
      result.ach.matches.map(({ provider, fed, direction }) => [direction, provider.amount / 100, provider.settlement,
        provider.periodDate, fed.date, provider.sourceRows.join(", "), fed.row, fed.reference]), [16, 18, 21, 25, 17, 37, 13, 24]));
    const { groups, unknown } = result.purpose;
    // Same block layout as the monthly workbook's Purpose Matchup: batch-level rows by product with the statement tie-outs.
    const purposeBlock = lendPurposeBlock(result, lines, { label, number, date, formula });
    const purposeRows = [...purposeBlock.grid].sort((a, b) => a[0] - b[0]).map(([r, cells]) =>
      row(r, [...cells].sort((a, b) => columnNumber(a[0]) - columnNumber(b[0])).map(([, xml]) => xml).join("")));
    summary.push(["Purpose Matchup paste block", "C1:N" + purposeBlock.last]);
    const hilLateFees = result.purpose.groups.find(group => group.code === "HIL")?.inputs["late-fees"];
    summary.push(["HIL late fees", hilLateFees == null ? "" : hilLateFees / 100],
      ["HIL late fees source", result.lateFees ? result.lateFees.source : hilLateFees == null ? "Not entered" : "Entered on the page"]);
    sheets.push(["Purpose Matchup", worksheet("C1:N" + purposeBlock.last, [[3, 4, 14], [5, 5, 18], [6, 7, 14], [8, 8, 16], [9, 14, 14]], purposeRows.join(""))]);
    const dollar = amount => amount == null ? "" : amount / 100;
    sheets.push(table("Carryover review", ["Product", "Type", "Prior amount", "New statement amount", "Suggested ACH rows", "Status"],
      groups.flatMap(group => group.candidates.map(item => [group.name, item.type, dollar(item.target),
        dollar(group.inputs[item.type === "A/R" ? "new-ar" : "new-ap"]), item.items.map(source => source.row).join(", "), item.status])),
      [24, 12, 18, 24, 30, 43]));
    const purposeDetails = [...groups.flatMap(group => group.rows), ...unknown];
    const cashGroups = new Map();
    for (const item of purposeDetails) {
      const key = JSON.stringify([item.name, item.direction, item.purpose]);
      const group = cashGroups.get(key) ?? { count: 0, amount: 0 };
      group.count++;
      group.amount += item.amount;
      requireReport(Number.isSafeInteger(group.amount), "A purpose subtotal exceeds the supported amount range.");
      cashGroups.set(key, group);
    }
    sheets.push(table("Purpose totals", ["Product", "Direction", "Purpose", "Rows", "ACH amount"],
      [...cashGroups].map(([key, group]) => [...JSON.parse(key), group.count, group.amount / 100]), [24, 15, 23, 12, 18]));
    sheets.push(table("Purpose source review", ["ACH source row", "Product", "Transaction date", "Comparison date", "Type", "Reason", "Source status", "Amount", "Purpose", "Clearing candidate"],
      purposeDetails.map(item => [item.row, item.name, item.date, item.periodDate, item.direction, item.reason,
        item.status, item.amount / 100, item.purpose, item.clearing ?? ""]), [17, 24, 18, 18, 12, 22, 20, 17, 20, 29]));
    const hil = groups[0];
    summary.push(["HIL ACH funding debits", dollar(hil.funding)], ["HIL statement disbursements", dollar(hil.statementFunding)],
      ["HIL funding less statement", dollar(hil.statementFunding == null ? null : hil.funding - hil.statementFunding)]);
  }
  if (result.performance) {
    summary.push(["Statement input workbook", result.performanceFile]);
    for (const schedule of result.performance) {
      const rows = [row(2, label("B2", schedule.name, 1)),
        row(4, schedule.entries.map(entry => label(entry.column + "4", entry.product, 1)).join(""))];
      const data = new Map(schedule.fields.map(([r, name]) => [r, label("B" + r, name)
        + schedule.entries.map(entry => number(entry.column + r, entry.values.get(r) / 100)).join("")]));
      const formulaRows = [
        [schedule.totals[0], "Total additions", entry => "SUM(" + schedule.plus.map(r => entry.column + r).join(",") + ")", "plus"],
        [schedule.totals[1], "Total deductions", entry => "SUM(" + schedule.less.map(r => entry.column + r).join(",") + ")", "less"],
        [schedule.end, "Calculated ending balance", ({ column }) => "ROUND(" + column + "5+" + column + "6+" + column + schedule.totals[0]
          + "-" + column + schedule.totals[1] + schedule.other.map(r => "+" + column + r).join("") + ",2)", "end"],
        [schedule.change, "Change from opening balance", ({ column }) => "ROUND(" + column + schedule.end + "-" + column + "5,2)", "change"],
        [schedule.difference, "Calculated less statement", ({ column }) => "ROUND(" + column + schedule.end + "-" + column + schedule.statement + ",2)", "difference"],
      ];
      for (const [r, title, expression, key] of formulaRows) {
        data.set(r, label("B" + r, title) + schedule.entries.map(entry => formula(entry.column + r, expression(entry), entry[key])).join(""));
      }
      for (const [r, cells] of [...data].sort((a, b) => a[0] - b[0])) rows.push(row(r, cells));
      sheets.push([schedule.name, worksheet("B2:H" + schedule.difference, [[2, 2, 42], [3, 8, 17]], rows.join(""))]);
      for (const entry of schedule.entries) summary.push([entry.product + " · " + schedule.name + " variance", entry.difference / 100]);
    }
  }
  if (result.performance || result.portfolio || result.statements.size) {
    const paste = lendSchedules.flatMap((schedule, index) => schedule.fields.flatMap(([r, field]) => {
      const prepared = result.performance?.[index].entries ?? [];
      return ["CNL", "HIL", "STU"].flatMap((product, i) => {
        const entry = prepared.find(item => item.product === product);
        if (!entry && !result.statements.has(product) && !(product === "HIL" && result.portfolio)) return [];
        const provider = lendStatementInput(result.statements, result.portfolio, product, schedule.name, r);
        const amount = provider?.amount ?? entry?.values.get(r);
        return [[schedule.name, "DFH"[i] + r, product, field, amount === undefined ? "" : amount / 100,
          provider?.source ?? (entry ? "Prepared monthly workbook" : "Enter from reviewed statement")]];
      });
    }));
    if (result.portfolio) {
      summary.push(["Provider portfolio workbook", result.portfolioFile]);
      if (!result.performance) summary.push(["Pending inputs", "Unmapped Home Improvement items, Consolidation and Student"]);
    }
    if (result.statements.size) summary.push(["Statement PDF portfolios", [...result.statements.keys()].join(", ")]);
    sheets.push(table("Statement paste cells", ["Destination sheet", "Cell", "Product", "Input", "Value", "Source"], paste, [25, 12, 12, 38, 18, 49]));
  }
  if (result.cover) {
    summary.push(["Wisdom trial balance", result.wisdomFile]);
    sheets.push(table("Cover balance review", ["Product", "Balance", "G/L", "Cover Page cell", "Wisdom balance",
      "Wisdom source", "Statement balance", "CUSA entries", "Adjusted book", "Statement less book"],
      result.cover.map(item => [item.product, item.metric, item.account, item.ledgerCell, item.ledger / 100, item.source,
        item.statement / 100, item.adjustment / 100, item.adjusted / 100, item.difference / 100]),
      [12, 19, 13, 19, 18, 25, 19, 17, 18, 23]));
    const inputs = [10, 19].map(index => row(index, result.cover.filter(item => item.ledgerCell.endsWith(String(index)))
      .map(item => number(item.ledgerCell, item.ledger / 100)).join("")));
    sheets.push(["Cover Page inputs", worksheet("K10:AG19", [[11, 33, 15]], inputs.join(""))]);
  }
  return createExcelWorkbook([table("Summary", ["Field", "Value"], summary, [48, 54]), ...sheets,
    ...completionPasteSheets(result.completion, audit)], audit);
}

for (const kind of Object.keys(lendSources)) {
  const input = lendField(kind + "-file"), drop = lendField(kind + "-drop");
  input.addEventListener("change", () => selectLendFile(kind, input.files[0]));
  drop.addEventListener("dragover", event => { event.preventDefault(); drop.classList.add("active"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("active"));
  drop.addEventListener("drop", event => {
    event.preventDefault();
    drop.classList.remove("active");
    if (event.dataTransfer.files.length !== 1) return lendStatus("Drop one workbook at a time.", true);
    input.value = "";
    selectLendFile(kind, event.dataTransfer.files[0]);
  });
}
document.addEventListener("dragover", event => event.preventDefault());
document.addEventListener("drop", event => event.preventDefault());
const statementInput = lendField("statements-file"), statementDrop = lendField("statements-drop");
function selectLendStatements(files) {
  lendStatementFiles = Array.from(files);
  lendField("statements-status").textContent = lendStatementFiles.length ? lendStatementFiles.length + " PDF file(s) selected." : "";
  invalidateLend();
}
statementInput.addEventListener("change", () => selectLendStatements(statementInput.files));
statementDrop.addEventListener("dragover", event => { event.preventDefault(); statementDrop.classList.add("active"); });
statementDrop.addEventListener("dragleave", () => statementDrop.classList.remove("active"));
statementDrop.addEventListener("drop", event => {
  event.preventDefault();
  statementDrop.classList.remove("active");
  statementInput.value = "";
  selectLendStatements(event.dataTransfer.files);
});
for (const [code, name] of lendProducts) {
  const details = document.createElement("details"), summary = document.createElement("summary"), fields = document.createElement("div");
  summary.textContent = name;
  fields.className = "mapping";
  for (const [key, title] of lendPurposeFields) {
    const label = document.createElement("label"), input = document.createElement("input");
    input.id = code + "-" + key;
    input.inputMode = "decimal";
    label.append(title, input);
    fields.append(label);
  }
  details.append(summary, fields);
  lendField("purpose-inputs").append(details);
}
lendField("month").addEventListener("input", () => {
  const month = lendField("month").value;
  if (!month) return;
  const [year, number] = month.split("-").map(Number);
  lendField("from").value = month + "-01";
  lendField("through").value = new Date(Date.UTC(year, number, 0)).toISOString().slice(0, 10);
});
for (const input of document.querySelectorAll("input:not([type=file]):not(#confirm), select, textarea")) input.addEventListener("input", invalidateLend);
lendField("confirm").addEventListener("change", () => { if (!lendField("confirm").checked) invalidateLend(); });

lendField("run").addEventListener("click", async () => {
  const current = ++lendRevision;
  lendResult = undefined;
  lendField("results").hidden = true;
  try {
    const month = lendField("month").value;
    requireReport(month && lendField("confirm").checked, "Choose the month and verify the source inputs.");
    const achHasPortfolio = lendSources.ach.book?.sheets.has(lendPortfolioSheet);
    const portfolioSource = lendSources.portfolio.file ? lendSources.portfolio
      : !lendSources.fed.file && achHasPortfolio ? lendSources.ach : undefined;
    const wantsAch = lendSources.fed.file || (lendSources.ach.file && !achHasPortfolio);
    const wantsPerformance = lendSources.performance.file;
    const pdfFiles = lendStatementFiles;
    requireReport(wantsAch || wantsPerformance || portfolioSource || pdfFiles.length, "Load statement PDFs, a provider portfolio, a performance workbook, or the ACH and Fed exports.");
    requireReport(pdfFiles.length <= 3, "Select one monthly statement PDF per LendKey portfolio.");
    const from = lendField("from").value, through = lendField("through").value, direction = lendField("direction").value;
    let ach, performance, cover, portfolio, completion;
    const statements = new Map();
    if (wantsAch) {
      requireReport(lendSources.ach.book && lendSources.fed.book, "Load both ACH and Fed workbooks for the transaction comparison.");
      requireReport(from && through && from <= through && direction, "Choose the Fed date range and direction convention.");
      const overrides = readLendPostingDates();
      const [provider, fed] = await Promise.all([lendAchRows("ach", from, through, direction === "reverse", overrides), lendAchRows("fed", from, through, false)]);
      const unused = [...overrides.values()].find(item => !item.used);
      requireReport(!unused, unused && "No LendKey ACH rows match the override for " + unused.date + " " + unused.type + ".");
      // Each LendKey Fed entry ends with its LendKey date (",073126" = 7/31/26). An entry that settles in the window but
      // carries another month's date belongs to that month's reconciliation, so it is excluded here automatically.
      for (const item of fed) {
        if (!item.inRange || item.exclusion) continue;
        const reference = /,\s*(\d{2})(\d{2})(\d{2})(?!\d)/.exec(item.reference ?? "");
        if (!reference) continue;
        const [, mm, dd, yy] = reference;
        const lendKeyDate = "20" + yy + "-" + mm + "-" + dd;
        if (Number.isNaN(Date.parse(lendKeyDate + "T00:00:00Z")) || lendKeyDate.slice(0, 7) === month) continue;
        item.exclusion = (lendKeyDate < month ? "Prior" : "Next") + " month LendKey reference " + mm + "/" + dd + "/" + yy + " · excluded automatically";
        item.autoExcluded = true;
        item.included = false;
      }
      for (const { item, reason } of lendReviewedFedRows(fed.filter(item => item.inRange && !item.exclusion), lendField("fed-exclusions").value)) {
        item.exclusion = reason;
        item.included = false;
      }
      requireReport(provider.some(item => item.included) || fed.some(item => item.included), "Neither export has transactions in the selected period.");
      ach = compareLendAch(provider, fed);
      lendReviewMissingProvider(ach, lendField("fed-confirmed").value);
    }
    if (current !== lendRevision) return;
    if (portfolioSource) {
      requireReport(portfolioSource.book, "Load a readable provider portfolio workbook.");
      portfolio = await readLendPortfolio(portfolioSource.book, month);
    }
    if (current !== lendRevision) return;
    const statementCash = new Map();
    for (const file of pdfFiles) {
      const { product, fields, cash } = await readLendStatementPdf(file, month);
      if (current !== lendRevision) return;
      requireReport(!statements.has(product), "More than one PDF was selected for " + product + ".");
      statements.set(product, fields);
      statementCash.set(product, cash);
    }
    if (ach) lendAutoReviewProvider(ach, statementCash);
    const priorPurpose = lendSources.performance.book?.sheets.has("Purpose Matchup")
      ? lendPriorPurposeLabels(await readCells(lendSources.performance.book, "Purpose Matchup")) : undefined;
    if (wantsPerformance) {
      requireReport(lendSources.performance.book, "Load a readable performance workbook.");
      const book = lendSources.performance.book;
      [performance, completion] = await Promise.all([
        readLendPerformance(book, portfolio, statements),
        readCompletionTarget(book, "Cover Page", "B25", "B28", "D25", "D26"),
      ]);
    } else if (statements.size) {
      const rollForwards = await readLendPerformance(undefined, portfolio, statements);
      if (rollForwards.some(schedule => schedule.entries.length)) performance = rollForwards;
    }
    if (current !== lendRevision) return;
    if (lendSources.wisdom.file) {
      requireReport(wantsPerformance && performance && lendSources.wisdom.book, "Load the monthly performance workbook and a readable Wisdom trial balance to fill the cover.");
      cover = await readLendCover(lendSources.performance.book, performance, month);
    }
    if (current !== lendRevision) return;
    const purpose = ach ? buildLendPurpose(ach, performance) : undefined;
    // An entered Home Improvement late fee takes precedence over the provider workbook calculation.
    let lateFees;
    const hil = purpose?.groups.find(group => group.code === "HIL");
    if (hil && hil.inputs["late-fees"] == null) {
      // Use whichever loaded workbook has the provider servicing detail.
      lateFees = await readLendServicingLateFees(lendSources.portfolio.book) ?? await readLendServicingLateFees(lendSources.ach.book);
      if (current !== lendRevision) return;
      if (lateFees) hil.inputs["late-fees"] = lateFees.amount;
    }
    const template = lendSources.performance.file;
    lendResult = { month, from, through, reverse: direction === "reverse", ach, performance, cover, purpose, portfolio, completion, statements, template,
      statementCash, priorPurpose, lateFees,
      portfolioFile: portfolioSource?.file?.name,
      achFile: lendSources.ach.file?.name, fedFile: lendSources.fed.file?.name, performanceFile: lendSources.performance.file?.name,
      wisdomFile: lendSources.wisdom.file?.name };
    showLend(lendResult);
    const canFill = Boolean(template && (portfolio || statements.size || cover || completion));
    lendField("download-template").hidden = !canFill;
    lendField("template-download-note").hidden = !canFill;
    lendStatus("");
  } catch (error) {
    if (current === lendRevision) lendStatus(error.message, true);
  }
});

function fillLendTemplate(result) {
  const audit = reconciliationAudit();
  const updates = result.performance.flatMap(schedule => schedule.entries.flatMap(entry => schedule.fields.flatMap(([row]) => {
    const provider = lendStatementInput(result.statements, result.portfolio, entry.product, schedule.name, row);
    return provider ? [[schedule.name, entry.column + row, provider.amount / 100]] : [];
  })));
  for (const item of result.cover ?? []) updates.push(["Cover Page", item.ledgerCell, item.ledger / 100]);
  if (result.completion) {
    const { sheet, initials, date } = result.completion;
    updates.push([sheet, initials, audit.by], [sheet, date, { date: audit.on }]);
  }
  return fillExcelTemplate(result.template, updates);
}

async function downloadLend(button, build, name) {
  const result = lendResult;
  if (!result) return;
  button.disabled = true;
  try {
    const blob = await build(result);
    if (result !== lendResult) return;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = name(result);
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 30000);
  } catch (error) {
    if (result === lendResult) lendStatus(error.message, true);
  } finally {
    button.disabled = false;
  }
}

lendField("download").addEventListener("click", event =>
  downloadLend(event.currentTarget, buildLend, result => "LendKey " + result.month + ".xlsx"));
lendField("download-template").addEventListener("click", event =>
  downloadLend(event.currentTarget, fillLendTemplate,
    result => "LendKey " + result.month + " filled template." + result.template.name.split(".").at(-1)));
