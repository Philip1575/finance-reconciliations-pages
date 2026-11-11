"use strict";

async function readLoanStreetProfile(file) {
  if (/\.xlsx$/i.test(file.name)) return readLoanStreetWorkbookProfile(file);
  requireReport(/\.json$/i.test(file.name) && file.size <= 65536, "Choose the local LoanStreet profile in JSON format.");
  let profile;
  try {
    profile = JSON.parse(await file.text());
  } catch {
    throw new Error("The profile is not valid JSON.");
  }
  return validateLoanStreetProfile(profile);
}

function validateLoanStreetProfile(profile) {
  const validSheet = name => typeof name === "string" && name.length > 0 && name.length <= 31
    && !/[\[\]:*?/\\]/.test(name) && !name.startsWith("'") && !name.endsWith("'");
  const validAddress = address => typeof address === "string" && /^[A-J][1-9]\d?$/.test(address);
  const validAccount = account => typeof account === "string" && /^\d{6}$/.test(account);
  const validCover = entry => Array.isArray(entry) && entry.length === 3 && validAccount(entry[0])
    && entry.slice(1).every(validAddress) && entry[1] !== entry[2];
  const { pools, covers, details, payable } = profile ?? {};
  requireReport(Array.isArray(pools) && pools.length === 8 && new Set(pools).size === pools.length
    && pools.every(pool => typeof pool === "string" && /^[A-Za-z0-9_-]{1,40}$/.test(pool)),
    "The profile must list eight unique pool identifiers in worksheet order.");
  requireReport(Array.isArray(covers) && covers.length === pools.length && covers.every(entry =>
    Array.isArray(entry) && entry.length === 3 && validSheet(entry[0]) && validCover(entry[1]) && validCover(entry[2])
      && new Set([...entry[1].slice(1), ...entry[2].slice(1)]).size === 4),
    "The profile is missing a valid cover mapping for a pool.");
  requireReport(Array.isArray(details) && details.length === pools.length && details.every(entry =>
    Array.isArray(entry) && entry.length === 5 && validSheet(entry[0])
      && entry.slice(1).every(column => typeof column === "string" && /^[A-U]$/.test(column))
      && new Set(entry.slice(1, 4)).size === 3),
    "The profile is missing a valid detail mapping for a pool.");
  requireReport(Array.isArray(payable) && payable.length === 4 && validSheet(payable[0]) && validCover(payable.slice(1)),
    "The profile is missing the accounts payable cover mapping.");
  const sheetNames = [...covers.map(entry => entry[0]), ...details.map(entry => entry[0]), payable[0]]
    .map(name => name.toLowerCase());
  requireReport(new Set(sheetNames).size === sheetNames.length && sheetNames.every(name =>
    !name.startsWith("key info ") && !["cover inputs", "detail inputs", "cover period inputs"].includes(name)),
    "Worksheet names in the profile must be unique and cannot use a download summary name.");
  const accounts = [...covers.flatMap(([, principal, interest]) => [principal[0], interest[0]]), payable[1]];
  requireReport(new Set(accounts).size === accounts.length, "Each cover mapping must use a distinct G/L account.");
  return { pools, covers, details, payable, accounts };
}

async function readLoanStreetWorkbookProfile(file) {
  const book = await openExcel(file);
  const names = [...book.sheets.keys()];
  const pairs = names.flatMap(cover => {
    const pool = /^([A-Za-z0-9_-]+)\s+cover$/i.exec(cover)?.[1];
    if (!pool) return [];
    const details = names.filter(sheet => sheet.toLowerCase() === pool.toLowerCase() + " detail");
    requireReport(details.length === 1, "Each pool cover needs one matching detail sheet.");
    return [{ pool, cover, detail: details[0] }];
  });
  requireReport(pairs.length === 8, "Choose a consolidated workbook with eight paired pool cover and detail sheets.");
  const apSheet = names.find(name => name.toLowerCase() === "ap cover page");
  requireReport(apSheet, "The consolidated workbook needs an AP Cover Page.");
  const readCover = async (sheet, count) => {
    const formulas = new Set(), cells = await readCells(book, sheet, formulas);
    const anchors = [...cells].flatMap(([address, value]) => {
      const row = Number(/^B([1-9]\d?)$/.exec(address)?.[1]);
      const account = normalizeAccount(value);
      const label = String(cells.get("A" + row) ?? "").trim();
      return row && account && (count === 1 ? row === 1 : /^(?:cusa|portico|wisdom):?$/i.test(label)) ? [{ row, account }] : [];
    }).sort((a, b) => a.row - b.row);
    requireReport(anchors.length === count, sheet + " has an unexpected number of account labels.");
    const amountAt = row => {
      const amounts = "BCDEFGHIJ".split("").map(column => column + row)
        .filter(address => typeof cells.get(address) === "number" && Number.isFinite(cells.get(address)) && !formulas.has(address));
      requireReport(amounts.length === 1, sheet + " needs one numeric input on balance row " + row + ". Use a completed workbook.");
      return amounts[0];
    };
    return anchors.map(({ row, account }, index) => {
      const next = anchors[index + 1]?.row ?? 100;
      const labels = [...cells].flatMap(([address, value]) => {
        const line = Number(/^A([1-9]\d?)$/.exec(address)?.[1]);
        const text = String(value).trim().replace(/\s+/g, " ");
        return line >= row && line < next ? [[line, text]] : [];
      });
      const find = pattern => {
        const matches = labels.filter(([, value]) => pattern.test(value));
        requireReport(matches.length === 1, sheet + " needs one provider and one ledger balance per account.");
        return amountAt(matches[0][0]);
      };
      const provider = find(/^balance per portfolio statement$/i), ledger = find(/^balance per (?:cusa|portico|portcio|wisdom) g\/l$/i);
      requireReport(provider[0] === ledger[0], sheet + " has inconsistent balance columns.");
      return [account, provider, ledger];
    });
  };
  const covers = [], details = [];
  const detailPatterns = [/^principal payment\b/i, /^net interest payment\b/i, /^change in a\/?p$/i, /^charge[ -]offs$/i];
  for (const pair of pairs) {
    covers.push([pair.cover, ...await readCover(pair.cover, 2)]);
    const cells = await readCells(book, pair.detail);
    const columns = detailPatterns.map(pattern => {
      const matches = new Set([...cells].flatMap(([address, value]) =>
        typeof value === "string" && pattern.test(value.trim()) ? [/^([A-U])\d+$/.exec(address)?.[1]] : []));
      requireReport(matches.size === 1 && !matches.has(undefined), pair.detail + " has an ambiguous detail column.");
      return [...matches][0];
    });
    details.push([pair.detail, ...columns]);
  }
  const [payable] = await readCover(apSheet, 1);
  return validateLoanStreetProfile({ pools: pairs.map(pair => pair.pool), covers, details, payable: [apSheet, ...payable] });
}
