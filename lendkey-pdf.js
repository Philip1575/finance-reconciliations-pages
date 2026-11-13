"use strict";

const lendPerformancePdfFields = [
  [
    [5, "Principal Ending Balance"], [6, "Prior Period Adjustment"], [9, /^Disb\b/i],
    [10, "Adjustments", 0, true], [11, "Origination Fee"], [12, "Adjustments", 1, true],
    [13, "Interest Capitalized"], [15, "Loan Principal Purchased"],
    [19, "Borrower Principal Repayment, Orig"], [21, "Borrower Principal Repayment, Part"],
    [23, "Loan Principal Cancelled/Defaulted"], [25, "Adjustments", 2], [27, "Loan Principal Sold"],
    [29, "Recovery Principal"], [31, "Principal Adjustment"], [37, "Ending Principal Balance"],
  ],
  [
    [5, "Accr. Int. Ending Balance"], [6, "Prior Period Adjustment"], [9, "Interest Earned, Origination"],
    [11, "Interest Earned, Participation"], [13, "Loan Accrued Purchased"], [18, "Interest Capitalized"],
    [20, "Interest Paid"], [22, "Interest Reversal"], [24, "Interest Loan Cancel / Default"],
    [26, "Loan Accrued Interest Sold"], [33, "Ending Accrued Interest"],
  ],
];
const lendHomePdfFields = [
  [
    [5, "PR Ending Prin Balance"], [9, "DISB"], [10, "PMT Reversals (PRIN)"], [12, "Refunds Applied to PRIN"],
    [13, "Capitalized INT"], [21, "PMT Applied to PRIN"], [27, "PRIN Claim Payment"], [37, "Ending PRIN Balance"],
  ],
  [
    [5, "Accr INT Ending Balance"], [6, "Prior Period Adjustments"], [11, "INT ACCR During Period"],
    [18, "Capitalized INT"], [20, "PMT Applied to INT"], [22, "PMT Reversals (INT)", undefined, false, -1],
    [26, "INT Claim Payment"], [33, "Ending INT Balance"],
  ],
];

// Cash-section lines used by the monthly Purpose Matchup: settlements, remittances and in-transit balances.
const lendPdfCashLines = {
  HIL: [["priorAR", "Prior AR Settlement"], ["priorAP", "Prior AP Settlement"], ["remittance", "PMT Remittance"],
    ["reversal", "PMT Remittance Reversal"], ["newAR", "Accounts Receivable (To Lender)"], ["newAP", "Accounts Payable (From Lender)"]],
  other: [["remittance", "Remittance to Lender"], ["newAR", "Accounts Receivable (to Lender)"], ["newAP", "Accounts Payable (from Lender)"]],
};

function lendPdfCash(items, product) {
  const cash = new Map();
  for (const [key, label] of lendPdfCashLines[product === "HIL" ? "HIL" : "other"]) {
    const labels = items.filter(item => item.text.trim().toLowerCase() === label.toLowerCase());
    if (labels.length !== 1) continue;
    const [line] = labels;
    const [amount] = items.filter(item => item.x > line.x && Math.abs(item.y - line.y) < 2.5 && Number.isFinite(readBalance(item.text)))
      .sort((a, b) => a.x - b.x);
    if (amount) cash.set(key, cents(amount.text));
  }
  return cash;
}

function lendPdfSection(items, start, end) {
  const section = items.filter(item => item.x >= start - 1 && item.x < end - 1);
  return (selector, occurrence) => {
    const labels = section.filter(item => typeof selector === "string"
      ? item.text.toLowerCase() === selector.toLowerCase() : selector.test(item.text)).sort((a, b) => b.y - a.y);
    const line = occurrence === undefined ? (labels.length === 1 ? labels[0] : undefined)
      : (labels.length === 3 ? labels[occurrence] : undefined);
    requireReport(line, "A required statement line could not be identified: " + selector + ".");
    const amounts = section.filter(item => item.x > line.x && Math.abs(item.y - line.y) < 2 && Number.isFinite(readBalance(item.text)));
    requireReport(amounts.length === 1, "The amount on the " + line.text + " line is ambiguous.");
    return { label: line.text, amount: cents(amounts[0].text) };
  };
}

async function readLendStatementPdf(file, month) {
  try {
    const text = await readLocalPdfText(file);
    const titles = text.filter(item => /^(?:loan principal|prin|int|interest) reconciliation$/i.test(item.text)).sort((a, b) => a.x - b.x);
    requireReport(titles.length === 2, "The PDF does not have the expected principal and interest sections.");
    const home = titles[0].text.toLowerCase() === "prin reconciliation";
    const header = text.find(item => home ? /^reconciliation report as of$/i.test(item.text) : /^month end performance report\s*-/i.test(item.text));
    const product = home ? "HIL" : /consolidation$/i.test(header?.text ?? "") ? "CNL" : /student$/i.test(header?.text ?? "") ? "STU" : "";
    requireReport(header && product, "The PDF does not identify a supported LendKey portfolio.");
    const dateText = text.find(item => home
      ? item.x > header.x && Math.abs(item.y - header.y) < 2 && /^[a-z]+ \d{1,2},? \d{4}$/i.test(item.text)
      : item.y < header.y && header.y - item.y < 30 && /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(item.text))?.text;
    const [year, period] = month.split("-").map(Number);
    const monthEnd = new Date(Date.UTC(year, period, 0));
    const [fullMonth, day] = new Intl.DateTimeFormat("en-US", { month: "long", day: "numeric", timeZone: "UTC" }).format(monthEnd).split(" ");
    const allowedDates = home ? [fullMonth + " " + day + " " + year]
      : [period + "/" + day + "/" + year, period + "/" + day + "/" + String(year).slice(2)];
    requireReport(dateText && allowedDates.includes(dateText.replaceAll(",", "").replace(/\b0(?=\d\/)/g, "")),
      "The statement PDF does not show the selected month-end.");
    const readers = [lendPdfSection(text, titles[0].x, titles[1].x), lendPdfSection(text, titles[1].x, Infinity)];
    const layouts = home ? lendHomePdfFields : lendPerformancePdfFields;
    const fields = new Map(lendSchedules.map((schedule, index) => {
      const read = readers[index], values = new Map();
      for (const [destination, selector, occurrence, zeroOnly = false, sign = 1] of layouts[index]) {
        const entry = read(selector, occurrence);
        if (!zeroOnly || entry.amount === 0) values.set(destination, {
          amount: entry.amount * sign, source: file.name + " · page 1 · " + entry.label + (sign === -1 ? " · sign reversed" : ""),
        });
      }
      if (home) {
        // The Home Improvement statement reports non-cash reductions as credits, charge-offs and cancelled
        // disbursements; they belong on the cancelled/defaulted line. Lines it never reports are zero for this portfolio,
        // so the statement alone fills every line and nothing is carried over from a prior workbook.
        const [destination, labels] = index === 0 ? [23, ["PRIN Credits", "Charge Off (PRIN)", "Cancelled DISB"]] : [24, ["INT Credits", "Charge Offs (INT)"]];
        const entries = labels.map(label => read(label));
        values.set(destination, { amount: entries.reduce((sum, entry) => sum + entry.amount, 0),
          source: file.name + " · page 1 · " + entries.map(entry => entry.label).join(" + ") });
        for (const [row] of schedule.fields) {
          if (!values.has(row)) values.set(row, { amount: 0, source: file.name + " · not reported on the Home Improvement statement (zero)" });
        }
      }
      return [schedule.name, values];
    }));
    return { product, fields, cash: lendPdfCash(text, product) };
  } catch (error) {
    throw new Error(file.name + ": " + error.message);
  }
}

// The prior month's LendKey billing invoice: line items, tax and amount due (shared by the VIZO pages).
async function readVizoLendKeyInvoice(file, month) {
  const text = await readLocalPdfText(file);
  requireReport(text.some(item => /^lendkey$/i.test(item.text)) && text.some(item => /^technologies,?$/i.test(item.text)),
    "The PDF does not identify a LendKey invoice.");
  const named = value => text.filter(item => item.text.toLowerCase() === value.toLowerCase());
  const sameRow = (a, b) => Math.abs(a.y - b.y) < 2;
  const pair = (first, second) => {
    const candidates = named(second).filter(b => named(first).some(a => a.x < b.x && sameRow(a, b)));
    requireReport(candidates.length === 1, "Check the " + first + " " + second + " invoice label.");
    return candidates[0];
  };
  const rightAmount = line => {
    const values = text.filter(item => item.x > line.x && sameRow(item, line) && Number.isFinite(readBalance(item.text)));
    requireReport(values.length === 1, "An invoice amount is missing or ambiguous.");
    return cents(values[0].text);
  };
  const dateLabel = pair("Invoice", "Date:");
  const dateMatch = text.find(item => item.x > dateLabel.x && sameRow(item, dateLabel)
    && /^\d{1,2}-[A-Za-z]{3}-\d{4}$/.test(item.text));
  requireReport(dateMatch, "The LendKey invoice date could not be verified.");
  const [day, abbreviation, year] = dateMatch.text.split("-");
  const months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const invoiceDate = sourceDate(year + "-" + String(months.indexOf(abbreviation.toLowerCase()) + 1).padStart(2, "0") + "-" + day.padStart(2, "0"));
  const priorMonth = dateOffset(month + "-01", -1).slice(0, 7);
  requireReport(invoiceDate.startsWith(priorMonth), "The manual calls for the " + priorMonth + " invoice for this reconciliation.");
  const subtotalLabel = pair("Invoice", "Total:"), subtotal = rightAmount(subtotalLabel);
  const tax = rightAmount(pair("Sales", "Tax:")), due = rightAmount(pair("Amount", "Due:"));
  const number = named("NO.")[0], quantity = named("QTY")[0], amount = named("AMOUNT").find(item => number && sameRow(item, number));
  requireReport(number && quantity && amount && sameRow(number, quantity) && number.x < quantity.x && quantity.x < amount.x,
    "The invoice item columns could not be verified.");
  const items = text.filter(item => Math.abs(item.x - number.x) < 12 && item.y < number.y
    && item.y > subtotalLabel.y && /^\d+$/.test(item.text)).sort((a, b) => b.y - a.y);
  requireReport(items.length && items.every((item, i) => Number(item.text) === i + 1), "The invoice item numbers could not be verified.");
  const lines = items.map((item, i) => {
    const description = text.filter(word => word.x > number.x + 12 && word.x < quantity.x - 5
      && word.y <= item.y + 2 && word.y > Math.max(item.y - 19, (items[i + 1]?.y ?? subtotalLabel.y) + 2))
      .sort((a, b) => Math.abs(a.y - b.y) < 2 ? a.x - b.x : b.y - a.y).map(word => word.text).join(" ");
    requireReport(description, "An invoice item description is missing.");
    return { number: Number(item.text), description, amount: rightAmount({ x: amount.x - 3, y: item.y }) };
  });
  requireReport(lines.reduce((sum, item) => sum + item.amount, 0) === subtotal && subtotal + tax === due && due > 0,
    "The line items, tax and invoice amount due do not agree.");
  return { file: file.name, invoiceDate, lines, subtotal, tax, due };
}
