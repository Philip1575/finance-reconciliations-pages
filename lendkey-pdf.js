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
      return [schedule.name, values];
    }));
    return { product, fields, cash: lendPdfCash(text, product) };
  } catch (error) {
    throw new Error(file.name + ": " + error.message);
  }
}
