"use strict";

const vizoAccounts = {
  overnight: {
    name: "Overnight Settlement", cover: "Overnight cover", groups: "OSL posting review", ledger: "trial",
    paste: { ledger: "L24", initials: "L45", date: "L46", difference: "L41" },
    note: "Review the opposite-signed posting groups and verify the receiving accounts before preparing month-end entries.",
    tabs: [],
  },
  psa: {
    name: "PSA", cover: "PSA cover", groups: "PSA groups", ledger: "history", interest: true,
    paste: { ledger: "L24", initials: "L46", date: "L45", difference: "L41" },
    note: "Compare transfers and overdraft protection with the matching accounts. Review the interest transfer pair.",
    tabs: [["PSA transfers", "Transfer"], ["PSA overdraft", "Overdraft"]],
  },
  pcc: {
    name: "PCC", cover: "PCC cover", groups: "PCC groups", ledger: "trial", interest: true,
    paste: { ledger: "L24", initials: "L48", date: "L49", difference: "L44" },
    note: "Compare the statement and ledger balances. Review any transfers and verify the interest transfer pair.",
    tabs: [["PCC transfers", "Transfer"]],
  },
  misc: {
    name: "MISC", cover: "MISC cover", groups: "MISC groups", ledger: "history",
    paste: { ledger: "L24", initials: "L52", date: "L53", difference: "L48" },
    note: "Review card funding suspense, overdraft protection, and transfers. Enter the signed total of reviewed in-transit items.",
    tabs: [["Card funding suspense", "Card funding suspense"], ["MISC overdraft", "Overdraft"], ["MISC transfers", "Transfer"]],
  },
  charge: {
    name: "Charge Cards", cover: "Charge Cards cover", groups: "Charge Cards groups", ledger: "history",
    paste: { ledger: "L24", initials: "L49", date: "L50", difference: "L45" },
    note: "Complete LoanStreet first. Review money order batches, cash advances, invoices, transfers and overdraft items.",
    tabs: [["Billing review", "Billing"], ["Cash advances", "Cash advance"], ["Money orders", "Money order"],
      ["Card transfers", "Transfer"], ["Card overdraft", "Overdraft"]],
  },
  settlement: {
    name: "Settlement", cover: "Settlement cover", groups: "Settlement groups", ledger: "history",
    paste: { sheet: "Reconciliation Settlement Acct", ledger: "L24", initials: "L61", date: "L62", difference: "L49" },
    note: "Review this account after the related accounts. Compare share drafts, branch capture and mobile capture, then confirm the branch carryover.",
    tabs: [["Sharedraft review", "Share draft"], ["Branch Capture review", "Branch capture"], ["Mobile Capture review", "Mobile capture"],
      ["Returned deposit review", "Returned deposit"],
      ["Settlement transfers", "Transfer"], ["Settlement overdraft", "Overdraft"], ["Investment deposits", "External deposit"]],
  },
  atm: {
    name: "ATM Debit Card Settlement", cover: "ATM cover", groups: "ATM provider groups", ledger: "history",
    paste: { sheet: "Reconciliation ATM", ledger: "L26", initials: "L69", date: "L70", difference: "L62" },
    note: "Use IF-67 for weekend and holiday batches and any possible ISA fee differences. Verify the final carryover.",
    tabs: [["ATM transfers", "Transfer"], ["ATM overdraft", "Overdraft"]],
  },
};

function vizoAccountCategory(description) {
  const value = description.toUpperCase();
  const base = vizoCategory(description);
  if (base === "ATM daily") return base;
  if (/LOAN[ -]*STREET/.test(value)) return "LoanStreet";
  if (/FUNDING.*SUSPENSE|SUSPENSE.*FUNDING/.test(value)) return "Card funding suspense";
  if (/OVERDRAFT/.test(value)) return "Overdraft";
  if (/\b(?:SD|SHARE[ -]*DRAFT) SUSPENSE\b/.test(value)) return "Share draft suspense";
  if (/SHARE[ -]*DRAFT|CASH LETTER|\bSD CLEARING\b/.test(value)) return "Share draft";
  if (/\b(?:(?:FORWARD|FWD) (?:CHECK )?COLLECTION.*RET(?:URN(?:ED)?)?|RETURNED DEPOSITS?|RDI)\b/.test(value)) return "Returned deposit";
  if (/\b(?:BRANCH CAPTURE|BRANCH DEPOSITS?|DAILY DEPOSITS?|PHYSICAL CHECKS?)\b/.test(value)) return "Branch capture";
  if (/\b(?:MOBILE CAPT(?:URE)?|MOBILE DEPOSITS?|RDC)\b/.test(value)) return "Mobile capture";
  if (/\b(?:FORWARD|FWD) (?:CHECK )?COLLECTION\b/.test(value)) return "Unclassified forward collection";
  if (/MONEY[ -]*GRAM|MONEY[ -]*ORDER/.test(value)) return "Money order";
  if (/CASH ADVANCE|CREDIT CARD CHARGE/.test(value)) return "Cash advance";
  if (/\bBILL(?:ING)?\b|\bINVOICE\b/.test(value)) return "Billing";
  if (/INTEREST|DIVIDEND/.test(value)) return "Interest or dividend";
  if (/^EXTERNAL DEPOSIT/.test(value) && !/TRANSFER/.test(value)) return "External deposit";
  return base;
}

function vizoForwardDepositCandidates(provider, ledger, sign, offset) {
  const groups = new Map();
  for (const item of ledger) {
    const category = vizoRowCategory(item);
    if (!["Branch capture", "Mobile capture"].includes(category)) continue;
    const key = item.date + "|" + category;
    const group = groups.get(key) ?? { amount: 0, rows: [] };
    group.amount += item.amount * sign;
    group.rows.push(item.row);
    groups.set(key, group);
  }
  return provider.map(item => {
    const dates = new Set([dateOffset(item.date, offset)]);
    const day = new Date(item.date + "T00:00:00Z").getUTCDay();
    if (offset === -1 && day === 1) dates.add(dateOffset(item.date, -3));
    const candidates = [...dates].flatMap(date => {
      const branch = groups.get(date + "|Branch capture");
      const mobile = groups.get(date + "|Mobile capture");
      const choices = [["Branch capture", branch, branch?.amount, 0], ["Mobile capture", mobile, 0, mobile?.amount]];
      if (branch?.amount && mobile?.amount) choices.push(["Branch and mobile capture", {
        amount: branch.amount + mobile.amount, rows: [...branch.rows, ...mobile.rows],
      }, branch.amount, mobile.amount]);
      return choices.filter(([, group]) => group?.amount === item.amount)
        .map(([category, group, branchAmount, mobileAmount]) => ({ date, category, rows: group.rows, branchAmount, mobileAmount }));
    });
    return { item, candidates };
  });
}

function vizoRowCategory(item) {
  return item.reviewedCategory ?? vizoAccountCategory(item.description);
}

function vizoReviewedDeposits(text, provider, ledger, end) {
  const eligible = provider.filter(item => item.amount > 0 && vizoRowCategory(item) === "Unclassified forward collection");
  const ledgerDates = new Set(ledger.filter(item => item.date <= end).map(item => item.date));
  return readReviewedDepositAllocations(text, eligible, ledgerDates);
}

function vizoAccountComparison(provider, ledger, category, sign, offset) {
  const buckets = new Map();
  for (const [side, source] of [["provider", provider], ["ledger", ledger]]) {
    for (const item of source) {
      if (vizoRowCategory(item) !== category) continue;
      const amount = side === "provider" ? item.amount : item.amount * sign;
      const direction = amount < 0 ? "Debit" : "Credit";
      const date = side === "provider" && !item.ledgerDate ? item.date : dateOffset(item.ledgerDate ?? item.date, -offset);
      const key = date + direction;
      const bucket = buckets.get(key) ?? { date, direction, provider: 0, ledger: 0, providerRows: [], ledgerRows: [], providerDates: new Set() };
      bucket[side] += amount;
      bucket[side + "Rows"].push(item.reviewRow ?? item.row);
      if (side === "provider") bucket.providerDates.add(item.date);
      requireReport(Number.isSafeInteger(bucket[side]), "An account comparison total exceeds the supported range.");
      buckets.set(key, bucket);
    }
  }
  return [...buckets.values()].sort((a, b) => a.date.localeCompare(b.date) || a.direction.localeCompare(b.direction));
}
