// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Decides every figure and date of a Dutch collection-costs statement.
 * Installed by `dxcl add nl-collection-costs`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The
 * thresholds are named locals at the top of `run` so an owner can correct them
 * without reading the rest.
 */
export default deriver({
  name: "nlCollectionCosts",
  run: (collection: unknown) => {
    const col = collection as any;

    // Thresholds, in cents unless stated. Not checked against primary sources.
    const FIXED_PER_INVOICE = 4000; // art. 6:96 lid 4 BW, business debtors
    const INSTALMENT_CEILING = 26667; // art. 2a Besluit BIK: below this, capped
    const INSTALMENT_FIRST = 4000;
    const INSTALMENT_LATER = 2000;
    const INSTALMENT_WINDOW_MONTHS = 6;
    const POST_DAYS_DEFAULT = 4;
    const NOTICE_DAYS = 15; // fourteen days plus the day after receipt
    const RECEIVED_DAYS = 14;
    const VAT_PERCENT = 21;
    const STAFFEL_MIN = 4000;
    const STAFFEL_CAP = 677500;

    const fail = (field: string, why: string): never => {
      throw new Error("collection." + field + " " + why);
    };
    const isoOf = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        return fail(field, "must be an ISO 8601 date");
      }
      return value.slice(0, 10);
    };
    // UTC day indexes, so a day does not shift with the engine's time zone.
    const dayOf = (iso: string) => {
      const [y, m, d] = iso.split("-").map(Number);
      return Date.UTC(y, m - 1, d) / 86_400_000;
    };
    const dateOf = (value: unknown, field: string) => dayOf(isoOf(value, field));
    const euros = (value: unknown, field: string, min: number): number => {
      if (typeof value !== "number" || !Number.isFinite(value) || value < min) {
        return fail(field, min > 0 ? "must be a finite number above 0" : "must be a finite number, not negative");
      }
      return Math.round(value * 100);
    };
    const named = (value: unknown, field: string): string => {
      if (typeof value !== "string" || value.trim() === "") return fail(field, "is required");
      return value;
    };

    const money = new Intl.NumberFormat("nl-NL", { style: "currency", currency: "EUR" });
    const fmtMoney = (cents: number) => money.format(cents / 100);
    const dates = new Intl.DateTimeFormat("nl-NL", { dateStyle: "full", timeZone: "UTC" });
    const fmtDay = (day: number) => dates.format(new Date(day * 86_400_000));
    const half = (n: number, d: number) => Math.floor((n + d / 2) / d);

    // Anonymous Gregorian algorithm (Meeus/Jones/Butcher).
    const easter = (year: number) => {
      const a = year % 19;
      const b = Math.floor(year / 100);
      const c = year % 100;
      const d = Math.floor(b / 4);
      const e = b % 4;
      const f = Math.floor((b + 8) / 25);
      const g = Math.floor((b - f + 1) / 3);
      const h = (19 * a + b - d - g + 15) % 30;
      const i = Math.floor(c / 4);
      const k = c % 4;
      const l = (32 + 2 * e + 2 * i - h - k) % 7;
      const m = Math.floor((a + 11 * h + 22 * l) / 451);
      const month = Math.floor((h + l - 7 * m + 114) / 31);
      const dayOfMonth = ((h + l - 7 * m + 114) % 31) + 1;
      return Date.UTC(year, month - 1, dayOfMonth) / 86_400_000;
    };
    // Algemene termijnenwet holidays. Koningsdag on a Sunday moves to a
    // Saturday, which is rolled anyway.
    const isHoliday = (day: number) => {
      const date = new Date(day * 86_400_000);
      const year = date.getUTCFullYear();
      const at = (month: number, dom: number) => Date.UTC(year, month - 1, dom) / 86_400_000;
      const e = easter(year);
      return [at(1, 1), e + 1, e + 39, e + 50, at(4, 27), at(5, 5), at(12, 25), at(12, 26)].includes(day);
    };
    const roll = (from: number) => {
      let day = from;
      while ([0, 6].includes(new Date(day * 86_400_000).getUTCDay()) || isHoliday(day)) day += 1;
      return day;
    };

    // The staffel in thousandths of a cent, so nothing is floating.
    const staffel = (p: number) => {
      const band = (from: number, width: number) => Math.min(Math.max(p - from, 0), width);
      const raw = band(0, 250000) * 150 + band(250000, 250000) * 100 + band(500000, 500000) * 50 +
        band(1000000, 19000000) * 10 + Math.max(p - 20000000, 0) * 5;
      return Math.min(STAFFEL_CAP, Math.max(STAFFEL_MIN, Math.floor((raw + 500) / 1000)));
    };

    // Validate first: every rule gives different answers if a field is wrong.
    if (col === null || typeof col !== "object") fail("", "is required");
    const debtor = col.debtorType;
    if (debtor !== "consumer" && debtor !== "business") fail("debtorType", 'must be "consumer" or "business"');
    const consumer = debtor === "consumer";
    const sendIso = isoOf(col.sendDate, "sendDate");
    const send = dayOf(sendIso);

    if (!Array.isArray(col.invoices) || col.invoices.length === 0) fail("invoices", "must list at least one invoice");
    const numbers: string[] = [];
    let principalCents = 0;
    col.invoices.forEach((invoice: any, i: number) => {
      const at = "invoices[" + i + "]";
      numbers.push(named(invoice?.number, at + ".number"));
      principalCents += euros(invoice?.principal, at + ".principal", 1e-9);
      if (dateOf(invoice?.dueDate, at + ".dueDate") >= send) fail(at + ".dueDate", "must be before sendDate");
    });

    const postDays = col.postDeliveryDays ?? POST_DAYS_DEFAULT;
    if (!Number.isInteger(postDays) || postDays < 0) fail("postDeliveryDays", "must be a non-negative integer");
    const prev = consumer ? col.previousLetter : undefined;
    if (consumer && !prev && col.channel !== "post" && col.channel !== "email") {
      fail("channel", 'must be "post" or "email" on a veertiendagenbrief');
    }

    const hasValue = (v: unknown) => v !== undefined && v !== null;
    const interestInput = hasValue(col.interestAccrued) ? euros(col.interestAccrued, "interestAccrued", 0) : 0;
    const sinceDay = hasValue(col.interestRunningSince)
      ? dateOf(col.interestRunningSince, "interestRunningSince")
      : undefined;

    const instalment = consumer ? col.instalment : undefined;
    let windowEnd = 0;
    if (instalment) {
      if (col.invoices.length !== 1) fail("instalment", "needs exactly one invoice");
      const prior = instalment.priorCharged;
      if (!Number.isInteger(prior) || prior < 0) fail("instalment.priorCharged", "must be a non-negative integer");
      if (prior > 0) {
        const [y, m, d] = isoOf(instalment.windowStart, "instalment.windowStart").split("-").map(Number);
        windowEnd = Date.UTC(y, m - 1 + INSTALMENT_WINDOW_MONTHS, d) / 86_400_000;
      }
    }

    const contractCents = !consumer && hasValue(col.contractualCosts)
      ? euros(col.contractualCosts, "contractualCosts", 0)
      : 0;

    let previousDeadline = 0;
    let statedCents = 0;
    let receivedEnd = 0;
    if (prev) {
      previousDeadline = dateOf(prev.deadline, "previousLetter.deadline");
      if (hasValue(prev.statedCosts)) statedCents = euros(prev.statedCosts, "previousLetter.statedCosts", 0);
      if (hasValue(prev.receivedDate)) {
        receivedEnd = roll(dateOf(prev.receivedDate, "previousLetter.receivedDate") + RECEIVED_DAYS);
      }
    }

    const agency = col.agency;
    if (agency) {
      named(agency.name, "agency.name");
      named(agency.registrationNumber, "agency.registrationNumber");
    }

    // Interest joins the base only once it has run a full calendar year.
    let interestCents = 0;
    if (interestInput > 0 && sinceDay !== undefined) {
      const since = new Date(sinceDay * 86_400_000);
      const anniversary = Date.UTC(since.getUTCFullYear() + 1, since.getUTCMonth(), since.getUTCDate()) / 86_400_000;
      if (anniversary <= send) interestCents = interestInput;
    }
    const baseCents = principalCents + interestCents;

    let outcome = "none";
    let costsCents = 0;
    let deadline = 0;
    if (consumer && !prev) {
      outcome = "consumer-notice";
      if (instalment && baseCents < INSTALMENT_CEILING) {
        costsCents = instalment.priorCharged > 0 && send < windowEnd ? INSTALMENT_LATER : INSTALMENT_FIRST;
      } else {
        costsCents = staffel(baseCents);
      }
      deadline = roll(send + (col.channel === "post" ? postDays : 1) + NOTICE_DAYS);
    } else if (consumer) {
      const due = send > previousDeadline && (receivedEnd === 0 || send > receivedEnd);
      if (due && statedCents > 0) {
        outcome = "consumer-due";
        costsCents = statedCents;
      }
    } else if (contractCents > 0) {
      outcome = "business-contract";
      costsCents = contractCents;
    } else {
      const fixed = FIXED_PER_INVOICE * col.invoices.length;
      const s = staffel(baseCents);
      outcome = fixed > s ? "business-fixed" : "business-staffel";
      costsCents = fixed > s ? fixed : s;
    }

    // The stated figure of an earlier letter already includes any btw.
    const vatCents = col.creditorCanReclaimVat === false && outcome !== "consumer-due" && outcome !== "none"
      ? half(costsCents * VAT_PERCENT, 100)
      : 0;
    const costsTotalCents = costsCents + vatCents;
    const totalCents = baseCents + costsTotalCents;

    const none = outcome === "none";
    const deadlineText = outcome === "consumer-notice" ? fmtDay(deadline) : "";
    return {
      outcome,
      invoices: new Intl.ListFormat("nl-NL", { type: "conjunction" }).format(numbers),
      invoiceCount: numbers.length,
      principal: fmtMoney(principalCents),
      interest: none || interestCents === 0 ? "" : fmtMoney(interestCents),
      payNow: none ? "" : fmtMoney(baseCents),
      costs: none ? "" : fmtMoney(costsCents),
      vat: vatCents === 0 ? "" : fmtMoney(vatCents),
      costsPhrase: none
        ? ""
        : vatCents === 0
        ? fmtMoney(costsCents)
        : fmtMoney(costsTotalCents) + " (inclusief " + fmtMoney(vatCents) + " btw)",
      total: none ? "" : fmtMoney(totalCents),
      totalLabel: none ? "" : deadlineText ? "Totaal als u niet uiterlijk op " + deadlineText + " betaalt" : "Totaal",
      deadline: deadlineText,
      previousDeadline: outcome === "consumer-due" ? fmtDay(previousDeadline) : "",
      baseCents,
      costsCents,
      vatCents,
      totalCents,
    };
  },
});
