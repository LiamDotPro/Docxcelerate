// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Works out every date and amount of a Belgian debt reminder: which stage the
 * law allows, the exact last day of the free period, and what may be added
 * after it. Installed by `dxcl add be-consumer-debt-reminder`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "beConsumerDebtReminder",
  run: (debt: unknown) => {
    const d = debt as any;

    // UTC day indexes, so a day does not shift with the engine's time zone.
    const isoDay = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      return value.slice(0, 10);
    };
    const dayIndex = (iso: string) => {
      const [y, m, day] = iso.split("-").map(Number);
      return Date.UTC(y, m - 1, day) / 86_400_000;
    };
    const dates = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" });
    const fmtDay = (index: number) => dates.format(new Date(index * 86_400_000));
    const money = new Intl.NumberFormat("en-IE", { style: "currency", currency: "EUR" });
    const fmtMoney = (cents: number) => money.format(cents / 100);
    const number = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isFinite(value)) throw new Error(field + " must be a finite number");
      return value;
    };
    const optional = (value: unknown, field: string): number | undefined =>
      value === undefined || value === null ? undefined : number(value, field);
    const nonNegative = (value: unknown, field: string): number | undefined => {
      const n = optional(value, field);
      if (n !== undefined && n < 0) throw new Error(field + " must not be negative");
      return n;
    };
    const present = (value: unknown) => typeof value === "string" && value !== "";

    // 1. Validate.
    const debtorType = d?.debtorType;
    if (debtorType !== "consumer" && debtorType !== "business") {
      throw new Error('debt.debtorType must be "consumer" or "business"');
    }
    const consumer = debtorType === "consumer";
    if (consumer && d.stage !== "first" && d.stage !== "subsequent") {
      throw new Error('debt.stage must be "first" or "subsequent"');
    }
    const asOf = dayIndex(isoDay(d.asOf, "debt.asOf"));
    const dueDate = dayIndex(isoDay(d.dueDate, "debt.dueDate"));
    const amountCents = Math.round(number(d.amountDue, "debt.amountDue") * 100);
    const agreedLumpSum = nonNegative(d.agreedLumpSum, "debt.agreedLumpSum");
    const agreedRate = nonNegative(d.agreedRate, "debt.agreedRate");
    const referenceRate = optional(d.referenceRate, "debt.referenceRate");
    const extraHolidays: number[] = [];
    if (d.holidays !== undefined && d.holidays !== null) {
      if (!Array.isArray(d.holidays)) throw new Error("debt.holidays must be a list of ISO 8601 dates");
      d.holidays.forEach((h: unknown, i: number) => extraHolidays.push(dayIndex(isoDay(h, "debt.holidays[" + i + "]"))));
    }
    let invoiceCount = 1;
    if (d.invoiceCount !== undefined && d.invoiceCount !== null) {
      if (!Number.isInteger(d.invoiceCount) || d.invoiceCount < 1) {
        throw new Error("debt.invoiceCount must be an integer of 1 or more");
      }
      invoiceCount = d.invoiceCount;
    }
    const fr = d.firstReminder;
    let sent = 0;
    let channel = "";
    if (fr !== undefined && fr !== null) {
      sent = dayIndex(isoDay(fr.sentDate, "debt.firstReminder.sentDate"));
      channel = fr.channel;
      if (channel !== "paper" && channel !== "electronic") {
        throw new Error('debt.firstReminder.channel must be "paper" or "electronic"');
      }
    }
    const hasFirst = channel !== "";

    const out = {
      outcome: "nothing-due" as string,
      amountDue: "",
      dueDate: "",
      periodStart: "",
      lastDay: "",
      lumpSum: "",
      rate: "",
      interestFrom: "",
      interest: "",
      interestDays: 0,
      interestIncomplete: false,
      charged: false,
      addsNothing: false,
      missing: "",
      compensation: "",
      invoicesLabel: "",
      total: "",
    };

    // 3. Nothing due.
    if (amountCents <= 0) return out;
    out.amountDue = fmtMoney(amountCents);
    out.dueDate = fmtDay(dueDate);

    // 4. Not yet due.
    if (asOf <= dueDate) {
      out.outcome = "not-yet-due";
      return out;
    }

    // 5. A business debtor: the fixed compensation of the Act of 2 August 2002.
    if (!consumer) {
      const compensation = 4000 * invoiceCount;
      out.outcome = "business";
      out.compensation = fmtMoney(compensation);
      out.invoicesLabel = invoiceCount === 1 ? "1 invoice" : invoiceCount + " invoices";
      out.total = fmtMoney(amountCents + compensation);
      return out;
    }

    // 6. The 14-day period. Easter by the anonymous Gregorian algorithm.
    const legalHolidays = (year: number): number[] => {
      const a = year % 19;
      const b = Math.floor(year / 100);
      const c = year % 100;
      const g = Math.floor((b - Math.floor((b + 8) / 25) + 1) / 3);
      const h = (19 * a + b - Math.floor(b / 4) - g + 15) % 30;
      const l = (32 + 2 * (b % 4) + 2 * Math.floor(c / 4) - h - (c % 4)) % 7;
      const m = Math.floor((a + 11 * h + 22 * l) / 451);
      const month = Math.floor((h + l - 7 * m + 114) / 31);
      const dayOfMonth = ((h + l - 7 * m + 114) % 31) + 1;
      const easter = Date.UTC(year, month - 1, dayOfMonth) / 86_400_000;
      return [
        Date.UTC(year, 0, 1), Date.UTC(year, 4, 1), Date.UTC(year, 6, 21), Date.UTC(year, 7, 15),
        Date.UTC(year, 10, 1), Date.UTC(year, 10, 11), Date.UTC(year, 11, 25),
      ].map((t) => t / 86_400_000).concat([easter + 1, easter + 39, easter + 50]);
    };
    let periodStart = 0;
    let lastDay = 0;
    if (hasFirst) {
      if (channel === "electronic") periodStart = sent + 1;
      else {
        // Saturdays do not count: of the two readings of "jour ouvrable" this gives the later date.
        let day = sent;
        let counted = 0;
        while (counted < 3) {
          day++;
          const weekday = new Date(day * 86_400_000).getUTCDay();
          const year = new Date(day * 86_400_000).getUTCFullYear();
          if (weekday !== 0 && weekday !== 6 && !legalHolidays(year).includes(day) && !extraHolidays.includes(day)) {
            counted++;
          }
        }
        periodStart = day;
      }
      lastDay = periodStart + 13;
      out.periodStart = fmtDay(periodStart);
      out.lastDay = fmtDay(lastDay);
    }

    // 7. The lump sum, floored to the cent: each amount is a maximum.
    const lumpCents = (): number => {
      if (d.contractProvidesDamages !== true) return 0;
      const c = amountCents;
      const cap = c <= 15000
        ? 2000
        : c <= 50000
        ? 3000 + Math.floor((c - 15000) / 10)
        : Math.min(200000, 6500 + Math.floor((c - 50000) / 20));
      return agreedLumpSum === undefined ? cap : Math.min(cap, Math.round(agreedLumpSum * 100));
    };
    // 8. The rate, capped at the reference rate plus 8 points.
    let rate = 0;
    if (d.contractProvidesInterest === true && referenceRate !== undefined) {
      const cap = referenceRate + 8;
      rate = Math.round((agreedRate === undefined ? cap : Math.min(agreedRate, cap)) * 100) / 100;
      if (rate < 0) rate = 0;
    }
    const interestFromDay = hasFirst && rate > 0 ? (d.creditorIsSme === true ? sent + 1 : lastDay + 1) : 0;

    if (d.stage === "first") {
      // 9. The four statutory mentions.
      const missing: string[] = [];
      if (!present(d.creditorName)) missing.push("the creditor's name");
      if (!present(d.creditorEnterpriseNumber)) missing.push("the creditor's enterprise number");
      if (!present(d.description)) missing.push("a description of what the debt is for");
      if (!hasFirst) missing.push("the date and channel the reminder is sent by");
      if (d.contractProvidesInterest === true && referenceRate === undefined) {
        missing.push("the reference rate that caps the late interest");
      }
      if (missing.length > 0) {
        out.outcome = "first-incomplete";
        out.missing = new Intl.ListFormat("en-GB", { type: "conjunction" }).format(missing);
        return out;
      }
      const lump = lumpCents();
      out.outcome = "first";
      out.lumpSum = lump > 0 ? fmtMoney(lump) : "";
      out.rate = rate > 0 ? rate + "%" : "";
      out.interestFrom = rate > 0 ? fmtDay(interestFromDay) : "";
      out.addsNothing = !out.lumpSum && !out.rate;
      return out;
    }

    // 10. A later reminder.
    if (!hasFirst) {
      out.outcome = "no-first-reminder";
      return out;
    }
    if (asOf <= lastDay) {
      out.outcome = "waiting";
      return out;
    }
    out.outcome = "lapsed";
    const lump = lumpCents();
    out.lumpSum = lump > 0 ? fmtMoney(lump) : "";
    let interestCents = 0;
    if (rate > 0) {
      const days = asOf - (interestFromDay - 1);
      // The epsilon stops a float such as 9000.999999 losing a cent.
      interestCents = Math.floor((amountCents * rate * days) / 36500 + 1e-9);
      out.rate = rate + "%";
      out.interestFrom = fmtDay(interestFromDay);
      out.interestDays = days;
      out.interest = interestCents > 0 ? fmtMoney(interestCents) : "";
    }
    out.interestIncomplete = d.contractProvidesInterest === true && referenceRate === undefined;
    out.charged = Boolean(out.lumpSum || out.interest);
    out.total = fmtMoney(amountCents + lump + interestCents);
    return out;
  },
});
