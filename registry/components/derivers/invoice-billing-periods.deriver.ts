import { deriver } from "docxcelerate/template";

/**
 * Works out the periods a subscription invoice bills for. Installed by
 * `dxcl add invoice-billing-periods`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The month
 * names, the helpers and the validation all live inside `run`.
 */
export default deriver({
  name: "invoiceBillingPeriods",
  run: (billing: unknown) => {
    // deno-lint-ignore no-explicit-any
    const b = billing as any;
    type Period = { start: string; end: string; part: boolean };

    if (typeof b !== "object" || b === null) throw new Error("billing must be an object");

    const intervals = ["week", "month", "year"];
    if (!intervals.includes(b.interval)) {
      throw new Error("billing.interval must be one of " + intervals.join(", "));
    }
    const interval: string = b.interval;
    const count: number = b.intervalCount === undefined || b.intervalCount === null ? 1 : b.intervalCount;
    if (!Number.isInteger(count) || count < 1) {
      throw new Error("billing.intervalCount must be a whole number of at least 1");
    }
    const timeZone = b.timeZone;
    if (typeof timeZone !== "string" || timeZone === "") {
      throw new Error("billing.timeZone is required: an IANA zone such as Europe/London");
    }
    try {
      new Intl.DateTimeFormat("en-US", { timeZone });
    } catch {
      throw new Error("billing.timeZone is not a time zone this runtime knows: " + timeZone);
    }
    const advanceItems = b.hasAdvanceItems === true;
    const arrearsItems = b.hasArrearsItems === true;
    if (!advanceItems && !arrearsItems) {
      throw new Error("billing.hasAdvanceItems or billing.hasArrearsItems must be true");
    }

    // Day arithmetic runs on serial day numbers. UTC is used here only as a
    // calendar with no daylight saving, so a day is always exactly 86400000 ms.
    // No instant is ever formatted in UTC: instants are turned into calendar
    // dates in the billing time zone, below, before they get here.
    const serial = (y: number, m: number, d: number) => Date.UTC(y, m - 1, d) / 86400000;
    const daysIn = (y: number, m: number) => new Date(Date.UTC(y, m, 0)).getUTCDate();
    const parts = (n: number) => {
      const dt = new Date(n * 86400000);
      return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
    };
    const iso = (n: number) => {
      const p = parts(n);
      return String(p.y).padStart(4, "0") + "-" + String(p.m).padStart(2, "0") + "-" + String(p.d).padStart(2, "0");
    };

    const toDay = (value: unknown, path: string): number => {
      if (typeof value !== "string") throw new Error(path + " is required, as a string");
      if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
        const y = Number(value.slice(0, 4));
        const m = Number(value.slice(5, 7));
        const d = Number(value.slice(8, 10));
        if (m < 1 || m > 12 || d < 1 || d > daysIn(y, m)) throw new Error(path + " is not a real date: " + value);
        return serial(y, m, d);
      }
      if (/^\d{4}-\d{2}-\d{2}T.+(Z|[+-]\d{2}:?\d{2})$/.test(value)) {
        const instant = new Date(value);
        if (Number.isNaN(instant.getTime())) throw new Error(path + " is not a real date: " + value);
        const got: Record<string, number> = {};
        for (
          const part of new Intl.DateTimeFormat("en-US", {
            timeZone,
            year: "numeric",
            month: "numeric",
            day: "numeric",
          }).formatToParts(instant)
        ) {
          got[part.type] = Number(part.value);
        }
        return serial(got.year, got.month, got.day);
      }
      // A date-time with no offset would be read in the engine's own zone, which is unknown.
      throw new Error(path + " must be YYYY-MM-DD, or a date-time with Z or an offset: " + value);
    };

    const anchor = toDay(b.anchor, "billing.anchor");
    const invoice = toDay(b.invoiceDate, "billing.invoiceDate");
    const signup = b.signupDate === undefined || b.signupDate === null ? null : toDay(b.signupDate, "billing.signupDate");
    if (signup !== null && signup > invoice) throw new Error("billing.signupDate is after billing.invoiceDate");

    const a = parts(anchor);
    // The start of period k, for any integer k. Always counted from the
    // original anchor, so 31 Mar follows 28 Feb rather than 28 Mar.
    const occ = (k: number): number => {
      if (interval === "week") return anchor + 7 * count * k;
      if (interval === "year") {
        const y = a.y + k * count;
        return serial(y, a.m, Math.min(a.d, daysIn(y, a.m)));
      }
      const t = (a.m - 1) + k * count;
      const y = a.y + Math.floor(t / 12);
      const m = ((t % 12) + 12) % 12 + 1;
      return serial(y, m, Math.min(a.d, daysIn(y, m)));
    };
    const periodIndex = (day: number): number => {
      if (interval === "week") return Math.floor((day - anchor) / (7 * count));
      const p = parts(day);
      const diff = interval === "year" ? p.y - a.y : (p.y * 12 + p.m) - (a.y * 12 + a.m);
      let k = Math.floor(diff / count);
      while (occ(k) > day) k--;
      while (occ(k + 1) <= day) k++;
      return k;
    };
    const periodAt = (k: number): Period => {
      const start = occ(k);
      const end = occ(k + 1) - 1;
      // A subscription that began inside the period is billed from its start.
      if (signup !== null && start < signup && signup <= end) {
        return { start: iso(signup), end: iso(end), part: true };
      }
      return { start: iso(start), end: iso(end), part: false };
    };

    const k = periodIndex(invoice);
    const advance = advanceItems ? periodAt(k) : null;
    let arrears: Period | null = null;
    if (arrearsItems) {
      if (signup !== null && signup > occ(k) - 1) {
        throw new Error("billing.signupDate leaves no earlier period to bill usage for");
      }
      arrears = periodAt(k - 1);
    }

    const stated = [advance, arrears].filter((p): p is Period => p !== null);
    const header = {
      start: stated.map((p) => p.start).reduce((x, y) => (y < x ? y : x)),
      end: stated.map((p) => p.end).reduce((x, y) => (y > x ? y : x)),
    };

    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const range = (start: string, end: string) => {
      const s = { y: Number(start.slice(0, 4)), m: Number(start.slice(5, 7)), d: Number(start.slice(8, 10)) };
      const e = { y: Number(end.slice(0, 4)), m: Number(end.slice(5, 7)), d: Number(end.slice(8, 10)) };
      const sm = months[s.m - 1];
      const em = months[e.m - 1];
      if (start === end) return s.d + " " + sm + " " + s.y;
      if (s.y === e.y && s.m === e.m) return s.d + "–" + e.d + " " + em + " " + e.y;
      if (s.y === e.y) return s.d + " " + sm + " – " + e.d + " " + em + " " + e.y;
      return s.d + " " + sm + " " + s.y + " – " + e.d + " " + em + " " + e.y;
    };
    const say = (label: string, p: Period, how: string) =>
      label + " " + range(p.start, p.end) + (p.part ? " (part period)" : "") + ", billed " + how + ".";

    const sentences: string[] = [];
    if (advance) sentences.push(say("Subscription", advance, "in advance"));
    if (arrears) sentences.push(say("Usage", arrears, "in arrears"));

    return {
      statement: sentences.join(" "),
      // With one period the header would only repeat the statement.
      headerLine: advance && arrears ? "Invoicing period: " + range(header.start, header.end) + "." : null,
      header,
      advance,
      arrears,
    };
  },
});
