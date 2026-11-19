import { deriver } from "docxcelerate/template";

/**
 * Works out the Construction Act payment timetable for one interim payment.
 * Installed by `dxcl add construction-payment-timetable`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level helper or constant would be a ReferenceError there. The Easter
 * calculation, the day counting and the formatter all live inside `run`.
 */
export default deriver({
  name: "constructionPaymentTimetable",
  run: (contract: unknown, payment: unknown, bankHolidays: unknown) => {
    // deno-lint-ignore no-explicit-any
    const con = contract as any;
    // deno-lint-ignore no-explicit-any
    const pay = payment as any;
    // deno-lint-ignore no-explicit-any
    const bh = bankHolidays as any;

    const regions = ["EW", "SC", "NI"];
    const region: string = con?.region;
    if (!regions.includes(region)) throw new Error("contract.region must be one of " + regions.join(", "));
    const act = region === "NI"
      ? "Construction Contracts (Northern Ireland) Order 1997"
      : "Housing Grants, Construction and Regeneration Act 1996";
    const regionName = { EW: "England and Wales", SC: "Scotland", NI: "Northern Ireland" }[region];
    const scheme = "the Scheme for Construction Contracts (" + regionName + ")";

    if (con.residentialOccupier === true) {
      return {
        applies: false,
        notApplicable: "This contract is with a residential occupier, so the " + act +
          " does not apply to it, and neither does the payment timetable it sets.",
        basis: null,
        dueDate: null,
        paymentNoticeBy: null,
        finalDate: null,
        payLessBy: null,
        noticeLate: null,
        noticeGiven: null,
        applicationStands: null,
        defaultNotice: null,
        defaultOpen: null,
        reckoning: null,
      };
    }

    const DAY = 86400000;
    const missing = (v: unknown) => v === undefined || v === null || v === "";
    // UTC midnight, so no local time or daylight saving ever enters the sum.
    const parse = (v: unknown, path: string): number => {
      const text = typeof v === "string" ? v.slice(0, 10) : "";
      const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(text);
      const t = m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) : NaN;
      if (!m || new Date(t).toISOString().slice(0, 10) !== text) {
        throw new Error(path + " must be an ISO 8601 date");
      }
      return t;
    };
    const days = (v: unknown, name: string, min: number, fallback: number | undefined): number | undefined => {
      if (v === undefined || v === null) return fallback;
      if (!Number.isInteger(v) || (v as number) < min) {
        throw new Error("contract." + name + " must be a whole number of days, at least " + min);
      }
      return v as number;
    };
    const dueDays = days(con.dueDateDays, "dueDateDays", 0, undefined);
    const finalDays = days(con.finalDateDays, "finalDateDays", 1, 17) as number;
    const payLessDays = days(con.payLessDays, "payLessDays", 0, 7) as number;

    const list = bh?.[region];
    if (!Array.isArray(list)) {
      throw new Error("bankHolidays." + region + " is required for a contract in " + regionName);
    }
    const holidays = new Set<number>(
      list.map((d: unknown, i: number) => parse(d, "bankHolidays." + region + "[" + i + "]")),
    );

    // Anonymous Gregorian algorithm (Meeus/Jones/Butcher).
    const goodFriday = (year: number): number => {
      const a = year % 19, b = Math.floor(year / 100), c = year % 100;
      const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
      const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
      const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
      const m = Math.floor((a + 11 * h + 22 * l) / 451);
      const month = Math.floor((h + l - 7 * m + 114) / 31);
      const day = ((h + l - 7 * m + 114) % 31) + 1;
      return Date.UTC(year, month - 1, day) - 2 * DAY;
    };
    const excluded = (t: number): boolean => {
      const d = new Date(t);
      return (d.getUTCMonth() === 11 && d.getUTCDate() === 25) ||
        t === goodFriday(d.getUTCFullYear()) || holidays.has(t);
    };
    const step = (from: number, n: number, dir: number): number => {
      let t = from;
      while (n > 0) {
        t += dir * DAY;
        if (!excluded(t)) n--;
      }
      return t;
    };
    const addDays = (t: number, n: number) => step(t, n, 1);
    const subDays = (t: number, n: number) => step(t, n, -1);
    const daysBetween = (a: number, b: number): number => {
      let n = 0;
      for (let t = a + DAY; t <= b; t += DAY) if (!excluded(t)) n++;
      return n;
    };
    const fmt = new Intl.DateTimeFormat("en-GB", {
      weekday: "long",
      day: "numeric",
      month: "long",
      year: "numeric",
      timeZone: "UTC",
    });
    const show = (t: number) => fmt.format(new Date(t));
    const iso = (t: number) => new Date(t).toISOString().slice(0, 10);

    const V = parse(pay?.valuationPeriodEnd, "payment.valuationPeriodEnd");
    const hasClaim = !missing(pay.claimReceived);
    const C = hasClaim ? parse(pay.claimReceived, "payment.claimReceived") : undefined;

    let D0: number;
    if (dueDays !== undefined) {
      D0 = addDays(V, dueDays);
    } else {
      if (C === undefined) {
        throw new Error("payment.claimReceived is required when contract.dueDateDays is not given");
      }
      D0 = Math.max(addDays(V, 7), C);
    }
    const PN = addDays(D0, 5);
    const F0 = addDays(D0, finalDays);

    const P = missing(pay.paymentNoticeIssued) ? undefined : parse(pay.paymentNoticeIssued, "payment.paymentNoticeIssued");
    const valid = P !== undefined && P <= PN;
    let F = F0;
    let latest = F0;
    let noticeLate: string | null = null;
    let noticeGiven: string | null = null;
    let applicationStands: string | null = null;
    let defaultNotice: string | null = null;
    let defaultOpen: string | null = null;
    let open = false;

    if (P !== undefined && !valid) {
      noticeLate = "The payment notice given on " + show(P) + " came after the deadline of " + show(PN) +
        ", so it does not count as one.";
    }
    if (valid) {
      noticeGiven = "The payer gave its payment notice on " + show(P) +
        ". The sum it states is the notified sum.";
    } else if (C !== undefined && C < PN) {
      applicationStands = "No payment notice was given by " + show(PN) + ", so the payment claim received on " +
        show(C) + " stands as the notified sum.";
    } else if (!missing(pay.defaultNoticeIssued)) {
      const N = parse(pay.defaultNoticeIssued, "payment.defaultNoticeIssued");
      if (N <= PN) {
        throw new Error("payment.defaultNoticeIssued must be after the payment notice deadline, " + iso(PN));
      }
      const k = daysBetween(PN, N);
      F = addDays(F0, k);
      latest = Math.max(F, N);
      const lead = "No payment notice was given by " + show(PN) + ". The payee's notice of " + show(N) +
        " stands as the notified sum";
      defaultNotice = k === 0
        ? lead + ". The final date for payment does not move."
        : lead + ", and moves the final date for payment back " + k + (k === 1 ? " day" : " days") +
          ", from " + show(F0) + " to " + show(F) + ".";
    } else {
      open = true;
      defaultOpen = "No payment notice was given by " + show(PN) +
        ", and the payee has not yet given a notice in its place. When it does, the final date for payment moves back one day for each day after " +
        show(PN) + " the notice is given, and the pay less notice deadline with it.";
    }

    const PL = subDays(F, payLessDays);
    if (PL < PN) {
      throw new Error(
        "the pay less notice deadline falls before the payment notice deadline; check contract.finalDateDays and contract.payLessDays",
      );
    }

    // A year with no bank holidays means the list is stale, and every count
    // through it would be quietly wrong.
    const years = new Set<number>([...holidays].map((t) => new Date(t).getUTCFullYear()));
    const last = new Date(latest).getUTCFullYear();
    for (let year = new Date(V).getUTCFullYear(); year <= last; year++) {
      if (!years.has(year)) throw new Error("bankHolidays." + region + " has no dates in " + year);
    }

    const earliest = open ? " at the earliest" : "";
    const basis = dueDays === undefined
      ? "Under " + scheme + ", payment falls due on the later of 7 days after the valuation period ended on " +
        show(V) + " and receipt of the payment claim on " + show(C as number) + "."
      : dueDays === 0
      ? "Under the contract, payment falls due on the day the valuation period ended, " + show(V) + "."
      : "Under the contract, payment falls due " + dueDays + (dueDays === 1 ? " day" : " days") +
        " after the valuation period ended on " + show(V) + ".";
    const reckoning = region === "NI"
      ? "Days are counted under the " + act +
        ": Christmas Day, Good Friday and bank holidays in Northern Ireland are not counted. Weekends are."
      : "Days are counted under section 116 of the " + act + ": Christmas Day, Good Friday and bank holidays in " +
        regionName + " are not counted. Weekends are.";

    return {
      applies: true,
      notApplicable: null,
      basis,
      dueDate: show(D0),
      paymentNoticeBy: show(PN),
      finalDate: show(F) + earliest,
      payLessBy: show(PL) + earliest,
      noticeLate,
      noticeGiven,
      applicationStands,
      defaultNotice,
      defaultOpen,
      reckoning,
    };
  },
});
