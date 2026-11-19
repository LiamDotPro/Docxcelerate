// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

type Window = {
  anchor: "renewal" | "cancel-by" | "trial-end" | "price-change";
  from: string;
  to: string;
  status: "early" | "ok" | "late";
};

type Result = {
  regime: "us-ca-arl" | "us-ny-5-903" | "de-consumer" | "uk-dmcc" | "contractual";
  heading: string;
  statement: string;
  renewal: string;
  trial: string | null;
  priceChange: string | null;
  previousPayment: string | null;
  cancelBy: string | null;
  consequence: string | null;
  warning: string | null;
  timing: "early" | "ok" | "late" | null;
  cancelByIso: string | null;
  windows: Window[];
};

/**
 * Decides everything a renewal or price-change notice says: which law applies,
 * the last day to cancel, whether the notice is early or late for that law's
 * window, and the wording each case gets. Installed by
 * `dxcl add renewal-and-price-change-notice`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper and threshold is declared inside `run`.
 */
export default deriver({
  name: "renewalAndPriceChangeNotice",
  run: (input: unknown): Result => {
    const r = input as any;
    const DAY = 86_400_000;
    const dayOf = (iso: string) => {
      const [y, m, d] = iso.split("-").map(Number);
      return Date.UTC(y, m - 1, d) / DAY;
    };
    const isoOf = (day: number) => new Date(day * DAY).toISOString().slice(0, 10);
    const isoDate = (value: unknown, field: string): string => {
      const iso = typeof value === "string" ? value.slice(0, 10) : "";
      if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) {
        const [y, m, d] = iso.split("-").map(Number);
        const t = new Date(Date.UTC(y, m - 1, d));
        if (t.getUTCFullYear() === y && t.getUTCMonth() === m - 1 && t.getUTCDate() === d) return iso;
      }
      throw new Error(field + " must be an ISO 8601 date");
    };
    const optDate = (value: unknown, field: string) =>
      value === undefined || value === null ? undefined : isoDate(value, field);
    const optMoney = (value: unknown, field: string) => {
      if (value === undefined || value === null) return undefined;
      if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
        throw new Error(field + " must be a finite number of 0 or more");
      }
      return value;
    };
    const wholeMonths = (value: unknown, field: string) => {
      if (typeof value !== "number" || !Number.isInteger(value) || value < 1) {
        throw new Error(field + " must be a positive whole number of months");
      }
      return value;
    };
    // Month arithmetic clamps to the end of the month: 31 March less a month is 28 or 29 February.
    const addMonths = (iso: string, n: number) => {
      const [y, m, d] = iso.split("-").map(Number);
      const total = y * 12 + (m - 1) + n;
      const ny = Math.floor(total / 12);
      const nm = ((total % 12) + 12) % 12;
      const last = new Date(Date.UTC(ny, nm + 1, 0)).getUTCDate();
      return isoOf(Date.UTC(ny, nm, Math.min(d, last)) / DAY);
    };

    // 1. Validate, in the order the fields are documented.
    const jurisdictions = ["US-CA", "US-NY", "DE", "UK", "other"];
    if (!jurisdictions.includes(r?.jurisdiction)) {
      throw new Error("renewal.jurisdiction must be one of " + jurisdictions.join(", "));
    }
    if (r.customerType !== "consumer" && r.customerType !== "business") {
      throw new Error("renewal.customerType must be consumer or business");
    }
    const contractStart = isoDate(r.contractStart, "renewal.contractStart");
    const amendedOn = optDate(r.amendedOn, "renewal.amendedOn");
    const initialTermMonths = wholeMonths(r.initialTermMonths, "renewal.initialTermMonths");
    const renewalTermMonths = wholeMonths(r.renewalTermMonths, "renewal.renewalTermMonths");
    const renewalDate = isoDate(r.renewalDate, "renewal.renewalDate");
    if (renewalDate < contractStart) throw new Error("renewal.renewalDate is before renewal.contractStart");
    const np = r.noticePeriod;
    if (
      typeof np?.value !== "number" || !Number.isInteger(np.value) || np.value < 0 ||
      (np.unit !== "days" && np.unit !== "months")
    ) {
      throw new Error("renewal.noticePeriod must be a non-negative whole value in days or months");
    }
    const notice: { value: number; unit: "days" | "months" } = { value: np.value, unit: np.unit };
    const sendDate = isoDate(r.sendDate, "renewal.sendDate");
    const jurisdiction: string = r.jurisdiction;
    const consumer = r.customerType === "consumer";
    const locale = jurisdiction === "US-CA" || jurisdiction === "US-NY" ? "en-US" : "en-GB";
    let money: Intl.NumberFormat;
    try {
      money = new Intl.NumberFormat(locale, { style: "currency", currency: r.currency });
    } catch {
      throw new Error("renewal.currency must be an ISO 4217 code");
    }
    const exp = money.resolvedOptions().maximumFractionDigits as number;
    const currentAmount = optMoney(r.currentAmount ?? NaN, "renewal.currentAmount") as number;
    const previousAmount = optMoney(r.previousAmount, "renewal.previousAmount");
    const pc = r.priceChange;
    if (pc !== undefined && pc !== null) {
      optMoney(pc.newAmount ?? NaN, "renewal.priceChange.newAmount");
      isoDate(pc.effectiveDate, "renewal.priceChange.effectiveDate");
    }
    const trialEnds = optDate(r.trialEnds, "renewal.trialEnds");
    if (typeof r.cancelInstructions !== "string" || r.cancelInstructions.trim() === "") {
      throw new Error("renewal.cancelInstructions is required");
    }
    const commencement = optDate(r.commencementDate, "renewal.commencementDate") ?? "2027-01-01";

    const dates = new Intl.DateTimeFormat(locale, { dateStyle: "long", timeZone: "UTC" });
    const fmt = (iso: string) => dates.format(new Date(iso + "T00:00:00Z"));
    const toMinor = (x: number) => Math.round(Number((x * 10 ** exp).toPrecision(15)));
    const fmtMinor = (minor: number) => money.format(minor / 10 ** exp);
    const noun = consumer ? "subscription" : "contract";
    const frequency = renewalTermMonths === 1
      ? "every month"
      : renewalTermMonths === 12
      ? "every year"
      : renewalTermMonths % 12 === 0
      ? "every " + renewalTermMonths / 12 + " years"
      : "every " + renewalTermMonths + " months";
    // The direction is a word, so "-0.0%" cannot appear.
    const pctText = (diff: number, old: number) => {
      if (old === 0) return null;
      const p = Math.round(Number(((Math.abs(diff) / old) * 100 * 10).toPrecision(15))) / 10;
      return p === 0 && diff !== 0 ? "less than 0.1%" : p.toFixed(1) + "%";
    };
    const periodText = (p: { value: number; unit: string }) =>
      p.value === 0 || (p.unit === "days" && p.value === 1)
        ? "1 day"
        : p.unit === "days"
        ? p.value + " days"
        : p.value === 1
        ? "one month"
        : p.value + " months";
    const minus = (iso: string, p: { value: number; unit: string }) =>
      p.unit === "days" || p.value === 0 ? isoOf(dayOf(iso) - Math.max(p.value, 1)) : addMonths(iso, -p.value);

    // 2. Regime: the first rule that matches.
    const lastChange = amendedOn && amendedOn > contractStart ? amendedOn : contractStart;
    const regime: Result["regime"] = jurisdiction === "US-NY" && renewalTermMonths > 1
      ? "us-ny-5-903"
      : jurisdiction === "US-CA" && consumer && lastChange >= "2025-07-01"
      ? "us-ca-arl"
      : jurisdiction === "DE" && consumer && contractStart >= "2022-03-01"
      ? "de-consumer"
      : jurisdiction === "UK" && consumer && contractStart >= commencement
      ? "uk-dmcc"
      : "contractual";

    // A German contract past its minimum term has no renewal date to measure from.
    const termEnd = addMonths(contractStart, initialTermMonths);
    const indefinite = regime === "de-consumer" && sendDate >= termEnd;
    if (!indefinite && renewalDate < sendDate) throw new Error("renewal.renewalDate is before renewal.sendDate");

    // 3-4. The configured deadline, and the German cap on it.
    let period = notice;
    const reference = indefinite ? termEnd : renewalDate;
    let cancelByRaw = minus(reference, period);
    const warnings: string[] = [];
    if (regime === "de-consumer" && cancelByRaw < addMonths(reference, -1)) {
      warnings.push(
        "Check before sending: the contract's notice period of " + periodText(notice) +
          " is longer than the one month German law allows for a consumer contract (BGB §309 No. 9); this notice uses one month.",
      );
      period = { value: 1, unit: "months" };
      cancelByRaw = addMonths(reference, -1);
    }
    // 5. The deadline in use.
    const cancelByIso = indefinite ? null : cancelByRaw;

    // 6. The lines that decide windows.
    const current = toMinor(currentAmount);
    const change = pc !== undefined && pc !== null && toMinor(pc.newAmount) !== current ? pc : null;
    const amountOn = (iso: string) => change && change.effectiveDate.slice(0, 10) <= iso ? toMinor(change.newAmount) : current;
    const renewalMinor = indefinite ? amountOn(sendDate) : amountOn(renewalDate);
    const trialShown = trialEnds !== undefined && trialEnds >= sendDate;
    const trial = trialShown ? fmt(trialEnds) + "; from then " + fmtMinor(amountOn(trialEnds)) + " " + frequency : null;
    let priceChange: string | null = null;
    if (change) {
      const diff = toMinor(change.newAmount) - current;
      const pct = pctText(diff, current);
      priceChange = "From " + fmt(change.effectiveDate.slice(0, 10)) + ": " + fmtMinor(current) + " to " +
        fmtMinor(toMinor(change.newAmount)) + ", " + (diff > 0 ? "an increase of " : "a decrease of ") +
        fmtMinor(Math.abs(diff)) + (pct ? " (" + pct + ")" : "");
    }
    let previousPayment: string | null = null;
    if (previousAmount !== undefined) {
      const prev = toMinor(previousAmount);
      previousPayment = fmtMinor(prev);
      if (renewalMinor > prev) {
        const pct = pctText(renewalMinor - prev, prev);
        previousPayment += ". This renewal is " + fmtMinor(renewalMinor - prev) + (pct ? " (" + pct + ")" : "") + " more";
      }
    } else if (regime === "uk-dmcc") {
      warnings.push(
        "Check before sending: a UK subscription reminder must show the customer's previous payment, and none was given.",
      );
    }

    // 7. Windows and timing. Each is measured from its own anchor.
    const windows: Window[] = [];
    const phrase: Record<string, string> = {
      renewal: "the renewal",
      "cancel-by": "the cancellation deadline",
      "trial-end": "your introductory period ends",
      "price-change": "the new price takes effect",
    };
    const bounds: Record<string, [number, number]> = {
      renewal: [45, 15],
      "cancel-by": [30, 15],
      "trial-end": [21, 3],
      "price-change": [30, 7],
    };
    const anchors: Record<string, string> = {};
    const addWindow = (anchor: Window["anchor"], date: string) => {
      const [far, near] = bounds[anchor];
      const from = isoOf(dayOf(date) - far);
      const to = isoOf(dayOf(date) - near);
      anchors[anchor] = date;
      windows.push({ anchor, from, to, status: sendDate < from ? "early" : sendDate > to ? "late" : "ok" });
    };
    if (regime === "us-ca-arl") {
      if (initialTermMonths >= 12) addWindow("renewal", renewalDate);
      if (trialShown && dayOf(trialEnds) - dayOf(contractStart) > 31) addWindow("trial-end", trialEnds);
      if (change) addWindow("price-change", change.effectiveDate.slice(0, 10));
    } else if (regime === "us-ny-5-903") {
      addWindow("cancel-by", cancelByRaw);
    }
    const rank = { ok: 0, early: 1, late: 2 };
    const timing = windows.length === 0
      ? null
      : windows.reduce((worst, w) => rank[w.status] > rank[worst] ? w.status : worst, "ok" as Window["status"]);

    // 8. Statement, cancel-by, consequence and warning.
    const nyLate = regime === "us-ny-5-903" && timing === "late";
    const statement = regime === "de-consumer"
      ? indefinite
        ? "Your " + noun + " has passed its minimum term and continues until you cancel it."
        : "Your " + noun + "'s minimum term ends on " + fmt(renewalDate) +
          ". Unless you cancel, it then continues until you cancel it, with " + periodText(period) + " notice."
      : nyLate
      ? "Your " + noun + " was set to renew automatically on " + fmt(renewalDate) + "."
      : "Your " + noun + " renews automatically on " + fmt(renewalDate) + " unless you cancel it.";
    const renewal = indefinite
      ? "Continues until you cancel, at " + fmtMinor(renewalMinor) + " " + frequency
      : fmt(renewalDate) + ", for " + fmtMinor(renewalMinor) + " " + frequency;

    const passed = cancelByIso !== null && cancelByIso < sendDate;
    const cancelBy = indefinite
      ? "None. You can cancel at any time; the " + noun + " then ends " + periodText(period) + " later."
      : timing === "late" || passed
      ? null
      : "On or before " + fmt(cancelByRaw);

    let consequence: string | null = null;
    if (nyLate) {
      consequence =
        "This notice is reaching you less than 15 days before the date you would have had to cancel by, so under New York General Obligations Law §5-903 the automatic renewal cannot be enforced against you. The " +
        noun + " will not renew on " + fmt(renewalDate) + " unless you agree to continue.";
    } else if (regime === "us-ca-arl" && timing === "late") {
      consequence = windows.filter((w) => w.status === "late").map((w) =>
        "California law requires notice " + bounds[w.anchor][1] + " to " + bounds[w.anchor][0] + " days before " +
        phrase[w.anchor] + " on " + fmt(anchors[w.anchor]) +
        ", and this notice is later than that. No cancellation deadline applies: you can cancel at any time before " +
        fmt(anchors[w.anchor]) + "."
      ).join(" ");
    } else if (passed) {
      consequence = "The last day to cancel before the renewal on " + fmt(renewalDate) + " was " + fmt(cancelByRaw) +
        ", so the " + noun + " will renew. You can still cancel before the renewal after that.";
    }

    for (const w of windows) {
      if (w.status !== "early") continue;
      const n = dayOf(w.from) - dayOf(sendDate);
      warnings.push(
        "Check before sending: " + (regime === "us-ny-5-903" ? "New York" : "California") + " law requires this notice " +
          bounds[w.anchor][1] + " to " + bounds[w.anchor][0] + " days before " + phrase[w.anchor] + " on " +
          fmt(anchors[w.anchor]) + "; sent on " + fmt(sendDate) + " it is " + n + (n === 1 ? " day" : " days") + " early.",
      );
    }

    return {
      regime,
      heading: change ? "Price change notice" : trialShown ? "Your introductory period is ending" : "Renewal notice",
      statement,
      renewal,
      trial,
      priceChange,
      previousPayment,
      cancelBy,
      consequence,
      warning: warnings.length ? warnings.join(" ") : null,
      timing,
      cancelByIso,
      windows,
    };
  },
});
