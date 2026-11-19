import { deriver } from "docxcelerate/template";

/**
 * Validates a direct debit and works out what its advance notice may say: the
 * amount, the earliest lawful collection date for the scheme and the mandate
 * details. Installed by `dxcl add direct-debit-collection-notice`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The Easter
 * calculation, the TARGET2 test, the check digits and the scheme tables all
 * live inside `run`, and nothing in it needs BigInt.
 */
export default deriver({
  name: "directDebitCollectionNotice",
  run: (debitInput: unknown, issuedAtInput: unknown) => {
    // deno-lint-ignore no-explicit-any
    const debit = debitInput as any;
    const DAY = 86_400_000;

    const schemes = ["sepa-core", "bacs", "ach"];
    if (!schemes.includes(debit?.scheme)) {
      throw new Error("debit.scheme must be one of " + schemes.join(", "));
    }
    const scheme: string = debit.scheme;
    const sepa = scheme === "sepa-core";
    const bacs = scheme === "bacs";
    const ach = scheme === "ach";

    if (typeof debit.creditorName !== "string" || debit.creditorName.trim() === "") {
      throw new Error("debit.creditorName is required");
    }
    const creditorName: string = debit.creditorName.trim();

    // A UTC day number, or a throw naming the field. The round trip is what
    // refuses 2027-02-30, which Date.UTC would quietly roll into March.
    const dayOf = (value: unknown, field: string): number => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      const y = Number(value.slice(0, 4));
      const m = Number(value.slice(5, 7));
      const d = Number(value.slice(8, 10));
      const ms = Date.UTC(y, m - 1, d);
      const back = new Date(ms);
      if (back.getUTCFullYear() !== y || back.getUTCMonth() !== m - 1 || back.getUTCDate() !== d) {
        throw new Error(field + " is not a real date");
      }
      return ms / DAY;
    };
    const issued = dayOf(issuedAtInput, "document.issuedAt");
    const requested = dayOf(debit.requestedDate, "debit.requestedDate");

    if (typeof debit.amount !== "number" || !Number.isFinite(debit.amount)) {
      throw new Error("debit.amount must be a finite number");
    }
    const toMinor = (n: number) => Math.sign(n) * Math.round(Math.abs(n) * 100);
    if (ach && debit.previousAmount !== undefined && debit.previousAmount !== null) {
      if (typeof debit.previousAmount !== "number" || !Number.isFinite(debit.previousAmount)) {
        throw new Error("debit.previousAmount must be a finite number");
      }
    }

    let creditorIdLabel: string | null = null;
    let creditorId: string | null = null;
    if (sepa) {
      if (typeof debit.creditorId !== "string") throw new Error("debit.creditorId is required for sepa-core");
      const id = debit.creditorId.replace(/\s+/g, "").toUpperCase();
      if (!/^[A-Z]{2}\d{2}[A-Z0-9]{3}[A-Z0-9]{1,28}$/.test(id) || id.length > 35) {
        throw new Error("debit.creditorId is not a SEPA creditor identifier");
      }
      // The business code (characters 5 to 7) is left out of the check.
      const digits = (id.slice(7) + id.slice(0, 2) + "00")
        .replace(/[A-Z]/g, (c) => String(c.charCodeAt(0) - 55));
      let r = 0;
      for (const ch of digits) r = (r * 10 + Number(ch)) % 97;
      const expected = String(98 - r).padStart(2, "0");
      if (id.slice(2, 4) !== expected) {
        throw new Error("debit.creditorId has the wrong check digits: expected " + expected);
      }
      creditorIdLabel = "Creditor identifier";
      creditorId = id;
    } else if (bacs) {
      const sun = typeof debit.sun === "string" ? debit.sun.replace(/\s+/g, "") : "";
      if (!/^\d{6}$/.test(sun)) throw new Error("debit.sun must be a six-digit Service User Number");
      creditorIdLabel = "Service User Number";
      creditorId = sun;
    }

    if (typeof debit.mandateReference !== "string" || debit.mandateReference.trim() === "") {
      throw new Error("debit.mandateReference is required");
    }
    const mandateReference: string = debit.mandateReference.trim();
    const maxReference = sepa ? 35 : bacs ? 18 : Infinity;
    if (mandateReference.length > maxReference) {
      throw new Error("debit.mandateReference is longer than " + maxReference + " characters");
    }

    let account: string | null = null;
    const prefix = sepa ? "IBAN ending " : "Account ending ";
    if (debit.payerIban !== undefined && debit.payerIban !== null) {
      const iban = typeof debit.payerIban === "string" ? debit.payerIban.replace(/\s+/g, "").toUpperCase() : "";
      if (!/^[A-Z0-9]{4,}$/.test(iban)) throw new Error("debit.payerIban must have at least four letters or digits");
      account = prefix + iban.slice(-4);
    } else if (debit.payerAccountLast4 !== undefined && debit.payerAccountLast4 !== null) {
      if (typeof debit.payerAccountLast4 !== "string" || !/^[A-Za-z0-9]{4}$/.test(debit.payerAccountLast4)) {
        throw new Error("debit.payerAccountLast4 must be exactly four letters or digits");
      }
      account = prefix + debit.payerAccountLast4.toUpperCase();
    }

    const agreed = debit.agreedNoticeDays;
    if (agreed !== undefined && agreed !== null && !(Number.isInteger(agreed) && agreed >= 0)) {
      throw new Error("debit.agreedNoticeDays must be a non-negative integer");
    }

    const holidays: number[] = [];
    if (bacs && !Array.isArray(debit.bankHolidays)) {
      throw new Error("debit.bankHolidays is required for bacs");
    }
    if (!sepa && debit.bankHolidays !== undefined && debit.bankHolidays !== null) {
      if (!Array.isArray(debit.bankHolidays)) throw new Error("debit.bankHolidays must be an array");
      for (const h of debit.bankHolidays) holidays.push(dayOf(h, "debit.bankHolidays entry"));
    }

    const g = debit.guaranteeText;
    // deno-lint-ignore no-explicit-any
    const guaranteeMissing = bacs && (!Array.isArray(g) || g.length === 0 || g.some((e: any) => typeof e !== "string" || e.trim() === ""));

    const minor = toMinor(debit.amount);
    const outcome = minor > 0 ? "collect" : minor === 0 ? "clear" : "credit";
    const currency = sepa ? "EUR" : bacs ? "GBP" : "USD";
    const locale = ach ? "en-US" : "en-GB";
    const money = new Intl.NumberFormat(locale, { style: "currency", currency });

    const amountChanged = ach && debit.previousAmount !== undefined && debit.previousAmount !== null &&
      toMinor(debit.previousAmount) !== minor;
    const previousAmount = amountChanged ? money.format(Math.abs(toMinor(debit.previousAmount)) / 100) : null;

    // Notice length and its label. ACH labels "10 days" whether or not it applies.
    const plural = (n: number, unit: string) => n + " " + unit + (n === 1 ? "" : "s");
    const days: number = sepa
      ? (agreed ?? 14)
      : bacs
      ? Math.max(agreed ?? 10, 3)
      : amountChanged
      ? 10
      : 0;
    const noticeLabel = sepa ? plural(days, "calendar day") : bacs ? plural(days, "working day") : "10 days";

    // Whether the scheme settles on this day. TARGET2 recomputes Easter for the
    // candidate's own year, since rolling forward can cross New Year.
    const isOpen = (n: number): boolean => {
      const date = new Date(n * DAY);
      const dow = date.getUTCDay();
      if (dow === 0 || dow === 6) return false;
      if (!sepa) return !holidays.includes(n);
      const y = date.getUTCFullYear();
      const month = date.getUTCMonth() + 1;
      const dom = date.getUTCDate();
      if ((month === 1 && dom === 1) || (month === 5 && dom === 1) || (month === 12 && (dom === 25 || dom === 26))) {
        return false;
      }
      const a = y % 19;
      const b = Math.floor(y / 100);
      const c = y % 100;
      const d = Math.floor(b / 4);
      const e = b % 4;
      const f = Math.floor((b + 8) / 25);
      const gg = Math.floor((b - f + 1) / 3);
      const h = (19 * a + b - d - gg + 15) % 30;
      const i = Math.floor(c / 4);
      const k = c % 4;
      const l = (32 + 2 * e + 2 * i - h - k) % 7;
      const m = Math.floor((a + 11 * h + 22 * l) / 451);
      const easter = Date.UTC(y, Math.floor((h + l - 7 * m + 114) / 31) - 1, ((h + l - 7 * m + 114) % 31) + 1) / DAY;
      return n !== easter - 2 && n !== easter + 1;
    };

    let late = false;
    let date = "";
    if (outcome === "collect") {
      let earliest: number;
      if (sepa) {
        earliest = issued + days;
        while (!isOpen(earliest)) earliest += 1;
      } else if (bacs) {
        earliest = issued;
        for (let counted = 0; counted < days;) {
          earliest += 1;
          if (isOpen(earliest)) counted += 1;
        }
      } else {
        earliest = amountChanged ? issued + 10 : issued;
      }
      // Compared before rolling: a requested date that only needs rolling is
      // still the date the payer was promised, not a short notice.
      late = requested < earliest;
      let target = late ? earliest : requested;
      while (!isOpen(target)) target += 1;
      date = new Intl.DateTimeFormat(locale, { day: "numeric", month: "long", year: "numeric", timeZone: "UTC" })
        .format(new Date(target * DAY));
    }

    return {
      heading: sepa ? "Payment by SEPA Direct Debit" : bacs ? "Payment by Direct Debit" : "Payment by automatic debit (ACH)",
      outcome,
      schemeName: sepa ? "SEPA Direct Debit" : bacs ? "Direct Debit" : "ACH debit",
      mandateNoun: sepa ? "mandate" : bacs ? "Direct Debit Instruction" : "authorization",
      amount: money.format(Math.abs(minor) / 100),
      previousAmount,
      amountChanged,
      late,
      noticeLabel,
      date,
      creditorName,
      creditorIdLabel,
      creditorId,
      mandateLabel: sepa ? "Mandate reference" : bacs ? "Direct Debit reference" : "Authorization reference",
      mandateReference,
      account,
      guaranteeMissing,
    };
  },
});
