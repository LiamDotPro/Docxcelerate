// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Sets a customer's open credits against an invoice and works out the amount
 * due. Installed by `dxcl add credits-applied-to-invoice`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "creditsAppliedToInvoice",
  run: (invoice: unknown, credits: unknown) => {
    const inv = invoice as any;
    const cr = credits as any;

    const moneyFor = (code: unknown, field: string) => {
      let format: Intl.NumberFormat;
      try {
        format = new Intl.NumberFormat("en-GB", { style: "currency", currency: code as string });
      } catch {
        throw new Error(field + " must be an ISO 4217 code");
      }
      const exponent = format.resolvedOptions().maximumFractionDigits as number;
      return (minor: number) => format.format(minor / 10 ** exponent);
    };
    const minorOf = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isSafeInteger(value) || value < 0) {
        throw new Error(field + " must be a non-negative integer of minor units");
      }
      return value;
    };
    const isoDay = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) {
        throw new Error(field + " must be an ISO 8601 date");
      }
      return value.slice(0, 10);
    };
    const dates = new Intl.DateTimeFormat("en-GB", { dateStyle: "long", timeZone: "UTC" });
    const fmtDate = (iso: string) => dates.format(new Date(iso + "T00:00:00Z"));
    const code = (value: unknown) => String(value).trim().toUpperCase();

    // Invoice first: its currency and total decide how everything else is read.
    if (typeof inv?.currency !== "string") throw new Error("invoice.currency must be an ISO 4217 code");
    const currency = code(inv.currency);
    const money = moneyFor(currency, "invoice.currency");
    const issueDate = isoDay(inv?.issueDate, "invoice.issueDate");
    const totalMinor = minorOf(inv?.totalMinor, "invoice.totalMinor");
    let vat: string | null = null;
    if (inv.vatTotalMinor !== undefined && inv.vatTotalMinor !== null) {
      const vatMinor = minorOf(inv.vatTotalMinor, "invoice.vatTotalMinor");
      if (vatMinor > totalMinor) throw new Error("invoice.vatTotalMinor must not exceed invoice.totalMinor");
      vat = money(vatMinor);
    }

    // All validation runs before anything is decided, so bad data fails even
    // when nothing would print.
    const open: any[] = cr?.open ?? [];
    if (!Array.isArray(open)) throw new Error("credits.open must be an array");
    const seen = new Set<string>();
    const all = open.map((c: any, index: number) => {
      const at = "credits.open[" + index + "]";
      if (typeof c?.id !== "string" || c.id.trim() === "") throw new Error(at + ".id is required");
      if (seen.has(c.id)) throw new Error(at + ".id is a duplicate: " + c.id);
      seen.add(c.id);
      if (c.kind !== "credit-note" && c.kind !== "payment-on-account") {
        throw new Error(at + '.kind must be "credit-note" or "payment-on-account"');
      }
      const issuedOn = isoDay(c.issuedOn, at + ".issuedOn");
      if (typeof c.currency !== "string") throw new Error(at + ".currency must be an ISO 4217 code");
      const creditCurrency = code(c.currency);
      const fmt = moneyFor(creditCurrency, at + ".currency");
      const remainingMinor = minorOf(c.remainingMinor, at + ".remainingMinor");
      const hasExpiry = c.expiresOn !== undefined && c.expiresOn !== null;
      const expiresOn = hasExpiry ? isoDay(c.expiresOn, at + ".expiresOn") : null;
      let reason: string | null = null;
      if (expiresOn !== null && expiresOn < issueDate) reason = "Expired on " + fmtDate(expiresOn);
      else if (c.disputed === true) reason = "On hold – disputed";
      else if (creditCurrency !== currency) reason = "Not applied – different currency";
      return { id: c.id as string, kind: c.kind as string, index, issuedOn, remainingMinor, fmt, reason };
    });

    const allocation: any[] = cr?.allocation ?? [];
    if (!Array.isArray(allocation)) throw new Error("credits.allocation must be an array");
    const named = new Set<string>();
    allocation.forEach((a: any, index: number) => {
      const at = "credits.allocation[" + index + "]";
      if (!seen.has(a?.creditId)) throw new Error(at + ".creditId names no credit in credits.open: " + a?.creditId);
      if (named.has(a.creditId)) throw new Error(at + ".creditId is named twice: " + a.creditId);
      named.add(a.creditId);
      minorOf(a.amountMinor, at + ".amountMinor");
    });

    const live = all.filter((c) => c.remainingMinor > 0);
    const byDate = (a: any, b: any) => (a.issuedOn < b.issuedOn ? -1 : a.issuedOn > b.issuedOn ? 1 : a.index - b.index);
    const eligible = live.filter((c) => c.reason === null).sort(byDate);

    // Credits reduce what is owed, never the total, so every step is min or subtraction.
    const appliedById = new Map<string, number>();
    let owed = totalMinor;
    const take = (c: any, limit: number) => {
      const applied = Math.min(c.remainingMinor, limit, owed);
      owed -= applied;
      appliedById.set(c.id, applied);
    };
    const allocated = allocation.length > 0;
    if (allocated) {
      for (const a of allocation) {
        const c = eligible.find((e) => e.id === a.creditId);
        if (c) take(c, a.amountMinor);
      }
    } else {
      for (const c of eligible) take(c, Infinity);
    }

    const rowOf = (c: any) => {
      const appliedMinor = appliedById.get(c.id) ?? 0;
      const carriedForwardMinor = c.remainingMinor - appliedMinor;
      return {
        reference: c.id,
        issued: fmtDate(c.issuedOn),
        applied: money(appliedMinor),
        carriedForward: money(carriedForwardMinor),
        appliedMinor,
        carriedForwardMinor,
      };
    };
    const creditNotes = eligible.filter((c) => c.kind === "credit-note").map(rowOf);
    const payments = eligible.filter((c) => c.kind === "payment-on-account").map(rowOf);
    const excluded = live
      .filter((c) => c.reason !== null)
      .sort(byDate)
      .map((c) => ({
        reference: c.id,
        type: c.kind === "credit-note" ? "Credit note" : "Payment on account",
        amount: c.fmt(c.remainingMinor),
        reason: c.reason as string,
      }));

    const creditNotesAppliedMinor = creditNotes.reduce((s, r) => s + r.appliedMinor, 0);
    const paymentsAppliedMinor = payments.reduce((s, r) => s + r.appliedMinor, 0);
    const appliedMinor = creditNotesAppliedMinor + paymentsAppliedMinor;
    const dueMinor = totalMinor - appliedMinor;
    const carriedForwardMinor = [...creditNotes, ...payments].reduce((s, r) => s + r.carriedForwardMinor, 0);
    if (dueMinor + appliedMinor !== totalMinor || dueMinor < 0) throw new Error("credits do not tie out");

    const outcome = appliedMinor === 0 ? "none" : dueMinor === 0 ? "covered" : "part";
    const intro = appliedMinor > 0
      ? allocated
        ? "Credits on your account have been set against this invoice as allocated."
        : "Open credits on your account have been set against this invoice, oldest first."
      : totalMinor === 0
      ? "There is nothing on this invoice to set credits against."
      : "None of the credits on your account could be set against this invoice.";
    const less = (minor: number) => (minor > 0 ? "−" + money(minor) : null);

    return {
      show: live.length > 0,
      outcome,
      intro,
      hasCreditNotes: creditNotes.length > 0,
      creditNotes,
      hasPayments: payments.length > 0,
      payments,
      hasExcluded: excluded.length > 0,
      excluded,
      total: money(totalMinor),
      totalMinor,
      lessCreditNotes: less(creditNotesAppliedMinor),
      creditNotesAppliedMinor,
      lessPayments: less(paymentsAppliedMinor),
      paymentsAppliedMinor,
      appliedMinor,
      due: money(dueMinor),
      dueMinor,
      vat,
      carriedForward: carriedForwardMinor > 0 ? money(carriedForwardMinor) : null,
      carriedForwardMinor,
    };
  },
});
