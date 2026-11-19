// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/** What `promptPaymentDiscount` returns: every string the block prints, already chosen. */
export interface PromptPaymentDiscountResult {
  /** False means print nothing. */
  show: boolean;
  /** The section title. */
  heading: string;
  /** The terms sentence, always printed when `show`. */
  terms: string;
  /** The UK credit-note sentence; empty everywhere else. */
  statement: string;
  /** Whether the per-band figures table is printed. */
  showTable: boolean;
  /** The table's column titles. */
  columns: { rate: string; first: string; second: string; third: string };
  /** Exactly six rows; unused slots are all empty strings. */
  rows: Array<{ label: string; first: string; second: string; third: string }>;
  /** The closing sentence with the early-payment amount. */
  payable: string;
  /** The early-payment amount in minor units, for tests. */
  payableMinor: number;
  /** The discount deadline, ISO YYYY-MM-DD; empty when nothing is shown. */
  deadline: string;
}

/**
 * Works the early-payment terms of an invoice per VAT band under its
 * jurisdiction's rule. Installed by `dxcl add prompt-payment-discount`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper and table is declared inside `run`.
 */
export default deriver({
  name: "promptPaymentDiscount",
  run: (invoiceInput: unknown): PromptPaymentDiscountResult => {
    const invoice = invoiceInput as any;
    const slot = () => ({ label: "", first: "", second: "", third: "" });
    const none = (): PromptPaymentDiscountResult => ({
      show: false,
      heading: "",
      terms: "",
      statement: "",
      showTable: false,
      columns: { rate: "", first: "", second: "", third: "" },
      rows: [slot(), slot(), slot(), slot(), slot(), slot()],
      payable: "",
      payableMinor: 0,
      deadline: "",
    });

    // Decisions that need nothing else come first: an invoice with no discount
    // must never throw, whatever else is missing from it.
    const terms = invoice?.promptPaymentDiscount;
    const percent = terms?.percent;
    if (percent === undefined || percent === null || percent === 0) return none();
    if (typeof percent !== "number" || !Number.isFinite(percent) || percent < 0 || percent >= 100) {
      throw new Error("invoice.promptPaymentDiscount.percent must be at least 0 and under 100");
    }
    if (Math.abs(percent * 100 - Math.round(percent * 100)) > 1e-6) {
      throw new Error("invoice.promptPaymentDiscount.percent has more than 2 decimal places");
    }
    const days = terms.days;
    if (days === undefined || days === null) return none();
    if (typeof days !== "number" || !Number.isInteger(days)) {
      throw new Error("invoice.promptPaymentDiscount.days must be a whole number of days");
    }
    if (days <= 0) return none();

    const issue = invoice.issueDate;
    if (typeof issue !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(issue)) {
      throw new Error("invoice.issueDate must be an ISO 8601 date");
    }
    const [y, m, d] = issue.slice(0, 10).split("-").map(Number);
    const issueMs = Date.UTC(y, m - 1, d);
    if (new Date(issueMs).toISOString().slice(0, 10) !== issue.slice(0, 10)) {
      throw new Error("invoice.issueDate must be an ISO 8601 date");
    }

    const currency = invoice.currency;
    let exponent: number;
    try {
      // Intl accepts any three letters, so a code it does not list is rejected here.
      if (!(Intl as any).supportedValuesOf("currency").includes(currency)) throw new Error();
      exponent = new Intl.NumberFormat("en-GB", { style: "currency", currency }).resolvedOptions()
        .maximumFractionDigits as number;
    } catch {
      throw new Error("invoice.currency must be an ISO 4217 code");
    }

    const jurisdiction = invoice.jurisdiction;
    if (!["UK", "EU", "DE", "FR", "US"].includes(jurisdiction)) {
      throw new Error('invoice.jurisdiction must be "UK", "EU", "DE", "FR" or "US"');
    }

    // Sales tax on the early-payment amount, per state. The table is the
    // owner's to extend; an unlisted state throws rather than being guessed.
    const usStates: Record<string, "full" | "discounted"> = {
      // New York: a prompt-payment discount does not reduce the tax base (NY Tax Law s.1101(b)(3)).
      NY: "full",
      // Massachusetts: sales tax is on the price before a prompt-payment discount (830 CMR 64H.1.1).
      MA: "full",
      // California: the measure is reduced when the discount is taken (Regulation 1655); the amount shown assumes it is.
      CA: "discounted",
      // Texas: a true cash discount reduces the sales price (34 TAC s.3.291).
      TX: "discounted",
    };
    let usRule: "full" | "discounted" = "full";
    if (jurisdiction === "US") {
      const state = typeof invoice.usState === "string" ? invoice.usState.trim().toUpperCase() : "";
      if (state === "") throw new Error("invoice.usState is required for a US invoice");
      if (!Object.prototype.hasOwnProperty.call(usStates, state)) {
        throw new Error(
          "invoice.usState " + state + " is not in the prompt-payment sales-tax table; add it to the deriver",
        );
      }
      usRule = usStates[state];
    }

    const dueDays = invoice.dueDays;
    if (dueDays !== undefined && dueDays !== null && (typeof dueDays !== "number" || !Number.isInteger(dueDays) || dueDays <= 0)) {
      throw new Error("invoice.dueDays must be a positive whole number");
    }

    const toMinor = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
        throw new Error(field + " must be a finite number of 0 or more");
      }
      const scaled = value * 10 ** exponent;
      if (Math.abs(scaled - Math.round(scaled)) > 1e-6) {
        throw new Error(field + " has more decimals than " + currency + " allows");
      }
      return Math.round(scaled);
    };
    const bands = invoice.bands;
    if (!Array.isArray(bands) || bands.length === 0) throw new Error("invoice.bands must not be empty");
    const treatments = ["standard", "reduced", "zero", "exempt", "reverse-charge"];
    const seen = new Set<string>();
    const all = bands.map((b: any, i: number) => {
      const at = "invoice.bands[" + i + "]";
      if (!treatments.includes(b?.treatment)) throw new Error(at + ".treatment is not a known treatment");
      const taxed = b.treatment === "standard" || b.treatment === "reduced";
      let rate4 = 0;
      if (taxed) {
        const r = b.rate;
        if (typeof r !== "number" || !Number.isFinite(r) || r <= 0 || r >= 100) {
          throw new Error(at + ".rate must be above 0 and under 100");
        }
        if (Math.abs(r * 10000 - Math.round(r * 10000)) > 1e-6) {
          throw new Error(at + ".rate has more than 4 decimal places");
        }
        rate4 = Math.round(r * 10000);
      }
      const key = taxed ? b.treatment + ":" + rate4 : b.treatment;
      if (seen.has(key)) throw new Error(at + " repeats an earlier band with the same treatment and rate");
      seen.add(key);
      const net = toMinor(b.net, at + ".net");
      const vat = toMinor(b.vat, at + ".vat");
      if (!taxed && vat !== 0) throw new Error(at + ".vat must be 0 on a " + b.treatment + " band");
      return { treatment: b.treatment as string, taxed, rate: b.rate as number, rate4, net, vat };
    });

    if (dueDays !== undefined && dueDays !== null && days > dueDays) {
      throw new Error(
        "invoice.promptPaymentDiscount.days (" + days + ") is after the due date (dueDays " + dueDays + ")",
      );
    }
    const active = all.filter((b) => b.net > 0);
    if (active.length === 0) return none();
    if (active.length > 6) throw new Error("invoice.bands has more than 6 bands with a net; there are 6 slots");

    // Products can pass 2^53 (a 1bn net times a rate scaled by 10,000), so they
    // are BigInt. Amounts are never negative, so half up is half away from zero.
    const mulDivHalfUp = (a: number, b: number, den: number): number => {
      const n = BigInt(a) * BigInt(b);
      const D = BigInt(den);
      const q = n / D;
      return Number((n % D) * BigInt(2) >= D ? q + BigInt(1) : q);
    };
    const pctBp = Math.round(percent * 100);

    const locale = { UK: "en-GB", EU: "en-GB", DE: "de-DE", FR: "fr-FR", US: "en-US" }[jurisdiction as string] as string;
    const money = new Intl.NumberFormat(locale, { style: "currency", currency });
    const fmt = (minor: number) => money.format(minor / 10 ** exponent);
    const pct = new Intl.NumberFormat(locale, { maximumFractionDigits: 2 }).format(percent);
    const rateFmt = new Intl.NumberFormat(locale, { maximumFractionDigits: 4 });
    const spaced = jurisdiction === "DE" || jurisdiction === "FR";
    const pctSign = spaced ? " %" : "%";
    const exemptLabel = jurisdiction === "DE" ? "steuerfrei" : jurisdiction === "FR" ? "Exonéré" : "Exempt";
    const reverseLabel =
      jurisdiction === "DE" ? "Reverse-Charge" : jurisdiction === "FR" ? "Autoliquidation" : "Reverse charge";
    const labelOf = (b: { treatment: string; rate: number }) =>
      b.treatment === "standard" || b.treatment === "reduced"
        ? rateFmt.format(b.rate) + pctSign
        : b.treatment === "zero"
          ? "0" + pctSign
          : b.treatment === "exempt"
            ? exemptLabel
            : reverseLabel;

    const deadlineMs = issueMs + days * 86_400_000;
    const deadline = new Date(deadlineMs).toISOString().slice(0, 10);
    const dateOptions: Intl.DateTimeFormatOptions =
      jurisdiction === "DE"
        ? { day: "2-digit", month: "2-digit", year: "numeric", timeZone: "UTC" }
        : jurisdiction === "FR"
          ? { timeZone: "UTC" }
          : { dateStyle: "long", timeZone: "UTC" };
    const deadlineText = new Intl.DateTimeFormat(locale, dateOptions).format(new Date(deadlineMs));

    let payableMinor = 0;
    let showTable = false;
    let statement = "";
    const lines: Array<{ label: string; first: string; second: string; third: string }> = [];
    const result = none();

    if (jurisdiction === "UK") {
      showTable = true;
      for (const b of active) {
        const discountedNet = b.net - mulDivHalfUp(b.net, pctBp, 10000);
        // One rounding per band, never per line, and never for an untaxed band.
        const vat = b.taxed ? mulDivHalfUp(discountedNet, b.rate4, 1_000_000) : 0;
        payableMinor += discountedNet + vat;
        lines.push({ label: labelOf(b), first: fmt(discountedNet), second: fmt(vat), third: fmt(discountedNet + vat) });
      }
      result.heading = "Prompt payment discount";
      result.columns = { rate: "Rate", first: "Net after discount", second: "VAT on discounted net", third: "Payable if paid early" };
      result.terms =
        "A discount of " + pct + "% of the full price applies if payment is made within " + days +
        " days of the invoice date.";
      statement = terms.issuesCreditNote === true
        ? "If payment is made within " + days + " days, we will issue a credit note for the discount and the VAT on it."
        : "No credit note will be issued. Following payment you must ensure you have only recovered the VAT actually paid.";
      result.payable = "Total payable if paid by " + deadlineText + ": " + fmt(payableMinor) + ".";
    } else if (jurisdiction === "US") {
      let taxMinor = 0;
      for (const b of active) {
        const discountedNet = b.net - mulDivHalfUp(b.net, pctBp, 10000);
        const tax = !b.taxed ? 0 : usRule === "full" ? b.vat : mulDivHalfUp(discountedNet, b.rate4, 1_000_000);
        payableMinor += discountedNet + tax;
        taxMinor += tax;
      }
      result.heading = "Early payment discount";
      result.terms =
        "Terms: " + pct + "/" + days + (typeof dueDays === "number" ? ", net " + dueDays : "") + ". A " + pct +
        "% discount applies if payment is made by " + deadlineText + ".";
      result.payable = "Amount due if paid by " + deadlineText + ": " + fmt(payableMinor) +
        (taxMinor > 0 ? ", including " + fmt(taxMinor) + " sales tax on the " + usRule + " price." : ".");
    } else {
      // The invoice's own VAT stays on the full price; the reduction is taken
      // off each band's gross, split by largest remainder so the cents add up.
      const gross = active.map((b) => b.net + b.vat);
      const G = gross.reduce((a, b) => a + b, 0);
      const totalDiscount = mulDivHalfUp(G, pctBp, 10000);
      const products = gross.map((g) => BigInt(totalDiscount) * BigInt(g));
      const shares = products.map((p) => Number(p / BigInt(G)));
      let leftover = totalDiscount - shares.reduce((a, b) => a + b, 0);
      const order = products
        .map((p, i) => ({ i, rem: p % BigInt(G) }))
        .sort((a, b) => (a.rem === b.rem ? a.i - b.i : a.rem > b.rem ? -1 : 1));
      for (const o of order) {
        if (leftover <= 0) break;
        shares[o.i] += 1;
        leftover -= 1;
      }
      active.forEach((b, i) => {
        lines.push({ label: labelOf(b), first: fmt(gross[i]), second: fmt(shares[i]), third: fmt(gross[i] - shares[i]) });
      });
      payableMinor = G - totalDiscount;
      // A single rate needs no split: the terms line is enough.
      showTable = active.length > 1;
      if (jurisdiction === "DE") {
        result.heading = "Skonto";
        result.columns = { rate: "Steuersatz", first: "Brutto", second: "Skonto", third: "Brutto nach Skonto" };
        result.terms = pct + " % Skonto bei Zahlung bis " + deadlineText + ".";
        result.payable = "Zahlbetrag abzüglich Skonto: " + fmt(payableMinor) + ".";
      } else if (jurisdiction === "FR") {
        result.heading = "Escompte pour paiement anticipé";
        result.columns = { rate: "Taux", first: "Montant TTC", second: "Escompte", third: "TTC après escompte" };
        result.terms = "Escompte de " + pct + " % pour paiement au plus tard le " + deadlineText + ".";
        result.payable = "Net à payer après escompte : " + fmt(payableMinor) + ".";
      } else {
        result.heading = "Prompt payment discount";
        result.columns = { rate: "Rate", first: "Gross", second: "Discount", third: "Gross after discount" };
        result.terms = "A discount of " + pct + "% applies if payment is made by " + deadlineText + ".";
        result.payable = "Amount payable if paid by " + deadlineText + ": " + fmt(payableMinor) + ".";
      }
    }

    if (!showTable) result.columns = { rate: "", first: "", second: "", third: "" };
    result.show = true;
    result.statement = statement;
    result.showTable = showTable;
    result.rows = [0, 1, 2, 3, 4, 5].map((i) => (showTable && lines[i]) || slot());
    result.payableMinor = payableMinor;
    result.deadline = deadline;
    return result;
  },
});
