import { deriver } from "docxcelerate/template";

/**
 * Decides whether a payment-method surcharge may be charged, how much, and
 * what it is taxed at. Installed by `dxcl add payment-surcharge-line`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. The rule
 * tables, the formatter and the integer helpers are all declared inside it.
 */
export default deriver({
  name: "paymentSurchargeLine",
  run: (surcharge: unknown) => {
    // deno-lint-ignore no-explicit-any
    const s = surcharge as any;
    const bad = (field: string, why: string) => new Error("surcharge." + field + " " + why);
    const present = (v: unknown) => v !== undefined && v !== null;
    const nonNegative = (v: unknown) => typeof v === "number" && Number.isFinite(v) && v >= 0;

    const jurisdictions = ["GB", "EU", "US"];
    const kinds = ["credit-card", "debit-card", "prepaid-card", "commercial-card", "sepa", "bank-transfer"];
    if (!jurisdictions.includes(s?.jurisdiction)) {
      throw bad("jurisdiction", "must be one of " + jurisdictions.join(", "));
    }
    if (present(s.customer) && s.customer !== "consumer" && s.customer !== "business") {
      throw bad("customer", "must be consumer or business");
    }
    if (!kinds.includes(s.method?.kind)) throw bad("method.kind", "must be one of " + kinds.join(", "));
    if (present(s.percent) && !nonNegative(s.percent)) throw bad("percent", "must be a non-negative number");
    if (present(s.costOfAcceptance) && !nonNegative(s.costOfAcceptance)) {
      throw bad("costOfAcceptance", "must be a non-negative number");
    }
    if (typeof s.currency !== "string" || s.currency === "") throw bad("currency", "is required");
    if (!Array.isArray(s.bands)) throw bad("bands", "must be an array");
    for (const band of s.bands) {
      if (!nonNegative(band?.ratePercent)) throw bad("bands.ratePercent", "must be a non-negative number");
      if (!nonNegative(band.net)) throw bad("bands.net", "must be a non-negative number");
      if (present(band.vat) && !nonNegative(band.vat)) throw bad("bands.vat", "must be a non-negative number");
    }

    const jurisdiction: string = s.jurisdiction;
    const us = jurisdiction === "US";
    const locale = us ? "en-US" : "en-GB";
    const currencyFormat = (() => {
      try {
        return new Intl.NumberFormat(locale, { style: "currency", currency: s.currency });
      } catch {
        throw bad("currency", "is not an ISO 4217 code: " + s.currency);
      }
    })();
    const d: number = currencyFormat.resolvedOptions().maximumFractionDigits ?? 2;
    const scale = 10 ** d;

    // Integer arithmetic only: a float product rounded for money is wrong at
    // exactly the halves these rules turn on.
    const floorDiv = (a: number, b: number) => (a - a % b) / b;
    const roundHalfUp = (a: number, b: number) => floorDiv(2 * a + b, 2 * b);
    const milli = (percent: number) => Math.round(percent * 1000);

    const bands = (s.bands as Array<{ ratePercent: number; net: number; vat?: number }>).map((band) => ({
      netMinor: Math.round(band.net * scale),
      vatMinor: Math.round((band.vat ?? 0) * scale),
      rateMilli: milli(band.ratePercent),
    }));
    const base = bands.reduce((sum, band) => sum + band.netMinor, 0);
    const preGross = base + bands.reduce((sum, band) => sum + band.vatMinor, 0);
    const pMilli = present(s.percent) ? milli(s.percent) : 0;

    const none = {
      surchargeLabel: null,
      surchargeAmount: null,
      taxLabel: null,
      taxAmount: null,
      taxNote: null,
      totalLabel: null,
      totalAmount: null,
      disclosure: null,
      warning: null,
    };
    if (pMilli === 0 || base === 0) return none;

    const refuse = (why: string) => ({
      ...none,
      warning: "Surcharge not applied: " + why +
        " Remove the surcharge or correct the data before this document is issued.",
    });
    const kind: string = s.method.kind;
    const state = typeof s.usState === "string" ? s.usState.trim().toUpperCase() : "";

    // The numbers an owner checks and edits. Percent of the net total.
    // Banned: Connecticut and Massachusetts forbid card surcharges outright.
    const banned = ["CT", "MA"];
    // Colorado caps a surcharge at 2%.
    const stateCap: Record<string, number> = { CO: 2 };
    // Visa 3%, Mastercard 4%. An unknown network gets the lowest.
    const networkCap: Record<string, number> = { visa: 3, mastercard: 4 };

    if (jurisdiction === "GB" && s.customer !== "business") {
      return refuse(
        "surcharging a consumer for any payment method is banned in the UK " +
          "(Consumer Rights (Payment Surcharges) Regulations 2012, as amended from 13 January 2018).",
      );
    }
    if (!us && ["credit-card", "debit-card", "prepaid-card"].includes(kind)) {
      return refuse(
        "consumer cards (credit, debit and prepaid) may not be surcharged in the UK or EU, " +
          "whoever is paying (PSD2 Art. 62(4) and the UK equivalent).",
      );
    }
    if (jurisdiction === "EU" && kind === "sepa") {
      return refuse("SEPA credit transfers and direct debits may not be surcharged (PSD2 Art. 62(4)).");
    }
    if (us && (kind === "debit-card" || kind === "prepaid-card")) {
      return refuse("debit and prepaid cards may never be surcharged in the US.");
    }
    if (us && (kind === "sepa" || kind === "bank-transfer")) {
      return refuse("only credit cards, commercial cards included, may be surcharged in the US.");
    }
    if (us && state === "") {
      return refuse("the US state decides whether a surcharge is allowed, and no usState was given.");
    }
    if (us && banned.includes(state)) return refuse("surcharges are banned in " + state + ".");
    if (!present(s.costOfAcceptance)) {
      return refuse("every rule that allows a surcharge caps it at the cost of acceptance, and no costOfAcceptance was given.");
    }

    // Every regime caps at the cost of acceptance; the US adds the network's
    // and the state's.
    let capMilli = milli(s.costOfAcceptance);
    if (us) {
      const network = typeof s.method.network === "string" ? s.method.network.trim().toLowerCase() : "";
      const networkMilli = milli(
        Object.prototype.hasOwnProperty.call(networkCap, network)
          ? networkCap[network]
          : Math.min(...Object.values(networkCap)),
      );
      capMilli = Math.min(capMilli, networkMilli);
      if (Object.prototype.hasOwnProperty.call(stateCap, state)) capMilli = Math.min(capMilli, milli(stateCap[state]));
    }
    const appliedMilli = Math.min(pMilli, capMilli);
    // Rounded once, and down at the cap so the cap is never exceeded.
    const S = Math.min(roundHalfUp(base * pMilli, 100000), floorDiv(base * capMilli, 100000));
    if (S === 0) return none;

    // Largest remainder: shares sum to exactly S, ties to the earlier band.
    const shares = bands.map((band) => floorDiv(S * band.netMinor, base));
    const order = bands.map((band, i) => ({ i, rem: (S * band.netMinor) % base }))
      .sort((a, b) => b.rem - a.rem || a.i - b.i);
    let leftover = S - shares.reduce((sum, share) => sum + share, 0);
    for (const { i } of order) {
      if (leftover === 0) break;
      shares[i] += 1;
      leftover -= 1;
    }

    const taxApplies = !us || s.usTaxable !== false;
    const T = taxApplies
      ? bands.reduce((sum, band, i) => sum + roundHalfUp(shares[i] * band.rateMilli, 100000), 0)
      : 0;

    const fmt = (minor: number) => currencyFormat.format(minor / scale);
    const percentFormat = new Intl.NumberFormat(locale, { maximumFractionDigits: 3 });
    const pct = (m: number) => percentFormat.format(m / 1000) + "%";
    const methods: Record<string, string> = {
      "credit-card": "credit card",
      "commercial-card": "commercial card",
      "bank-transfer": "bank transfer",
      sepa: jurisdiction === "GB" ? "bank transfer" : "SEPA transfer",
    };
    const taxName = us ? "tax" : "VAT";
    const split = bands.map((band, i) => ({ share: shares[i], rateMilli: band.rateMilli })).filter((x) => x.share > 0);

    return {
      surchargeLabel: "Surcharge for payment by " + methods[kind] + " (" + pct(appliedMilli) + ")",
      surchargeAmount: fmt(S),
      taxLabel: T > 0 ? (us ? "Sales tax on surcharge" : "VAT on surcharge") : null,
      taxAmount: T > 0 ? fmt(T) : null,
      taxNote: T > 0 && split.length >= 2
        ? "The surcharge is split across the invoice's " + taxName + " rates in proportion to their net: " +
          split.map((x) => fmt(x.share) + " at " + pct(x.rateMilli)).join(", ") + "."
        : null,
      totalLabel: "Total including surcharge",
      totalAmount: fmt(preGross + S + T),
      disclosure: us
        ? "We impose a surcharge of " + pct(appliedMilli) +
          " on credit card payments, which is not greater than our cost of acceptance. " +
          "No surcharge applies to debit card payments."
        : "A surcharge of " + pct(appliedMilli) + " has been added for payment by " + methods[kind] +
          ". It is no more than the cost to us of accepting that payment method.",
      warning: null,
    };
  },
});
