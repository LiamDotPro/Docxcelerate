// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Reconciles a marketplace seller's period to the balance that reaches their
 * bank, one signed row per step. Installed by
 * `dxcl add marketplace-payout-statement`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper is declared inside `run`.
 */
export default deriver({
  name: "marketplacePayoutStatement",
  run: (payout: unknown) => {
    const p = payout as any;

    // Currency first: the exponent decides how every amount is read and printed.
    // narrowSymbol so a dollar or a yen prints as one, not "US$" or "JP¥".
    const currency = p?.currency;
    const options = { style: "currency", currency, currencyDisplay: "narrowSymbol" } as const;
    let plain: Intl.NumberFormat;
    try {
      plain = new Intl.NumberFormat("en-GB", options);
    } catch {
      throw new Error("payout.currency must be an ISO 4217 code");
    }
    const exponent = plain.resolvedOptions().maximumFractionDigits as number;
    const scale = 10 ** exponent;
    const signedFmt = new Intl.NumberFormat("en-GB", { ...options, signDisplay: "exceptZero" });
    const autoFmt = new Intl.NumberFormat("en-GB", { ...options, signDisplay: "auto" });
    const unsignedFmt = new Intl.NumberFormat("en-GB", { ...options, signDisplay: "never" });
    const fmt = (minor: number) => unsignedFmt.format(minor / scale);

    const money = (value: unknown, field: string, signed = false): number => {
      if (typeof value !== "number" || !Number.isSafeInteger(value) || (!signed && value < 0)) {
        throw new Error(field + " must be " + (signed ? "an" : "a non-negative") + " integer of minor units");
      }
      return value;
    };
    const optional = (value: unknown, field: string) => value === undefined || value === null ? 0 : money(value, field);
    const list = (value: unknown): any[] => Array.isArray(value) ? value : [];
    // Half away from zero on exact integers; no float rate is ever formed.
    const roundHalfAway = (numerator: number, denominator: number) =>
      Math.sign(numerator) * Math.floor((2 * Math.abs(numerator) + denominator) / (2 * denominator));

    const openingBalance = money(p?.openingBalance, "payout.openingBalance", true);

    let sales = 0;
    let tax = 0;
    list(p.orders).forEach((o, i) => {
      const f = "payout.orders[" + i + "]";
      sales += money(o?.itemTotal, f + ".itemTotal") + optional(o?.shipping, f + ".shipping");
      tax += optional(o?.facilitatorTax, f + ".facilitatorTax");
    });

    let refunds = 0;
    let commissionReturned = 0;
    let adminFees = 0;
    list(p.refunds).forEach((r, i) => {
      const f = "payout.refunds[" + i + "]";
      refunds += money(r?.amount, f + ".amount");
      const commission = money(r?.commissionRefunded, f + ".commissionRefunded");
      const admin = money(r?.adminFee, f + ".adminFee");
      if (admin > commission) throw new Error(f + ".adminFee is larger than " + f + ".commissionRefunded");
      commissionReturned += commission;
      adminFees += admin;
    });

    let fees = 0;
    let feeVat = 0;
    let reverseCharged = 0;
    list(p.fees).forEach((fee, i) => {
      const f = "payout.fees[" + i + "]";
      if (typeof fee?.label !== "string" || fee.label.trim() === "") throw new Error(f + ".label is required");
      const amount = money(fee.amount, f + ".amount", true);
      const rate = fee.vatPercent;
      if (rate !== undefined && rate !== null && (typeof rate !== "number" || !Number.isFinite(rate) || rate < 0 || rate > 100)) {
        throw new Error(f + ".vatPercent must be a number from 0 to 100");
      }
      fees += amount;
      if (fee.reverseCharge === true) reverseCharged += amount;
      else if (typeof rate === "number") feeVat += roundHalfAway(amount * Math.round(rate * 100), 10000);
    });

    let chargebacks = 0;
    let disputeFees = 0;
    let disputesWon = 0;
    list(p.chargebacks).forEach((c, i) => {
      const f = "payout.chargebacks[" + i + "]";
      const amount = money(c?.amount, f + ".amount");
      const fee = optional(c?.disputeFee, f + ".disputeFee");
      if (c.outcome !== "open" && c.outcome !== "won" && c.outcome !== "lost") {
        throw new Error(f + ".outcome must be open, won or lost");
      }
      if (c.raisedEarlier !== true) {
        chargebacks += amount;
        disputeFees += fee;
      }
      if (c.outcome === "won") disputesWon += amount;
    });

    const held = optional(p.reserve?.held, "payout.reserve.held");
    const released = optional(p.reserve?.released, "payout.reserve.released");
    const days = p.reserve?.days;
    if (days !== undefined && days !== null && (typeof days !== "number" || !Number.isSafeInteger(days) || days < 1)) {
      throw new Error("payout.reserve.days must be a positive integer");
    }

    const method = p.negativeBalanceMethod;
    if (method !== undefined && method !== null && method !== "carry" && method !== "debit") {
      throw new Error("payout.negativeBalanceMethod must be carry or debit");
    }

    const row = (label: string, minor: number, printed: boolean) => {
      const on = printed || minor !== 0;
      return {
        label,
        amount: on ? (printed === true && minor === 0 ? autoFmt : signedFmt).format(minor / scale) : "",
        paidIn: on && minor >= 0 ? fmt(minor) : "",
        takenOut: on && minor < 0 ? fmt(minor) : "",
        minor,
      };
    };
    const heldLabel = typeof days === "number"
      ? "Held in reserve for " + days + (days === 1 ? " day" : " days")
      : "Held in reserve";

    const rows = {
      broughtForward: row("Brought forward from your last statement", openingBalance, false),
      sales: row("Sales (items and shipping)", sales, true),
      taxCollected: row("Sales tax collected from buyers", tax, false),
      taxRemitted: row("Sales tax remitted by the marketplace", -tax, false),
      refunds: row("Refunds to buyers", -refunds, false),
      commissionReturned: row("Commission returned on refunds", commissionReturned, false),
      refundAdminFees: row("Refund administration fees", -adminFees, false),
      fees: row("Commission and fees", -fees, false),
      feeVat: row("VAT on fees", -feeVat, false),
      chargebacks: row("Chargebacks", -chargebacks, false),
      disputeFees: row("Dispute fees", -disputeFees, false),
      disputesWon: row("Chargebacks reversed (disputes won)", disputesWon, false),
      reserveHeld: row(heldLabel, -held, false),
      reserveReleased: row("Reserve released", released, false),
    };
    // The closing is the sum of the rows as printed, so they cannot disagree with it.
    let closing = 0;
    for (const key of Object.keys(rows)) closing += (rows as any)[key].minor;
    if (!Number.isSafeInteger(closing)) throw new Error("payout: the rows sum to more than a safe integer");

    const stated = p.closingBalance;
    if (stated !== undefined && stated !== null) {
      money(stated, "payout.closingBalance", true);
      if (stated !== closing) {
        throw new Error("payout.closingBalance is " + stated + " but the rows sum to " + closing);
      }
    }

    // Closing always prints, and a zero is the currency's zero in "Paid in".
    const closingRow = {
      label: "Closing balance",
      amount: autoFmt.format(closing / scale),
      paidIn: closing >= 0 ? fmt(closing) : "",
      takenOut: closing < 0 ? fmt(closing) : "",
      minor: closing,
    };

    const outcome = closing > 0 ? "payout" : closing === 0 ? "nil" : method === "debit" ? "debit" : "carry";
    const digits = typeof p.accountLast4 === "string" ? p.accountLast4.replace(/\D/g, "") : "";

    return {
      rows: { ...rows, closing: closingRow },
      outcome,
      outcomeAmount: fmt(Math.abs(closing)),
      payTo: digits.length >= 4 ? "your account ending " + digits.slice(-4) : "your bank account",
      reverseChargeNote: reverseCharged !== 0
        ? "No VAT has been charged on fees of " + fmt(Math.abs(reverseCharged)) +
          ", which are supplied under the reverse charge. Account for the VAT on them in your own VAT return."
        : "",
      closingMinor: closing,
    };
  },
});
