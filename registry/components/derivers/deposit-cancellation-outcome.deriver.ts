// deno-lint-ignore-file no-explicit-any
import { deriver } from "docxcelerate/template";

/**
 * Decides, for a cancelled order, what happens to each kind of money paid in
 * advance and how VAT or sales tax treats each part. Installed by
 * `dxcl add deposit-cancellation-outcome`.
 *
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level constant or helper would be a ReferenceError there. Every
 * helper, and the cap of six deduction slots, is declared inside `run`.
 *
 * Money never meets a float. Each amount is read from the number's own decimal
 * string into integer minor units, a rate into basis points, and every share
 * is an integer division rounded half up. The VAT on a partly retained advance
 * is worked on the retained part and the refunded side takes whatever is left
 * of the VAT charged, so the two halves always add back to it.
 */
export default deriver({
  name: "depositCancellationOutcome",
  run: (input: unknown) => {
    const P = "depositCancellationOutcome: ";
    const fail = (message: string): never => {
      throw new Error(P + message);
    };
    const order = input as any;
    if (order === null || typeof order !== "object" || Array.isArray(order)) fail("order must be an object");

    const reference = typeof order.reference === "string" ? order.reference.trim() : "";
    if (reference === "") fail("order.reference is required");

    let probe: Intl.NumberFormat | undefined;
    try {
      probe = new Intl.NumberFormat("en", { style: "currency", currency: order.currency });
    } catch {
      fail("order.currency must be an ISO 4217 code");
    }
    const exponent = probe!.resolvedOptions().maximumFractionDigits as number;

    const jurisdiction = order.jurisdiction;
    if (jurisdiction !== "UK" && jurisdiction !== "DE" && jurisdiction !== "US") {
      fail('order.jurisdiction must be "UK", "DE" or "US"');
    }
    const US = jurisdiction === "US";
    const locale = US ? "en-US" : "en-GB";
    const money = new Intl.NumberFormat(locale, { style: "currency", currency: order.currency });
    const fmt = (minor: number) => money.format(minor / 10 ** exponent);
    const dates = new Intl.DateTimeFormat(locale, { dateStyle: "long", timeZone: "UTC" });

    const isoDay = (value: unknown, field: string): string => {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}/.test(value)) fail(field + " must be an ISO 8601 date");
      const day = (value as string).slice(0, 10);
      if (new Date(day + "T00:00:00Z").toISOString().slice(0, 10) !== day) fail(field + " is not a real date");
      return day;
    };
    const fmtDay = (iso: string) => dates.format(new Date(iso + "T00:00:00Z"));

    const cancelledIso = isoDay(order.cancelledOn, "order.cancelledOn");
    if (jurisdiction === "UK" && cancelledIso < "2019-03-01") {
      fail("order.cancelledOn is before 2019-03-01, when the current UK rule began; the earlier rule is not covered");
    }
    const hasPaidOn = order.refundPaidOn !== undefined && order.refundPaidOn !== null;
    const paidIso = hasPaidOn ? isoDay(order.refundPaidOn, "order.refundPaidOn") : "";

    // Parsed from the number's own digits: 1250.5 is 125050 whatever the float says.
    const minorOf = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isFinite(value) || value < 0 || Object.is(value, -0)) {
        fail(field + " must be a non-negative number");
      }
      const text = String(value);
      if (/e/i.test(text)) fail(field + " is too large or too small to read exactly");
      const [whole, fraction = ""] = text.split(".");
      if (fraction.length > exponent) fail(field + " has more decimals than " + order.currency + " allows");
      const minor = Number(whole + fraction.padEnd(exponent, "0"));
      if (!Number.isSafeInteger(minor)) fail(field + " is too large");
      return minor;
    };
    const bpOf = (value: unknown, field: string): number => {
      if (typeof value !== "number" || !Number.isFinite(value) || value < 0 || value > 100) {
        fail(field + " must be a number from 0 to 100");
      }
      const bp = Math.round((value as number) * 100);
      if (bp / 100 !== value) fail(field + " has more than two decimals");
      return bp;
    };
    // Half away from zero, on non-negative integers.
    const roundHalf = (n: number, d: number): number => {
      let q = Math.floor(n / d);
      while (q * d > n) q--;
      const r = n - q * d;
      return 2 * r >= d ? q + 1 : q;
    };

    // Pooling, in input order. Each band keeps the payment it came from.
    const payments = order.payments === undefined || order.payments === null ? [] : order.payments;
    if (!Array.isArray(payments)) fail("order.payments must be an array");
    let S = 0;
    let R = 0;
    let N = 0;
    const refundable: Array<{ gross: number; vat: number; rateBp: number; payment: number }> = [];
    let nonRefundableVat = 0;
    payments.forEach((payment: any, i: number) => {
      const at = "order.payments[" + i + "]";
      const kind = payment?.kind;
      if (kind !== "security" && kind !== "refundable-advance" && kind !== "non-refundable-advance") {
        fail(at + '.kind must be "security", "refundable-advance" or "non-refundable-advance"');
      }
      const gross = minorOf(payment.gross, at + ".gross");
      let bands = payment.bands;
      if (bands === undefined || bands === null) bands = [{ rate: 0, net: gross / 10 ** exponent, vat: 0 }];
      if (!Array.isArray(bands)) fail(at + ".bands must be an array");
      let sum = 0;
      const parsed = bands.map((band: any, j: number) => {
        const bandAt = at + ".bands[" + j + "]";
        const rateBp = bpOf(band?.rate, bandAt + ".rate");
        const net = minorOf(band.net, bandAt + ".net");
        const vat = minorOf(band.vat, bandAt + ".vat");
        if (vat > 0 && kind === "security") fail(bandAt + ".vat: a security deposit carries no VAT");
        if (vat > 0 && US) fail(bandAt + ".vat: refunding sales tax collected on an advance is not covered");
        sum += net + vat;
        return { gross: net + vat, vat, rateBp };
      });
      if (sum !== gross) fail(at + ".bands do not add up to " + at + ".gross");
      if (kind === "security") S += gross;
      else if (kind === "refundable-advance") {
        R += gross;
        for (const band of parsed) refundable.push({ ...band, payment: i });
      } else {
        N += gross;
        for (const band of parsed) nonRefundableVat += band.vat;
      }
      if (payment.invoiceNumber !== undefined && payment.invoiceNumber !== null && typeof payment.invoiceNumber !== "string") {
        fail(at + ".invoiceNumber must be a string");
      }
    });
    const paid = S + R + N;

    // Retention from the refundable advances.
    let T = 0;
    let reason = "";
    const retention = order.retention;
    if (retention !== undefined && retention !== null) {
      if (typeof retention !== "object") fail("order.retention must be an object");
      const hasPercent = retention.percent !== undefined && retention.percent !== null;
      const hasAmount = retention.amount !== undefined && retention.amount !== null;
      if (hasPercent === hasAmount) fail("order.retention needs exactly one of percent and amount");
      if (typeof retention.reason !== "string" || retention.reason.trim() === "") {
        fail("order.retention.reason is required");
      }
      reason = retention.reason.trim();
      if (hasPercent) T = roundHalf(R * bpOf(retention.percent, "order.retention.percent"), 10000);
      else {
        T = minorOf(retention.amount, "order.retention.amount");
        if (T > R) fail("order.retention.amount is more than the refundable advances paid");
      }
    }

    // Largest remainder: floor shares, then the spare minor units go one each
    // to the biggest remainders, the earlier band on a tie.
    const shares = refundable.map((band) =>
      R === 0 ? 0 : Number((BigInt(T) * BigInt(band.gross)) / BigInt(R))
    );
    const remainders = refundable.map((band) => R === 0 ? BigInt(0) : (BigInt(T) * BigInt(band.gross)) % BigInt(R));
    let spare = T - shares.reduce((a, b) => a + b, 0);
    const byRemainder = refundable.map((_, k) => k).sort((a, b) =>
      remainders[a] === remainders[b] ? a - b : remainders[a] > remainders[b] ? -1 : 1
    );
    for (const k of byRemainder) {
      if (spare <= 0) break;
      shares[k] += 1;
      spare -= 1;
    }

    let retainedVatRefundable = 0;
    let refundedVat = 0;
    const refundedByPayment = new Map<number, number>();
    refundable.forEach((band, k) => {
      const t = shares[k];
      const retainedVat = t === band.gross
        ? band.vat
        : Math.min(band.vat, Math.max(0, roundHalf(t * band.rateBp, 10000 + band.rateBp)));
      retainedVatRefundable += retainedVat;
      refundedVat += band.vat - retainedVat;
      refundedByPayment.set(band.payment, (refundedByPayment.get(band.payment) ?? 0) + band.gross - t);
    });
    const retainedVat = retainedVatRefundable + nonRefundableVat;

    // Deductions from the security deposit.
    const deductions = order.deductions === undefined || order.deductions === null ? [] : order.deductions;
    if (!Array.isArray(deductions)) fail("order.deductions must be an array");
    if (deductions.length > 6) fail("order.deductions has more than 6 entries");
    let D = 0;
    const slots: Array<{ description: string | null; amount: string | null; line: string | null }> = [];
    deductions.forEach((item: any, i: number) => {
      const at = "order.deductions[" + i + "]";
      if (typeof item?.label !== "string" || item.label.trim() === "") fail(at + ".label is required");
      const label = item.label.trim();
      const amount = minorOf(item.amount, at + ".amount");
      if (amount === 0) fail(at + ".amount must be more than zero");
      D += amount;
      let treatment = "";
      if (!US) {
        if (typeof item.isSupply !== "boolean") fail(at + ".isSupply is required: VAT depends on it");
        if (item.isSupply) {
          if (item.vatRate === undefined || item.vatRate === null) fail(at + ".vatRate is required when isSupply is true");
          const rateBp = bpOf(item.vatRate, at + ".vatRate");
          treatment = "including " + fmt(roundHalf(amount * rateBp, 10000 + rateBp)) + " VAT";
        } else treatment = "compensation, outside the scope of VAT";
      }
      slots.push({
        description: label + (treatment ? " (" + treatment + ")" : ""),
        amount: fmt(amount),
        line: label + ": " + fmt(amount) + (treatment ? ", " + treatment : ""),
      });
    });
    if (D > S) fail(S === 0 ? "order.deductions given but no security deposit was paid" : "order.deductions exceed the security deposit");
    while (slots.length < 6) slots.push({ description: null, amount: null, line: null });

    const kept = T + N;
    if (US && kept > 0) {
      if (typeof order.usState !== "string" || order.usState.trim() === "") fail("order.usState is required when an advance is retained");
      if (typeof order.usForfeitTaxable !== "boolean") {
        fail("order.usForfeitTaxable is required when an advance is retained");
      }
    }

    const refunded = S - D + (R - T);
    const retained = D + T + N;
    if (refunded + retained !== paid) fail("internal error: refunded and retained do not add up to what was paid");

    const nothingPaid = paid === 0;
    const some = (minor: number) => (!nothingPaid && minor > 0 ? fmt(minor) : null);
    const lines = {
      paid: some(paid),
      security: some(S),
      deductions: slots,
      refundable: some(R),
      retentionLabel: some(T) ? "Retained: " + reason : null,
      retention: some(T),
      nonRefundable: some(N),
      refunded: nothingPaid ? null : fmt(refunded),
      retained: some(retained),
    };

    const refundedGross = R - T;
    const invoices: string[] = [];
    payments.forEach((payment: any, i: number) => {
      const number = typeof payment.invoiceNumber === "string" ? payment.invoiceNumber.trim() : "";
      if (payment.kind === "refundable-advance" && number && (refundedByPayment.get(i) ?? 0) > 0 && !invoices.includes(number)) {
        invoices.push(number);
      }
    });
    const against = invoices.length === 0
      ? ""
      : invoices.length === 1
      ? " against invoice " + invoices[0]
      : " against invoices " + invoices.slice(0, -1).join(", ") + " and " + invoices[invoices.length - 1];

    const paidOn = hasPaidOn ? fmtDay(paidIso) : "";
    const creditsVat = !US && refundedVat > 0;
    return {
      nothingPaid,
      cancelledOn: fmtDay(cancelledIso),
      reference,
      lines,
      refundPaid: nothingPaid || refunded === 0
        ? null
        : hasPaidOn
        ? "The refund of " + fmt(refunded) + " was paid on " + paidOn + "."
        : "The refund of " + fmt(refunded) + " has not been paid yet.",
      creditNote: !creditsVat || nothingPaid
        ? null
        : order.creditNoteNumber
        ? "Credit note " + String(order.creditNoteNumber).trim() + " issued for " + fmt(refundedGross) + ", including " +
          fmt(refundedVat) + " VAT" + against + "."
        : "A credit note for " + fmt(refundedGross) + ", including " + fmt(refundedVat) + " VAT" + against +
          ", will follow.",
      vatCorrection: jurisdiction !== "DE" || refundedVat === 0
        ? null
        : hasPaidOn
        ? "The " + fmt(refundedVat) + " VAT on the refunded amount is corrected in the period the refund was paid, " +
          paidOn + "."
        : "The " + fmt(refundedVat) + " VAT on the refunded amount is corrected when the refund is paid, not on cancellation.",
      retainedVat: US || retainedVat === 0
        ? null
        : jurisdiction === "UK"
        ? "The " + fmt(kept) + " retained includes " + fmt(retainedVat) +
          " VAT, accounted for when you paid. It remains due, so no credit note is issued for it."
        : "The " + fmt(kept) + " retained is the cancellation fee agreed in your contract and includes " +
          fmt(retainedVat) + " VAT.",
      securityVat: US || S === 0 ? null : "The " + fmt(S) + " security deposit carried no VAT.",
      salesTax: !US || kept === 0
        ? null
        : order.usForfeitTaxable
        ? "Under " + order.usState.trim() + " rules, sales tax is due on the " + fmt(kept) +
          " retained. It is not included in the amounts above."
        : order.usState.trim() + " charges no sales tax on the " + fmt(kept) + " retained.",
      totals: {
        paid,
        security: S,
        deductions: D,
        refundable: R,
        retention: T,
        nonRefundable: N,
        refunded,
        retained,
        refundedVat,
        retainedVat,
      },
    };
  },
});
