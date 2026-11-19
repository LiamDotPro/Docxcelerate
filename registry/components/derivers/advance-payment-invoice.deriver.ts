import { deriver } from "docxcelerate/template";

/** What the deriver hands the node: every display string, and the same figures as integers. */
export interface AdvancePaymentInvoiceFigures {
  /** The section heading: "Advance payment", "Anzahlung" or "Security deposit". */
  title: string;
  /** The order reference, trimmed. */
  orderReference: string;
  /** The order date, formatted, or null when none was given. */
  orderDate: string | null;
  /** The order gross, formatted. */
  orderTotal: string;
  /** One per previous deposit, in input order. */
  previous: Array<{
    /** The earlier deposit's invoice number, trimmed. */
    number: string;
    /** The earlier deposit, formatted. */
    amount: string;
  }>;
  /**
   * `previous` padded with empty entries (`number` is "") to 8 slots. The page
   * prints fixed slots because the resolver cannot loop over a derived list.
   */
  previousSlots: Array<{ number: string; amount: string }>;
  /** What the deposit line says, naming the percentage when one was given. */
  depositLabel: string;
  /** This request, formatted. */
  deposit: string;
  /** What is left to invoice, formatted. */
  leftToInvoice: string;
  /** "Date of payment" or "Zeitpunkt der Vereinnahmung"; null when no date prints. */
  paymentDateLabel: string | null;
  /** The date payment was received, formatted; null when no date prints. */
  paymentDate: string | null;
  /** True when the VAT analysis is printed: a prepayment outside the US. */
  vatApplies: boolean;
  /** This deposit's VAT analysis, one per rate; empty when `vatApplies` is false. */
  bands: Array<{
    /** The band's rate, formatted: "20%". */
    label: string;
    /** This deposit's net in the band, formatted. */
    net: string;
    /** This deposit's VAT in the band, formatted. */
    vat: string;
    /** This deposit's share of the band, gross, formatted. */
    gross: string;
    /** This deposit's share of the band, gross, in minor units. */
    shareMinor: number;
    /** This deposit's net in the band, in minor units. */
    netMinor: number;
    /** This deposit's VAT in the band, in minor units. */
    vatMinor: number;
  }>;
  /**
   * `bands` padded with empty entries (`label` is "") to 4 slots, for the same
   * reason as `previousSlots`.
   */
  bandSlots: Array<{ label: string; net: string; vat: string; gross: string }>;
  /** Net total of the analysis, formatted; "" when `vatApplies` is false. */
  netTotal: string;
  /** VAT total of the analysis, formatted; "" when `vatApplies` is false. */
  vatTotal: string;
  /** The order gross in minor units, so tests assert on integers rather than Intl strings. */
  orderTotalMinor: number;
  /** This deposit in minor units. */
  depositMinor: number;
  /** All previous deposits in minor units. */
  previousMinor: number;
  /** What is left to invoice, in minor units. */
  leftMinor: number;
  /** The analysis net total in minor units; 0 when `vatApplies` is false. */
  netTotalMinor: number;
  /** The analysis VAT total in minor units; 0 when `vatApplies` is false. */
  vatTotalMinor: number;
}

/**
 * Every figure on a deposit invoice, worked once, so the page adds up to the
 * penny: previous deposits plus this one plus what is left is the order total.
 *
 * Money is integer minor units at the currency's own exponent (taken from
 * `Intl`), and every decimal is parsed exactly into a BigInt: no float is ever
 * multiplied and `toFixed` is never used. A percentage is taken of the order
 * gross, never of net plus VAT, because that is the figure on the customer's
 * quote. A mixed-rate deposit is spread across the rates by largest remainder,
 * so the shares add up to the deposit, and VAT is backed out of each share.
 *
 * It lives in a deriver because arithmetic on request data throws at publish.
 * It is self-contained on purpose: it travels to the engine as source text, so
 * a module-level helper or constant would be a ReferenceError there.
 *
 * Installed by `dxcl add advance-payment-invoice`.
 */
export default deriver({
  name: "advancePaymentInvoice",
  run: (order: unknown, deposit: unknown, previousDeposits: unknown, invoice: unknown): AdvancePaymentInvoiceFigures => {
    // deno-lint-ignore no-explicit-any
    const ord = order as any;
    // deno-lint-ignore no-explicit-any
    const dep = deposit as any;
    // deno-lint-ignore no-explicit-any
    const inv = invoice as any;

    const DECIMAL = /^(-?)(\d+)(?:\.(\d+))?(?:e([+-]?\d+))?$/;
    // String(n) is the shortest round-trip form, which is what the caller wrote.
    const decimal = (n: number): { int: bigint; scale: number } => {
      const match = DECIMAL.exec(String(n));
      if (!match) throw new Error("cannot read " + n + " as a decimal");
      const [, sign, whole, frac = "", exp = "0"] = match;
      let int = BigInt(whole + frac);
      let scale = frac.length - Number(exp);
      if (scale < 0) {
        int *= 10n ** BigInt(-scale);
        scale = 0;
      }
      return { int: sign ? -int : int, scale };
    };
    /** round(n / d) half away from zero, for d > 0. */
    const round = (n: bigint, d: bigint): bigint => {
      const abs = n < 0n ? -n : n;
      const out = (2n * abs + d) / (2n * d);
      return n < 0n ? -out : out;
    };
    const isDay = (v: unknown): v is string => {
      return typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v) &&
        !Number.isNaN(new Date(v.slice(0, 10) + "T00:00:00Z").getTime());
    };
    const finite = (v: unknown, path: string): number => {
      if (typeof v !== "number" || !Number.isFinite(v)) throw new Error(path + " must be a finite number");
      return v;
    };
    const present = (v: unknown) => v !== undefined && v !== null && v !== "";

    // 1. Validate, in the documented order.
    const jurisdictions = ["UK", "EU", "DE", "FR", "US"];
    if (!jurisdictions.includes(inv?.jurisdiction)) {
      throw new Error("invoice.jurisdiction must be one of " + jurisdictions.join(", "));
    }
    if (!isDay(inv.issueDate)) throw new Error("invoice.issueDate must be an ISO 8601 date");
    if (typeof ord?.reference !== "string" || ord.reference.trim() === "") {
      throw new Error("order.reference is required");
    }
    const jurisdiction: string = inv.jurisdiction;
    const us = jurisdiction === "US";
    const de = jurisdiction === "DE";
    let exponent = 2;
    try {
      if (typeof ord.currency !== "string") throw new Error("not a string");
      exponent = new Intl.NumberFormat("en-GB", { style: "currency", currency: ord.currency }).resolvedOptions()
        .maximumFractionDigits as number;
    } catch {
      throw new Error("order.currency must be an ISO 4217 code");
    }
    if (!Array.isArray(ord.lines) || ord.lines.length === 0) {
      throw new Error("order.lines is required and must not be empty");
    }
    const unit = 10n ** BigInt(exponent);
    /** A major-unit amount as minor units, rounded once. */
    const minor = (n: number): bigint => {
      const d = decimal(n);
      return round(d.int * unit, 10n ** BigInt(d.scale));
    };

    const lineRows: Array<{ gross: bigint | null; net: bigint | null; rate: number }> = [];
    ord.lines.forEach((line: { gross?: unknown; net?: unknown; vatRate?: unknown }, i: number) => {
      const at = "order.lines[" + i + "]";
      const hasGross = present(line?.gross);
      const hasNet = present(line?.net);
      if (hasGross === hasNet) throw new Error(at + " needs exactly one of gross or net");
      const amount = finite(hasGross ? line.gross : line.net, at + (hasGross ? ".gross" : ".net"));
      let rate = 0;
      if (!us) {
        rate = finite(line.vatRate, at + ".vatRate");
        if (rate < 0) throw new Error(at + ".vatRate must not be negative");
      }
      lineRows.push({ gross: hasGross ? minor(amount) : null, net: hasNet ? minor(amount) : null, rate });
    });

    const kinds = ["prepayment", "security"];
    if (!kinds.includes(dep?.kind)) throw new Error("deposit.kind must be one of " + kinds.join(", "));
    const prepayment = dep.kind === "prepayment";
    const hasPercent = present(dep.percent);
    const hasAmount = present(dep.amount);
    if (hasPercent === hasAmount) throw new Error("deposit needs exactly one of percent or amount");
    if (hasPercent) {
      const p = finite(dep.percent, "deposit.percent");
      if (!(p > 0 && p <= 100)) throw new Error("deposit.percent must be over 0 and at most 100");
    } else if (!(finite(dep.amount, "deposit.amount") > 0)) {
      throw new Error("deposit.amount must be greater than 0");
    }

    const prevInput: unknown[] = present(previousDeposits) ? (previousDeposits as unknown[]) : [];
    if (!Array.isArray(prevInput)) throw new Error("previousDeposits must be a list");
    const previous = prevInput.map((entry, i) => {
      // deno-lint-ignore no-explicit-any
      const p = entry as any;
      const at = "previousDeposits[" + i + "]";
      if (typeof p?.number !== "string" || p.number.trim() === "") throw new Error(at + ".number is required");
      if (!(finite(p.amount, at + ".amount") > 0)) throw new Error(at + ".amount must be greater than 0");
      return { number: p.number.trim() as string, amount: minor(p.amount) };
    });

    const dayOf = (v: unknown, path: string): string | null => {
      if (!present(v)) return null;
      if (!isDay(v)) throw new Error(path + " must be an ISO 8601 date");
      return v.slice(0, 10);
    };
    const orderDay = dayOf(ord.date, "order.date");
    const receivedDay = dayOf(dep.receivedDate, "deposit.receivedDate");
    const issueDay = inv.issueDate.slice(0, 10) as string;

    // 3. Bands, in the order their rate first appears.
    const bandGross: Array<{ rate: number; gross: bigint }> = [];
    if (us) {
      let sum = 0n;
      for (const l of lineRows) sum += (l.net ?? l.gross) as bigint;
      bandGross.push({ rate: 0, gross: sum });
    } else {
      const groups = new Map<number, { gross: bigint; net: bigint }>();
      for (const l of lineRows) {
        const g = groups.get(l.rate) ?? { gross: 0n, net: 0n };
        g.gross += l.gross ?? 0n;
        g.net += l.net ?? 0n;
        groups.set(l.rate, g);
      }
      for (const [rate, g] of groups) {
        const r = decimal(rate);
        // VAT on net lines is worked once per band, on the band's net sum.
        const vat = round(g.net * r.int, 100n * 10n ** BigInt(r.scale));
        const total = g.gross + g.net + vat;
        if (total < 0n) throw new Error("order.lines: the " + rate + "% band is negative");
        if (total > 0n) bandGross.push({ rate, gross: total });
      }
    }
    let G = 0n;
    for (const b of bandGross) G += b.gross;
    if (G <= 0n) throw new Error("order.lines must total more than zero");

    // 4. This deposit.
    let D: bigint;
    if (hasPercent) {
      const p = decimal(dep.percent);
      D = round(G * p.int, 100n * 10n ** BigInt(p.scale));
    } else {
      D = minor(dep.amount);
    }

    // Formatting needs the locale, so the cap message can print money.
    const locale = us ? "en-US" : "en-GB";
    const moneyFormat = new Intl.NumberFormat(locale, {
      style: "currency",
      currency: ord.currency,
      minimumFractionDigits: exponent,
      maximumFractionDigits: exponent,
    });
    const money = (n: bigint) => moneyFormat.format(Number(n) / Number(unit));
    if (D <= 0n) throw new Error("deposit rounds to zero: it must be at least one minor unit");

    // 5. Cap. The page prints 8 previous deposits and 4 rate bands in fixed slots.
    if (previous.length > 8) throw new Error("previousDeposits has " + previous.length + " entries; at most 8 print");
    if (bandGross.length > 4) throw new Error("order.lines has " + bandGross.length + " VAT rates; at most 4 print");
    let P = 0n;
    for (const p of previous) P += p.amount;
    if (P > G) {
      throw new Error("deposits requested (" + money(P) + ") exceed the order total (" + money(G) + ")");
    }
    if (prepayment && P + D > G) {
      throw new Error("deposits requested (" + money(P + D) + ") exceed the order total (" + money(G) + ")");
    }
    const L = prepayment ? G - P - D : G - P;

    // 6. Label.
    const pct = (n: number) => new Intl.NumberFormat("en-GB", { maximumFractionDigits: 4 }).format(n);
    const base = prepayment ? "Deposit requested" : "Security deposit";
    const depositLabel = hasPercent ? base + " (" + pct(dep.percent) + "% of the order total)" : base;

    // 7. Spread the deposit across the bands, by largest remainder.
    const vatApplies = prepayment && !us;
    const bands: AdvancePaymentInvoiceFigures["bands"] = [];
    let netTotal = 0n;
    let vatTotal = 0n;
    if (vatApplies) {
      const shares = bandGross.map((b) => (D * b.gross) / G);
      const rems = bandGross.map((b) => (D * b.gross) % G);
      let leftover = D;
      for (const s of shares) leftover -= s;
      const order = rems.map((_, i) => i).sort((a, b) => (rems[b] > rems[a] ? 1 : rems[b] < rems[a] ? -1 : a - b));
      for (let k = 0; k < Number(leftover); k++) shares[order[k]] += 1n;
      bandGross.forEach((b, i) => {
        const r = decimal(b.rate);
        const vat = round(shares[i] * r.int, 100n * 10n ** BigInt(r.scale) + r.int);
        const net = shares[i] - vat;
        netTotal += net;
        vatTotal += vat;
        bands.push({
          label: pct(b.rate) + "%",
          net: money(net),
          vat: money(vat),
          gross: money(shares[i]),
          shareMinor: Number(shares[i]),
          netMinor: Number(net),
          vatMinor: Number(vat),
        });
      });
    }

    // 8. Dates, in UTC so the day cannot shift on the engine.
    const dateFormat = new Intl.DateTimeFormat(locale, { dateStyle: "long", timeZone: "UTC" });
    const fmtDay = (d: string) => dateFormat.format(new Date(d + "T00:00:00Z"));
    const showPayment = vatApplies && receivedDay !== null && receivedDay !== issueDay;

    const slots = <T extends { number?: string; label?: string }>(rows: T[], n: number, empty: T): T[] => {
      const out = rows.slice();
      while (out.length < n) out.push(empty);
      return out;
    };

    return {
      title: !prepayment ? "Security deposit" : de ? "Anzahlung" : "Advance payment",
      orderReference: ord.reference.trim(),
      orderDate: orderDay === null ? null : fmtDay(orderDay),
      orderTotal: money(G),
      previous: previous.map((p) => ({ number: p.number, amount: money(p.amount) })),
      previousSlots: slots(previous.map((p) => ({ number: p.number, amount: money(p.amount) })), 8, { number: "", amount: "" }),
      depositLabel,
      deposit: money(D),
      leftToInvoice: money(L),
      paymentDateLabel: showPayment ? (de ? "Zeitpunkt der Vereinnahmung" : "Date of payment") : null,
      paymentDate: showPayment ? fmtDay(receivedDay as string) : null,
      vatApplies,
      bands,
      bandSlots: slots(bands.map((b) => ({ label: b.label, net: b.net, vat: b.vat, gross: b.gross })), 4, {
        label: "",
        net: "",
        vat: "",
        gross: "",
      }),
      netTotal: vatApplies ? money(netTotal) : "",
      vatTotal: vatApplies ? money(vatTotal) : "",
      orderTotalMinor: Number(G),
      depositMinor: Number(D),
      previousMinor: Number(P),
      leftMinor: Number(L),
      netTotalMinor: Number(netTotal),
      vatTotalMinor: Number(vatTotal),
    };
  },
});
