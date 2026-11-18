import { deriver } from "docxcelerate/template";

/**
 * Every figure on an invoice, worked once, so the page matches the customer's
 * accounting system to the penny.
 *
 * Money is integer minor units at the currency's own exponent (yen 0, pounds 2,
 * dinars 3, taken from `Intl`), and every decimal is parsed exactly into a
 * BigInt: no float is ever multiplied and `toFixed` is never used. A line is
 * rounded once, half away from zero. Lines are grouped by tax rate and VAT is
 * worked once per band on the band's sum (rate-level rounding, which HMRC and
 * the EU both accept); summing per-line VAT drifts a penny from that figure.
 * Inclusive prices have VAT backed out of the band's gross, so the total is
 * exactly what was charged. An exempt band keeps its own label and is never
 * a 0% band, because zero-rated is a different treatment.
 *
 * It lives in a deriver because arithmetic on request data throws at publish.
 *
 * Installed by `dxcl add invoice-totals`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface InvoiceTotalsData {
  invoice: {
    /**
     * ISO 4217 code: "GBP", "JPY", "KWD". Required: an invoice that does not
     * say its currency is not a valid invoice, so there is no default. Its
     * exponent (decimal places) comes from Intl. A code Intl rejects throws,
     * naming `invoice.currency`.
     */
    currency: string;
    /**
     * Whether `unitPrice` excludes VAT ("exclusive", B2B) or includes it
     * ("inclusive", retail). Defaults to "exclusive". Any other value throws.
     */
    pricing?: "exclusive" | "inclusive";
    /**
     * One entry per line, in the order they are printed. Missing or empty
     * prints the "no charges" paragraph and no tables.
     */
    lines: Array<{
      /** What was supplied, as printed. Wraps; never truncated. */
      description: string;
      /** How many. May be fractional (1.25 hours), zero, or negative. Not finite throws. */
      quantity: number;
      /** Printed after the quantity: "h", "visits", "kWh". Omitted, the quantity prints bare. */
      unit?: string;
      /**
       * Price per unit in major units (pounds, not pence), net or gross as
       * `pricing` says. Up to four decimal places are honoured exactly.
       * Negative for a discount or goodwill credit. Not finite throws.
       */
      unitPrice: number;
      /**
       * VAT rate as a percentage: 20, 5, 12.5, 0. 0 is zero-rated and prints
       * as "0%". `null` is exempt or outside the scope of VAT, with no rate at
       * all. Required: `undefined`, a negative rate or a non-finite rate
       * throws rather than being guessed as exempt.
       */
      taxRate: number | null;
      /**
       * Read only when `taxRate` is null: the band's label, such as "Exempt"
       * or "Outside the scope of VAT". Defaults to "No VAT". Ignored on a
       * rated line, which is always labelled by its rate.
       */
      taxLabel?: string;
    }>;
  };
}

/** What the deriver hands the node: every display value already formatted. */
export interface InvoiceTotals {
  empty: boolean;
  multipleBands: boolean;
  amountHeading: string;
  lines: Array<{ description: string; quantity: string; unitPrice: string; tax: string; amount: string }>;
  bands: Array<{ label: string; net: string; vat: string; netMinor: number; vatMinor: number }>;
  subtotalLabel: string;
  subtotal: string;
  vatLabel: string;
  vat: string;
  total: string;
  subtotalMinor: number;
  vatMinor: number;
  totalMinor: number;
}

interface Decimal {
  /** The value times 10^scale. */
  int: bigint;
  scale: number;
}

const DECIMAL = /^(-?)(\d+)(?:\.(\d+))?(?:e([+-]?\d+))?$/;

// String(n) is the shortest round-trip form, which is what the caller wrote.
function decimal(n: number): Decimal {
  const match = DECIMAL.exec(String(n));

  if (!match) throw new Error(`cannot read ${n} as a decimal`);

  const [, sign, whole, frac = "", exp = "0"] = match;
  let int = BigInt(whole + frac);
  let scale = frac.length - Number(exp);

  if (scale < 0) {
    int *= 10n ** BigInt(-scale);
    scale = 0;
  }

  return { int: sign ? -int : int, scale };
}

/** round(n / d) half away from zero, for d > 0. */
function round(n: bigint, d: bigint): bigint {
  const abs = n < 0n ? -n : n;
  const out = (2n * abs + d) / (2n * d);

  return n < 0n ? -out : out;
}

function finite(value: unknown, path: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${path} is not a finite number`);
  }

  return value;
}

interface Band {
  label: string;
  rate: Decimal | null;
  gross: bigint;
}

function invoiceTotalsOf(invoice: InvoiceTotalsData["invoice"]): InvoiceTotals {
  let exponent: number;

  try {
    exponent = new Intl.NumberFormat("en-GB", { style: "currency", currency: invoice?.currency })
      .resolvedOptions().maximumFractionDigits!;
  } catch {
    throw new Error(`invoice.currency is not a valid ISO 4217 code: ${String(invoice?.currency)}`);
  }

  const pricing = invoice.pricing ?? "exclusive";

  if (pricing !== "exclusive" && pricing !== "inclusive") {
    throw new Error(`invoice.pricing must be "exclusive" or "inclusive", not ${String(pricing)}`);
  }

  const inclusive = pricing === "inclusive";
  const input = invoice.lines ?? [];

  if (!Array.isArray(input)) throw new Error("invoice.lines is not an array");

  const scaleUp = 10n ** BigInt(exponent);
  const toMoney = (minor: bigint) =>
    new Intl.NumberFormat("en-GB", {
      style: "currency",
      currency: invoice.currency,
      minimumFractionDigits: exponent,
      maximumFractionDigits: exponent,
    }).format(Number(minor) / Number(scaleUp));
  const toQuantity = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 4 });

  const rated = new Map<number, Band>();
  const unrated = new Map<string, Band>();

  const lines = input.map((line, i) => {
    const at = `invoice.lines[${i}]`;
    const q = decimal(finite(line.quantity, `${at}.quantity`));
    const p = decimal(finite(line.unitPrice, `${at}.unitPrice`));
    const amount = round(q.int * p.int * scaleUp, 10n ** BigInt(q.scale + p.scale));
    let band: Band;

    if (line.taxRate === null) {
      const label = line.taxLabel ?? "No VAT";

      band = unrated.get(label) ?? { label, rate: null, gross: 0n };
      unrated.set(label, band);
    } else {
      const rate = finite(line.taxRate, `${at}.taxRate`);

      if (rate < 0) throw new Error(`${at}.taxRate is negative`);

      band = rated.get(rate) ?? { label: `${rate}%`, rate: decimal(rate), gross: 0n };
      rated.set(rate, band);
    }

    band.gross += amount;

    const priceDigits = Math.max(exponent, Math.min(4, p.scale));

    return {
      description: line.description,
      quantity: toQuantity.format(Number(line.quantity)) + (line.unit ? ` ${line.unit}` : ""),
      unitPrice: new Intl.NumberFormat("en-GB", {
        style: "currency",
        currency: invoice.currency,
        minimumFractionDigits: exponent,
        maximumFractionDigits: priceDigits,
      }).format(line.unitPrice),
      tax: band.label,
      amount: toMoney(amount),
    };
  });

  const ordered = [
    ...[...rated.entries()].sort((a, b) => b[0] - a[0]).map(([, band]) => band),
    ...unrated.values(),
  ];

  let subtotal = 0n;
  let vat = 0n;

  const bands = ordered.map((band) => {
    let bandVat = 0n;

    if (band.rate) {
      const { int: r, scale } = band.rate;
      const hundred = 100n * 10n ** BigInt(scale);

      bandVat = inclusive ? round(band.gross * r, hundred + r) : round(band.gross * r, hundred);
    }

    const net = inclusive ? band.gross - bandVat : band.gross;

    subtotal += net;
    vat += bandVat;

    return {
      label: band.label,
      net: toMoney(net),
      vat: band.rate ? toMoney(bandVat) : "—",
      netMinor: Number(net),
      vatMinor: Number(bandVat),
    };
  });

  const single = bands.length === 1 ? ordered[0] : undefined;

  return {
    empty: input.length === 0,
    multipleBands: bands.length > 1,
    amountHeading: inclusive ? "Amount (incl. VAT)" : "Amount (excl. VAT)",
    lines,
    bands,
    subtotalLabel: inclusive ? "Total excluding VAT" : "Subtotal",
    subtotal: toMoney(subtotal),
    vatLabel: single ? (single.rate ? `VAT at ${single.label}` : single.label) : "VAT",
    vat: single && !single.rate ? "—" : toMoney(vat),
    total: toMoney(subtotal + vat),
    subtotalMinor: Number(subtotal),
    vatMinor: Number(vat),
    totalMinor: Number(subtotal + vat),
  };
}

export default deriver({ name: "invoiceTotals", run: invoiceTotalsOf });
