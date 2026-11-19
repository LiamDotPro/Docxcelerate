/**
 * What the payment surcharge line hands each of its arrangements: every line
 * it may print, already decided. An arrangement places these and decides
 * nothing. A `null` is a line this document does not carry; test it in the
 * JSX (`{lines.taxAmount && …}`) so it compiles to a condition the engine
 * settles per document.
 *
 * Installed by `dxcl add payment-surcharge-line`, with the arrangements.
 */
export interface PaymentSurchargeLineLines {
  /** "Surcharge for payment by commercial card (1.95%)", or null when none is charged. */
  surchargeLabel: string | null;
  /** The surcharge, formatted in the currency, or null when none is charged. */
  surchargeAmount: string | null;
  /** "VAT on surcharge" or "Sales tax on surcharge", or null when no tax is due on it. */
  taxLabel: string | null;
  /** The tax on the surcharge, formatted, or null when no tax is due on it. */
  taxAmount: string | null;
  /** How the surcharge was split across tax rates, when it touched more than one; else null. */
  taxNote: string | null;
  /** "Total including surcharge", or null when none is charged. */
  totalLabel: string | null;
  /** Pre-surcharge gross + surcharge + its tax, formatted, or null when none is charged. */
  totalAmount: string | null;
  /** The disclosure the rules require, with the rate charged, or null when none is charged. */
  disclosure: string | null;
  /** Why a requested surcharge was not applied, or null. Never set together with `surchargeAmount`. */
  warning: string | null;
}
