/**
 * What the advance payment invoice hands each of its arrangements: every line
 * it may print, already decided. An arrangement places these and decides
 * nothing, so both print the same lines under the same conditions.
 *
 * A field that is `null` is a line this invoice does not carry. Test it in the
 * JSX, `{lines.orderDate && …}`, so the test is compiled into a condition the
 * engine settles per document. The two lists are mapped as they are: an empty
 * one prints nothing, so neither needs a guard.
 *
 * Installed by `dxcl add advance-payment-invoice`, with the arrangements.
 */
export interface AdvancePaymentInvoiceLines {
  /** The order the deposit is against, as given (trimmed). */
  orderReference: string;
  /** The order date, formatted, or null when none was given. */
  orderDate: string | null;
  /** The order gross, formatted. */
  orderTotal: string;
  /** One per deposit already requested, in the order they were issued. */
  previous: Array<{ number: string; amount: string }>;
  /** What the deposit line says: "Deposit requested (30% of the order total)", "Security deposit". */
  depositLabel: string;
  /** This request, formatted. */
  deposit: string;
  /** What is left to invoice, formatted. */
  leftToInvoice: string;
  /** "Date of payment" or "Zeitpunkt der Vereinnahmung"; null when no date prints. */
  paymentDateLabel: string | null;
  /** The date payment was received, formatted; null when no date prints. */
  paymentDate: string | null;
  /** Whether the VAT analysis is printed. */
  vatApplies: boolean;
  /** This deposit's VAT analysis, one per rate; empty when `vatApplies` is false. */
  bands: Array<{ label: string; net: string; vat: string; gross: string }>;
  /** The analysis net total; "" when `vatApplies` is false. */
  netTotal: string;
  /** The analysis VAT total; "" when `vatApplies` is false. */
  vatTotal: string;
}
