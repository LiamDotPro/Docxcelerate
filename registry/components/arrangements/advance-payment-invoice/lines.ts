/**
 * What the advance payment invoice hands each of its arrangements: every line
 * it may print, already decided. An arrangement places these and decides
 * nothing, so both print the same lines under the same conditions.
 *
 * A field that is `null` is a line this invoice does not carry. Test it in the
 * JSX, `{lines.orderDate && …}`, so the test is compiled into a condition the
 * engine settles per document. The two lists are fixed slots, 8 previous
 * deposits and 4 rate bands, each printed when its first field is not empty:
 * the resolver reads a loop's collection from request data and ignores the
 * derived scope, so a `.map()` over them would publish and then print no rows.
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
  /** Eight slots: a deposit already requested, in the order issued, then empty ones (`number` is ""). */
  previous: Array<{
    /** The earlier deposit's invoice number; "" in an unused slot. */
    number: string;
    /** The earlier deposit, formatted. */
    amount: string;
  }>;
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
  /** Four slots: this deposit's VAT analysis, one per rate, then empty ones (`label` is ""). */
  bands: Array<{
    /** The band's rate, formatted; "" in an unused slot. */
    label: string;
    /** This deposit's net in the band, formatted. */
    net: string;
    /** This deposit's VAT in the band, formatted. */
    vat: string;
    /** This deposit's share of the band, formatted. */
    gross: string;
  }>;
  /** The analysis net total; "" when `vatApplies` is false. */
  netTotal: string;
  /** The analysis VAT total; "" when `vatApplies` is false. */
  vatTotal: string;
}
