/**
 * What the sales-tax block hands each of its arrangements: every line it may
 * print, already decided. An arrangement places these and decides nothing, so
 * both print the same lines under the same conditions.
 *
 * A field that is `null` is a line this invoice does not carry. Test it in the
 * JSX, `{lines.federalLabel && …}`, so the test is compiled into a condition
 * the engine settles per document.
 *
 * Installed by `dxcl add canadian-sales-tax`, with the arrangements.
 */
export interface CanadianSalesTaxLines {
  /** Sum of all line amounts, formatted: "$1,307,465.50". */
  subtotal: string;
  /** "GST 5% on $1,305,015.50" or "HST 14% on …"; null when the seller is not GST/HST-registered. */
  federalLabel: string | null;
  /** That tax, formatted: "$65,250.78". Null with the label. */
  federalAmount: string | null;
  /** "QST 9.975% on $1,303,175.50", "PST 7% on …", "RST 7% on …"; null when none is charged. */
  provincialLabel: string | null;
  /** That tax, formatted. Null with the label. */
  provincialAmount: string | null;
  /** Subtotal plus every charged tax, formatted. */
  total: string;
  /** Normalised valid GST/HST number, or null. */
  bn: string | null;
  /** Normalised valid QST number, or null. */
  qst: string | null;
  /** The small-supplier sentence, or null. */
  smallSupplierNote: string | null;
  /** "No GST on: …" sentences for mixed-status invoices, or null. */
  statusNote: string | null;
  /** Every warning sentence joined by a space, or null. */
  warning: string | null;
}
