/**
 * What the GST statement hands each of its arrangements: items and totals,
 * already decided and formatted. An arrangement places these and decides
 * nothing, so both print the same lines under the same conditions.
 *
 * A `null` is a line this invoice does not carry. Test it in the JSX,
 * `{lines.subtotal && …}`, so the test compiles into a condition the engine
 * settles per document.
 *
 * Installed by `dxcl add gst-inclusive-statement`, with the arrangements.
 */
export interface GstInclusiveStatementLines {
  /** Every line in input order; a taxable line on a mixed invoice already ends " *". */
  items: Array<{ description: string; amount: string }>;
  /** Total of the lines before GST; exclusive basis only, otherwise null. */
  subtotal: string | null;
  /** The GST figure when it is shown, otherwise null. */
  gst: string | null;
  /** "GST" or "GST included"; "" when gst is null. */
  gstLabel: string;
  /** The amount payable. */
  total: string;
  /** "Total including GST" or "Total". */
  totalLabel: string;
}
