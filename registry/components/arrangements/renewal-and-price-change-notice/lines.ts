/**
 * What the renewal and price change notice hands each of its arrangements:
 * every line it may print, already decided. An arrangement places these and
 * decides nothing, so both print the same lines under the same conditions.
 *
 * A field that is `null` is a line this notice does not carry. Test it in the
 * JSX, `{lines.trial && …}`, so the test is compiled into a condition the
 * engine settles per document.
 *
 * Installed by `dxcl add renewal-and-price-change-notice`, with the arrangements.
 */
export interface RenewalAndPriceChangeNoticeLines {
  /** The opening sentence: when the contract renews, or that it continues until cancelled. */
  statement: string;
  /** The value of the "Renewal" line: the date and what is charged. */
  renewal: string;
  /** The value of the "Introductory period ends" line, or null when there is no trial. */
  trial: string | null;
  /** The value of the "Price change" line, or null when the price does not change. */
  priceChange: string | null;
  /** The value of the "Previous payment" line, or null when none was given. */
  previousPayment: string | null;
  /** The value of the "Cancellation deadline" line, or null when no fixed deadline applies. */
  cancelBy: string | null;
  /** How to cancel, as the document's data gives it. */
  howToCancel: string;
  /** What the law gives the customer when the deadline is gone or the notice is late, or null. */
  consequence: string | null;
  /** A "Check before sending" note for whoever issues the notice, or null. */
  warning: string | null;
}
