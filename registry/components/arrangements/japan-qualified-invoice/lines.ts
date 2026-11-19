/**
 * What the Japanese qualified invoice hands each of its arrangements: the
 * deriver's output, already final text, plus the buyer. An arrangement places
 * these and decides nothing, so both print the same lines under the same
 * conditions.
 *
 * A `null` field is a line this invoice does not carry. Test it in the JSX,
 * `{lines.standard.tax && …}`, so the test is compiled into a condition the
 * engine settles per document.
 *
 * Installed by `dxcl add japan-qualified-invoice`, with the arrangements.
 */
export interface JapanQualifiedInvoiceLines {
  /** The buyer's name as given, honorific included; missing on some simplified invoices. */
  buyer?: string;
  /** The transaction date, "2026年11月30日". */
  date: string;
  /** The normalised T-number, or null when the seller is not registered. */
  registration: string | null;
  /** The amount column's header: "金額（税抜）", "金額（税込）" or "金額". */
  amountHeader: string;
  /** Always ten slots; an unused one is `{ description: "", amount: "" }`. */
  items: Array<{ description: string; amount: string }>;
  /** The 10% band. `total` and `tax` are null when it has no lines or no tax is stated. */
  standard: { label: string; total: string | null; taxLabel: string; tax: string | null };
  /** The 8% band, the same shape. */
  reduced: { label: string; total: string | null; taxLabel: string; tax: string | null };
  /** "合計（税込）" or "合計". */
  totalLabel: string;
  /** The grand total, "251,872,208円". */
  total: string;
  /** "※は軽減税率対象" when any reduced-rate line is present. */
  legend: string | null;
  /** The transitional-measure note, for an unregistered seller. */
  note: string | null;
}
