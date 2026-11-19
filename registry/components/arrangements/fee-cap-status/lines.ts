/**
 * What the fee cap status hands each of its arrangements: every figure it may
 * print, already decided and formatted. An arrangement places these and
 * decides nothing, so both print the same rows under the same conditions.
 *
 * A field that is `null` is a row this invoice does not carry. Test the value
 * in the JSX, `{lines.writtenOff && …}`, so the test is compiled into a
 * condition the engine settles per document.
 *
 * Installed by `dxcl add fee-cap-status`, with the arrangements.
 */
export interface FeeCapStatusLines {
  /** The cap row's label: "Agreed cap", "Ceiling", "Estimate" or "Revised estimate". */
  capLabel: string;
  /** The limit fees are measured against. */
  capValue: string;
  /** Fees charged on earlier invoices. */
  billedToDate: string;
  /** Fees for this period, before the cap and withholding. */
  thisPeriod: string;
  /** "Above the cap, not charged" or "Above the ceiling, not charged"; null when nothing is written off. */
  writtenOffLabel: string | null;
  /** The fees above the limit that are not charged, or null. */
  writtenOff: string | null;
  /** What this invoice charges; null unless something is written off, since otherwise it equals the period's fees. */
  charged: string | null;
  /** "Retained (5%)" or "Retained (to the agreed limit)"; null when nothing is retained. */
  withheldLabel: string | null;
  /** The retention held back on this invoice, or null. */
  withheld: string | null;
  /** What is payable for these fees after retention, or null. */
  payable: string | null;
  /** "Remaining under the cap" and so on; null once the limit is reached or passed. */
  remainingLabel: string | null;
  /** What is left under the limit, or null. */
  remaining: string | null;
}
