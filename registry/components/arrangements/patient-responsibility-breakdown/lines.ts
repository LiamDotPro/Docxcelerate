/**
 * What the patient responsibility breakdown hands each of its arrangements:
 * every line already decided and formatted. An arrangement places these and
 * decides nothing, so both print every amount, $0.00 included, and the note
 * under the same condition.
 *
 * Installed by `dxcl add patient-responsibility-breakdown`, with the arrangements.
 */
export interface PatientResponsibilityBreakdownLines {
  /** One per claim line, in service-date order. A published loop over the deriver's rows. */
  rows: Array<{
    /** The description, as given. */
    service: string;
    /** "Sep 14, 2026 · 27447", or just the date when there is no code. */
    detail: string;
    /** Formatted money, such as "$248,750.00". */
    billed: string;
    /** $0.00 on a non-covered or denied line. */
    allowed: string;
    /** The deductible taken on this line. */
    deductible: string;
    /** The copay taken on this line. */
    copay: string;
    /** The coinsurance taken on this line. */
    coinsurance: string;
    /** Deductible + copay + coinsurance + any non-covered charge + any balance bill. */
    youOwe: string;
    /** One sentence explaining the line, or null for an ordinary covered line. */
    note: string | null;
  }>;
  /** The column totals, each the sum of the rows above it, formatted. */
  total: {
    /** The sum of the billed charges. */
    billed: string;
    /** The sum of the allowed amounts. */
    allowed: string;
    /** The sum of the deductible portions. */
    deductible: string;
    /** The sum of the copays. */
    copay: string;
    /** The sum of the coinsurance amounts. */
    coinsurance: string;
    /** The sum of what the patient owes, never recomputed from rates. */
    youOwe: string;
  };
}
