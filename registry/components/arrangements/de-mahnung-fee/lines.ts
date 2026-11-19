/**
 * What the Mahnung fee hands each of its arrangements: every line, already
 * decided. An arrangement places these and decides nothing, so both print the
 * same lines under the same conditions.
 *
 * Installed by `dxcl add de-mahnung-fee`, with the arrangements.
 */
export interface DeMahnungFeeLines {
  /** Whether the debtor is a business: only then is there a Pauschale line. */
  business: boolean;
  /** The proof-of-lower-damage clause, or "" when this letter charges a consumer nothing. */
  proofLine: string;
  /** One entry per invoice, oldest due date first. */
  invoices: Array<{
    /** The invoice number as given. */
    number: string;
    /** "Rechnung … über …, fällig am …", with the due date as given, not shifted. */
    heading: string;
    /** Since when the invoice is in default, and why. */
    verzugLine: string;
    /** What this letter charges for this invoice. */
    feeLine: string;
    /** "" for a consumer. */
    pauschaleLine: string;
  }>;
}
