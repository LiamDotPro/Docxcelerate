/**
 * What the direct debit collection notice hands each of its arrangements: the
 * mandate details, already decided. An arrangement places these and decides
 * nothing, so both print the same lines under the same conditions.
 *
 * A field that is `null` is a line this debit does not carry. Test it in the
 * JSX, `{lines.account && …}`, so the test is compiled into a condition the
 * engine settles per document.
 *
 * Installed by `dxcl add direct-debit-collection-notice`, with the arrangements.
 */
export interface DirectDebitCollectionNoticeLines {
  /** Who collects, trimmed. */
  creditorName: string;
  /** "Creditor identifier" | "Service User Number", or null for ACH. */
  creditorIdLabel: string | null;
  /** The normalised identifier, or null for ACH. */
  creditorId: string | null;
  /** "Mandate reference" | "Direct Debit reference" | "Authorization reference". */
  mandateLabel: string;
  /** The mandate reference, trimmed. */
  mandateReference: string;
  /** "IBAN ending 3456" / "Account ending 1234", or null when none was given. */
  account: string | null;
}
