/**
 * What the invoice document type hands each of its arrangements: every line
 * it may print, already decided. An arrangement places these and decides
 * nothing, so all three print the same lines under the same conditions.
 *
 * A field that is `null` (or a missing purchase order) is a line this invoice
 * does not carry. Test it in the JSX, `{lines.supplyLine && …}`, so the test
 * is compiled into a condition the engine settles per document.
 *
 * Installed by `dxcl add invoice-document-type`, with the arrangements.
 */
export interface InvoiceDocumentTypeLines {
  /** "This is not a VAT invoice.", a self-billing statement, or null. */
  statement: string | null;
  /** The invoice number, as given. */
  number: string;
  /** The date of issue, formatted for the jurisdiction. */
  issueDate: string;
  /** The supply date or period as one line, or null when none is printed. */
  supplyLine: string | null;
  /** The same, its label alone: "Date of supply", "Leistungszeitraum". */
  supplyLabel: string | null;
  /** The same, its value alone. */
  supplyValue: string | null;
  /** The customer's purchase order, when there is one. */
  purchaseOrder?: string;
  /** The seller's VAT number, or null when no VAT applies. */
  vatNumber: string | null;
}
