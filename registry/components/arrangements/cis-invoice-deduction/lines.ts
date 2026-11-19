/**
 * What the CIS invoice deduction hands each of its arrangements: every line it
 * may print, already worked out. An arrangement places these and decides
 * nothing, so both print the same lines under the same conditions.
 *
 * A field that is `""` is a line this invoice does not carry. Test it in the
 * JSX, `{lines.plantHire && …}`, so the test is compiled into a condition the
 * engine settles per document.
 *
 * Installed by `dxcl add cis-invoice-deduction`, with the arrangements.
 */
export interface CisInvoiceDeductionLines {
  /** The invoice value before VAT, formatted. */
  gross: string;
  /** Materials at cost, formatted. */
  materials: string;
  /** Plant hire, formatted, or `""` when there is none. */
  plantHire: string;
  /** Consumables, formatted, or `""` when there are none. */
  consumables: string;
  /** The labour CIS is taken from, formatted. */
  labour: string;
  /** The VAT line's label, or `""` when no VAT line is printed. */
  vatLabel: string;
  /** The VAT line's value. */
  vatValue: string;
  /** The CIS line's label, naming the rate. */
  cisLabel: string;
  /** The CIS deduction, formatted. */
  cisDeduction: string;
  /** The amount payable, formatted. */
  payable: string;
}
