/**
 * What the Dutch collection-costs component hands each of its arrangements:
 * every figure it may print, already decided and formatted. An arrangement
 * places these and decides nothing, so both print the same lines under the
 * same conditions.
 *
 * An empty `interest` or `vat` is a line this letter does not carry. Test it
 * in the JSX, `{lines.interest && …}`, so the test is compiled into a
 * condition the engine settles per document.
 *
 * Installed by `dxcl add nl-collection-costs`, with the arrangements.
 */
export interface NlCollectionCostsLines {
  /** The invoice numbers, joined: "A, B en C". */
  invoices: string;
  /** The sum of the invoice principals. */
  principal: string;
  /** Statutory interest added to the base, or "" when none is. */
  interest: string;
  /** The costs excluding btw. */
  costs: string;
  /** The btw on the costs, or "" when none is added. */
  vat: string;
  /** What is owed in all. */
  total: string;
  /** Its label, which carries the payment date on a veertiendagenbrief. */
  totalLabel: string;
}
