/**
 * What the dunning reminder hands each of its arrangements: the invoice lines
 * it has already decided on. An arrangement places these and decides nothing,
 * so both print the same slots, overflow and total under the same conditions.
 *
 * There are always eight slots. An unused slot has every field `""`, so test
 * `{lines.invoices[0].number && …}` in the JSX, which compiles into a condition
 * the engine settles per document.
 *
 * Installed by `dxcl add dunning-reminder`, with the arrangements.
 */
export interface DunningReminderLines {
  /** The eight invoice slots, oldest first. */
  invoices: Array<{
    /** The invoice number, as given. `""` when the slot is unused. */
    number: string;
    /** The due date, formatted. */
    dueDate: string;
    /** Days overdue, as "1 day" or "N days". */
    overdue: string;
    /** The open amount, formatted. */
    amount: string;
    /** The whole invoice as one sentence, for a stacked layout. */
    line: string;
  }>;
  /** The sentence covering invoices beyond the eighth, or `""` when there are none. */
  overflow: string;
  /** The formatted sum of every included invoice, overflow too. */
  total: string;
}
