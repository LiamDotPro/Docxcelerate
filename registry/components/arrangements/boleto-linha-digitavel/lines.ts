/**
 * What the boleto component hands each of its arrangements: every line it may
 * print, already decided. An arrangement places these and decides nothing, so
 * both print the same lines under the same conditions.
 *
 * A field that is `null` is a line this document does not carry. Test it in
 * the JSX, `{lines.line && …}`, so the test is compiled into a condition the
 * engine settles per document.
 *
 * Installed by `dxcl add boleto-linha-digitavel`, with the arrangements.
 */
export interface BoletoLinhaDigitavelLines {
  /** The grouped 54-character line, or null when withheld. */
  line: string | null;
  /** "15/12/2026", or "Contra apresentação"; null when withheld. */
  dueDate: string | null;
  /** "R$ 8.750.432,19"; null when withheld. */
  amount: string | null;
  /** The whole "unavailable" sentence, or null when a line is printed. */
  withheld: string | null;
}
