/**
 * What the net payable section hands each of its arrangements: every row it
 * may print, already decided. An arrangement places these and decides
 * nothing, so both print the same rows under the same conditions.
 *
 * A `null` label is a row this invoice does not carry. Test it in the JSX,
 * `{lines.vatLabel && …}`, so the test is compiled into a condition the engine
 * settles per document.
 *
 * Installed by `dxcl add net-payable-withholdings`, with the arrangements.
 */
export interface NetPayableWithholdingsLines {
  /** The base row's label and figure: always present. */
  baseLabel: string;
  base: string;
  rivalsaLabel: string | null;
  rivalsa: string | null;
  cassaLabel: string | null;
  cassa: string | null;
  vatLabel: string | null;
  vat: string | null;
  reLabel: string | null;
  re: string | null;
  exemptLabel: string | null;
  exempt: string | null;
  stampLabel: string | null;
  stamp: string | null;
  /** The document total: always present. */
  totalLabel: string;
  total: string;
  /** A deduction prints with a leading minus. */
  withholdingLabel: string | null;
  withholding: string | null;
  splitPaymentLabel: string | null;
  splitPayment: string | null;
  /** What the customer transfers: always present. */
  netLabel: string;
  net: string;
}
