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
  /** The base amount, formatted. */
  base: string;
  /** The rivalsa INPS row's label, or null when the supplier charges none. */
  rivalsaLabel: string | null;
  /** The rivalsa INPS amount, or null when there is no row. */
  rivalsa: string | null;
  /** The cassa previdenza row's label, or null when none applies. */
  cassaLabel: string | null;
  /** The cassa previdenza amount, or null when there is no row. */
  cassa: string | null;
  /** The VAT row's label, or null when no VAT is charged. */
  vatLabel: string | null;
  /** The VAT amount, or null when there is no row. */
  vat: string | null;
  /** The recargo di equivalencia row's label, or null when none applies. */
  reLabel: string | null;
  /** The recargo di equivalencia amount, or null when there is no row. */
  re: string | null;
  /** The VAT-exempt row's label, or null when nothing is exempt. */
  exemptLabel: string | null;
  /** The VAT-exempt amount, or null when there is no row. */
  exempt: string | null;
  /** The imposta di bollo row's label, or null when no stamp is charged. */
  stampLabel: string | null;
  /** The imposta di bollo amount, or null when there is no row. */
  stamp: string | null;
  /** The document total: always present. */
  totalLabel: string;
  /** The document total, formatted. */
  total: string;
  /** A deduction prints with a leading minus. */
  withholdingLabel: string | null;
  /** The withholding amount, or null when nothing is withheld. */
  withholding: string | null;
  /** The split-payment row's label, or null when it does not apply. */
  splitPaymentLabel: string | null;
  /** The IVA the customer pays directly, or null when it does not apply. */
  splitPayment: string | null;
  /** What the customer transfers: always present. */
  netLabel: string;
  /** The net payable, formatted. */
  net: string;
}
