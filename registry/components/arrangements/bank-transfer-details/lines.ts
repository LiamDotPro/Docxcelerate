/**
 * What bank transfer details hands each of its arrangements: every line it may
 * print, already decided. An arrangement places these and decides nothing, so
 * all three print the same lines under the same conditions.
 *
 * A field that is `null` is a line this payer does not get. Test it in the JSX,
 * `{lines.iban && …}`, so the test is compiled into a condition the engine
 * settles per document.
 *
 * Installed by `dxcl add bank-transfer-details`, with the arrangements.
 */
export interface BankTransferDetailsLines {
  /** The changed-details warning, or null. Always placed first. */
  warning: string | null;
  /** One sentence naming how to pay, or saying the details cannot be shown. Always present. */
  method: string;
  /** The account name exactly as held, or null when no details are printed. */
  accountName: string | null;
  /** `60-16-13`, UK rail only. */
  sortCode: string | null;
  /** "Routing number (ACH and wire)" or "ACH routing number", US rail only. */
  routingLabel: string | null;
  /** The ACH routing number, US rail only. */
  routing: string | null;
  /** The wire routing number, only when valid and different from the ACH one. */
  wireRouting: string | null;
  /** The account number, UK or US rail. */
  accountNumber: string | null;
  /** The IBAN in groups of four, SEPA or international rail. */
  iban: string | null;
  /** The BIC, SEPA (when valid) or international rail. */
  bic: string | null;
  /** The reference to quote, already fitted to the rail, or null when no details are printed. */
  reference: string | null;
}
