/**
 * What the VAT in local currency hands each of its arrangements: the bands,
 * totals and advances, already converted and formatted. An arrangement places
 * these and decides nothing, so both print the same lines under the same
 * conditions.
 *
 * Installed by `dxcl add vat-in-local-currency`, with the arrangements.
 */
export interface VatInLocalCurrencyLines {
  /** The currency the invoice is issued in, "EUR". */
  invoiceCurrency: string;
  /** The currency the VAT is owed in, "GBP". */
  vatCurrency: string;
  /** One entry per rate band, in the order the invoice gives them. */
  bands: Array<{ rate: string; net: string; vat: string; localNet: string; localVat: string }>;
  /** The sums of the bands, in both currencies. Never a converted gross. */
  total: { net: string; vat: string; localNet: string; localVat: string };
  /** Whether advances were deducted, which the loop below is empty without. */
  hasAdvances: boolean;
  /** One entry per advance, each at the rate it was declared at. */
  advances: Array<{ label: string; rateLine: string; net: string; vat: string; localNet: string; localVat: string }>;
}
