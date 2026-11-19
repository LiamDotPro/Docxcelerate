/**
 * What the Finnish virtual barcode hands each of its arrangements: every line
 * it may print, already decided. An arrangement places these and decides
 * nothing, so both print the same lines under the same conditions.
 *
 * An empty string is a line this document does not carry. Test it in the JSX,
 * `{lines.code && …}`, so the test is compiled into a condition the engine
 * settles per document.
 *
 * Installed by `dxcl add finnish-virtual-barcode`, with the arrangements.
 */
export interface FinnishVirtualBarcodeLines {
  /** The IBAN as printed (grouped in fours when valid), or "" for no line. */
  iban: string;
  /** The reference as printed (grouped when valid), or "" for no line. */
  reference: string;
  /** The 54 digits, unbroken, or "" when the code is withheld or nothing is owed. */
  code: string;
}
