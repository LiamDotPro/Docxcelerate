import { Paragraph, type Yield } from "docxcelerate/template";
import type { FinnishVirtualBarcodeLines } from "./lines.ts";

/**
 * Each labelled line under the last. The code sits on its own paragraph under
 * its label: at 11pt its 54 digits need the whole text column, and a soft wrap
 * in the middle would make a payer doubt they had copied all of it.
 */
export const stacked = (lines: FinnishVirtualBarcodeLines): Yield => (
  <>
    {lines.iban && <Paragraph id="virtual-barcode-iban">Saajan tilinumero (IBAN): {lines.iban}</Paragraph>}
    {lines.reference && <Paragraph id="virtual-barcode-reference">Viitenumero: {lines.reference}</Paragraph>}
    {lines.code && <Paragraph id="virtual-barcode-label">Virtuaaliviivakoodi</Paragraph>}
    {lines.code && <Paragraph id="virtual-barcode-code">{lines.code}</Paragraph>}
  </>
);
