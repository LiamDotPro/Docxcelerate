import { Paragraph, type Yield } from "docxcelerate/template";
import type { PaymentSurchargeLineLines } from "./lines.ts";

/** Each line as a sentence, one under another, down the text column. */
export const stacked = (lines: PaymentSurchargeLineLines): Yield => (
  <>
    {lines.surchargeAmount && (
      <Paragraph id="surcharge-amount">{lines.surchargeLabel}: {lines.surchargeAmount}</Paragraph>
    )}
    {lines.taxAmount && <Paragraph id="surcharge-tax">{lines.taxLabel}: {lines.taxAmount}</Paragraph>}
    {lines.totalAmount && <Paragraph id="surcharge-total">{lines.totalLabel}: {lines.totalAmount}</Paragraph>}
    {lines.taxNote && <Paragraph id="surcharge-tax-note">{lines.taxNote}</Paragraph>}
    {lines.disclosure && <Paragraph id="surcharge-disclosure">{lines.disclosure}</Paragraph>}
    {lines.warning && <Paragraph id="surcharge-warning">{lines.warning}</Paragraph>}
  </>
);
