import { Paragraph, type Yield } from "docxcelerate/template";
import type { RenewalAndPriceChangeNoticeLines } from "./lines.ts";

/**
 * One labelled line under another, down the text column. The default: it
 * reads in any width, and nothing has to line up with anything.
 */
export const stacked = (lines: RenewalAndPriceChangeNoticeLines): Yield => (
  <>
    <Paragraph id="renewal-notice-statement">{lines.statement}</Paragraph>
    <Paragraph id="renewal-notice-renewal">Renewal: {lines.renewal}</Paragraph>
    {lines.trial && <Paragraph id="renewal-notice-trial">Introductory period ends: {lines.trial}</Paragraph>}
    {lines.priceChange && <Paragraph id="renewal-notice-price-change">Price change: {lines.priceChange}</Paragraph>}
    {lines.previousPayment && (
      <Paragraph id="renewal-notice-previous-payment">Previous payment: {lines.previousPayment}</Paragraph>
    )}
    {lines.cancelBy && <Paragraph id="renewal-notice-cancel-by">Cancellation deadline: {lines.cancelBy}</Paragraph>}
    <Paragraph id="renewal-notice-how-to-cancel">How to cancel: {lines.howToCancel}</Paragraph>
    {lines.consequence && <Paragraph id="renewal-notice-consequence">{lines.consequence}</Paragraph>}
    {lines.warning && <Paragraph id="renewal-notice-warning">{lines.warning}</Paragraph>}
  </>
);
