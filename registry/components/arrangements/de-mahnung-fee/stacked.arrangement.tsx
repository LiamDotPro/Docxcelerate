import { Paragraph, type Yield } from "docxcelerate/template";
import type { DeMahnungFeeLines } from "./lines.ts";

/**
 * Each invoice as a heading with its lines beneath, as letter prose. The
 * business and consumer arms are chosen once, outside the loop, so no body
 * holds a condition and every string in one is always non-empty.
 */
export const stacked = (lines: DeMahnungFeeLines): Yield => (
  <>
    {lines.business
      ? lines.invoices.map((invoice) => (
        <>
          <Paragraph id="de-mahnung-fee-business-invoice">{invoice.heading}</Paragraph>
          <Paragraph id="de-mahnung-fee-business-verzug">{invoice.verzugLine}</Paragraph>
          <Paragraph id="de-mahnung-fee-business-fee">{invoice.feeLine}</Paragraph>
          <Paragraph id="de-mahnung-fee-business-pauschale">{invoice.pauschaleLine}</Paragraph>
        </>
      ))
      : lines.invoices.map((invoice) => (
        <>
          <Paragraph id="de-mahnung-fee-consumer-invoice">{invoice.heading}</Paragraph>
          <Paragraph id="de-mahnung-fee-consumer-verzug">{invoice.verzugLine}</Paragraph>
          <Paragraph id="de-mahnung-fee-consumer-fee">{invoice.feeLine}</Paragraph>
        </>
      ))}
    {lines.proofLine && <Paragraph id="de-mahnung-fee-proof">{lines.proofLine}</Paragraph>}
  </>
);
