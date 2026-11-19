import { Paragraph, type Yield } from "docxcelerate/template";
import type { FeeCapStatusLines } from "./lines.ts";

/**
 * One labelled figure under another, down the text column. Pick it where the
 * figures sit in running text and nothing has to line up with anything.
 */
export const stacked = (lines: FeeCapStatusLines): Yield => (
  <>
    <Paragraph id="fee-cap-amount">{lines.capLabel}: {lines.capValue}</Paragraph>
    <Paragraph id="fee-cap-billed">Billed before this invoice: {lines.billedToDate}</Paragraph>
    <Paragraph id="fee-cap-period">Fees this period: {lines.thisPeriod}</Paragraph>
    {lines.writtenOff && <Paragraph id="fee-cap-written-off">{lines.writtenOffLabel}: {lines.writtenOff}</Paragraph>}
    {lines.charged && <Paragraph id="fee-cap-charged">Charged on this invoice: {lines.charged}</Paragraph>}
    {lines.withheld && <Paragraph id="fee-cap-withheld">{lines.withheldLabel}: {lines.withheld}</Paragraph>}
    {lines.payable && <Paragraph id="fee-cap-payable">Payable for these fees: {lines.payable}</Paragraph>}
    {lines.remaining && <Paragraph id="fee-cap-remaining">{lines.remainingLabel}: {lines.remaining}</Paragraph>}
  </>
);
