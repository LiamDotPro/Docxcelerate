import { Paragraph, type Yield } from "docxcelerate/template";
import type { DirectDebitCollectionNoticeLines } from "./lines.ts";

/**
 * One labelled mandate detail under another, down the text column. The
 * default: it reads in any width, and nothing has to line up with anything.
 */
export const stacked = (lines: DirectDebitCollectionNoticeLines): Yield => (
  <>
    <Paragraph id="dd-notice-creditor">Collected by: {lines.creditorName}</Paragraph>
    {lines.creditorId && (
      <Paragraph id="dd-notice-creditor-id">{lines.creditorIdLabel}: {lines.creditorId}</Paragraph>
    )}
    <Paragraph id="dd-notice-mandate">{lines.mandateLabel}: {lines.mandateReference}</Paragraph>
    {lines.account && <Paragraph id="dd-notice-account">From account: {lines.account}</Paragraph>}
  </>
);
