import { Paragraph, type Yield } from "docxcelerate/template";
import type { InvoiceDocumentTypeLines } from "./lines.ts";

/**
 * One labelled line under another, down the text column. The default: it
 * reads in any width, and nothing has to line up with anything.
 */
export const stacked = (lines: InvoiceDocumentTypeLines): Yield => (
  <>
    {lines.statement && <Paragraph id="invoice-type-statement">{lines.statement}</Paragraph>}
    <Paragraph id="invoice-type-number">Invoice number: {lines.number}</Paragraph>
    <Paragraph id="invoice-type-issue-date">Date of issue: {lines.issueDate}</Paragraph>
    {lines.supplyLine && <Paragraph id="invoice-type-supply">{lines.supplyLine}</Paragraph>}
    {lines.purchaseOrder && (
      <Paragraph id="invoice-type-purchase-order">Purchase order: {lines.purchaseOrder}</Paragraph>
    )}
    {lines.vatNumber && <Paragraph id="invoice-type-vat-number">VAT number: {lines.vatNumber}</Paragraph>}
  </>
);
