import { Paragraph, type Yield } from "docxcelerate/template";
import type { CanadianSalesTaxLines } from "./lines.ts";

/**
 * One labelled line under another, for a narrow column or a plain letter:
 * nothing has to line up with anything.
 */
export const stacked = (lines: CanadianSalesTaxLines): Yield => (
  <>
    <Paragraph id="canadian-sales-tax-subtotal">Subtotal: {lines.subtotal}</Paragraph>
    {lines.federalLabel && (
      <Paragraph id="canadian-sales-tax-federal">{lines.federalLabel}: {lines.federalAmount}</Paragraph>
    )}
    {lines.provincialLabel && (
      <Paragraph id="canadian-sales-tax-provincial">{lines.provincialLabel}: {lines.provincialAmount}</Paragraph>
    )}
    <Paragraph id="canadian-sales-tax-total">Total including tax: {lines.total}</Paragraph>
    {lines.bn && <Paragraph id="canadian-sales-tax-bn">GST/HST registration number: {lines.bn}</Paragraph>}
    {lines.qst && <Paragraph id="canadian-sales-tax-qst">QST registration number: {lines.qst}</Paragraph>}
    {lines.smallSupplierNote && <Paragraph id="canadian-sales-tax-small-supplier">{lines.smallSupplierNote}</Paragraph>}
    {lines.statusNote && <Paragraph id="canadian-sales-tax-status">{lines.statusNote}</Paragraph>}
    {lines.warning && <Paragraph id="canadian-sales-tax-warning">{lines.warning}</Paragraph>}
  </>
);
