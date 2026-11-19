import { Paragraph, type Yield } from "docxcelerate/template";
import type { CisInvoiceDeductionLines } from "./lines.ts";

/** One "Label: figure" line under another, for a narrow column. */
export const stacked = (lines: CisInvoiceDeductionLines): Yield => (
  <>
    <Paragraph id="cis-deduction-gross">Gross amount, excluding VAT: {lines.gross}</Paragraph>
    <Paragraph id="cis-deduction-materials">Less materials at cost: {lines.materials}</Paragraph>
    {lines.plantHire && <Paragraph id="cis-deduction-plant-hire">Less plant hire: {lines.plantHire}</Paragraph>}
    {lines.consumables && (
      <Paragraph id="cis-deduction-consumables">Less consumables: {lines.consumables}</Paragraph>
    )}
    <Paragraph id="cis-deduction-labour">Labour subject to CIS: {lines.labour}</Paragraph>
    {lines.vatLabel && <Paragraph id="cis-deduction-vat">{lines.vatLabel}: {lines.vatValue}</Paragraph>}
    <Paragraph id="cis-deduction-cis">Less {lines.cisLabel}: {lines.cisDeduction}</Paragraph>
    <Paragraph id="cis-deduction-payable">Amount payable: {lines.payable}</Paragraph>
  </>
);
