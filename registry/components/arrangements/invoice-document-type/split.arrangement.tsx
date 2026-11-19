import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { InvoiceDocumentTypeLines } from "./lines.ts";

/**
 * What identifies the invoice on the left (number, purchase order, VAT
 * number) and its dates ranged right, the way an invoice head often faces
 * the letterhead.
 */
export const split = (lines: InvoiceDocumentTypeLines): Yield => (
  <>
    {lines.statement && <Paragraph id="invoice-type-statement">{lines.statement}</Paragraph>}
    {/* The dates get what their longest line needs; the identifiers, which
        run long (a PO with a cost centre on it), get the rest. */}
    <Table id="invoice-type-details" columns={[{ width: "auto" }, { width: 70, align: "right" }]}>
      <Row id="invoice-type-details-row">
        <Cell>
          <Paragraph id="invoice-type-number">Invoice number: {lines.number}</Paragraph>
          {lines.purchaseOrder && (
            <Paragraph id="invoice-type-purchase-order">Purchase order: {lines.purchaseOrder}</Paragraph>
          )}
          {lines.vatNumber && <Paragraph id="invoice-type-vat-number">VAT number: {lines.vatNumber}</Paragraph>}
        </Cell>
        <Cell>
          <Paragraph id="invoice-type-issue-date">Date of issue: {lines.issueDate}</Paragraph>
          {lines.supplyLine && <Paragraph id="invoice-type-supply">{lines.supplyLine}</Paragraph>}
        </Cell>
      </Row>
    </Table>
  </>
);
