import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { InvoiceDocumentTypeLines } from "./lines.ts";

/**
 * Labels in a narrow column with the values beside them, so the values line
 * up down the page and a reader scanning for the number finds it in one place.
 */
export const table = (lines: InvoiceDocumentTypeLines): Yield => (
  <>
    {lines.statement && <Paragraph id="invoice-type-statement">{lines.statement}</Paragraph>}
    {/* 45mm holds the longest label, "Leistungszeitraum", on one line. */}
    <Table id="invoice-type-details" columns={[{ width: 45 }, { width: "auto" }]}>
      <Row id="invoice-type-number">
        <Cell>Invoice number</Cell>
        <Cell>{lines.number}</Cell>
      </Row>
      <Row id="invoice-type-issue-date">
        <Cell>Date of issue</Cell>
        <Cell>{lines.issueDate}</Cell>
      </Row>
      {lines.supplyLabel && (
        <Row id="invoice-type-supply">
          <Cell>{lines.supplyLabel}</Cell>
          <Cell>{lines.supplyValue}</Cell>
        </Row>
      )}
      {lines.purchaseOrder && (
        <Row id="invoice-type-purchase-order">
          <Cell>Purchase order</Cell>
          <Cell>{lines.purchaseOrder}</Cell>
        </Row>
      )}
      {lines.vatNumber && (
        <Row id="invoice-type-vat-number">
          <Cell>VAT number</Cell>
          <Cell>{lines.vatNumber}</Cell>
        </Row>
      )}
    </Table>
  </>
);
