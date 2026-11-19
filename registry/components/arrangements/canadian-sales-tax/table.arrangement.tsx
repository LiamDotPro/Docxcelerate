import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { CanadianSalesTaxLines } from "./lines.ts";

/**
 * Labels on the left, amounts right-aligned beside them. The default: this is
 * a totals block that sits under line items, and the amounts have to line up.
 */
export const table = (lines: CanadianSalesTaxLines): Yield => (
  <>
    <Table id="canadian-sales-tax-summary" columns={[{}, { align: "right" }]}>
      <Row id="canadian-sales-tax-subtotal">
        <Cell>Subtotal</Cell>
        <Cell>{lines.subtotal}</Cell>
      </Row>
      {lines.federalLabel && (
        <Row id="canadian-sales-tax-federal">
          <Cell>{lines.federalLabel}</Cell>
          <Cell>{lines.federalAmount}</Cell>
        </Row>
      )}
      {lines.provincialLabel && (
        <Row id="canadian-sales-tax-provincial">
          <Cell>{lines.provincialLabel}</Cell>
          <Cell>{lines.provincialAmount}</Cell>
        </Row>
      )}
      <Row id="canadian-sales-tax-total">
        <Cell>Total including tax</Cell>
        <Cell>{lines.total}</Cell>
      </Row>
      {lines.bn && (
        <Row id="canadian-sales-tax-bn">
          <Cell>GST/HST registration number</Cell>
          <Cell>{lines.bn}</Cell>
        </Row>
      )}
      {lines.qst && (
        <Row id="canadian-sales-tax-qst">
          <Cell>QST registration number</Cell>
          <Cell>{lines.qst}</Cell>
        </Row>
      )}
    </Table>
    {lines.smallSupplierNote && <Paragraph id="canadian-sales-tax-small-supplier">{lines.smallSupplierNote}</Paragraph>}
    {lines.statusNote && <Paragraph id="canadian-sales-tax-status">{lines.statusNote}</Paragraph>}
    {lines.warning && <Paragraph id="canadian-sales-tax-warning">{lines.warning}</Paragraph>}
  </>
);
