import { Cell, Row, Table, type Yield } from "docxcelerate/template";
import type { CisInvoiceDeductionLines } from "./lines.ts";

/**
 * Labels in a column with the figures right-aligned beside them, so the sums
 * can be checked down the page. The default.
 */
export const table = (lines: CisInvoiceDeductionLines): Yield => (
  <Table id="cis-deduction-lines" columns={[{ width: "auto" }, { width: 45, align: "right" }]}>
    <Row id="cis-deduction-gross">
      <Cell>Gross amount, excluding VAT</Cell>
      <Cell>{lines.gross}</Cell>
    </Row>
    <Row id="cis-deduction-materials">
      <Cell>Less materials at cost</Cell>
      <Cell>{lines.materials}</Cell>
    </Row>
    {lines.plantHire && (
      <Row id="cis-deduction-plant-hire">
        <Cell>Less plant hire</Cell>
        <Cell>{lines.plantHire}</Cell>
      </Row>
    )}
    {lines.consumables && (
      <Row id="cis-deduction-consumables">
        <Cell>Less consumables</Cell>
        <Cell>{lines.consumables}</Cell>
      </Row>
    )}
    <Row id="cis-deduction-labour">
      <Cell>Labour subject to CIS</Cell>
      <Cell>{lines.labour}</Cell>
    </Row>
    {lines.vatLabel && (
      <Row id="cis-deduction-vat">
        <Cell>{lines.vatLabel}</Cell>
        <Cell>{lines.vatValue}</Cell>
      </Row>
    )}
    <Row id="cis-deduction-cis">
      <Cell>Less {lines.cisLabel}</Cell>
      <Cell>{lines.cisDeduction}</Cell>
    </Row>
    <Row id="cis-deduction-payable">
      <Cell>Amount payable</Cell>
      <Cell>{lines.payable}</Cell>
    </Row>
  </Table>
);
