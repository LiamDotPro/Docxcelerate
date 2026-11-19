import { Cell, Row, Table, Text, type Yield } from "docxcelerate/template";
import type { NetPayableWithholdingsLines } from "./lines.ts";

/**
 * Labels in an auto-width column with the figures right-aligned in a 40 mm
 * column, so the bridge adds up down the page. The default.
 */
export const table = (lines: NetPayableWithholdingsLines): Yield => (
  <Table id="net-payable-lines" columns={[{ width: "auto" }, { width: 40, align: "right" }]}>
    <Row id="net-payable-base">
      <Cell>{lines.baseLabel}</Cell>
      <Cell>{lines.base}</Cell>
    </Row>
    {lines.rivalsaLabel && (
    <Row id="net-payable-rivalsa">
      <Cell>{lines.rivalsaLabel}</Cell>
      <Cell>{lines.rivalsa}</Cell>
    </Row>
    )}
    {lines.cassaLabel && (
    <Row id="net-payable-cassa">
      <Cell>{lines.cassaLabel}</Cell>
      <Cell>{lines.cassa}</Cell>
    </Row>
    )}
    {lines.vatLabel && (
    <Row id="net-payable-vat">
      <Cell>{lines.vatLabel}</Cell>
      <Cell>{lines.vat}</Cell>
    </Row>
    )}
    {lines.reLabel && (
    <Row id="net-payable-re">
      <Cell>{lines.reLabel}</Cell>
      <Cell>{lines.re}</Cell>
    </Row>
    )}
    {lines.exemptLabel && (
    <Row id="net-payable-exempt">
      <Cell>{lines.exemptLabel}</Cell>
      <Cell>{lines.exempt}</Cell>
    </Row>
    )}
    {lines.stampLabel && (
    <Row id="net-payable-stamp">
      <Cell>{lines.stampLabel}</Cell>
      <Cell>{lines.stamp}</Cell>
    </Row>
    )}
    <Row id="net-payable-total">
      <Cell>{lines.totalLabel}</Cell>
      <Cell>{lines.total}</Cell>
    </Row>
    {lines.withholdingLabel && (
    <Row id="net-payable-withholding">
      <Cell>{lines.withholdingLabel}</Cell>
      <Cell>{lines.withholding}</Cell>
    </Row>
    )}
    {lines.splitPaymentLabel && (
    <Row id="net-payable-split-payment">
      <Cell>{lines.splitPaymentLabel}</Cell>
      <Cell>{lines.splitPayment}</Cell>
    </Row>
    )}
    <Row id="net-payable-net">
      <Cell>{lines.netLabel}</Cell>
      <Cell><Text bold>{lines.net}</Text></Cell>
    </Row>
  </Table>
);
