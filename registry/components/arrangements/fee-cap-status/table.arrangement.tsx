import { Cell, Row, Table, type Yield } from "docxcelerate/template";
import type { FeeCapStatusLines } from "./lines.ts";

/**
 * Labels in a column with the amounts right-aligned beside them, so a reader
 * can check cap, billed and charged against each other by eye. The default.
 */
export const table = (lines: FeeCapStatusLines): Yield => (
  // 40mm holds "£2,500,000.00" on one line.
  <Table id="fee-cap-figures" columns={[{ width: "auto" }, { width: 40, align: "right" }]}>
    <Row id="fee-cap-amount">
      <Cell>{lines.capLabel}</Cell>
      <Cell>{lines.capValue}</Cell>
    </Row>
    <Row id="fee-cap-billed">
      <Cell>Billed before this invoice</Cell>
      <Cell>{lines.billedToDate}</Cell>
    </Row>
    <Row id="fee-cap-period">
      <Cell>Fees this period</Cell>
      <Cell>{lines.thisPeriod}</Cell>
    </Row>
    {lines.writtenOff && (
      <Row id="fee-cap-written-off">
        <Cell>{lines.writtenOffLabel}</Cell>
        <Cell>{lines.writtenOff}</Cell>
      </Row>
    )}
    {lines.charged && (
      <Row id="fee-cap-charged">
        <Cell>Charged on this invoice</Cell>
        <Cell>{lines.charged}</Cell>
      </Row>
    )}
    {lines.withheld && (
      <Row id="fee-cap-withheld">
        <Cell>{lines.withheldLabel}</Cell>
        <Cell>{lines.withheld}</Cell>
      </Row>
    )}
    {lines.payable && (
      <Row id="fee-cap-payable">
        <Cell>Payable for these fees</Cell>
        <Cell>{lines.payable}</Cell>
      </Row>
    )}
    {lines.remaining && (
      <Row id="fee-cap-remaining">
        <Cell>{lines.remainingLabel}</Cell>
        <Cell>{lines.remaining}</Cell>
      </Row>
    )}
  </Table>
);
