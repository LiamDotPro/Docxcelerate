import { Cell, Row, Table, type Yield } from "docxcelerate/template";
import type { DirectDebitCollectionNoticeLines } from "./lines.ts";

/**
 * Labels in a narrow column with the mandate details beside them, so the
 * values line up and a payer checking a reference finds it in one place.
 */
export const table = (lines: DirectDebitCollectionNoticeLines): Yield => (
  // 45mm holds the longest label, "Authorization reference", on one line.
  <Table id="dd-notice-details" columns={[{ width: 45 }, { width: "auto" }]}>
    <Row id="dd-notice-creditor">
      <Cell>Collected by</Cell>
      <Cell>{lines.creditorName}</Cell>
    </Row>
    {lines.creditorId && (
      <Row id="dd-notice-creditor-id">
        <Cell>{lines.creditorIdLabel}</Cell>
        <Cell>{lines.creditorId}</Cell>
      </Row>
    )}
    <Row id="dd-notice-mandate">
      <Cell>{lines.mandateLabel}</Cell>
      <Cell>{lines.mandateReference}</Cell>
    </Row>
    {lines.account && (
      <Row id="dd-notice-account">
        <Cell>From account</Cell>
        <Cell>{lines.account}</Cell>
      </Row>
    )}
  </Table>
);
