import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { ExemptionCertificateCheckLines } from "./lines.ts";

/**
 * One row per line with its amount, status and reason, so the statuses line up
 * down the page. The default.
 */
export const table = (lines: ExemptionCertificateCheckLines): Yield => (
  <>
    <Paragraph id="exemption-check-summary">{lines.summary}</Paragraph>
    <Table
      id="exemption-check-lines"
      columns={[{ width: "auto" }, { width: 30, align: "right" }, { width: 20 }, { width: 70 }]}
    >
      <Row id="exemption-check-lines-header" header>
        <Cell>Line</Cell>
        <Cell>Amount</Cell>
        <Cell>Status</Cell>
        <Cell>Reason</Cell>
      </Row>
      {lines.rows[0].description && (
        <Row id="exemption-check-line-1">
          <Cell>{lines.rows[0].description}</Cell>
          <Cell>{lines.rows[0].amount}</Cell>
          <Cell>{lines.rows[0].status}</Cell>
          <Cell>{lines.rows[0].reason}</Cell>
        </Row>
      )}
      {lines.rows[1].description && (
        <Row id="exemption-check-line-2">
          <Cell>{lines.rows[1].description}</Cell>
          <Cell>{lines.rows[1].amount}</Cell>
          <Cell>{lines.rows[1].status}</Cell>
          <Cell>{lines.rows[1].reason}</Cell>
        </Row>
      )}
      {lines.rows[2].description && (
        <Row id="exemption-check-line-3">
          <Cell>{lines.rows[2].description}</Cell>
          <Cell>{lines.rows[2].amount}</Cell>
          <Cell>{lines.rows[2].status}</Cell>
          <Cell>{lines.rows[2].reason}</Cell>
        </Row>
      )}
      {lines.rows[3].description && (
        <Row id="exemption-check-line-4">
          <Cell>{lines.rows[3].description}</Cell>
          <Cell>{lines.rows[3].amount}</Cell>
          <Cell>{lines.rows[3].status}</Cell>
          <Cell>{lines.rows[3].reason}</Cell>
        </Row>
      )}
      {lines.rows[4].description && (
        <Row id="exemption-check-line-5">
          <Cell>{lines.rows[4].description}</Cell>
          <Cell>{lines.rows[4].amount}</Cell>
          <Cell>{lines.rows[4].status}</Cell>
          <Cell>{lines.rows[4].reason}</Cell>
        </Row>
      )}
      {lines.rows[5].description && (
        <Row id="exemption-check-line-6">
          <Cell>{lines.rows[5].description}</Cell>
          <Cell>{lines.rows[5].amount}</Cell>
          <Cell>{lines.rows[5].status}</Cell>
          <Cell>{lines.rows[5].reason}</Cell>
        </Row>
      )}
      {lines.rows[6].description && (
        <Row id="exemption-check-line-7">
          <Cell>{lines.rows[6].description}</Cell>
          <Cell>{lines.rows[6].amount}</Cell>
          <Cell>{lines.rows[6].status}</Cell>
          <Cell>{lines.rows[6].reason}</Cell>
        </Row>
      )}
      {lines.rows[7].description && (
        <Row id="exemption-check-line-8">
          <Cell>{lines.rows[7].description}</Cell>
          <Cell>{lines.rows[7].amount}</Cell>
          <Cell>{lines.rows[7].status}</Cell>
          <Cell>{lines.rows[7].reason}</Cell>
        </Row>
      )}
      {lines.rows[8].description && (
        <Row id="exemption-check-line-9">
          <Cell>{lines.rows[8].description}</Cell>
          <Cell>{lines.rows[8].amount}</Cell>
          <Cell>{lines.rows[8].status}</Cell>
          <Cell>{lines.rows[8].reason}</Cell>
        </Row>
      )}
      {lines.rows[9].description && (
        <Row id="exemption-check-line-10">
          <Cell>{lines.rows[9].description}</Cell>
          <Cell>{lines.rows[9].amount}</Cell>
          <Cell>{lines.rows[9].status}</Cell>
          <Cell>{lines.rows[9].reason}</Cell>
        </Row>
      )}
      {lines.rows[10].description && (
        <Row id="exemption-check-line-11">
          <Cell>{lines.rows[10].description}</Cell>
          <Cell>{lines.rows[10].amount}</Cell>
          <Cell>{lines.rows[10].status}</Cell>
          <Cell>{lines.rows[10].reason}</Cell>
        </Row>
      )}
      {lines.rows[11].description && (
        <Row id="exemption-check-line-12">
          <Cell>{lines.rows[11].description}</Cell>
          <Cell>{lines.rows[11].amount}</Cell>
          <Cell>{lines.rows[11].status}</Cell>
          <Cell>{lines.rows[11].reason}</Cell>
        </Row>
      )}
      {lines.rows[12].description && (
        <Row id="exemption-check-line-13">
          <Cell>{lines.rows[12].description}</Cell>
          <Cell>{lines.rows[12].amount}</Cell>
          <Cell>{lines.rows[12].status}</Cell>
          <Cell>{lines.rows[12].reason}</Cell>
        </Row>
      )}
      {lines.rows[13].description && (
        <Row id="exemption-check-line-14">
          <Cell>{lines.rows[13].description}</Cell>
          <Cell>{lines.rows[13].amount}</Cell>
          <Cell>{lines.rows[13].status}</Cell>
          <Cell>{lines.rows[13].reason}</Cell>
        </Row>
      )}
      {lines.rows[14].description && (
        <Row id="exemption-check-line-15">
          <Cell>{lines.rows[14].description}</Cell>
          <Cell>{lines.rows[14].amount}</Cell>
          <Cell>{lines.rows[14].status}</Cell>
          <Cell>{lines.rows[14].reason}</Cell>
        </Row>
      )}
      {lines.rows[15].description && (
        <Row id="exemption-check-line-16">
          <Cell>{lines.rows[15].description}</Cell>
          <Cell>{lines.rows[15].amount}</Cell>
          <Cell>{lines.rows[15].status}</Cell>
          <Cell>{lines.rows[15].reason}</Cell>
        </Row>
      )}
    </Table>
    <Paragraph id="exemption-check-basis">{lines.basis}</Paragraph>
  </>
);
