import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { PatientResponsibilityBreakdownLines } from "./lines.ts";

/**
 * One row per claim line with the amounts in columns, the layout of a real EOB.
 * Each money column is 21mm, which holds "$248,750.00" and the "Coinsurance"
 * header at body size.
 */
export const table = (lines: PatientResponsibilityBreakdownLines): Yield => (
  <Table
    id="patient-responsibility-lines"
    columns={[
      { width: "auto" },
      { width: 21, align: "right" },
      { width: 21, align: "right" },
      { width: 21, align: "right" },
      { width: 21, align: "right" },
      { width: 21, align: "right" },
      { width: 21, align: "right" },
    ]}
  >
    <Row id="patient-responsibility-lines-header" header>
      <Cell>Service</Cell>
      <Cell>Billed</Cell>
      <Cell>Allowed</Cell>
      <Cell>Deductible</Cell>
      <Cell>Copay</Cell>
      <Cell>Coinsurance</Cell>
      <Cell>You owe</Cell>
    </Row>
    {lines.rows.map((row) => (
      <Row id="patient-responsibility-line">
        <Cell>
          <Paragraph id="patient-responsibility-line-service">{row.service}</Paragraph>
          <Paragraph id="patient-responsibility-line-detail">{row.detail}</Paragraph>
          {row.note && <Paragraph id="patient-responsibility-line-note">{row.note}</Paragraph>}
        </Cell>
        <Cell>{row.billed}</Cell>
        <Cell>{row.allowed}</Cell>
        <Cell>{row.deductible}</Cell>
        <Cell>{row.copay}</Cell>
        <Cell>{row.coinsurance}</Cell>
        <Cell>{row.youOwe}</Cell>
      </Row>
    ))}
    <Row id="patient-responsibility-total">
      <Cell>Total</Cell>
      <Cell>{lines.total.billed}</Cell>
      <Cell>{lines.total.allowed}</Cell>
      <Cell>{lines.total.deductible}</Cell>
      <Cell>{lines.total.copay}</Cell>
      <Cell>{lines.total.coinsurance}</Cell>
      <Cell>{lines.total.youOwe}</Cell>
    </Row>
  </Table>
);
