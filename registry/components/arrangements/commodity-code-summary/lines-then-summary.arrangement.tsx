import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { CommodityCodeSummaryLines } from "./lines.ts";

/**
 * Every goods line, then the summary under it, the way the invoice reads
 * top to bottom and the declaration is keyed last.
 */
export const linesThenSummary = (lines: CommodityCodeSummaryLines): Yield => (
  <>
    <Paragraph id="commodity-code-summary-basis">{lines.basis}</Paragraph>
    {lines.attention && <Paragraph id="commodity-code-summary-attention">{lines.attention}</Paragraph>}
    {/* Item auto; 24mm code (10 digits), 14 origin, 12 pref., 20 quantity, 22 supp. unit, 26 mass ("12,000.000"), 34 value ("£1,485,000.00"). */}
    <Table
      id="commodity-code-summary-lines"
      columns={[
        { width: "auto" },
        { width: 24 },
        { width: 14 },
        { width: 12 },
        { width: 20, align: "right" },
        { width: 22, align: "right" },
        { width: 26, align: "right" },
        { width: 34, align: "right" },
      ]}
    >
      <Row id="commodity-code-summary-lines-header" header>
        <Cell>Item</Cell>
        <Cell>Tariff code</Cell>
        <Cell>Origin</Cell>
        <Cell>Pref.</Cell>
        <Cell>Quantity</Cell>
        <Cell>Supp. unit</Cell>
        <Cell>Net mass (kg)</Cell>
        <Cell>Value</Cell>
      </Row>
      {lines.lines.map((row) => (
        <Row id="commodity-code-summary-line">
          <Cell>
            <Paragraph>{row.item}</Paragraph>
            {row.warning && <Paragraph id="commodity-code-summary-line-warning">{row.warning}</Paragraph>}
          </Cell>
          <Cell>{row.tariffCode}</Cell>
          <Cell>{row.origin}</Cell>
          <Cell>{row.preference}</Cell>
          <Cell>{row.quantity}</Cell>
          <Cell>{row.supplementary}</Cell>
          <Cell>{row.netMass}</Cell>
          <Cell>{row.value}</Cell>
        </Row>
      ))}
    </Table>
    {/* Code auto (it carries the warning); 14 origin, 12 pref., 20 lines, 22 supp. unit, 26 mass, 34 value. */}
    <Table
      id="commodity-code-summary-groups"
      columns={[
        { width: "auto" },
        { width: 14 },
        { width: 12 },
        { width: 20, align: "right" },
        { width: 22, align: "right" },
        { width: 26, align: "right" },
        { width: 34, align: "right" },
      ]}
    >
      <Row id="commodity-code-summary-groups-header" header>
        <Cell>Tariff code</Cell>
        <Cell>Origin</Cell>
        <Cell>Pref.</Cell>
        <Cell>Lines</Cell>
        <Cell>Supp. unit</Cell>
        <Cell>Net mass (kg)</Cell>
        <Cell>Value</Cell>
      </Row>
      {lines.groups.map((group) => (
        <Row id="commodity-code-summary-group">
          <Cell>
            <Paragraph>{group.tariffCode}</Paragraph>
            {group.warning && <Paragraph id="commodity-code-summary-group-warning">{group.warning}</Paragraph>}
          </Cell>
          <Cell>{group.origin}</Cell>
          <Cell>{group.preference}</Cell>
          <Cell>{group.lineCount}</Cell>
          <Cell>{group.supplementary}</Cell>
          <Cell>{group.netMass}</Cell>
          <Cell>{group.value}</Cell>
        </Row>
      ))}
      <Row id="commodity-code-summary-total">
        <Cell>Total</Cell>
        <Cell></Cell>
        <Cell></Cell>
        <Cell></Cell>
        <Cell></Cell>
        <Cell>{lines.totalNetMass}</Cell>
        <Cell>{lines.totalValue}</Cell>
      </Row>
    </Table>
    {lines.grossMass && (
      <Paragraph id="commodity-code-summary-gross">Declared gross mass: {lines.grossMass} kg.</Paragraph>
    )}
    {lines.grossWarning && <Paragraph id="commodity-code-summary-gross-warning">{lines.grossWarning}</Paragraph>}
  </>
);
