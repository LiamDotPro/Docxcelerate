import { Cell, type Nodes, Paragraph, Row, Table, useState } from "docxcelerate/template";
import { arrowed, pounds } from "../format.ts";
import { Gap, KeyFigures, Opening, Rubric } from "./parts.tsx";
import { Written } from "./written.node.tsx";
import type { Holding, PortfolioReviewData } from "../types.ts";

/** The note's node id for each holding, as the canvas names them. */
const NOTE_IDS: Record<string, string> = {
  "Global Equity Index Fund": "global-equity-note",
  "UK Equity Income Fund": "uk-income-note",
  "Emerging Markets Fund": "emerging-note",
  "Smaller Companies Fund": "smaller-cos-note",
  "Global Bond Fund (hedged)": "global-bond-note",
  "UK Gilt 4.25% 2031": "gilt-2031-note",
  "Global Property Securities": "property-note",
  "Cash": "cash-note",
};

const COLUMNS = [
  { width: 44 },
  { width: 12, align: "right" as const },
  { width: 18, align: "right" as const },
  { width: 14, align: "right" as const },
  { width: 13, align: "right" as const },
  { width: "auto" as const },
];

/** One holding's row: name and kind, figures, and why it moved. */
const HoldingRow: Nodes<{ holding: Holding; shaded: boolean }> = ({ holding, shaded }) => (
  <Row variant={shaded ? "holdingAlt" : "holding"}>
    <Cell variant="holdingName">
      <Paragraph variant="holdingTitle">{holding.name}</Paragraph>
      <Paragraph variant="holdingKind">{holding.kind}</Paragraph>
    </Cell>
    <Cell variant="holdingFigure">{`${holding.weight.toFixed(1)}%`}</Cell>
    <Cell variant="holdingFigure">{pounds(holding.value)}</Cell>
    <Cell variant={holding.quarter < 0 ? "holdingDown" : "holdingUp"}>{arrowed(holding.quarter)}</Cell>
    <Cell variant={holding.year < 0 ? "holdingYearDown" : "holdingYearUp"}>{arrowed(holding.year)}</Cell>
    <Cell variant="holdingNote">
      <Written
        id={NOTE_IDS[holding.name] ?? `${holding.name}-note`}
        variant="holdingWhy"
        ask="One sentence on why this holding moved this quarter, from its return and its manager's note."
        placeholder={`Why ${holding.name} moved this quarter.`}
        from={{ fund: holding.name, quarter: holding.quarter, note: holding.managerNote }}
      />
    </Cell>
  </Row>
);

/** Page 7: everything held, and the accounts it is held in. */
export const Holdings: Nodes = () => {
  const [held] = useState((data: PortfolioReviewData) => ({
    holdings: data.holdings,
    total: data.closingValue,
    quarter: data.quarterPercent,
    year: data.portfolioIndex[data.portfolioIndex.length - 1]! - 100,
    accounts: data.accounts,
    valuedAt: data.period.valuedAt,
  }));

  return (
    <>
      <Opening id="holdings" number="5" title="Holdings" standfirst={`Everything you own at ${held.valuedAt}, and why each part moved.`} />
      <Gap pt={14} />
      <Table id="holdings-table" variant="holdings" columns={COLUMNS}>
        <Row variant="holdingsHead">
          <Cell variant="holdingName">Holding</Cell>
          <Cell>Weight</Cell>
          <Cell>Value</Cell>
          <Cell>Quarter</Cell>
          <Cell>1 year</Cell>
          <Cell variant="holdingNote">Why it moved</Cell>
        </Row>
        {held.holdings.map((holding, index) => <HoldingRow holding={holding} shaded={index % 2 === 1} />)}
        <Row variant="holdingsTotal">
          <Cell variant="holdingName">Total</Cell>
          <Cell>100.0%</Cell>
          <Cell>{pounds(held.total)}</Cell>
          <Cell variant="holdingYearUp">{arrowed(held.quarter)}</Cell>
          <Cell variant="holdingYearUp">{arrowed(held.year)}</Cell>
          <Cell></Cell>
        </Row>
      </Table>
      <Paragraph id="holdings-note" variant="tableNote">
        Prices at close of business on 30 September 2026. Returns are total returns, net of fund charges.
      </Paragraph>

      <Rubric id="accounts" label="Account types" />
      <KeyFigures
        id="account-figures"
        columns={3}
        gutterPt={14}
        valueVariant="keyValueSmall"
        figures={held.accounts.map((account) => ({ label: account.name, value: pounds(account.value), note: account.note }))}
      />
    </>
  );
};
