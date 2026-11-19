import { Cell, Graph, type Nodes, Paragraph, Row, Table, useState } from "docxcelerate/template";
import { pounds } from "../format.ts";
import { Gap, KeyFigures, Opening, Rubric } from "./parts.tsx";
import type { PortfolioReviewData } from "../types.ts";

/**
 * Page 12: the regulated costs disclosure.
 *
 * Every figure is data and every sentence is fixed wording: what a client is
 * charged is not a place for a model to choose its words.
 */
export const Costs: Nodes = () => {
  const [costs] = useState((data: PortfolioReviewData) => ({
    rows: data.costs,
    year: data.yearCosts,
    effect: data.costEffect,
    value: data.closingValue,
  }));

  const total = costs.rows.reduce((sum, row) => sum + row.amount, 0);
  const totalPercent = costs.rows.reduce((sum, row) => sum + row.percent, 0);
  const yearTotal = costs.year.reduce((sum, row) => sum + row.amount, 0);
  const before = costs.effect.before[costs.effect.before.length - 1]!;
  const after = costs.effect.after[costs.effect.after.length - 1]!;
  const share = (amount: number) => `${((amount / costs.value) * 100).toFixed(2)}%`;

  return (
    <>
      <Opening id="costs" number="9" title="Costs and charges" standfirst="What you paid, in pounds and as a share of your portfolio." />
      <Rubric id="quarter-costs" label="This quarter" variant="rubricTight" />
      <Table
        id="costs-table"
        variant="transactions"
        columns={[{ width: "auto" }, { width: 58.12 }, { width: 20.12, align: "right" }, { width: 16, align: "right" }]}
      >
        <Row variant="transactionsHead">
          <Cell>Charge</Cell>
          <Cell>How it is taken</Cell>
          <Cell>£</Cell>
          <Cell>%</Cell>
        </Row>
        {costs.rows.map((row) => (
          <Row variant="costRow">
            <Cell>{row.charge}</Cell>
            <Cell variant="returnMuted">{row.how}</Cell>
            <Cell>{pounds(row.amount)}</Cell>
            <Cell>{`${row.percent.toFixed(2)}%`}</Cell>
          </Row>
        ))}
        <Row variant="costTotal">
          <Cell>Total</Cell>
          <Cell></Cell>
          <Cell>{pounds(total)}</Cell>
          <Cell>{`${totalPercent.toFixed(2)}%`}</Cell>
        </Row>
      </Table>

      <Rubric id="year-costs" label="Last twelve months" />
      <KeyFigures
        id="year-figures"
        columns={5}
        gutterPt={12}
        cellVariant="costFigure"
        labelVariant="costLabel"
        valueVariant="costValue"
        noteVariant="costNote"
        figures={[
          ...costs.year.map((row) => ({ label: row.label, value: pounds(row.amount), note: share(row.amount) })),
          { label: "Total", value: pounds(yearTotal), note: `${share(yearTotal)} of value`, cell: "costFigureTotal" },
        ]}
      />

      <Rubric id="cost-effect" label="The effect of costs over ten years" note="╌ Before costs   ━ After costs" />
      <Table id="cost-effect-grid" variant="leadersGrid" columns={[{ width: 114.9 }, { width: 5.64 }, { width: "auto" }]}>
        <Row>
          <Cell variant="centred">
            <Graph
              id="cost-effect-chart"
              alt="The effect of costs over ten years: before and after costs"
              graphType="line"
              legend="none"
              markers={false}
              textSizePt={7}
              // The plot rectangle the canvas draws in, measured off its SVG.
              plotArea={{ leftPt: 35.8, topPt: 8.4, rightPt: 8.1, bottomPt: 16.3 }}
              fromEdge
              width={325.7}
              height={154.7}
              valueAxisMin={400}
              valueAxisMax={700}
              valueAxisStep={100}
              numberFormat={'"£"0"k"'}
              data={{
                categories: costs.effect.years.map((year, i) => (i % 2 === 0 ? year : "")),
                series: [
                  { label: "What costs take", values: costs.effect.before, band: costs.effect.after, color: "DBEEE5" },
                  { label: "Before costs", values: costs.effect.before, color: "5C6B64", dash: true, widthPt: 1 },
                  { label: "After costs", values: costs.effect.after, color: "17624A", widthPt: 1.8 },
                ],
              }}
            />
          </Cell>
          <Cell></Cell>
          <Cell variant="centred">
            <Paragraph variant="effectFigure">{`about ${pounds(Math.round((before - after)))},000`}</Paragraph>
            <Paragraph variant="effectText">
              {`Over ten years, at today's charges, costs would reduce the value of your portfolio from ${pounds(before)},000 to ${pounds(after)},000.`}
            </Paragraph>
            <Paragraph variant="effectNote">
              An illustration, not a forecast. It assumes growth of 5% a year before costs, and today's charges throughout.
            </Paragraph>
          </Cell>
        </Row>
      </Table>
      <Gap pt={2} />
    </>
  );
};
