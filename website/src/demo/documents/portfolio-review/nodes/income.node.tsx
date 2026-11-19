import { Cell, type Nodes, Paragraph, Row, Table, Text, useState } from "docxcelerate/template";
import { pounds } from "../format.ts";
import { Bar, Opening, Rubric, Uprights } from "./parts.tsx";
import { Written } from "./written.node.tsx";
import type { PortfolioReviewData } from "../types.ts";

/** Twelve columns 6pt apart across 178mm. */
const GAP = 2.12;
const COLUMN = (178 - 11 * GAP) / 12;
/** The canvas draws a pound as 0.2302px: £655 stands 150.8px tall. */
const PT_PER_POUND = 0.2302 * 0.75;
/** The plot is 240px deep, less the hairline it stands on. */
const PLOT_PT = 180;

/** Page 10: twelve months of income, where this quarter's came from, and what the cash is for. */
export const Income: Nodes = () => {
  const [income] = useState((data: PortfolioReviewData) => ({
    months: data.income,
    byHolding: data.incomeByHolding,
    cash: data.cash,
    total: data.income.reduce((sum, month) => sum + month.amount, 0),
    quarter: data.income.filter((month) => month.inQuarter).reduce((sum, month) => sum + month.amount, 0),
    value: data.closingValue,
  }));

  const columns = income.months.flatMap((_, i) => (i === 0 ? [{ width: COLUMN }] : [{ width: GAP }, { width: COLUMN }]));
  const cashTotal = income.cash.reduce((sum, line) => sum + line.amount, 0);
  const swatches = ["barDeep", "barMid", "barRule"];
  let at = 0;

  return (
    <>
      <Opening id="income" number="7" title="Income and cash" standfirst="What your investments paid you, and what the cash is for." />
      <Rubric
        id="income-months"
        label="Income received, last twelve months"
        note={`${pounds(income.total)} in total · this quarter in dark green`}
        variant="rubricTight"
      />
      <Table id="income-chart" variant="incomeChart" columns={columns}>
        <Row height={PLOT_PT} heightRule="exact">
          {income.months.map((month, i) => (
            <>
              {i > 0 && <Cell></Cell>}
              <Cell>
                <Paragraph variant={month.inQuarter ? "incomeFigureNow" : "incomeFigure"}>{pounds(month.amount)}</Paragraph>
                <Uprights
                  id={`income-${month.month}`}
                  slots={[{ width: COLUMN, bar: { heightPt: month.amount * PT_PER_POUND, variant: month.inQuarter ? "barGreen" : "barRule" } }]}
                />
              </Cell>
            </>
          ))}
        </Row>
      </Table>
      <Table id="income-months-axis" variant="incomeAxis" columns={columns}>
        <Row>
          {income.months.map((month, i) => (
            <>
              {i > 0 && <Cell></Cell>}
              <Cell>{month.month}</Cell>
            </>
          ))}
        </Row>
      </Table>
      <Written
        id="income-summary"
        variant="caption"
        ask="Two sentences on the year's income: the total, its share of the portfolio, and what drove this quarter's."
        placeholder="A sentence on the year's income."
        from={{ months: income.months, quarter: income.quarter, value: income.value }}
      />

      <Table id="income-tables" variant="leadersGrid" columns={[{ width: 82.65 }, { width: 8.47 }, { width: "auto" }]}>
        <Row>
          <Cell>
            <Paragraph variant="rubricInCell">Income by holding this quarter</Paragraph>
            <Table id="income-by-holding" variant="leaders" columns={[{ width: "auto" }, { width: 18, align: "right" }]}>
              {income.byHolding.map((line) => (
                <Row>
                  <Cell>{line.label}</Cell>
                  <Cell>{pounds(line.amount)}</Cell>
                </Row>
              ))}
              <Row variant="leadersTotal">
                <Cell>Total, paid to your account on 30 September</Cell>
                <Cell>{pounds(income.quarter)}</Cell>
              </Row>
            </Table>
          </Cell>
          <Cell></Cell>
          <Cell>
            <Paragraph variant="rubricInCell">Your cash</Paragraph>
            <Bar
              id="cash-bar"
              width={86.88}
              heightPt={10}
              track="barRule"
              marks={income.cash.map((line, i) => {
                const from = at;
                at += (line.amount / cashTotal) * 86.88;
                return { from, to: at, variant: swatches[i]! };
              })}
            />
            <Paragraph variant="gap6" />
            <Table id="cash-lines" variant="leaders" columns={[{ width: 6.6 }, { width: "auto" }, { width: 18, align: "right" }]}>
              {income.cash.map((line, i) => (
                <Row>
                  <Cell>
                    <Text variant={`swatch${i}`}>■</Text>
                  </Cell>
                  <Cell>{line.label}</Cell>
                  <Cell>{pounds(line.amount)}</Cell>
                </Row>
              ))}
              <Row variant="leadersTotal">
                <Cell></Cell>
                <Cell>Cash at 30 September</Cell>
                <Cell>{pounds(cashTotal)}</Cell>
              </Row>
            </Table>
          </Cell>
        </Row>
      </Table>
    </>
  );
};
