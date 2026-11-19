import { Cell, Graph, type Nodes, Paragraph, Row, Table, useState } from "docxcelerate/template";
import { Opening, Rubric, Uprights } from "./parts.tsx";
import { Written } from "./written.node.tsx";
import type { PortfolioReviewData } from "../types.ts";

/** Six years 10pt apart across the column; in each, the two bars at 22% and 52% of its width. */
const GROUP = (178 - 5 * 3.53) / 6;
const calendarColumns = Array.from({ length: 11 }, (_, i) => ({ width: i % 2 === 0 ? GROUP : 3.53 }));
/** The canvas draws a percentage point 8px tall. */
const PT_PER_POINT = 6;
const pair = (portfolio: number, benchmark: number) => [
  { width: GROUP * 0.22 },
  { width: GROUP * 0.26, bar: { heightPt: portfolio * PT_PER_POINT, variant: "barGreen" } },
  { width: GROUP * 0.04 },
  { width: GROUP * 0.26, bar: { heightPt: benchmark * PT_PER_POINT, variant: "barRule" } },
  { width: GROUP * 0.22 },
];

/** Page 4: twelve months against the benchmark, the returns table, and the calendar years. */
export const Performance: Nodes = () => {
  const [perf] = useState((data: PortfolioReviewData) => ({
    months: data.months,
    portfolio: data.portfolioIndex,
    benchmark: data.benchmarkIndex,
    columns: data.returnColumns,
    returns: data.returns,
    years: data.calendarYears,
  }));

  const tone = ["returnStrong", "returnMuted", "returnGreen"];

  return (
    <>
      <Opening id="performance" number="2" title="Performance" standfirst="How your portfolio has done against its benchmark." />

      <Rubric id="twelve-months" label="Twelve months, rebased to 100" note="━ Your portfolio   ╌ Benchmark" variant="rubricTight" />
      <Graph
        id="performance-chart"
        alt="Twelve months, rebased to 100: your portfolio against its benchmark"
        graphType="line"
        legend="none"
        markers={false}
        textSizePt={7}
        width={504.6}
        height={171.9}
        // The plot rectangle the canvas draws in, measured off its SVG.
        plotArea={{ leftPt: 41, topPt: 15.8, rightPt: 53.6, bottomPt: 17.3 }}
        fromEdge
        valueAxisMin={98}
        valueAxisMax={110}
        valueAxisStep={4}
        data={{
          categories: perf.months,
          series: [
            // The ground under the portfolio's line, in the canvas's 8% green.
            { label: "Your portfolio", values: perf.portfolio, band: perf.portfolio.map(() => 98), color: "ECF2F0" },
            { label: perf.benchmark[perf.benchmark.length - 1]!.toFixed(1), values: perf.benchmark, color: "5C6B64", dash: true, widthPt: 1, endLabel: true },
            // Named for the figure it ends on, which is what the canvas prints there.
            { label: perf.portfolio[perf.portfolio.length - 1]!.toFixed(1), values: perf.portfolio, color: "17624A", widthPt: 1.8, endLabel: true },
          ],
        }}
      />
      <Written
        id="performance-caption"
        variant="caption"
        ask="One or two sentences comparing the year's return with the benchmark's, from the same series the chart plots."
        placeholder="A sentence comparing the year with the benchmark."
        from={{ portfolio: perf.portfolio, benchmark: perf.benchmark }}
      />

      <Rubric id="returns" label="Returns" />
      <Table
        id="returns-table"
        variant="returns"
        columns={[{ width: 34 }, ...perf.columns.map(() => ({ width: 28.8, align: "right" as const }))]}
      >
        <Row variant="returnsHead">
          <Cell></Cell>
          {perf.columns.map((column) => <Cell>{column}</Cell>)}
        </Row>
        {perf.returns.map((row, index) => (
          <Row variant={tone[index]}>
            <Cell>{row.label}</Cell>
            {row.values.map((value) => <Cell>{value}</Cell>)}
          </Row>
        ))}
      </Table>
      <Paragraph id="returns-note" variant="tableNote">
        Net of fund charges. The three-year figure is annualised; the portfolio was opened in March 2021.
      </Paragraph>

      <Rubric id="calendar" label="Calendar years" note="■ Your portfolio   ■ Benchmark" />
      <Table id="calendar-chart" variant="calendar" columns={calendarColumns}>
        <Row height={63.75} heightRule="exact">
          {perf.years.map((year, i) => (
            <>
              {i > 0 && <Cell></Cell>}
              <Cell variant="calendarUp">
                <Paragraph variant="calendarFigure">{year.portfolio > 0 ? `+${year.portfolio.toFixed(1)}` : ""}</Paragraph>
                <Uprights id={`calendar-up-${i}`} slots={pair(Math.max(0, year.portfolio), Math.max(0, year.benchmark))} />
              </Cell>
            </>
          ))}
        </Row>
        <Row height={0.75} heightRule="exact">
          {perf.years.map((_, i) => (
            <>
              {i > 0 && <Cell></Cell>}
              <Cell variant="barAxis"></Cell>
            </>
          ))}
        </Row>
        <Row height={51} heightRule="exact">
          {perf.years.map((year, i) => (
            <>
              {i > 0 && <Cell></Cell>}
              <Cell variant="calendarDown">
                <Uprights id={`calendar-down-${i}`} hanging slots={pair(Math.max(0, -year.portfolio), Math.max(0, -year.benchmark))} />
                <Paragraph variant="calendarFigure">{year.portfolio < 0 ? `−${(-year.portfolio).toFixed(1)}` : ""}</Paragraph>
              </Cell>
            </>
          ))}
        </Row>
        <Row>
          {perf.years.map((year, i) => (
            <>
              {i > 0 && <Cell></Cell>}
              <Cell variant="calendarYear">{year.year}</Cell>
            </>
          ))}
        </Row>
      </Table>
      <Paragraph id="calendar-note" variant="tableNote">* 2021 from March, when the portfolio was opened.</Paragraph>
    </>
  );
};
