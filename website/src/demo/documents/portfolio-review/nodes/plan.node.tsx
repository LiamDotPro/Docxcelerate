import { Cell, Graph, type Nodes, Paragraph, Row, Table, Text, useState } from "docxcelerate/template";
import { pounds } from "../format.ts";
import { BarList, Gap, Opening, Rubric } from "./parts.tsx";
import { StampWords } from "./stamp.node.tsx";
import { Written } from "./written.node.tsx";
import type { PortfolioReviewData } from "../types.ts";

/** 178mm less the 34mm name, the 56mm figure and two 8pt gaps. */
const ALLOWANCE_TRACK = 82.36;

/** Page 13: the projections against the goal, the decision on them, allowances, and the small print. */
export const Plan: Nodes = () => {
  const [plan] = useState((data: PortfolioReviewData) => ({
    projection: data.projection,
    allowances: data.allowances,
    disclosures: data.disclosures,
    glossary: data.glossary,
    notes: data.clientNotes,
  }));

  const p = plan.projection;
  const end = (series: number[]) => series[series.length - 1]!;

  return (
    <>
      <Opening id="plan" number="10" title="Your plan" standfirst="Whether you remain on track for the retirement you have planned." />
      <Rubric id="projection" label="Projected value to 2036, in today's money" variant="rubricTight" />
      <Graph
        id="projection-chart"
        alt="Projected value to 2036, in today's money: weak, central and strong projections against the goal"
        graphType="line"
        legend="none"
        markers={false}
        textSizePt={9}
        // The plot rectangle the canvas draws in, measured off its SVG.
        plotArea={{ leftPt: 51.7, topPt: 12.1, rightPt: 85.5, bottomPt: 24.6 }}
        fromEdge
        width={504.6}
        height={246.1}
        valueAxisMin={400}
        valueAxisMax={1000}
        valueAxisStep={200}
        numberFormat={'[>=1000]"£1m";"£"0"k"'}
        data={{
          // Every other year named, as the canvas names them.
          categories: p.years.map((year, i) => (i % 2 === 0 ? year : "")),
          series: [
            { label: `Weak £${end(p.weak)}k to strong £${end(p.strong)}k`, values: p.strong, band: p.weak, color: "DBEEE5" },
            // The band's two edges, named where they end and not drawn.
            { label: `Strong £${end(p.strong)}k`, values: p.strong, color: "5C6B64", widthPt: 0, endLabel: true },
            { label: `Weak £${end(p.weak)}k`, values: p.weak, color: "5C6B64", widthPt: 0, endLabel: true },
            { label: `Goal £${p.goal}k`, values: p.years.map(() => p.goal), color: "16211C", dash: true, widthPt: 1.2, endLabel: true },
            { label: `Central £${end(p.central)}k`, values: p.central, color: "17624A", widthPt: 2.8, endLabel: true },
          ],
        }}
      />

      <Gap pt={10} />
      <Table id="plan-decision-box" variant="leadersGrid" columns={[{ width: 40 }, { width: 4.94 }, { width: "auto" }]}>
        <Row>
          <Cell variant="stamp">
            <StampWords
              id="plan-decision"
              ask="Whether the client is on track for their goal, as On track or Review."
              placeholder="On track or review"
              noteAsk="A few words under the decision naming the goal it was decided against."
              notePlaceholder="decided per client"
              from={{ projection: plan.projection, notes: plan.notes }}
            />
          </Cell>
          <Cell></Cell>
          <Cell>
            <Written
              id="plan-check"
              variant="body"
              ask="Two sentences on whether the projections still meet the client's goal, and why."
              placeholder="Whether the projections still meet the client's goal, and why."
              from={{ projection: plan.projection, notes: plan.notes }}
            />
          </Cell>
        </Row>
      </Table>

      <Rubric id="allowances" label="Allowances 2026/27" />
      <BarList
        id="allowance-bars"
        labelWidth={34}
        valueWidth={56}
        heightPt={9}
        gapPt={6}
        lines={plan.allowances.map((allowance) => ({
          label: allowance.name,
          labelVariant: "barLabelStrong",
          valueVariant: "barValuePlain",
          value: allowance.name === "ISA"
            ? `${pounds(allowance.used)} of ${pounds(allowance.limit)} used · ${pounds(allowance.limit - allowance.used)} left`
            : `${pounds(allowance.used)} of ${pounds(allowance.limit)} used`,
          marks: [{ from: 0, to: (allowance.used / allowance.limit) * ALLOWANCE_TRACK, variant: "barGreen" }],
        }))}
      />

      <Rubric id="important" label="11  Important information" />
      <Table id="disclosures" variant="disclosures" columns={[{ width: 56.04 }, { width: 4.94 }, { width: 56.04 }, { width: 4.94 }, { width: 56.04 }]}>
        <Row>
          {plan.disclosures.map((text, index) => (
            <>
              {index > 0 && <Cell></Cell>}
              <Cell>{text}</Cell>
            </>
          ))}
        </Row>
      </Table>
      <Gap pt={10} />
      <Table id="glossary" variant="glossary" columns={[{ width: 85.82 }, { width: 6.35 }, { width: 85.82 }]}>
        {[0, 2].map((start) => (
          <Row>
            {plan.glossary.slice(start, start + 2).map((term, index) => (
              <>
                {index > 0 && <Cell></Cell>}
                <Cell>
                  <Paragraph>
                    <Text variant="glossaryTerm">{`${term.term}.`}</Text>
                    <Text variant="glossaryMeaning">{` ${term.meaning}`}</Text>
                  </Paragraph>
                </Cell>
              </>
            ))}
          </Row>
        ))}
      </Table>
    </>
  );
};
