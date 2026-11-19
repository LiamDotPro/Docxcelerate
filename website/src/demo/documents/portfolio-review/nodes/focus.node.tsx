import { Cell, Graph, type Nodes, PageBreak, Paragraph, Row, Table, useState } from "docxcelerate/template";
import { pounds } from "../format.ts";
import { Bar, Gap, Opening } from "./parts.tsx";
import { Written } from "./written.node.tsx";
import type { FocusHolding, PortfolioReviewData } from "../types.ts";

/** The "why we hold it" node id for each fund, as the canvas names them. */
const WHY_IDS: Record<string, string> = {
  "Global Equity Index Fund": "global-equity-why",
  "Global Bond Fund (hedged)": "global-bond-why",
  "UK Equity Income Fund": "uk-income-why",
  "Smaller Companies Fund": "smaller-cos-why",
  "Emerging Markets Fund": "emerging-why",
  "UK Gilt 4.25% 2031": "gilt-2031-why",
};

/** 178mm less the 60mm fact column and a 16pt gutter. */
const RIGHT = 112.36;
const POSITION_BAR = 22;

/** One fund: its figure and twelve months on the left, what is inside it and why on the right. */
const FocusPanel: Nodes<{ fund: FocusHolding }> = ({ fund }) => (
  <Table id={`focus-${fund.name}`} variant="focus" columns={[{ width: 60 }, { width: 5.64 }, { width: RIGHT }]}>
    <Row>
      <Cell>
        <Paragraph variant="focusName">{fund.name}</Paragraph>
        <Paragraph variant="focusKind">{fund.kind}</Paragraph>
        <Table id={`focus-value-${fund.name}`} variant="focusValueRow" columns={[{ width: 34 }, { width: "auto", align: "right" }]}>
          <Row>
            <Cell variant="focusValue">{pounds(fund.value)}</Cell>
            <Cell variant="focusShare">{`${fund.weight}% of portfolio`}</Cell>
          </Row>
        </Table>
        <Paragraph variant="gap6" />
        <Table id={`spark-frame-${fund.name}`} variant="sparkFrame" columns={[{ width: 60 }]}>
          <Row>
            <Cell>
              <Graph
                id={`spark-${fund.name}`}
                alt={`${fund.name}, twelve months`}
                graphType="line"
                legend="none"
                axes={false}
                markers={false}
                width={170}
                height={30}
                data={{
                  categories: fund.spark.map((_, i) => String(i + 1)),
                  series: [
                    { label: "Where it started", values: fund.spark.map(() => fund.spark[0]!), color: "C8D3CE", dash: true, widthPt: 0.75 },
                    { label: fund.name, values: fund.spark, color: "17624A", widthPt: 1.3 },
                  ],
                }}
              />
            </Cell>
          </Row>
        </Table>
        <Paragraph variant="focusSparkNote">Twelve months; the dashed line is where it started</Paragraph>
        <Table id={`facts-${fund.name}`} variant="focusFacts" columns={[{ width: "auto" }, { width: 26, align: "right" }]}>
          {fund.facts.map((fact) => (
            <Row>
              <Cell variant="factLabel">{fact.label}</Cell>
              <Cell variant="factValue">{fact.value}</Cell>
            </Row>
          ))}
        </Table>
      </Cell>
      <Cell></Cell>
      <Cell>
        <Paragraph variant="focusLabel">Largest positions</Paragraph>
        <Table
          id={`positions-${fund.name}`}
          variant="positions"
          columns={[{ width: RIGHT - POSITION_BAR - 10 - 4.24 }, { width: 2.12 }, { width: POSITION_BAR }, { width: 2.12 }, { width: 10, align: "right" }]}
        >
          {fund.positions.map((position, index) => (
            <Row>
              <Cell>{position.name}</Cell>
              <Cell></Cell>
              <Cell variant="positionBar">
                <Bar
                  id={`position-${fund.name}-${index}`}
                  width={POSITION_BAR}
                  heightPt={5}
                  track="barTrack"
                  marks={[{ from: 0, to: Math.min(1, position.percent / 10) * POSITION_BAR, variant: "barPale" }]}
                />
              </Cell>
              <Cell></Cell>
              <Cell>{`${position.percent}%`}</Cell>
            </Row>
          ))}
        </Table>
        <Paragraph variant="focusWhyLabel">Why we hold it</Paragraph>
        <Written
          id={WHY_IDS[fund.name] ?? `${fund.name}-why`}
          variant="focusWhy"
          ask="Two sentences on what this fund is for in this client's portfolio, and how it did."
          placeholder={`Why we hold ${fund.name}, for this client.`}
          from={{ fund: fund.name, purpose: fund.purpose, oneYear: fund.facts[2]?.value ?? "" }}
        />
      </Cell>
    </Row>
  </Table>
);

/** Pages 8 and 9: the six largest holdings, three to a page. */
export const Focus: Nodes = () => {
  const [focus] = useState((data: PortfolioReviewData) => ({ funds: data.focus }));

  return (
    <>
      <Opening id="focus" number="6" title="Holdings in focus" standfirst="Your six largest holdings: what each is for, and what is inside it." />
      {focus.funds.slice(0, 3).map((fund) => (
        <>
          <Gap pt={10} />
          <FocusPanel fund={fund} />
        </>
      ))}
      <PageBreak id="focus-continued" />
      <Opening id="focus-2" number="6" title="Holdings in focus, continued" />
      {focus.funds.slice(3, 6).map((fund) => (
        <>
          <Gap pt={10} />
          <FocusPanel fund={fund} />
        </>
      ))}
    </>
  );
};
