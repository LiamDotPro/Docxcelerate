import { Cell, type Nodes, Paragraph, Row, Table, useState } from "docxcelerate/template";
import { signed } from "../format.ts";
import { type BarLine, BarList, Opening, Rubric } from "./parts.tsx";
import { Written } from "./written.node.tsx";
import type { Contribution, PortfolioReviewData } from "../types.ts";

/** The bar column: 178mm less the 44mm label, the 14mm figure and two 8pt gaps. */
const TRACK = 114.36;
/** Where zero sits along it, and how much of it the largest contribution fills. */
const AXIS = 0.14;
const REACH = 0.8;

/** Contributions as diverging bars from a common axis, scaled to the largest. */
function diverging(rows: Contribution[]): BarLine[] {
  const largest = Math.max(...rows.map((row) => Math.abs(row.points)));
  const axis = TRACK * AXIS;
  return rows.map((row) => {
    const length = (Math.abs(row.points) / largest) * TRACK * REACH;
    const negative = row.points < 0;
    return {
      label: row.name,
      value: signed(row.points),
      valueVariant: negative ? "barValueMuted" : "barValue",
      marks: [
        negative
          ? { from: axis - length, to: axis, variant: "barPale", inset: true }
          : { from: axis, to: axis + length, variant: "barGreen", inset: true },
        { from: axis, to: axis + 0.26, variant: "barAxis" },
      ],
    };
  });
}

/** A two-column list of names and figures under a coloured label. */
const Leaders: Nodes<{ id: string; title: string; titleVariant: string; rows: Contribution[] }> = (
  { id, title, titleVariant, rows },
) => (
  <>
    <Paragraph variant={titleVariant}>{title}</Paragraph>
    <Table id={id} variant="leaders" columns={[{ width: "auto" }, { width: 14, align: "right" }]}>
      {rows.map((row) => (
        <Row>
          <Cell>{row.name}</Cell>
          <Cell>{signed(row.points)}</Cell>
        </Row>
      ))}
    </Table>
  </>
);

/** Page 5: what added to the quarter's return, by asset class and by holding. */
export const Attribution: Nodes = () => {
  const [att] = useState((data: PortfolioReviewData) => ({
    byClass: data.byClass,
    byHolding: data.byHolding,
    quarter: data.quarterPercent,
  }));

  const ranked = [...att.byHolding].sort((a, b) => b.points - a.points);
  const top = ranked.slice(0, 3);
  const bottom = [...ranked].reverse().slice(0, 3);

  return (
    <>
      <Opening
        id="attribution"
        number="3"
        title="Where the return came from"
        standfirst={`The quarter's ${att.quarter.toFixed(1)}% return, broken down by asset class and by holding.`}
      />
      <Rubric id="by-class" label="Contribution by asset class" note="Percentage points" variant="rubricTight" />
      <BarList id="class-bars" lines={diverging(att.byClass)} labelWidth={44} valueWidth={14} heightPt={15} insetPt={3} gapPt={7} />
      <Written
        id="return-drivers"
        variant="caption"
        ask="Two or three sentences on which parts of the portfolio added to the return and which held it back."
        placeholder="Which parts of the portfolio added to the return, and which held it back."
        from={{ byClass: att.byClass, byHolding: att.byHolding }}
      />
      <Rubric id="by-holding" label="Contribution by holding" note="Percentage points" />
      <BarList id="holding-bars" lines={diverging(att.byHolding)} labelWidth={44} valueWidth={14} heightPt={15} insetPt={3} gapPt={7} />

      <Rubric id="leaders" label="Top and bottom contributors" />
      <Table id="leaders-grid" variant="leadersGrid" columns={[{ width: 80.82 }, { width: 5.65 }, { width: "auto" }]}>
        <Row>
          <Cell>
            <Leaders id="added-most" title="Added most" titleVariant="leadersUp" rows={top} />
          </Cell>
          <Cell></Cell>
          <Cell>
            <Leaders id="held-back" title="Held back" titleVariant="leadersDown" rows={bottom} />
          </Cell>
        </Row>
      </Table>
    </>
  );
};
