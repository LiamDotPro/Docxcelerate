import { Cell, type Nodes, Paragraph, Row, Table, useState } from "docxcelerate/template";
import { pounds, signedPounds } from "../format.ts";
import { type BarLine, BarList, Opening, Rubric } from "./parts.tsx";
import type { BridgeStep, PortfolioReviewData } from "../types.ts";

/** Where the bridge's axis starts and how many pounds its whole track spans. */
const AXIS_FROM = 390000;
const AXIS_SPAN = 30000;
/** 178mm less the 40mm label, the 20mm figure and two 8pt gaps. */
const TRACK = 112.36;

/**
 * Opening value to closing value as a waterfall: each step starts where the
 * last one ended. The steps are checked against the closing value, because a
 * bridge that does not arrive is the one chart on the page that is plainly wrong.
 */
function bridge(steps: BridgeStep[]): BarLine[] {
  const x = (pounds: number) => ((pounds - AXIS_FROM) / AXIS_SPAN) * TRACK;
  let running = 0;
  const lines = steps.map((step) => {
    if (step.total) {
      running = step.amount;
      return {
        label: step.label,
        value: pounds(step.amount),
        labelVariant: "barLabelStrong",
        valueVariant: "bridgeTotal",
        marks: [{ from: 0, to: x(step.amount), variant: "barDeep", inset: true }],
      };
    }
    const start = running;
    running += step.amount;
    return {
      label: step.label,
      value: signedPounds(step.amount),
      valueVariant: step.amount < 0 ? "barValueMuted" : "barValuePlain",
      marks: [{
        from: x(Math.min(start, running)),
        to: x(Math.max(start, running)),
        variant: step.amount < 0 ? "barPale" : "barGreen",
        inset: true,
      }],
    };
  });
  const closing = steps[steps.length - 1]!;
  const moved = steps.slice(1, -1).reduce((sum, step) => sum + step.amount, steps[0]!.amount);
  if (Math.abs(moved - closing.amount) > 0.5) {
    throw new Error(`The value bridge does not arrive: the steps reach ${moved}, the closing value is ${closing.amount}.`);
  }
  return lines;
}

/** Page 11: how the value changed, and every transaction in the quarter. */
export const Activity: Nodes = () => {
  const [activity] = useState((data: PortfolioReviewData) => ({
    bridge: data.bridge,
    transactions: data.transactions,
  }));

  return (
    <>
      <Opening id="activity" number="8" title="Activity" standfirst="How the value changed, and every transaction in the quarter." />
      <Rubric id="bridge" label="How the value changed" note="The axis starts at £390,000" variant="rubricTight" />
      <BarList id="bridge-bars" lines={bridge(activity.bridge)} labelWidth={40} valueWidth={20} heightPt={17} insetPt={3} gapPt={8} />

      <Rubric id="transactions" label="Transactions" />
      <Table
        id="transactions-table"
        variant="transactions"
        columns={[{ width: 18.12 }, { width: "auto" }, { width: 48.12 }, { width: 22.12, align: "right" }]}
      >
        <Row variant="transactionsHead">
          <Cell>Date</Cell>
          <Cell>Description</Cell>
          <Cell>Holding</Cell>
          <Cell>Amount</Cell>
        </Row>
        {activity.transactions.map((line) => (
          <Row>
            <Cell variant="returnMuted">{line.date}</Cell>
            <Cell>{line.description}</Cell>
            <Cell variant="returnMuted">{line.holding}</Cell>
            <Cell variant={line.amount < 0 ? "returnMuted" : "plain"}>{signedPounds(line.amount)}</Cell>
          </Row>
        ))}
      </Table>
      <Paragraph id="transactions-note" variant="tableNote">
        Contract notes for every transaction are in your online account. Income was collected in cash and paid to your nominated account on 30 September.
      </Paragraph>
    </>
  );
};
