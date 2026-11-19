import { Cell, type Nodes, Paragraph, Row, Table, Text, useState } from "docxcelerate/template";
import { pounds } from "../format.ts";
import { Gap, KeyFigures, Opening, Rubric } from "./parts.tsx";
import { StampWords } from "./stamp.node.tsx";
import { Written } from "./written.node.tsx";
import type { PortfolioReviewData } from "../types.ts";


/** Page 3: the figures, the decision, what happens next, and the quarter in three points. */
export const Summary: Nodes = () => {
  const [review] = useState((data: PortfolioReviewData) => ({
    figures: data.keyFigures,
    trades: data.trades.map((trade) => ({
      side: trade.side,
      fund: trade.fund,
      amount: pounds(trade.amount),
    })),
    steps: data.nextSteps,
    facts: { drift: data.drift, trades: data.trades, notes: data.clientNotes },
    valuedAt: data.period.valuedAt,
  }));

  return (
    <>
      <Opening
        id="summary"
        number="1"
        title="Summary"
        standfirst={`Where your portfolio stands at ${review.valuedAt}, and what we recommend.`}
      />

      <Gap pt={16} />
      <KeyFigures id="key-figures" figures={review.figures} columns={3} gutterPt={14} />

      <Rubric id="recommendation" label="Our recommendation" />
      <Table id="recommendation-box" variant="recBox" columns={[{ width: 34 }, { width: 86 }, { width: 58 }]}>
        <Row height={2} heightRule="exact">
          <Cell variant="barDeep"></Cell>
          <Cell variant="barDeep"></Cell>
          <Cell variant="barDeep"></Cell>
        </Row>
        <Row>
          <Cell variant="stamp">
            <StampWords
              id="decision"
              ask="Decide whether this client should rebalance or hold this quarter. Answer with the one word Rebalance or Hold."
              placeholder="Rebalance or hold"
              noteAsk="A few words under the decision saying what kind of decision it is."
              notePlaceholder="decided per client"
              from={review.facts}
            />
          </Cell>
          <Cell variant="recMiddle">
            <Written
              id="rebalance"
              variant="recHead"
              ask="One sentence giving the reason for the decision, naming the asset class and how far it is from target."
              placeholder="The reason for the decision."
              from={review.facts}
            />
            <Written
              id="rebalance-why"
              variant="recWhy"
              ask="Two sentences on why the decision follows from the drift and the client's plans."
              placeholder="Written by the engine from drift, agreed ranges and planned withdrawals."
              from={review.facts}
            />
          </Cell>
          <Cell variant="trades">
            <Paragraph variant="tradesLabel">Proposed trades</Paragraph>
            <Table id="trades" columns={[{ width: 8.47 }, { width: 28.47 }, { width: 14, align: "right" }]}>
              {review.trades.map((trade) => (
                <Row>
                  <Cell variant="trade">
                    <Text variant="tradeSide">{trade.side}</Text>
                  </Cell>
                  <Cell variant="trade">{trade.fund}</Cell>
                  <Cell variant="trade">
                    <Text variant="tradeAmount">{trade.amount}</Text>
                  </Cell>
                </Row>
              ))}
            </Table>
          </Cell>
        </Row>
      </Table>

      <Rubric id="next" label="What happens next" />
      <Table
        id="next-steps"
        variant="stepCell"
        columns={[
          { width: 7.05 }, { width: 48.99 }, { width: 4.94 },
          { width: 7.05 }, { width: 48.99 }, { width: 4.94 },
          { width: 7.05 }, { width: "auto" },
        ]}
      >
        <Row>
          {review.steps.map((step, index) => (
            <>
              {index > 0 && <Cell></Cell>}
              <Cell variant="stepNumber">{String(index + 1)}</Cell>
              <Cell>
                <Paragraph variant="stepTitle">{step.title}</Paragraph>
                <Paragraph variant="stepBody">{step.body}</Paragraph>
              </Cell>
            </>
          ))}
        </Row>
      </Table>

      <Rubric id="points" label="The quarter in three points" />
      <Written
        id="highlights"
        variant="point"
        list="number"
        ask="Three points from this client's quarter, one sentence each, separated by blank lines."
        placeholder="Three points from the quarter, written for this client."
        from={review.facts}
      />
    </>
  );
};
