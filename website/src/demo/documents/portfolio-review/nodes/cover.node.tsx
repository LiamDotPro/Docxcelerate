import { Cell, Image, type Nodes, PageNumber, Paragraph, Row, Table, useState } from "docxcelerate/template";
import { coverJpg } from "../assets.ts";
import { pounds, signedPounds } from "../format.ts";
import { Gap, Rubric } from "./parts.tsx";
import { StampWords } from "./stamp.node.tsx";
import { Written } from "./written.node.tsx";
import type { PortfolioReviewData } from "../types.ts";

/** The sections the contents lists, with the page each opens on. */
const CONTENTS: [string, string, string][] = [
  ["—", "A letter from your adviser", "2"],
  ["1", "Summary", "3"],
  ["2", "Performance", "4"],
  ["3", "Where the return came from", "5"],
  ["4", "Asset allocation", "6"],
  ["5", "Holdings", "7"],
  ["6", "Holdings in focus", "8"],
  ["7", "Income and cash", "10"],
  ["8", "Activity", "11"],
  ["9", "Costs and charges", "12"],
  ["10", "Your plan", "13"],
  ["11", "Important information", "13"],
];

/**
 * Page 1: the photograph, who this is for, what it is worth, and the decision.
 *
 * The photograph floats behind the page from the paper's corner; the words are
 * ordinary paragraphs set on top of it. A table holds them because a row of an
 * exact height with its text at the bottom is how the title block sits on the
 * foot of the picture whatever the client's name runs to.
 */
export const Cover: Nodes = () => {
  const [cover] = useState((data: PortfolioReviewData) => ({
    firm: data.firm.name.toUpperCase(),
    initial: data.firm.initial,
    period: data.period,
    client: data.client,
    value: pounds(data.closingValue),
    change: `▲ ${signedPounds(data.quarterReturn).slice(1)} investment return · ${data.quarterPercent.toFixed(1)}% this quarter`,
    facts: { quarter: data.quarterPercent, benchmark: data.benchmarkPercent, drift: data.drift },
    prepared: `Prepared by ${data.adviser.name}, ${data.adviser.role}, on ${data.period.preparedOn}.`,
    address: `${data.firm.name} Ltd · ${data.firm.address} · ${data.firm.phone}`,
  }));

  return (
    <>
      <Paragraph id="cover-anchor" variant="anchor">
        <Image id="cover-photo" src={coverJpg} alt="" width={595.3} height={470.7} float={{ xMm: 0, yMm: 0, behind: true }} />
      </Paragraph>
      <Table id="cover" variant="coverBand" columns={[{ width: 178 }]}>
        <Row height={18} heightRule="exact">
          <Cell>
            <Table id="cover-letterhead" variant="coverBand" columns={[{ width: 6.35 }, { width: 2.82 }, { width: 100 }, { width: "auto", align: "right" }]}>
              <Row height={18} heightRule="exact">
                <Cell variant="coverMark">{cover.initial}</Cell>
                <Cell></Cell>
                <Cell variant="coverFirm">{cover.firm}</Cell>
                <Cell variant="coverPrivate">Private and confidential</Cell>
              </Row>
            </Table>
          </Cell>
        </Row>
        <Row height={415.7} heightRule="exact">
          <Cell variant="coverFoot">
            <Paragraph variant="coverKicker">Quarterly portfolio review</Paragraph>
            <Paragraph variant="coverTitle">{cover.period.title}</Paragraph>
            <Paragraph variant="coverDates">{cover.period.range}</Paragraph>
            <Paragraph variant="coverGap" />
            <Table id="cover-for" variant="coverFor" columns={[{ width: 90 }, { width: "auto", align: "right" }]}>
              <Row>
                <Cell variant="coverForCell">
                  <Paragraph variant="coverLabel">Prepared for</Paragraph>
                  <Paragraph variant="coverName">{cover.client.name}</Paragraph>
                  <Paragraph variant="coverSub">
                    {`${cover.client.portfolio} · ${cover.client.account}`}
                  </Paragraph>
                </Cell>
                <Cell variant="coverForCell">
                  <Paragraph variant="coverLabelValue">{`Value at ${cover.period.valuedAt}`}</Paragraph>
                  <Paragraph variant="coverValue">{cover.value}</Paragraph>
                  <Paragraph variant="coverChange">{cover.change}</Paragraph>
                </Cell>
              </Row>
            </Table>
            <Paragraph variant="coverFootGap" />
          </Cell>
        </Row>
      </Table>

      <Gap pt={14} />
      <Table id="cover-decision-box" variant="coverDecision" columns={[{ width: 34 }, { width: "auto" }]}>
        <Row>
          <Cell variant="stamp">
            <StampWords
              id="cover-decision"
              ask="The decision for this quarter, as the one word Rebalance or Hold."
              placeholder="Rebalance or hold"
              noteAsk="A few words under the decision saying what kind of decision it is."
              notePlaceholder="decided per client"
              from={cover.facts}
            />
          </Cell>
          <Cell variant="coverHeadlineCell">
            <Paragraph variant="keyLabel">This quarter</Paragraph>
            <Written
              id="cover-headline"
              variant="coverHeadline"
              ask="One or two sentences summing up this client's quarter and the decision."
              placeholder="A one-sentence summary of the quarter, written for this client."
              from={cover.facts}
            />
          </Cell>
        </Row>
      </Table>

      <Rubric id="contents" label="Contents" variant="rubricTight" />
      <Table
        id="contents-list"
        variant="contents"
        columns={[
          { width: 7.06 }, { width: 72.3 }, { width: 6.5, align: "right" },
          { width: 6.35 },
          { width: 7.06 }, { width: 72.3 }, { width: 6.5, align: "right" },
        ]}
      >
        {[0, 1, 2, 3, 4, 5].map((row) => (
          <Row>
            {[CONTENTS[row * 2]!, CONTENTS[row * 2 + 1]!].map(([number, title, page], side) => (
              <>
                {side === 1 && <Cell variant="contentsGap"></Cell>}
                <Cell variant="contentsNumber">{number}</Cell>
                <Cell>{title}</Cell>
                <Cell variant="contentsPage">{page}</Cell>
              </>
            ))}
          </Row>
        ))}
      </Table>

      {/*
        Who prepared it, under the contents. The canvas pins it 17mm off the
        foot of the page; as body text it lands there too, and Word does not
        carve a two-line footer out of the cover's body to make room for it.
      */}
      <Gap pt={16} />
      <Table id="signoff" variant="signoff" columns={[{ width: 86.8 }, { width: 6.35 }, { width: "auto", align: "right" }]}>
        <Row>
          <Cell>{cover.prepared}</Cell>
          <Cell></Cell>
          <Cell>{cover.address}</Cell>
        </Row>
      </Table>
    </>
  );
};

/** The cover's foot: the regulator's line and the page count, like every other page. */
export const CoverFooter: Nodes = () => {
  const [foot] = useState((data: PortfolioReviewData) => ({
    line: `${data.firm.name} Ltd · ${data.firm.regulation}`,
  }));

  return (
    <Table id="cover-foot" variant="foot" columns={[{ width: "auto" }, { width: 9, align: "right" }, { width: 13, align: "right" }]}>
      <Row>
        <Cell>{foot.line}</Cell>
        <Cell variant="footPage">PAGE</Cell>
        <Cell variant="footPage">
          <PageNumber id="cover-page-number" format="currentOfTotal" separator=" OF " />
        </Cell>
      </Row>
    </Table>
  );
};
