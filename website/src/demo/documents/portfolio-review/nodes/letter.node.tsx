import { Cell, type Nodes, Paragraph, Row, Table, useState } from "docxcelerate/template";
import { Gap } from "./parts.tsx";
import { Written } from "./written.node.tsx";
import type { PortfolioReviewData } from "../types.ts";

/**
 * Page 2: the adviser's letter.
 *
 * The engine writes it as one node in four paragraphs, from the same facts the rest of the
 * review prints, so the letter cannot claim a figure the tables contradict.
 * The address, the sign-off and the contact strip are data.
 */
export const Letter: Nodes = () => {
  const [letter] = useState((data: PortfolioReviewData) => ({
    client: data.client,
    adviser: data.adviser,
    date: data.period.preparedOn,
    reference: data.period.reference,
    nextReview: data.nextReview,
    facts: {
      quarter: data.quarterPercent,
      benchmark: data.benchmarkPercent,
      value: data.closingValue,
      drift: data.drift,
      trades: data.trades,
      notes: data.clientNotes,
    },
  }));

  return (
    <>
      <Table id="letter-head" variant="letterHead" columns={[{ width: 110 }, { width: "auto", align: "right" }]}>
        <Row>
          <Cell>
            <Paragraph>{letter.client.name}</Paragraph>
            {letter.client.addressLines.map((line) => <Paragraph>{line}</Paragraph>)}
          </Cell>
          <Cell variant="letterDate">
            <Paragraph>{letter.date}</Paragraph>
            <Paragraph>{`Our reference ${letter.reference}`}</Paragraph>
          </Cell>
        </Row>
      </Table>
      <Paragraph id="greeting" variant="greeting">{`Dear ${letter.client.salutation},`}</Paragraph>
      <Paragraph id="letter-title" variant="letterTitle">Your quarter in brief</Paragraph>
      <Written
        id="letter"
        variant="letterBody"
        ask={"Four paragraphs, separated by blank lines: how the portfolio did against its benchmark and why; " +
          "what the quarter did to the agreed mix and the trades you recommend; what has and has not changed " +
          "in the client's plans; and whether they remain on track, with an invitation to say if anything changes."}
        placeholder={`A letter from ${letter.adviser.name} on the quarter, written from the figures, the decision and ${letter.client.salutation}'s plans.`}
        from={letter.facts}
      />
      <Paragraph id="closing" variant="closing">Yours sincerely,</Paragraph>
      <Paragraph id="signature" variant="signature">{letter.adviser.signature}</Paragraph>
      <Paragraph id="signatory" variant="signatory">{letter.adviser.name}</Paragraph>
      <Paragraph id="signatory-role" variant="signatoryRole">{letter.adviser.role}</Paragraph>

      <Gap pt={45} />
      <Table id="contact" variant="contact" columns={[{ width: 56.04 }, { width: 4.94 }, { width: 56.04 }, { width: 4.94 }, { width: 56.04 }]}>
        <Row>
          <Cell>
            <Paragraph variant="contactLabel">Your adviser</Paragraph>
            <Paragraph>{`${letter.adviser.name} · ${letter.adviser.phone}`}</Paragraph>
          </Cell>
          <Cell></Cell>
          <Cell>
            <Paragraph variant="contactLabel">Next review</Paragraph>
            <Paragraph>{letter.nextReview}</Paragraph>
          </Cell>
          <Cell></Cell>
          <Cell>
            <Paragraph variant="contactLabel">If things change</Paragraph>
            <Paragraph>Tell us before your review, so our advice reflects it.</Paragraph>
          </Cell>
        </Row>
      </Table>
    </>
  );
};
