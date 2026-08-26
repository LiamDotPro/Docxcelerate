import { Cell, Image, type Nodes, Paragraph, Row, Table, useState } from "docxcelerate/template";
import { senderMarkPng, senderMarkSvg } from "../assets.ts";
import type { InvoiceData } from "../types.ts";

/**
 * Who sent this, and which invoice it is.
 *
 * A three-column table rather than a run of paragraphs, because the two halves
 * have to sit level: the mark and the name on the left, the wordmark and the
 * reference hard right. Paragraphs would stack them, and a letterhead that
 * stacks is a letterhead that has stopped being one.
 *
 * The reference is the string that has to be exact — it is repeated on the
 * payment page and quoted on every transfer.
 */
export const Letterhead: Nodes = () => {
  const [state] = useState((data: InvoiceData) => ({
    name: data.sender.name,
    trade: data.sender.trade,
    reference: data.reference,
  }));

  return (
    <>
      <Table
        id="letterhead"
        columns={[{ width: 14 }, { width: "auto" }, { width: 62, align: "right" }]}
      >
        <Row>
          <Cell id="sender-mark">
            <Image
              id="sender-mark-image"
              src={senderMarkSvg}
              fallbackSrc={senderMarkPng}
              alt={state.name}
              width={28}
              height={28}
            />
          </Cell>
          <Cell id="sender" variant="lineItem">
            <Paragraph variant="senderName">{state.name}</Paragraph>
            <Paragraph variant="muted">{state.trade}</Paragraph>
          </Cell>
          <Cell id="wordmark" variant="lineItem">
            <Paragraph variant="wordmark">Invoice</Paragraph>
            <Paragraph variant="reference">{state.reference}</Paragraph>
          </Cell>
        </Row>
      </Table>
      {/*
        The rule the running strip draws on every other page. Page one has no
        running header — its letterhead is the top of the page — so the line
        under it belongs to the letterhead rather than to the furniture.
      */}
      <Paragraph id="letterhead-rule" variant="rule" />
    </>
  );
};
