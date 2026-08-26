import { Cell, Image, type Nodes, Paragraph, Row, Table, useState } from "docxcelerate/template";
import { senderMarkPng, senderMarkSvg } from "../assets.ts";
import type { InvoiceData } from "../types.ts";

/**
 * The top of the payment page: the mark, the sender, and what this page is.
 *
 * Body content rather than running furniture, because it belongs to this page
 * alone — the running strip says which invoice every page belongs to, and
 * saying it twice on the one page that also carries a wordmark would be the
 * reference three times over.
 *
 * It reads PAYMENT rather than INVOICE: the sheet can be handed to whoever
 * pays without the sheet that says what for, and it should say which of the
 * two it is.
 */
export const PaymentLetterhead: Nodes = () => {
  const [state] = useState((data: InvoiceData) => ({
    name: data.sender.name,
    reference: data.reference,
  }));

  return (
    <Table
      id="payment-letterhead"
      columns={[{ width: 10 }, { width: "auto" }, { width: 62, align: "right" }]}
    >
      <Row>
        <Cell id="payment-mark">
          <Image
            id="payment-mark-image"
            src={senderMarkSvg}
            fallbackSrc={senderMarkPng}
            alt={state.name}
            width={18}
            height={18}
          />
        </Cell>
        <Cell id="payment-sender">
          <Paragraph variant="senderName">{state.name}</Paragraph>
        </Cell>
        <Cell id="payment-wordmark">
          <Paragraph variant="wordmark">Payment</Paragraph>
          <Paragraph variant="reference">{state.reference}</Paragraph>
        </Cell>
      </Row>
    </Table>
  );
};
