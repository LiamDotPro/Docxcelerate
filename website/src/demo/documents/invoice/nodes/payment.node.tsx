import {
  Cell,
  Paragraph,
  Row,
  Section,
  type Section as SectionComponent,
  Table,
  useDeriver,
  useState,
} from "docxcelerate/template";
import { invoiceTotals } from "../derivers.ts";
import type { InvoiceData } from "../types.ts";

/**
 * Where the money goes, and what to quote when sending it.
 *
 * The amount is repeated here on purpose. Someone paying an invoice is looking
 * at this page, not the one before it, and a payment page that makes them turn
 * back to find the figure is a payment page that gets the figure wrong.
 */
export const Payment: SectionComponent = async () => {
  const [state] = useState((data: InvoiceData) => ({
    bank: data.sender.bank,
    reference: data.reference,
    lines: data.lines,
    rate: data.vatRate,
  }));
  // The same one pass over the lines as the totals table: two derivations of
  // one figure are two figures that can disagree.
  const totals = await useDeriver(invoiceTotals, [state.lines, state.rate]);

  return (
    <Section id="payment" title="Pay by bank transfer">
      <Table id="bank-details" columns={[{ width: 40 }, { width: "auto" }]}>
        <Row>
          <Cell variant="lineItem">Account name</Cell>
          <Cell variant="lineItem">{state.bank.accountName}</Cell>
        </Row>
        <Row>
          <Cell>Sort code</Cell>
          <Cell variant="money">{state.bank.sortCode}</Cell>
        </Row>
        <Row>
          <Cell>Account no</Cell>
          <Cell variant="money">{state.bank.accountNumber}</Cell>
        </Row>
        <Row>
          <Cell>IBAN</Cell>
          <Cell variant="money">{state.bank.iban}</Cell>
        </Row>
        <Row>
          <Cell>BIC</Cell>
          <Cell variant="money">{state.bank.bic}</Cell>
        </Row>
        <Row>
          <Cell>Amount</Cell>
          <Cell variant="money">{totals.total}</Cell>
        </Row>
      </Table>
      {/*
        One cell, three paragraphs — not three shaded paragraphs. Consecutive
        shaded paragraphs each draw their own box with the paragraph gap
        showing between them, so the panel the design draws as one card comes
        out as three stacked tiles.
      */}
      <Table id="reference-panel" columns={[{ width: "auto" }]}>
        <Row>
          <Cell id="reference-panel-cell" variant="panel">
            <Paragraph variant="label">Payment reference</Paragraph>
            <Paragraph variant="money">{state.reference}</Paragraph>
            <Paragraph variant="muted">
              Quote this reference on every transfer, so the payment reconciles on receipt.
            </Paragraph>
          </Cell>
        </Row>
      </Table>
    </Section>
  );
};
