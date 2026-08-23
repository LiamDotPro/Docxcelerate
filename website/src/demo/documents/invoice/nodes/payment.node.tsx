import {
  Cell,
  Paragraph,
  Row,
  Section,
  type Section as SectionComponent,
  Table,
  useFormat,
  useState,
} from "docxcelerate/template";
import type { InvoiceData } from "../types.ts";

/**
 * Where the money goes, and what to quote when sending it.
 *
 * The amount is repeated here on purpose. Someone paying an invoice is looking
 * at this page, not the one before it, and a payment page that makes them turn
 * back to find the figure is a payment page that gets the figure wrong.
 */
export const Payment: SectionComponent = () => {
  const { currency } = useFormat("en-GB");
  const [state] = useState((data: InvoiceData) => ({
    bank: data.sender.bank,
    reference: data.reference,
    total: data.lines.reduce((sum, line) => sum + line.qty * line.rate, 0) * (1 + data.vatRate),
  }));

  return (
    <Section id="payment" title="Pay by bank transfer">
      <Table id="bank-details" columns={[{ width: 40 }, { width: "auto" }]}>
        <Row>
          <Cell>Account name</Cell>
          <Cell>{state.bank.accountName}</Cell>
        </Row>
        <Row>
          <Cell>Sort code</Cell>
          <Cell>{state.bank.sortCode}</Cell>
        </Row>
        <Row>
          <Cell>Account no</Cell>
          <Cell>{state.bank.accountNumber}</Cell>
        </Row>
        <Row>
          <Cell>IBAN</Cell>
          <Cell>{state.bank.iban}</Cell>
        </Row>
        <Row>
          <Cell>BIC</Cell>
          <Cell>{state.bank.bic}</Cell>
        </Row>
        <Row>
          <Cell>Amount</Cell>
          <Cell>{currency(state.total)}</Cell>
        </Row>
      </Table>
      <Paragraph id="payment-reference" variant="panel">
        Quote {state.reference} on every transfer, so the payment reconciles on receipt.
      </Paragraph>
    </Section>
  );
};
