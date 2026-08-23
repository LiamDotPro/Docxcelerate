import {
  Cell,
  Row,
  Section,
  type Section as SectionComponent,
  Table,
  useFormat,
  useState,
} from "docxcelerate/template";
import type { InvoiceData } from "../types.ts";

/**
 * The dates band, and the status the invoice carries.
 *
 * The status is a decision rather than a label: a settled invoice must not
 * print a payment deadline, and one still owing must not claim to be paid.
 * Each outcome is written as its own cell so that the build compiles the
 * ternary into a condition — a ternary picking between two *strings* would be
 * a value, and a value is settled once, at build time, for everybody. Both
 * cells travel, each carrying the test that selects it, and the engine chooses
 * per recipient.
 */
export const InvoiceMeta: SectionComponent = () => {
  const { date } = useFormat("en-GB");
  const [state] = useState((data: InvoiceData) => ({
    issueDate: data.issueDate,
    dueDate: data.dueDate,
    poReference: data.poReference,
    paid: data.paid,
  }));

  return (
    <Section id="invoice-meta" title="Invoice details">
      <Table
        id="meta-band"
        variant="band"
        columns={[{ width: 38 }, { width: 42 }, { width: 40 }, { width: "auto", align: "right" }]}
      >
        <Row header>
          <Cell>Issue date</Cell>
          <Cell>Due date</Cell>
          <Cell>PO reference</Cell>
          <Cell>Status</Cell>
        </Row>
        <Row>
          <Cell>{date(state.issueDate)}</Cell>
          <Cell>{date(state.dueDate)}</Cell>
          <Cell>{state.poReference}</Cell>
          {state.paid
            ? <Cell id="status-paid" variant="badge-done">Paid</Cell>
            : <Cell id="status-awaiting" variant="badge">Awaiting payment</Cell>}
        </Row>
      </Table>
    </Section>
  );
};
