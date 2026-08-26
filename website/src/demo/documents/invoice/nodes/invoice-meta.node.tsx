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
import { invoiceDates } from "../derivers.ts";
import type { InvoiceData } from "../types.ts";

/**
 * The dates band, and the status the invoice carries.
 *
 * One row, not two: each cell holds its label over its value, which is how the
 * design draws it. Two rows put every label on one line and every value on the
 * next, so a long PO reference pushed all four values down together.
 *
 * The status is a decision rather than a label: a settled invoice must not
 * print a payment deadline, and one still owing must not claim to be paid.
 * Each outcome is written as its own cell so that the build compiles the
 * ternary into a condition — a ternary picking between two *strings* would be
 * a value, and a value is settled once, at build time, for everybody. Both
 * cells travel, each carrying the test that selects it, and the engine chooses
 * per recipient.
 */
export const InvoiceMeta: SectionComponent = async () => {
  const [state] = useState((data: InvoiceData) => ({
    issueDate: data.issueDate,
    dueDate: data.dueDate,
    poReference: data.poReference,
    paid: data.paid,
  }));
  // Formatting a value read from the request is exactly what a deriver is for:
  // `useFormat` would settle it here, once, for everybody.
  const dates = await useDeriver(invoiceDates, [state.issueDate, state.dueDate]);

  return (
    <Section id="invoice-meta" title="Invoice details" showTitle={false}>
      <Table
        id="meta-band"
        variant="band"
        // The status is a pill, so its column is measured rather than left to
        // take up the slack: an `"auto"` column here made the box as wide as
        // whatever the three dates did not use. The slack goes to the PO
        // reference instead, which is plain text and does not mind.
        columns={[{ width: 38 }, { width: 42 }, { width: "auto" }, { width: 44, align: "right" }]}
      >
        <Row>
          <Cell id="issue" variant="bandCell">
            <Paragraph variant="label">Issue date</Paragraph>
            <Paragraph>{dates.issue}</Paragraph>
          </Cell>
          <Cell id="due" variant="bandCell">
            <Paragraph variant="label">Due date</Paragraph>
            <Paragraph>{dates.due}</Paragraph>
          </Cell>
          <Cell id="po" variant="bandCell">
            <Paragraph variant="label">PO reference</Paragraph>
            <Paragraph>{state.poReference}</Paragraph>
          </Cell>
          {state.paid
            ? <Cell id="status-paid" variant="badge-done">Paid</Cell>
            : <Cell id="status-awaiting" variant="badge">Awaiting payment</Cell>}
        </Row>
      </Table>
    </Section>
  );
};
