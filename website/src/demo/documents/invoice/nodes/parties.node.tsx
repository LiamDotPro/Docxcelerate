import {
  Cell,
  Paragraph,
  Row,
  Section,
  type Section as SectionComponent,
  Table,
  useState,
} from "docxcelerate/template";
import type { InvoiceData } from "../types.ts";

/**
 * Who is being billed, beside who is billing them.
 *
 * A two-column table rather than two runs of paragraphs, because the two
 * addresses have to sit level however many lines each one has — a five-line
 * address next to a three-line one is the case that breaks anything else.
 *
 * The section is titled for both parties. The columns are already headed
 * "Billed to" and "From", and a section heading repeating one of them printed
 * the same two words twice, a few millimetres apart.
 */
export const Parties: SectionComponent = () => {
  const [state] = useState((data: InvoiceData) => ({
    billedTo: data.billedTo,
    sender: data.sender,
  }));

  return (
    <Section id="parties" title="Parties" showTitle={false}>
      <Table id="parties-grid" columns={[{ width: "auto" }, { width: "auto" }]}>
        {/*
          `label`, not `header`. A header row draws the theme's navy bar, which
          is right for the charges table and wrong here — the design sets these
          two as small tracked capitals over their columns, with no bar. Naming
          a variant is what stops the bar: the navy is only the default for a
          header row that resolves to nothing else.
        */}
        <Row>
          <Cell variant="label">Billed to</Cell>
          <Cell variant="label">From</Cell>
        </Row>
        <Row>
          <Cell id="billed-to" variant="lineItem">
            <Paragraph>{state.billedTo.name}</Paragraph>
            <Paragraph>{state.billedTo.attn}</Paragraph>
            {state.billedTo.addressLines.map((line) => <Paragraph>{line}</Paragraph>)}
          </Cell>
          <Cell id="billed-from" variant="lineItem">
            <Paragraph>{state.sender.name}</Paragraph>
            {state.sender.addressLines.map((line) => <Paragraph>{line}</Paragraph>)}
            <Paragraph>{state.sender.email}</Paragraph>
          </Cell>
        </Row>
      </Table>
    </Section>
  );
};
