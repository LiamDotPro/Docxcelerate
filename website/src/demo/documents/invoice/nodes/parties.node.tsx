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
 */
export const Parties: SectionComponent = () => {
  const [state] = useState((data: InvoiceData) => ({
    billedTo: data.billedTo,
    sender: data.sender,
  }));

  return (
    <Section id="parties" title="Billed to">
      <Table id="parties-grid" columns={[{ width: "auto" }, { width: "auto" }]}>
        <Row header>
          <Cell>Billed to</Cell>
          <Cell>From</Cell>
        </Row>
        <Row>
          <Cell id="billed-to">
            <Paragraph>{state.billedTo.name}</Paragraph>
            <Paragraph>{state.billedTo.attn}</Paragraph>
            {state.billedTo.addressLines.map((line) => <Paragraph>{line}</Paragraph>)}
          </Cell>
          <Cell id="billed-from">
            <Paragraph>{state.sender.name}</Paragraph>
            {state.sender.addressLines.map((line) => <Paragraph>{line}</Paragraph>)}
            <Paragraph>{state.sender.email}</Paragraph>
          </Cell>
        </Row>
      </Table>
    </Section>
  );
};
