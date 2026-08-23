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
 * What is being charged for, line by line.
 *
 * The rows are a `.map()`, which is the whole point: built against real data it
 * walks the lines and the preview shows seven rows, and published it becomes
 * one loop the engine walks — an invoice with three lines and one with thirty
 * are the same document. Nothing here knows which of the two is happening.
 */
export const Charges: SectionComponent = () => {
  const { currency, number } = useFormat("en-GB");
  const [state] = useState((data: InvoiceData) => ({ lines: data.lines }));

  return (
    <Section id="charges" title="Charges">
      <Table
        id="lines"
        columns={[
          { width: "auto" },
          { width: 16, align: "right" },
          { width: 24, align: "right" },
          { width: 26, align: "right" },
        ]}
      >
        <Row header>
          <Cell>Description</Cell>
          <Cell>Qty</Cell>
          <Cell>Rate</Cell>
          <Cell>Amount</Cell>
        </Row>
        {state.lines.map((line) => (
          <Row>
            <Cell>
              <Paragraph>{line.desc}</Paragraph>
              <Paragraph variant="muted">{line.meta}</Paragraph>
            </Cell>
            <Cell>{number(line.qty, { minimumFractionDigits: 1 })}</Cell>
            <Cell>{currency(line.rate)}</Cell>
            <Cell>{currency(line.qty * line.rate)}</Cell>
          </Row>
        ))}
      </Table>
    </Section>
  );
};
