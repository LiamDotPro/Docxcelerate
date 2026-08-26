import {
  Cell,
  Row,
  Section,
  type Section as SectionComponent,
  Table,
  useState,
} from "docxcelerate/template";
import { ChargeRow } from "./charge-row.node.tsx";
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

  const [state] = useState((data: InvoiceData) => ({ lines: data.lines }));

  return (
    <Section id="charges" title="Charges" showTitle={false}>
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
        {state.lines.map((line) => <ChargeRow line={line} />)}
      </Table>
    </Section>
  );
};
