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
 * Subtotal, VAT, and what is actually owed.
 *
 * Every figure is computed in the one initializer, from the same lines the
 * table above prints. That is deliberate: a total worked out in a second place
 * is a total that can disagree with the first, and on an invoice that is the
 * error nobody forgives.
 */
export const Totals: SectionComponent = () => {
  const { currency, number } = useFormat("en-GB");
  const [state] = useState((data: InvoiceData) => {
    const subtotal = data.lines.reduce((total, line) => total + line.qty * line.rate, 0);
    const vat = subtotal * data.vatRate;

    return { subtotal, vat, rate: data.vatRate, total: subtotal + vat };
  });

  return (
    <Section id="totals" title="Total">
      <Table
        id="totals-table"
        columns={[{ width: "auto" }, { width: 50 }, { width: 34, align: "right" }]}
      >
        <Row>
          <Cell></Cell>
          <Cell>Subtotal</Cell>
          <Cell>{currency(state.subtotal)}</Cell>
        </Row>
        <Row>
          <Cell></Cell>
          <Cell>VAT ({number(state.rate, { style: "percent" })})</Cell>
          <Cell>{currency(state.vat)}</Cell>
        </Row>
        <Row header>
          <Cell></Cell>
          <Cell variant="totalRow">Total due</Cell>
          <Cell variant="totalRow">{currency(state.total)}</Cell>
        </Row>
      </Table>
    </Section>
  );
};
