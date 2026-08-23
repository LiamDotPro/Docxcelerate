import { Cell, type Nodes, Row, Table, useFormat, useState } from "docxcelerate/template";
import type { InvoiceData } from "../types.ts";

/**
 * Subtotal, VAT, and what is actually owed.
 *
 * Every figure is computed in the one initializer, from the same lines the
 * table above prints. That is deliberate: a total worked out in a second place
 * is a total that can disagree with the first, and on an invoice that is the
 * error nobody forgives.
 *
 * The block carries no heading of its own: the row a reader stops on says
 * TOTAL DUE, and a section titled "Total" above it is the same word twice.
 *
 * The last row is styled as a row rather than cell by cell. Naming only the
 * two cells with words in them left the empty one beside them drawing the
 * plain header fill, so the bar a reader's eye stops on came out in two
 * different navies.
 */
export const Totals: Nodes = () => {
  const { currency, number } = useFormat("en-GB");
  const [state] = useState((data: InvoiceData) => {
    const subtotal = data.lines.reduce((total, line) => total + line.qty * line.rate, 0);
    const vat = subtotal * data.vatRate;

    return { subtotal, vat, rate: data.vatRate, total: subtotal + vat };
  });

  return (
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
      <Row header variant="totalRow">
        <Cell></Cell>
        <Cell>Total due</Cell>
        <Cell>{currency(state.total)}</Cell>
      </Row>
    </Table>
  );
};
