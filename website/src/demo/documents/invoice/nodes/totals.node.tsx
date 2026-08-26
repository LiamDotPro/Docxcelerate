import { Cell, type Nodes, Row, Table, useDeriver, useState } from "docxcelerate/template";
import { invoiceTotals } from "../derivers.ts";
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
 * The bar is named cell by cell rather than as a row, and the row is not a
 * header. `header` is what sets a row in tracked capitals through `headingRun`
 * — which is a heading's treatment, not a total's — and naming the row would
 * paint the empty spacer navy too, running the bar across the whole page where
 * the design stops it above the figures it adds up.
 */
export const Totals: Nodes = async () => {
  const [state] = useState((data: InvoiceData) => ({
    lines: data.lines,
    rate: data.vatRate,
  }));
  // The arithmetic reaches the engine rather than being settled here: the
  // lines belong to a request nobody has made yet.
  const totals = await useDeriver(invoiceTotals, [state.lines, state.rate]);

  return (
    <Table
      id="totals-table"
      columns={[{ width: "auto" }, { width: 50 }, { width: 34, align: "right" }]}
    >
      <Row>
        <Cell></Cell>
        <Cell variant="panel">Subtotal</Cell>
        <Cell variant="panel">{totals.subtotal}</Cell>
      </Row>
      <Row>
        <Cell></Cell>
        <Cell variant="panel">VAT ({totals.rate})</Cell>
        <Cell variant="panel">{totals.vat}</Cell>
      </Row>
      <Row>
        <Cell></Cell>
        <Cell variant="totalRow">Total due</Cell>
        <Cell variant="totalRow">{totals.total}</Cell>
      </Row>
    </Table>
  );
};
