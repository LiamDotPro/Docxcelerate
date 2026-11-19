import { Cell, Row, Table, type Yield } from "docxcelerate/template";
import type { GstInclusiveStatementLines } from "./lines.ts";

/**
 * Items and totals in two columns with the amounts right-aligned, so a reader
 * adding them up finds every figure in one column. The header says nothing
 * about GST, which keeps an unregistered invoice clean.
 */
export const table = (lines: GstInclusiveStatementLines): Yield => (
  <Table id="gst-lines" columns={[{ width: "auto" }, { width: 36, align: "right" }]}>
    <Row id="gst-lines-header">
      <Cell>Item</Cell>
      <Cell>Amount</Cell>
    </Row>
    {lines.items.map((item) => (
      <Row id="gst-item">
        <Cell>{item.description}</Cell>
        <Cell>{item.amount}</Cell>
      </Row>
    ))}
    {lines.subtotal && (
      <Row id="gst-subtotal">
        <Cell>Subtotal excluding GST</Cell>
        <Cell>{lines.subtotal}</Cell>
      </Row>
    )}
    {lines.gst && (
      <Row id="gst-amount">
        <Cell>{lines.gstLabel}</Cell>
        <Cell>{lines.gst}</Cell>
      </Row>
    )}
    <Row id="gst-total">
      <Cell>{lines.totalLabel}</Cell>
      <Cell>{lines.total}</Cell>
    </Row>
  </Table>
);
