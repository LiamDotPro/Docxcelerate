import { Paragraph, type Yield } from "docxcelerate/template";
import type { GstInclusiveStatementLines } from "./lines.ts";

/** One line per item and per total, for an invoice laid out as text. */
export const stacked = (lines: GstInclusiveStatementLines): Yield => (
  <>
    {lines.items.map((item) => <Paragraph id="gst-item">{item.description}: {item.amount}</Paragraph>)}
    {lines.subtotal && <Paragraph id="gst-subtotal">Subtotal excluding GST: {lines.subtotal}</Paragraph>}
    {lines.gst && <Paragraph id="gst-amount">{lines.gstLabel}: {lines.gst}</Paragraph>}
    <Paragraph id="gst-total">{lines.totalLabel}: {lines.total}</Paragraph>
  </>
);
