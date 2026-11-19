import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { PaymentSurchargeLineLines } from "./lines.ts";

/**
 * The surcharge, its tax and the new total as label and amount rows, the
 * amounts in the second column so they line up with an invoice totals table
 * above. The notes follow beneath, full width.
 */
export const table = (lines: PaymentSurchargeLineLines): Yield => (
  <>
    {lines.surchargeAmount && (
      <Table id="surcharge-table" columns={[{ width: "auto" }, { width: 40, align: "right" }]}>
        <Row id="surcharge-amount">
          <Cell>{lines.surchargeLabel}</Cell>
          <Cell>{lines.surchargeAmount}</Cell>
        </Row>
        {lines.taxAmount && (
          <Row id="surcharge-tax">
            <Cell>{lines.taxLabel}</Cell>
            <Cell>{lines.taxAmount}</Cell>
          </Row>
        )}
        <Row id="surcharge-total">
          <Cell>{lines.totalLabel}</Cell>
          <Cell>{lines.totalAmount}</Cell>
        </Row>
      </Table>
    )}
    {lines.taxNote && <Paragraph id="surcharge-tax-note">{lines.taxNote}</Paragraph>}
    {lines.disclosure && <Paragraph id="surcharge-disclosure">{lines.disclosure}</Paragraph>}
    {lines.warning && <Paragraph id="surcharge-warning">{lines.warning}</Paragraph>}
  </>
);
