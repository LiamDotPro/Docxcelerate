import { Cell, Paragraph, Row, Section, type Section as SectionComponent, Table, useDeriver, useState } from "docxcelerate/template";
import invoiceTotals from "../derivers/invoice-totals.deriver.ts";
import type { InvoiceTotalsData } from "../derivers/invoice-totals.deriver.ts";

/**
 * An invoice's line table and the totals beneath it, from one calculation.
 *
 * The figures are worked once, in the deriver, so the page agrees with the
 * customer's accounting system to the penny. Money is integer minor units at
 * the currency's own exponent. A line is rounded once, lines are grouped by
 * rate, and VAT is worked once per band on the band's sum (rate-level rounding,
 * which HMRC and the EU accept); VAT rounded per line and summed drifts a penny
 * from that. Inclusive prices have their VAT backed out of the band's gross, so
 * the total is exactly what was charged. A discount reduces its own band and is
 * never pooled across rates. An exempt band prints under its own label, never
 * as 0%, because zero-rated is a different treatment. More than one band draws
 * the per-rate analysis an invoice is required to carry, and an invoice with
 * no lines says there are no charges rather than drawing a zero-total table.
 *
 * The node does no arithmetic: arithmetic on request data throws at publish, so
 * it only interpolates, loops and branches on what the deriver returned.
 *
 * Installed by `dxcl add invoice-totals`.
 */

export type { InvoiceTotalsData } from "../derivers/invoice-totals.deriver.ts";

export const InvoiceTotals: SectionComponent = async () => {
  const [invoice] = useState((d: InvoiceTotalsData) => d.invoice);
  const totals = await useDeriver(invoiceTotals, [invoice]);

  return (
    <Section id="invoice-totals" title="Charges">
      {totals.empty && <Paragraph id="invoice-totals-none">There are no charges on this invoice.</Paragraph>}
      {!totals.empty && (
        <Table
          id="invoice-totals-lines"
          columns={[
            { width: "auto" },
            { width: 22, align: "right" },
            { width: 26, align: "right" },
            { width: 18, align: "right" },
            { width: 28, align: "right" },
          ]}
        >
          <Row header>
            <Cell>Description</Cell>
            <Cell>Qty</Cell>
            <Cell>Unit price</Cell>
            <Cell>VAT</Cell>
            <Cell>{totals.amountHeading}</Cell>
          </Row>
          {totals.lines.map((line) => (
            <Row>
              <Cell>{line.description}</Cell>
              <Cell>{line.quantity}</Cell>
              <Cell>{line.unitPrice}</Cell>
              <Cell>{line.tax}</Cell>
              <Cell>{line.amount}</Cell>
            </Row>
          ))}
        </Table>
      )}
      {totals.multipleBands && (
        <Table
          id="invoice-totals-analysis"
          columns={[{ width: "auto" }, { width: 34, align: "right" }, { width: 34, align: "right" }]}
        >
          <Row header>
            <Cell>Rate</Cell>
            <Cell>Net</Cell>
            <Cell>VAT</Cell>
          </Row>
          {totals.bands.map((band) => (
            <Row>
              <Cell>{band.label}</Cell>
              <Cell>{band.net}</Cell>
              <Cell>{band.vat}</Cell>
            </Row>
          ))}
        </Table>
      )}
      {!totals.empty && (
        <Table id="invoice-totals-summary" columns={[{ width: "auto" }, { width: 40, align: "right" }]}>
          <Row id="invoice-totals-subtotal">
            <Cell>{totals.subtotalLabel}</Cell>
            <Cell>{totals.subtotal}</Cell>
          </Row>
          <Row id="invoice-totals-vat">
            <Cell>{totals.vatLabel}</Cell>
            <Cell>{totals.vat}</Cell>
          </Row>
          <Row id="invoice-totals-total">
            <Cell>Total</Cell>
            <Cell>{totals.total}</Cell>
          </Row>
        </Table>
      )}
    </Section>
  );
};
