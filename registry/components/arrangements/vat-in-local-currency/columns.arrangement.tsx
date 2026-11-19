import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { VatInLocalCurrencyLines } from "./lines.ts";

/**
 * One table with the local figures beside each band, then the total. The
 * default: a reader checking the figures against the invoice's own VAT
 * analysis looks along the row. Advances get a second table of the same shape.
 */
export const columns = (lines: VatInLocalCurrencyLines): Yield => (
  <>
    <Table
      id="vat-local-bands"
      columns={[{ width: "auto" }, { width: 30, align: "right" }, { width: 30, align: "right" }, {
        width: 30,
        align: "right",
      }, { width: 30, align: "right" }]}
    >
      <Row id="vat-local-header" header>
        <Cell>Rate</Cell>
        <Cell>Net ({lines.invoiceCurrency})</Cell>
        <Cell>VAT ({lines.invoiceCurrency})</Cell>
        <Cell>Net ({lines.vatCurrency})</Cell>
        <Cell>VAT ({lines.vatCurrency})</Cell>
      </Row>
      {lines.bands.map((band) => (
        <Row id="vat-local-band">
          <Cell>{band.rate}</Cell>
          <Cell>{band.net}</Cell>
          <Cell>{band.vat}</Cell>
          <Cell>{band.localNet}</Cell>
          <Cell>{band.localVat}</Cell>
        </Row>
      ))}
      <Row id="vat-local-total">
        <Cell>Total</Cell>
        <Cell>{lines.total.net}</Cell>
        <Cell>{lines.total.vat}</Cell>
        <Cell>{lines.total.localNet}</Cell>
        <Cell>{lines.total.localVat}</Cell>
      </Row>
    </Table>
    {lines.hasAdvances && (
      <Paragraph id="vat-local-advances-intro">
        Advances already invoiced, each at the rate of its own chargeable event:
      </Paragraph>
    )}
    {lines.hasAdvances && (
      <Table
        id="vat-local-advances"
        columns={[{ width: "auto" }, { width: 30, align: "right" }, { width: 30, align: "right" }, {
          width: 30,
          align: "right",
        }, { width: 30, align: "right" }]}
      >
        <Row id="vat-local-advances-header" header>
          <Cell>Advance</Cell>
          <Cell>Net ({lines.invoiceCurrency})</Cell>
          <Cell>VAT ({lines.invoiceCurrency})</Cell>
          <Cell>Net ({lines.vatCurrency})</Cell>
          <Cell>VAT ({lines.vatCurrency})</Cell>
        </Row>
        {lines.advances.map((advance) => (
          <Row id="vat-local-advance">
            <Cell>
              <Paragraph>{advance.label}</Paragraph>
              <Paragraph>{advance.rateLine}</Paragraph>
            </Cell>
            <Cell>{advance.net}</Cell>
            <Cell>{advance.vat}</Cell>
            <Cell>{advance.localNet}</Cell>
            <Cell>{advance.localVat}</Cell>
          </Row>
        ))}
      </Table>
    )}
  </>
);
