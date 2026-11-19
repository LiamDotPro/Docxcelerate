import { Paragraph, type Yield } from "docxcelerate/template";
import type { VatInLocalCurrencyLines } from "./lines.ts";

/**
 * A separate statement under the invoice's totals, one paragraph per band. It
 * suits an invoice whose own VAT table should not grow two columns.
 */
export const block = (lines: VatInLocalCurrencyLines): Yield => (
  <>
    {lines.bands.map((band) => (
      <Paragraph id="vat-local-band">
        {band.rate}: net {band.net} = {band.localNet}; VAT {band.vat} = {band.localVat}.
      </Paragraph>
    ))}
    <Paragraph id="vat-local-total">
      Total in {lines.vatCurrency}: net {lines.total.localNet}, VAT {lines.total.localVat} (net {lines.total.net}, VAT{" "}
      {lines.total.vat}).
    </Paragraph>
    {lines.hasAdvances && (
      <Paragraph id="vat-local-advances-intro">
        Advances already invoiced, each at the rate of its own chargeable event:
      </Paragraph>
    )}
    {lines.advances.map((advance) => (
      <Paragraph id="vat-local-advance">
        {advance.label}, at {advance.rateLine}: net {advance.net} = {advance.localNet}; VAT {advance.vat} ={" "}
        {advance.localVat}.
      </Paragraph>
    ))}
  </>
);
