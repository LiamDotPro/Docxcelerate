import { Cell, Paragraph, Row, Section, type Section as SectionComponent, Table, useDeriver, useState } from "docxcelerate/template";
import usSalesTaxBreakdown from "../derivers/us-sales-tax-breakdown.deriver.ts";

/**
 * The sales tax on a US invoice, one row per taxing jurisdiction: state, then
 * county, city and special district. Each row shows its own taxable amount,
 * rate and tax, and the rows add up exactly to the tax charged.
 *
 * Different bases: a city can exempt goods the state taxes, such as clothing,
 * and a county can cap the taxable price of a single item, as Tennessee and
 * Alabama do. One base applied to every layer would overtax or undertax, so
 * each row has its own.
 *
 * The residual goes to the state row: each row is rounded on its own base and
 * the total is rounded once from the exact sum, so the rows can differ from it
 * by a cent or two. That difference is put on the state row. Multiplying the
 * combined rate by the subtotal gives the right total but rows that do not add
 * up to it.
 *
 * Shipping is taxed proportionally, and only where delivery is taxed: it joins
 * a jurisdiction's base only when that jurisdiction taxes delivery, and only
 * for the share of the order that is taxable there.
 *
 * A combined shipping-and-handling charge warns. States that exempt delivery
 * stated on its own tax it when it is combined, so the warning says so and what
 * to change on the invoice.
 *
 * When a marketplace facilitator collected the tax, the rows are replaced by a
 * statement saying so.
 *
 * The rows are eight fixed slots and the exemptions four, rather than a
 * `.map()`: the resolver reads a published loop's entries from request data and
 * ignores the derived scope, so a loop over the deriver's rows would publish and
 * then resolve to nothing. More than eight jurisdictions throws; more than four
 * exemption groups fold into the last slot.
 *
 * All of it is decided in one deriver and the node only prints its output.
 * Per-request arithmetic cannot run at publish time, and branching per layer
 * would spend `branchLimit`. Tests on the deriver's output are written inline
 * (`tax.marketplace && ...`) and never hoisted into a `const`.
 *
 * Installed by `dxcl add us-sales-tax-breakdown`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface UsSalesTaxData {
  salesTax: {
    /**
     * ISO 4217 code that every `…Minor` amount is in. Missing means `"USD"`.
     * The number of minor digits comes from `Intl` (2 for USD), so
     * `amountMinor: 4699` in USD is $46.99. An unknown code throws.
     */
    currency?: string;
    /**
     * The goods and services on the invoice, in invoice order. Missing or empty
     * means no goods: every row prints a zero taxable amount and zero tax, and
     * shipping is not taxed, because there is nothing taxable for it to follow.
     */
    lines?: Array<{
      /**
       * What the line is, as printed on the invoice. Printed only when the line
       * is listed as exempt. Missing or blank on a line that has to be listed
       * throws.
       */
      description: string;
      /**
       * The line's net price in minor units (cents), after any line discount and
       * before tax. Must be a safe integer. A negative amount is a discount or a
       * return, and it reduces every base that includes the line. Anything that
       * is not an integer throws.
       */
      amountMinor: number;
      /**
       * Whether the state taxes this line at all. Required. Missing or not a
       * boolean throws: guessing taxability is wrong in both directions. `false`
       * leaves the line out of every base and lists it under `exemptReason`.
       */
      taxable: boolean;
      /**
       * A category matched exactly against each jurisdiction's
       * `exemptCategories`, for example `"clothing"`. Missing means no
       * jurisdiction exempts the line by category.
       */
      category?: string;
      /**
       * Why a `taxable: false` line is exempt, as printed. Missing prints
       * "Not taxable". Ignored on a taxable line.
       */
      exemptReason?: string;
    }>;
    /** The delivery charge. Missing means no shipping on the invoice. */
    shipping?: {
      /** Minor units, safe integer. Negative refunds shipping. Not an integer throws. */
      amountMinor: number;
      /**
       * Whether delivery is shown as its own line on the invoice. Missing means
       * `true`, since it arrived as its own field. `false` means it is folded
       * into the sale price, so it is taxable wherever the goods are.
       */
      separatelyStated?: boolean;
      /**
       * Whether this one charge covers shipping and handling together. In a
       * state that exempts delivery stated on its own, a combined charge is
       * taxable, and a warning prints. Missing means `false`.
       */
      combinedWithHandling?: boolean;
    };
    /**
     * Every jurisdiction that taxes this sale, in any order. They are printed
     * state, county, city, special, and in the given order within a level.
     * Required unless `collectedBy` is `"marketplace"`. Missing, empty, without
     * exactly one `state` entry, or with a level outside the four all throw.
     * A state with no state sales tax, such as Alaska, still sends a state row
     * with `rate: "0"`.
     */
    jurisdictions?: Array<{
      /** Which layer this is. */
      level: "state" | "county" | "city" | "special";
      /** What it is called, as printed: `"Travis County"`. Missing or blank throws. */
      name: string;
      /**
       * The rate as a decimal fraction string with at most 6 decimal places:
       * `"0.0625"`, `"0.0825"`, `"0.047125"`, `"0"`. Never a number and never a
       * percent. Not matching `^(0|[1-9]\d*)(\.\d{1,6})?$`, or a value of 1 or
       * more (which is how `"6.25"` gets caught), throws.
       */
      rate: string;
      /**
       * Whether this jurisdiction taxes delivery charges that are stated
       * separately. Missing means it follows the state row. Missing on the state
       * row means `false`.
       */
      taxesShipping?: boolean;
      /**
       * Line categories this jurisdiction does not tax, even though the line is
       * `taxable`. Matched exactly against `line.category`. Missing means none.
       */
      exemptCategories?: string[];
      /**
       * The most of any one line this jurisdiction taxes, in minor units. Each
       * line counts as a single item (see Out of scope). Missing means no cap.
       * Present but not a positive safe integer throws.
       */
      perItemCapMinor?: number;
    }>;
    /**
     * Who collected the tax. Missing means `"seller"`. With `"marketplace"`, no
     * rows, exemptions or warning are printed, only the facilitator statement,
     * and nothing else here is validated. Any other value throws.
     */
    collectedBy?: "seller" | "marketplace";
    /**
     * The facilitator's name, printed in the statement. Read only when
     * `collectedBy` is `"marketplace"`. Missing or blank prints "the marketplace
     * facilitator".
     */
    marketplaceName?: string;
  };
}

export const UsSalesTaxBreakdown: SectionComponent = async () => {
  const [salesTax] = useState((data: UsSalesTaxData) => data.salesTax);
  const tax = await useDeriver(usSalesTaxBreakdown, [salesTax]);

  return (
    <Section id="us-sales-tax" title="Sales tax">
      {tax.marketplace && <Paragraph id="us-sales-tax-marketplace">{tax.marketplace}</Paragraph>}
      {!tax.marketplace && (
        <Table
          id="us-sales-tax-table"
          columns={[{ width: "auto" }, { width: 32, align: "right" }, { width: 20, align: "right" }, { width: 28, align: "right" }]}
        >
          <Row id="us-sales-tax-header" header>
            <Cell>Jurisdiction</Cell>
            <Cell>Taxable amount</Cell>
            <Cell>Rate</Cell>
            <Cell>Tax</Cell>
          </Row>
          {tax.rowSlots[0].name && (
            <Row id="us-sales-tax-row-1">
              <Cell>{tax.rowSlots[0].name}</Cell>
              <Cell>{tax.rowSlots[0].taxable}</Cell>
              <Cell>{tax.rowSlots[0].rate}</Cell>
              <Cell>{tax.rowSlots[0].tax}</Cell>
            </Row>
          )}
          {tax.rowSlots[1].name && (
            <Row id="us-sales-tax-row-2">
              <Cell>{tax.rowSlots[1].name}</Cell>
              <Cell>{tax.rowSlots[1].taxable}</Cell>
              <Cell>{tax.rowSlots[1].rate}</Cell>
              <Cell>{tax.rowSlots[1].tax}</Cell>
            </Row>
          )}
          {tax.rowSlots[2].name && (
            <Row id="us-sales-tax-row-3">
              <Cell>{tax.rowSlots[2].name}</Cell>
              <Cell>{tax.rowSlots[2].taxable}</Cell>
              <Cell>{tax.rowSlots[2].rate}</Cell>
              <Cell>{tax.rowSlots[2].tax}</Cell>
            </Row>
          )}
          {tax.rowSlots[3].name && (
            <Row id="us-sales-tax-row-4">
              <Cell>{tax.rowSlots[3].name}</Cell>
              <Cell>{tax.rowSlots[3].taxable}</Cell>
              <Cell>{tax.rowSlots[3].rate}</Cell>
              <Cell>{tax.rowSlots[3].tax}</Cell>
            </Row>
          )}
          {tax.rowSlots[4].name && (
            <Row id="us-sales-tax-row-5">
              <Cell>{tax.rowSlots[4].name}</Cell>
              <Cell>{tax.rowSlots[4].taxable}</Cell>
              <Cell>{tax.rowSlots[4].rate}</Cell>
              <Cell>{tax.rowSlots[4].tax}</Cell>
            </Row>
          )}
          {tax.rowSlots[5].name && (
            <Row id="us-sales-tax-row-6">
              <Cell>{tax.rowSlots[5].name}</Cell>
              <Cell>{tax.rowSlots[5].taxable}</Cell>
              <Cell>{tax.rowSlots[5].rate}</Cell>
              <Cell>{tax.rowSlots[5].tax}</Cell>
            </Row>
          )}
          {tax.rowSlots[6].name && (
            <Row id="us-sales-tax-row-7">
              <Cell>{tax.rowSlots[6].name}</Cell>
              <Cell>{tax.rowSlots[6].taxable}</Cell>
              <Cell>{tax.rowSlots[6].rate}</Cell>
              <Cell>{tax.rowSlots[6].tax}</Cell>
            </Row>
          )}
          {tax.rowSlots[7].name && (
            <Row id="us-sales-tax-row-8">
              <Cell>{tax.rowSlots[7].name}</Cell>
              <Cell>{tax.rowSlots[7].taxable}</Cell>
              <Cell>{tax.rowSlots[7].rate}</Cell>
              <Cell>{tax.rowSlots[7].tax}</Cell>
            </Row>
          )}
          <Row id="us-sales-tax-total">
            <Cell>Total sales tax</Cell>
            <Cell>{tax.total.taxable}</Cell>
            <Cell>{tax.total.rate}</Cell>
            <Cell>{tax.total.tax}</Cell>
          </Row>
        </Table>
      )}
      {tax.exemptionSlots[0].reason && (
        <Paragraph id="us-sales-tax-exempt-1">{tax.exemptionSlots[0].reason}: {tax.exemptionSlots[0].items}.</Paragraph>
      )}
      {tax.exemptionSlots[1].reason && (
        <Paragraph id="us-sales-tax-exempt-2">{tax.exemptionSlots[1].reason}: {tax.exemptionSlots[1].items}.</Paragraph>
      )}
      {tax.exemptionSlots[2].reason && (
        <Paragraph id="us-sales-tax-exempt-3">{tax.exemptionSlots[2].reason}: {tax.exemptionSlots[2].items}.</Paragraph>
      )}
      {tax.exemptionSlots[3].reason && (
        <Paragraph id="us-sales-tax-exempt-4">{tax.exemptionSlots[3].reason}: {tax.exemptionSlots[3].items}.</Paragraph>
      )}
      {tax.shippingWarning && <Paragraph id="us-sales-tax-shipping-warning">{tax.shippingWarning}</Paragraph>}
    </Section>
  );
};
