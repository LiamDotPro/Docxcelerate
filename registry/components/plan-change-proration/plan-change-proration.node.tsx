import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import planChangeProration from "../derivers/plan-change-proration.deriver.ts";

/**
 * A subscription invoice section that turns each change made partway through a
 * billing period (to a plan, a quantity or a seat count) into the credit and
 * charge lines it causes, then says plainly what the net means.
 *
 * Decisions already made for you:
 *
 * - Days are counted, not assumed. The period's real length is counted in
 *   days, end date inclusive. It is never 30.
 * - Each line is rounded once, in integer minor units. The daily rate is never
 *   rounded first and then multiplied, so the pennies agree.
 * - The credit is what was actually paid: the amount charged after any
 *   discount, never the list price.
 * - Later changes chain. A second change in the same period credits the first
 *   change's new plan, not the original plan again, so nothing is credited twice.
 * - No zero or renewal-date lines. A change on the renewal date, a metered
 *   item, or a line that rounds to zero prints nothing.
 * - The net picks the outcome. Above zero: a charge included on this invoice.
 *   Below zero: a credit carried to the account balance, stated as not
 *   refunded unless `refundPolicy` is `"refund"`. Exactly zero: a lateral
 *   change, and the section only says so.
 *
 * All of it is decided in one deriver, because every step does arithmetic,
 * sorting or date maths on request data, and publishing refuses all of those.
 * The node computes nothing: it reads the deriver's result and branches on it,
 * with every test written inline so a publish build keeps each arm.
 *
 * The lines are eight fixed slots rather than a `.map()`. The resolver reads a
 * published loop's collection from request data and ignores the derived
 * scope, so a `.map()` over the deriver's lines would publish and then resolve
 * to no rows; conditions and `{{derived.…}}` tokens do honour the derived scope.
 *
 * The node publishes 14 branches against the default `branchLimit` of 32: the
 * section, the table, eight rows and four outcomes.
 *
 * Installed by `dxcl add plan-change-proration`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface PlanChangeProrationData {
  subscription: {
    /**
     * ISO 4217 code the amounts are in: `"GBP"`, `"EUR"`, `"JPY"`. It also sets
     * how many minor units make one major unit (two for GBP, none for JPY).
     * Required: the deriver throws when it is missing or not a code `Intl` knows.
     */
    currency: string;
    /**
     * The billing period this invoice covers, as ISO 8601 dates. `end` is
     * inclusive: 1 November to 30 November is 30 days. Anything after the tenth
     * character (a time) is ignored. Required: the deriver throws when either
     * end is missing or not YYYY-MM-DD, or when `end` is before `start`.
     */
    period: { start: string; end: string };
    /**
     * Every change made during the period, in any order. They are sorted by
     * date. Missing or empty means there is nothing to prorate, and the
     * component prints nothing.
     */
    changes: Array<{
      /**
       * The day the change took effect, ISO 8601. The change counts from the
       * start of that day, so that day is billed on the new plan. A time after
       * the date is ignored. Required: throws when it is missing or not a date.
       * A change on `period.start` (the renewal date), before it, or after
       * `period.end` produces no lines.
       */
      at: string;
      /**
       * Which subscription item changed: `"plan"`, `"seats"`. Changes chain
       * per item, so a seat change never credits the plan. Missing means every
       * change without an item is on the same single item.
       */
      item?: string;
      /**
       * True for an item billed on usage. Skipped entirely: no lines, and it
       * takes no part in chaining. Missing means false.
       */
      metered?: boolean;
      /** What the item was before this change. */
      from: {
        /** The plan's name as printed: `"Team (monthly)"`. Required. */
        label: string;
        /**
         * Seats or units. When it is not 1, the line prints `"12 × Team"`.
         * Missing, or 1, prints the label alone.
         */
        quantity?: number;
        /**
         * What the customer was actually charged for the whole period on this
         * plan, in integer minor units, after any discount (2400 for 24.00,
         * not the 30.00 list price). The credit is based on this. Required on
         * the first change for an item. On a later change for the same item it
         * is ignored, because the credit is based on the previous change's
         * `to.periodAmountMinor`. Throws when it is not a non-negative integer.
         */
        paidMinor: number;
      };
      /** What the item became. */
      to: {
        /** The new plan's name as printed. Required. */
        label: string;
        /** As `from.quantity`. */
        quantity?: number;
        /**
         * The new plan's price for a whole period, in integer minor units,
         * after any discount that applies to it. The charge is based on this.
         * Required: throws when it is not a non-negative integer.
         */
        periodAmountMinor: number;
      };
    }>;
    /**
     * What happens to a net credit. `"balance"`: it is carried to the account
     * balance and the section says it is not refunded. `"refund"`: the section
     * says it will be refunded. Missing means `"balance"`, because promising a
     * refund nobody will issue is worse than understating one. Ignored unless
     * the net is negative.
     */
    refundPolicy?: "balance" | "refund";
  };
}

/**
 * Prints the credit and charge lines for plan changes made mid-period, then what
 * the net means for the invoice. Prints nothing when no change left a line.
 */
export const PlanChangeProration: Nodes = async () => {
  const [subscription] = useState((data: PlanChangeProrationData) => data.subscription);
  const p = await useDeriver(planChangeProration, [subscription]);

  return (
    <>
      {p.outcome !== "none" && (
        <Section id="plan-change-proration" title="Changes to your plan">
          <Paragraph id="plan-change-proration-intro">
            Your plan changed during the billing period {p.periodLabel}. Each change is charged or credited for the
            days left in the period, out of the {p.days} days it actually has.
          </Paragraph>
          {p.showLines && (
            <Table
              id="plan-change-proration-lines"
              columns={[{ width: "auto" }, { width: 48 }, { width: 30, align: "right" }]}
            >
              <Row id="plan-change-proration-lines-header" header>
                <Cell>Description</Cell>
                <Cell>Period</Cell>
                <Cell>Amount</Cell>
              </Row>
              {p.lines[0].description && (
                <Row id="plan-change-proration-line-1">
                  <Cell>{p.lines[0].description}</Cell>
                  <Cell>{p.lines[0].period}</Cell>
                  <Cell>{p.lines[0].amount}</Cell>
                </Row>
              )}
              {p.lines[1].description && (
                <Row id="plan-change-proration-line-2">
                  <Cell>{p.lines[1].description}</Cell>
                  <Cell>{p.lines[1].period}</Cell>
                  <Cell>{p.lines[1].amount}</Cell>
                </Row>
              )}
              {p.lines[2].description && (
                <Row id="plan-change-proration-line-3">
                  <Cell>{p.lines[2].description}</Cell>
                  <Cell>{p.lines[2].period}</Cell>
                  <Cell>{p.lines[2].amount}</Cell>
                </Row>
              )}
              {p.lines[3].description && (
                <Row id="plan-change-proration-line-4">
                  <Cell>{p.lines[3].description}</Cell>
                  <Cell>{p.lines[3].period}</Cell>
                  <Cell>{p.lines[3].amount}</Cell>
                </Row>
              )}
              {p.lines[4].description && (
                <Row id="plan-change-proration-line-5">
                  <Cell>{p.lines[4].description}</Cell>
                  <Cell>{p.lines[4].period}</Cell>
                  <Cell>{p.lines[4].amount}</Cell>
                </Row>
              )}
              {p.lines[5].description && (
                <Row id="plan-change-proration-line-6">
                  <Cell>{p.lines[5].description}</Cell>
                  <Cell>{p.lines[5].period}</Cell>
                  <Cell>{p.lines[5].amount}</Cell>
                </Row>
              )}
              {p.lines[6].description && (
                <Row id="plan-change-proration-line-7">
                  <Cell>{p.lines[6].description}</Cell>
                  <Cell>{p.lines[6].period}</Cell>
                  <Cell>{p.lines[6].amount}</Cell>
                </Row>
              )}
              {p.lines[7].description && (
                <Row id="plan-change-proration-line-8">
                  <Cell>{p.lines[7].description}</Cell>
                  <Cell>{p.lines[7].period}</Cell>
                  <Cell>{p.lines[7].amount}</Cell>
                </Row>
              )}
              <Row id="plan-change-proration-net">
                <Cell>Net adjustment</Cell>
                <Cell></Cell>
                <Cell>{p.net}</Cell>
              </Row>
            </Table>
          )}
          {p.outcome === "upgrade" && (
            <Paragraph id="plan-change-proration-upgrade">
              The net charge of {p.netAbs} is included in the total of this invoice.
            </Paragraph>
          )}
          {p.outcome === "credit-balance" && (
            <Paragraph id="plan-change-proration-credit-balance">
              The net credit of {p.netAbs} has been added to your account balance and will be set against your next
              invoice. It is not refunded.
            </Paragraph>
          )}
          {p.outcome === "credit-refund" && (
            <Paragraph id="plan-change-proration-credit-refund">
              The net credit of {p.netAbs} will be refunded to your original payment method.
            </Paragraph>
          )}
          {p.outcome === "lateral" && (
            <Paragraph id="plan-change-proration-lateral">
              Your new plan costs the same as the old one for the rest of this period, so nothing is charged or
              credited.
            </Paragraph>
          )}
        </Section>
      )}
    </>
  );
};
