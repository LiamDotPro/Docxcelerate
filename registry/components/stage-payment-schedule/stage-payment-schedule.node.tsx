import { Cell, Paragraph, Row, Section, type Section as SectionComponent, Table, useDeriver, useState } from "docxcelerate/template";
import stagePaymentSchedule from "../derivers/stage-payment-schedule.deriver.ts";

/**
 * A construction or works contract's stage payments as one table that adds up
 * to the contract sum exactly, to the penny. For each stage it shows what is
 * paid, what is invoiced and what is still to come, and where UK VAT applies it
 * gives each stage the VAT tax point the law gives it.
 *
 * Decisions already made for you:
 *
 * - Rounding. Percentage stages are floored to the penny and the lost pennies
 *   go on the last percentage stage, so the stages never come to a penny or two
 *   more or less than the contract sum. A fixed stage is paid as stated and is
 *   never adjusted.
 * - Deposit. It is netted off by the contract's own rule: first stage, final
 *   stage or pro rata. A deposit bigger than the stage it comes off carries its
 *   excess to the next stage, so no stage is ever negative.
 * - Tax points, regulation 93. Under VAT Regulations 1995 reg. 93 (stage
 *   payments in the construction industry) a stage's tax point is the earlier
 *   of the date a VAT invoice is issued and the date payment is received.
 *   Reaching a milestone is never a tax point, so a stage that is reached but
 *   not invoiced shows no VAT due.
 * - VAT after the deposit. The deposit had its own tax point when it was
 *   received, so the VAT on a stage is worked on the stage after the deposit,
 *   never charged again on the full stage.
 * - Single-payment contracts. They fall outside reg. 93: the table gets no tax
 *   point column, and one line says the tax point is completion, unless the
 *   work is invoiced or paid for earlier.
 *
 * Every step does arithmetic, sorting or date logic on request data, so all of
 * it happens in one deriver and the node only prints what it returns.
 *
 * Notes:
 *
 * 1. The table has eight fixed stage slots, not a `.map()`. The resolver reads
 *    a published loop's collection from request data and ignores the derived
 *    scope, so a `.map()` over the deriver's rows would publish and then
 *    resolve to no rows. `plan-change-proration` has the same limit.
 * 2. The node publishes 21 branches against the default `branchLimit` of 32:
 *    the two table variants (with and without tax points), a deposit row in
 *    each, eight stage rows in each, and the tax note. The compile step counts
 *    the table choice as one decision, so it reports 20; the published
 *    document carries 21.
 * 3. All amounts are exclusive of VAT, in integer minor units.
 * 4. Installed by `dxcl add stage-payment-schedule`, which also installs its
 *    deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface StagePaymentScheduleData {
  contract: {
    /**
     * ISO 4217 code the amounts are in. It also sets how many minor units make
     * one major unit (two for GBP, none for JPY). Missing means "GBP". Throws
     * when present but not a code `Intl` knows.
     */
    currency?: string;
    /**
     * The contract sum, exclusive of VAT, in integer minor units
     * (4875250033 is £48,752,500.33). Every stage is measured against it, and
     * the schedule always adds up to it exactly. Required: throws when missing,
     * zero, negative or not a safe integer.
     */
    contractSumMinor: number;
    /**
     * The VAT the contract is under. Missing means no VAT applies: no VAT
     * column, no tax point column, no tax note. Only the UK's rules are
     * implemented.
     */
    vat?: {
      /** Whose rules apply. Only `"UK"`; anything else throws. */
      jurisdiction: "UK";
      /**
       * The rate as a percentage: 20 means 20%. At most two decimal places.
       * 0 (zero-rated) still shows the columns, because a zero-rated supply
       * still has tax points. Throws when not finite, below 0, above 100, or
       * given to more than two decimal places.
       */
      ratePercent: number;
    };
    /**
     * True for a contract paid in one sum on completion. Such a contract is
     * outside regulation 93. It must have exactly one stage (throws
     * otherwise). It gets no tax point column, and when `vat` is present it
     * gets one line saying the tax point is completion. Missing means false.
     */
    singlePayment?: boolean;
    /**
     * A deposit received before work started. Missing, or `amountMinor` 0,
     * means no deposit: no deposit row, no netting, and the other deposit
     * fields are not read.
     */
    deposit?: {
      /**
       * The deposit, exclusive of VAT, in integer minor units. Throws when
       * negative, not an integer, or more than `contractSumMinor`.
       */
      amountMinor: number;
      /**
       * When it was received, ISO 8601 (YYYY-MM-DD; anything after the tenth
       * character is ignored). This is the deposit's tax point. Required when
       * `amountMinor` is above 0: throws when missing or not a date.
       */
      receivedDate: string;
      /**
       * Which stages the contract nets the deposit off. `first`: off the first
       * stage, with any excess carried to the second, then the third. `final`:
       * off the last stage, with any excess carried back to the one before.
       * `proRata`: spread over every stage in proportion to its size, floored,
       * with the remaining pennies on the last stage. Required when
       * `amountMinor` is above 0: anything else throws, because a guessed rule
       * changes every amount on the page.
       */
      application: "first" | "final" | "proRata";
    };
    /**
     * The stages, in contract order. They are printed in this order and never
     * sorted. 1 to 8 entries: throws when missing, empty or longer than 8.
     */
    stages: Array<{
      /** The milestone as printed: "Roof structure topped out". Required: throws when missing or blank. */
      label: string;
      /**
       * A percentage of the contract sum, at most two decimal places: 12.5.
       * Floored to the minor unit. Give exactly one of `percent`,
       * `amountMinor` or `balance`, or it throws. Throws when not above 0,
       * above 100, or given to more than two decimal places.
       */
      percent?: number;
      /**
       * A fixed amount, exclusive of VAT, in integer minor units. Paid as
       * stated and never adjusted for rounding. Throws when not a positive
       * integer.
       */
      amountMinor?: number;
      /**
       * True for "the balance": whatever the contract sum less every earlier
       * stage comes to. Allowed on the last stage only (throws elsewhere).
       * Throws when the earlier stages already come to the whole contract sum
       * or more.
       */
      balance?: true;
      /**
       * Where the stage has got to. `pending`: not reached. `reached`: the
       * milestone is met but no VAT invoice has been issued. An application
       * for payment is not a VAT invoice. `invoiced`: a VAT invoice has been
       * issued but not paid. `paid`: payment received. Required: anything else
       * throws.
       */
      state: "pending" | "reached" | "invoiced" | "paid";
      /** When the milestone was met, ISO 8601. Printed in the status when given. Optional in every state. Throws when present but not a date. */
      reachedDate?: string;
      /** When the VAT invoice was issued, ISO 8601. Required when `state` is `invoiced`. Read when `paid` (see below). Ignored when `pending` or `reached`. Throws when present but not a date. */
      invoiceDate?: string;
      /** When payment was received, ISO 8601. Required when `state` is `paid`. Ignored otherwise. Throws when present but not a date. */
      paidDate?: string;
    }>;
  };
}

/** Prints the stage payment schedule, with VAT tax points where UK VAT applies. */
export const StagePaymentSchedule: SectionComponent = async () => {
  const [contract] = useState((data: StagePaymentScheduleData) => data.contract);
  const s = await useDeriver(stagePaymentSchedule, [contract]);

  return (
    <Section id="stage-payment-schedule" title="Stage payments">
      <Paragraph id="stage-payment-schedule-intro">{s.intro}</Paragraph>
      {s.taxPoints
        ? (
          <Table
            id="stage-payment-schedule-vat-table"
            columns={[{ width: "auto" }, { width: 34 }, { width: 28, align: "right" }, { width: 24, align: "right" }, { width: 28 }]}
          >
            <Row id="stage-payment-schedule-vat-header" header>
              <Cell>Stage</Cell>
              <Cell>Status</Cell>
              <Cell>Amount</Cell>
              <Cell>VAT</Cell>
              <Cell>Tax point</Cell>
            </Row>
            {s.deposit.description && (
              <Row id="stage-payment-schedule-vat-deposit">
                <Cell>{s.deposit.description}</Cell>
                <Cell>{s.deposit.status}</Cell>
                <Cell>{s.deposit.amount}</Cell>
                <Cell>{s.deposit.vat}</Cell>
                <Cell>{s.deposit.taxPoint}</Cell>
              </Row>
            )}
              {s.stages[0].description && (
                <Row id="stage-payment-schedule-vat-stage-1">
                  <Cell>{s.stages[0].description}</Cell>
                  <Cell>{s.stages[0].status}</Cell>
                  <Cell>{s.stages[0].amount}</Cell>
                  <Cell>{s.stages[0].vat}</Cell>
                  <Cell>{s.stages[0].taxPoint}</Cell>
                </Row>
              )}
              {s.stages[1].description && (
                <Row id="stage-payment-schedule-vat-stage-2">
                  <Cell>{s.stages[1].description}</Cell>
                  <Cell>{s.stages[1].status}</Cell>
                  <Cell>{s.stages[1].amount}</Cell>
                  <Cell>{s.stages[1].vat}</Cell>
                  <Cell>{s.stages[1].taxPoint}</Cell>
                </Row>
              )}
              {s.stages[2].description && (
                <Row id="stage-payment-schedule-vat-stage-3">
                  <Cell>{s.stages[2].description}</Cell>
                  <Cell>{s.stages[2].status}</Cell>
                  <Cell>{s.stages[2].amount}</Cell>
                  <Cell>{s.stages[2].vat}</Cell>
                  <Cell>{s.stages[2].taxPoint}</Cell>
                </Row>
              )}
              {s.stages[3].description && (
                <Row id="stage-payment-schedule-vat-stage-4">
                  <Cell>{s.stages[3].description}</Cell>
                  <Cell>{s.stages[3].status}</Cell>
                  <Cell>{s.stages[3].amount}</Cell>
                  <Cell>{s.stages[3].vat}</Cell>
                  <Cell>{s.stages[3].taxPoint}</Cell>
                </Row>
              )}
              {s.stages[4].description && (
                <Row id="stage-payment-schedule-vat-stage-5">
                  <Cell>{s.stages[4].description}</Cell>
                  <Cell>{s.stages[4].status}</Cell>
                  <Cell>{s.stages[4].amount}</Cell>
                  <Cell>{s.stages[4].vat}</Cell>
                  <Cell>{s.stages[4].taxPoint}</Cell>
                </Row>
              )}
              {s.stages[5].description && (
                <Row id="stage-payment-schedule-vat-stage-6">
                  <Cell>{s.stages[5].description}</Cell>
                  <Cell>{s.stages[5].status}</Cell>
                  <Cell>{s.stages[5].amount}</Cell>
                  <Cell>{s.stages[5].vat}</Cell>
                  <Cell>{s.stages[5].taxPoint}</Cell>
                </Row>
              )}
              {s.stages[6].description && (
                <Row id="stage-payment-schedule-vat-stage-7">
                  <Cell>{s.stages[6].description}</Cell>
                  <Cell>{s.stages[6].status}</Cell>
                  <Cell>{s.stages[6].amount}</Cell>
                  <Cell>{s.stages[6].vat}</Cell>
                  <Cell>{s.stages[6].taxPoint}</Cell>
                </Row>
              )}
              {s.stages[7].description && (
                <Row id="stage-payment-schedule-vat-stage-8">
                  <Cell>{s.stages[7].description}</Cell>
                  <Cell>{s.stages[7].status}</Cell>
                  <Cell>{s.stages[7].amount}</Cell>
                  <Cell>{s.stages[7].vat}</Cell>
                  <Cell>{s.stages[7].taxPoint}</Cell>
                </Row>
              )}
              <Row id="stage-payment-schedule-vat-paid">
                <Cell span={2}>Paid</Cell>
                <Cell>{s.totals.paid}</Cell>
                <Cell>{s.totals.paidVat}</Cell>
                <Cell></Cell>
              </Row>
              <Row id="stage-payment-schedule-vat-invoiced">
                <Cell span={2}>Invoiced, not yet paid</Cell>
                <Cell>{s.totals.invoiced}</Cell>
                <Cell>{s.totals.invoicedVat}</Cell>
                <Cell></Cell>
              </Row>
              <Row id="stage-payment-schedule-vat-not-invoiced">
                <Cell span={2}>Not yet invoiced</Cell>
                <Cell>{s.totals.notInvoiced}</Cell>
                <Cell>{s.totals.notInvoicedVat}</Cell>
                <Cell></Cell>
              </Row>
              <Row id="stage-payment-schedule-vat-contract-sum">
                <Cell span={2}>Contract sum</Cell>
                <Cell>{s.totals.contractSum}</Cell>
                <Cell></Cell>
                <Cell></Cell>
              </Row>
          </Table>
        )
        : (
          <Table
            id="stage-payment-schedule-table"
            columns={[{ width: "auto" }, { width: 40 }, { width: 30, align: "right" }]}
          >
            <Row id="stage-payment-schedule-header" header>
              <Cell>Stage</Cell>
              <Cell>Status</Cell>
              <Cell>Amount</Cell>
            </Row>
            {s.deposit.description && (
              <Row id="stage-payment-schedule-deposit">
                <Cell>{s.deposit.description}</Cell>
                <Cell>{s.deposit.status}</Cell>
                <Cell>{s.deposit.amount}</Cell>
              </Row>
            )}
              {s.stages[0].description && (
                <Row id="stage-payment-schedule-stage-1">
                  <Cell>{s.stages[0].description}</Cell>
                  <Cell>{s.stages[0].status}</Cell>
                  <Cell>{s.stages[0].amount}</Cell>
                </Row>
              )}
              {s.stages[1].description && (
                <Row id="stage-payment-schedule-stage-2">
                  <Cell>{s.stages[1].description}</Cell>
                  <Cell>{s.stages[1].status}</Cell>
                  <Cell>{s.stages[1].amount}</Cell>
                </Row>
              )}
              {s.stages[2].description && (
                <Row id="stage-payment-schedule-stage-3">
                  <Cell>{s.stages[2].description}</Cell>
                  <Cell>{s.stages[2].status}</Cell>
                  <Cell>{s.stages[2].amount}</Cell>
                </Row>
              )}
              {s.stages[3].description && (
                <Row id="stage-payment-schedule-stage-4">
                  <Cell>{s.stages[3].description}</Cell>
                  <Cell>{s.stages[3].status}</Cell>
                  <Cell>{s.stages[3].amount}</Cell>
                </Row>
              )}
              {s.stages[4].description && (
                <Row id="stage-payment-schedule-stage-5">
                  <Cell>{s.stages[4].description}</Cell>
                  <Cell>{s.stages[4].status}</Cell>
                  <Cell>{s.stages[4].amount}</Cell>
                </Row>
              )}
              {s.stages[5].description && (
                <Row id="stage-payment-schedule-stage-6">
                  <Cell>{s.stages[5].description}</Cell>
                  <Cell>{s.stages[5].status}</Cell>
                  <Cell>{s.stages[5].amount}</Cell>
                </Row>
              )}
              {s.stages[6].description && (
                <Row id="stage-payment-schedule-stage-7">
                  <Cell>{s.stages[6].description}</Cell>
                  <Cell>{s.stages[6].status}</Cell>
                  <Cell>{s.stages[6].amount}</Cell>
                </Row>
              )}
              {s.stages[7].description && (
                <Row id="stage-payment-schedule-stage-8">
                  <Cell>{s.stages[7].description}</Cell>
                  <Cell>{s.stages[7].status}</Cell>
                  <Cell>{s.stages[7].amount}</Cell>
                </Row>
              )}
              <Row id="stage-payment-schedule-paid">
                <Cell span={2}>Paid</Cell>
                <Cell>{s.totals.paid}</Cell>
              </Row>
              <Row id="stage-payment-schedule-invoiced">
                <Cell span={2}>Invoiced, not yet paid</Cell>
                <Cell>{s.totals.invoiced}</Cell>
              </Row>
              <Row id="stage-payment-schedule-not-invoiced">
                <Cell span={2}>Not yet invoiced</Cell>
                <Cell>{s.totals.notInvoiced}</Cell>
              </Row>
              <Row id="stage-payment-schedule-contract-sum">
                <Cell span={2}>Contract sum</Cell>
                <Cell>{s.totals.contractSum}</Cell>
              </Row>
          </Table>
        )}
      {s.taxNote && <Paragraph id="stage-payment-schedule-tax-note">{s.taxNote}</Paragraph>}
    </Section>
  );
};
