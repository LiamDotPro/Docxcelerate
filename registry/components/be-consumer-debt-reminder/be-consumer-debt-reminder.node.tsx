import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import beConsumerDebtReminder from "../derivers/be-consumer-debt-reminder.deriver.ts";

/**
 * A Belgian payment reminder under Book XIX of the Code of Economic Law (CEL),
 * already split into the stage the law allows.
 *
 * Decisions already made for you:
 *
 * - A consumer's first reminder is free. No cost, interest or indemnity is
 *   added to it, and it carries the four statutory mentions of art. XIX.2 §3:
 *   the amount, the creditor's name and enterprise number, what the debt is
 *   for and when it fell due, and the period to pay. If one is missing it
 *   prints an "incomplete" warning instead, because without them the 14-day
 *   period never starts and nothing can ever be charged.
 * - The exact last day of the period is printed. It is 14 calendar days from
 *   the third working day after a paper reminder was sent, or from the day
 *   after an electronic one, never from the send date. Saturdays are not
 *   working days, which gives the later date and so can never make the period
 *   too short. The last day is not moved off a weekend or holiday.
 * - Costs only after the period has lapsed unpaid, and only what the contract
 *   expressly provides: a lump sum capped by bands on the outstanding amount,
 *   and late interest capped at the reference rate plus 8 points. Nothing else
 *   is passed on (art. XIX.4). With no first reminder on record, a period that
 *   has not lapsed or no clause, the section adds nothing and says so.
 * - A business debtor falls under the Act of 2 August 2002 instead: a fixed
 *   EUR 40 per unpaid invoice, owed without notice.
 *
 * All of the dates and money are worked out in one deriver, because every step
 * is date or money arithmetic on request data and publishing refuses all of it.
 * The node computes nothing: it branches on the deriver's result, with every
 * test written inline so a publish build keeps each arm.
 *
 * The node publishes 16 branches against the default `branchLimit` of 32.
 *
 * Not covered: regular-supply contracts, whether the contract clauses are
 * valid, interest on a business debt, part payments, collection agencies,
 * languages other than English and delivery of the reminder.
 *
 * Installed by `dxcl add be-consumer-debt-reminder`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface BeConsumerDebtReminderData {
  debt: {
    /**
     * Who owes the debt. `"consumer"` applies Book XIX CEL; `"business"` applies the
     * Act of 2 August 2002 instead and ignores `stage`, `firstReminder` and the
     * contract fields. Required: throws when missing or outside the union.
     */
    debtorType: "consumer" | "business";
    /**
     * Which reminder this document is. `"first"`: the free first reminder, being
     * sent now. `"subsequent"`: any later one, which may carry costs if the first
     * reminder's period has lapsed. Read for consumers only. Required for a consumer:
     * throws when missing or outside the union.
     */
    stage: "first" | "subsequent";
    /**
     * The date this document is issued, ISO 8601. Interest runs to this day, and it is
     * what decides whether the period has lapsed. Required: throws when missing.
     */
    asOf: string;
    /**
     * What is still owed, in euros, before any cost: the principal left after part
     * payments. The lump-sum bands and the interest are both worked out on it. Zero or
     * negative prints "nothing outstanding". Required: throws when missing or not finite.
     */
    amountDue: number;
    /**
     * When the debt fell due, ISO 8601. Printed on the first reminder (a statutory
     * mention). On or after `asOf` prints "not yet due". Required: throws when missing.
     */
    dueDate: string;
    /**
     * The creditor's name, as the consumer knows it. A statutory mention of the first
     * reminder: missing there prints the "incomplete" warning. Not read otherwise.
     */
    creditorName?: string;
    /**
     * The creditor's enterprise number (KBO/BCE), printed as given: "0712.345.678".
     * A statutory mention of the first reminder: missing there prints the "incomplete"
     * warning. Not read otherwise. Not checked for well-formedness.
     */
    creditorEnterpriseNumber?: string;
    /**
     * What the debt is for: the product or service that gave rise to it, with any
     * order or invoice reference. A statutory mention of the first reminder: missing
     * there prints the "incomplete" warning. Not read otherwise.
     */
    description?: string;
    /**
     * When and how the free first reminder was (or, at stage "first", is being) sent.
     * The 14-day period is counted from it. Missing at stage "first" prints the
     * "incomplete" warning. Missing at "subsequent" means no first reminder is on
     * record: nothing may be charged, and a warning says so.
     */
    firstReminder?: {
      /** The day it was handed to the post or sent, ISO 8601. */
      sentDate: string;
      /**
       * `"paper"`: the period starts on the third working day after `sentDate`.
       * `"electronic"`: it starts the calendar day after.
       */
      channel: "paper" | "electronic";
    };
    /**
     * True when the consumer accepted a contract clause that expressly provides a
     * lump-sum indemnity for late payment. False or missing: no lump sum is ever
     * announced or charged.
     */
    contractProvidesDamages?: boolean;
    /**
     * The lump sum the clause states, in euros, when it states a figure. The amount
     * charged is the lower of this and the statutory cap. Missing means the clause
     * adopts the statutory scale, so the cap is charged. 0 means none. Ignored unless
     * `contractProvidesDamages`.
     */
    agreedLumpSum?: number;
    /**
     * True when the consumer accepted a clause providing late interest. False or
     * missing: no interest is ever announced or charged.
     */
    contractProvidesInterest?: boolean;
    /**
     * The annual rate the clause states, as a percentage: 10 means 10% a year. The rate
     * applied is the lower of this and `referenceRate + 8`. Missing means the clause
     * adopts the statutory cap. Ignored unless `contractProvidesInterest`.
     */
    agreedRate?: number;
    /**
     * The reference interest rate under art. 5 of the Act of 2 August 2002 for the
     * current half-year, as a percentage (published by the FPS Finance). Needed
     * whenever interest is provided for, because the cap is this plus 8 points.
     * Missing then: at stage "first" this is an "incomplete" warning; at "subsequent"
     * the interest is not charged and a warning says why.
     */
    referenceRate?: number;
    /**
     * True when the creditor is an SME under art. 1:24 §1 of the Code of Companies and
     * Associations. If the period lapses unpaid, it may let interest run from the
     * calendar day after the first reminder was sent (art. XIX.2 §4), instead of from
     * the day after the period ends. Missing means false: the later start, which can
     * never overcharge.
     */
    creditorIsSme?: boolean;
    /**
     * Extra non-working days, ISO 8601, added to the ten Belgian legal holidays the
     * deriver already knows (for example 11 July or 27 September for a community, or a
     * company closure). Read only when counting the three working days after a paper
     * reminder.
     */
    holidays?: string[];
    /**
     * Business debtors only: how many unpaid invoices `amountDue` covers. Each one owes
     * the fixed EUR 40. Missing means 1.
     */
    invoiceCount?: number;
  };
}

/** Prints the reminder the law allows for this debt's stage, or says why nothing can be added. */
export const BeConsumerDebtReminder: Nodes = async () => {
  const [debt] = useState((data: BeConsumerDebtReminderData) => data.debt);
  const r = await useDeriver(beConsumerDebtReminder, [debt]);

  return (
    <>
      {r.outcome === "nothing-due" && (
        <Section id="be-debt-nothing-due" title="Your account">
          <Paragraph id="be-debt-nothing-due-text">
            There is nothing outstanding on this account, so there is nothing to pay.
          </Paragraph>
        </Section>
      )}
      {r.outcome === "not-yet-due" && (
        <Section id="be-debt-not-yet-due" title="Your account">
          <Paragraph id="be-debt-not-yet-due-text">
            The amount of {r.amountDue} falls due on {r.dueDate}. It is not overdue.
          </Paragraph>
        </Section>
      )}
      {r.outcome === "first" && (
        <Section id="be-debt-first" title="Payment reminder">
          <Paragraph id="be-debt-first-creditor">
            This reminder is sent by {debt.creditorName}, enterprise number {debt.creditorEnterpriseNumber}.
          </Paragraph>
          <Paragraph id="be-debt-first-debt">
            Our records show that {r.amountDue} is still outstanding for {debt.description}, which fell due on{" "}
            {r.dueDate}.
          </Paragraph>
          <Paragraph id="be-debt-first-free">
            This first reminder is free of charge: no costs, interest or damages have been added to it.
          </Paragraph>
          <Paragraph id="be-debt-first-deadline">
            Please pay {r.amountDue} by {r.lastDay}. You have 14 calendar days, from {r.periodStart} to {r.lastDay}{" "}
            inclusive, to pay without any costs.
          </Paragraph>
          {r.lumpSum && (
            <Paragraph id="be-debt-first-lump-sum">
              If it is not paid by then, a lump-sum indemnity of {r.lumpSum} will be claimed, as your contract provides.
            </Paragraph>
          )}
          {r.rate && (
            <Paragraph id="be-debt-first-interest">
              Late interest at {r.rate} a year will then also be charged on the outstanding amount, from{" "}
              {r.interestFrom}.
            </Paragraph>
          )}
          {r.addsNothing && (
            <Paragraph id="be-debt-first-adds-nothing">
              If it is not paid by then, no costs, interest or damages will be added.
            </Paragraph>
          )}
        </Section>
      )}
      {r.outcome === "first-incomplete" && (
        <Section id="be-debt-first-incomplete" title="Payment reminder">
          <Paragraph id="be-debt-first-incomplete-text">
            Reminder incomplete: a first reminder under Book XIX of the Code of Economic Law must state {r.missing}, and
            this one does not. Without it the 14-day period does not start. Add it before this reminder is sent.
          </Paragraph>
        </Section>
      )}
      {r.outcome === "waiting" && (
        <Section id="be-debt-waiting" title="Payment reminder">
          <Paragraph id="be-debt-waiting-text">
            The amount of {r.amountDue}, due on {r.dueDate}, is still outstanding. You can pay it without any costs,
            interest or damages until {r.lastDay} inclusive, and none have been added.
          </Paragraph>
        </Section>
      )}
      {r.outcome === "no-first-reminder" && (
        <Section id="be-debt-no-first-reminder" title="Payment reminder">
          <Paragraph id="be-debt-no-first-reminder-amount">
            The amount of {r.amountDue}, due on {r.dueDate}, is still outstanding.
          </Paragraph>
          <Paragraph id="be-debt-no-first-reminder-warning">
            Reminder incomplete: no first reminder is on record for this debt, so no costs, interest or damages can be
            claimed, and none have been added. A free first reminder has to be sent, and its 14-day period has to
            lapse, first.
          </Paragraph>
        </Section>
      )}
      {r.outcome === "lapsed" && (
        <Section id="be-debt-lapsed" title="Reminder: payment overdue">
          <Paragraph id="be-debt-lapsed-text">
            The period to pay without costs ended on {r.lastDay}, and {r.amountDue} is still outstanding.
          </Paragraph>
          {r.charged && (
            <Table id="be-debt-lapsed-charges" columns={[{ width: "auto" }, { width: 35, align: "right" }]}>
              <Row id="be-debt-lapsed-outstanding">
                <Cell>Outstanding amount</Cell>
                <Cell>{r.amountDue}</Cell>
              </Row>
              {r.lumpSum && (
                <Row id="be-debt-lapsed-lump-sum">
                  <Cell>Lump-sum indemnity under your contract</Cell>
                  <Cell>{r.lumpSum}</Cell>
                </Row>
              )}
              {r.interest && (
                <Row id="be-debt-lapsed-interest">
                  <Cell>Late interest at {r.rate} a year from {r.interestFrom} ({r.interestDays} days)</Cell>
                  <Cell>{r.interest}</Cell>
                </Row>
              )}
              <Row id="be-debt-lapsed-total">
                <Cell>Total now due</Cell>
                <Cell>{r.total}</Cell>
              </Row>
            </Table>
          )}
          {!r.charged && !r.interestIncomplete && (
            <Paragraph id="be-debt-lapsed-no-charges">
              Your contract provides for no indemnity or late interest, so nothing has been added to this amount.
            </Paragraph>
          )}
          {r.interestIncomplete && (
            <Paragraph id="be-debt-lapsed-interest-incomplete">
              Late interest not charged: your contract provides for it, but no reference rate was given, so the
              statutory cap could not be applied.
            </Paragraph>
          )}
        </Section>
      )}
      {r.outcome === "business" && (
        <Section id="be-debt-business" title="Overdue payment">
          <Paragraph id="be-debt-business-amount">
            The amount of {r.amountDue}, due on {r.dueDate}, has not been paid.
          </Paragraph>
          <Paragraph id="be-debt-business-compensation">
            Under the Act of 2 August 2002 on combating late payment in commercial transactions, a fixed compensation
            of €40.00 per unpaid invoice is owed without notice: {r.compensation} for {r.invoicesLabel}. The total now
            due is {r.total}.
          </Paragraph>
        </Section>
      )}
    </>
  );
};
