import { type Nodes, Paragraph, useDeriver, useState } from "docxcelerate/template";
import financialAssistanceNotice from "../derivers/financial-assistance-notice.deriver.ts";

/**
 * The financial assistance block a US tax-exempt hospital's billing statement
 * carries under section 501(r) of the Internal Revenue Code: the notice, the
 * earliest date collections may start, the date applications close, and what
 * the patient's application changes. No heading, so it can sit on every
 * statement, and when it is suppressed it leaves nothing behind.
 *
 * Decisions already made for you:
 *
 * - Collections wait for the later of 120 days after the first post-discharge
 *   statement and 30 days after the written collections notice. With no notice
 *   yet the statement gives the 120-day date and says collection also waits
 *   for 30 days' written notice: it never prints a date the missing notice
 *   would make false.
 * - Applications are accepted until 240 days after the first statement, or the
 *   collections date if that is later. A later date than the law requires is
 *   always lawful; an earlier one is not.
 * - A pending application suspends collections and prints no dates.
 * - A partial approval caps the balance at the amount generally billed (AGB),
 *   worked out in whole cents and rounded half-up, and shows the working. The
 *   notice changes the balance, not just the wording. Free care says nothing is
 *   owed.
 * - A for-profit facility prints nothing, unless state law requires equivalent
 *   wording, and then only the availability notice: the timeline and the AGB
 *   cap are 501(r) commitments it has not made. A missing `taxExempt` is read
 *   as tax-exempt, because omitting the notice is the unlawful direction.
 * - A missing phone number or web address prints an "incomplete" warning in
 *   place of the notice; the timeline still prints.
 * - There is no arrangement. Every line is a sentence whose qualifying clause
 *   is what makes it true, and a label/value table would lift the bare dates
 *   and amounts out of those clauses.
 *
 * All of it is decided once, in one deriver, and the node prints what the
 * deriver returns. The dates and amounts are arithmetic on request data, which
 * cannot be done at publish time, and written as conditionals the cases come to
 * around twenty published arms. This publishes 8 branches against the default
 * `branchLimit` of 32.
 *
 * Not checked: that the phone number or address is well-formed or live
 * (presence after trimming is the only test), when a statement or notice counts
 * as "provided", state-specific wording, and the FAP's own eligibility rules.
 * Reading applications as open until the collections date, when that falls after
 * day 240, is the generous direction, but have someone with healthcare
 * compliance knowledge confirm the wording before release.
 *
 * Installed by `dxcl add financial-assistance-notice`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface FinancialAssistanceData {
  facility: {
    /**
     * Whether the hospital facility is tax-exempt under 501(c)(3), which is
     * what brings 501(r) into play. Missing is read as `true`: omitting the
     * notice from a tax-exempt hospital's statement is the unlawful direction,
     * printing it on a for-profit one is only generous.
     */
    taxExempt: boolean;
    /**
     * For a facility that is not tax-exempt: whether state law requires it to
     * give equivalent financial assistance wording anyway. `true` prints the
     * availability notice alone, with no timeline, suspension or AGB lines.
     * Missing means false. Ignored when `taxExempt` is not `false`.
     */
    stateNoticeRequired?: boolean;
    /**
     * The direct web address of the financial assistance policy, the
     * application form and the plain-language summary, printed as given
     * (trimmed). Missing or blank prints the "incomplete" warning in place of
     * the notice.
     */
    fapUrl: string;
    /**
     * The phone number of the office that explains the policy and the
     * application process, printed as given (trimmed). Missing or blank prints
     * the "incomplete" warning in place of the notice.
     */
    fapPhone: string;
  };
  billing: {
    /** ISO 4217 code for every amount printed. Defaults to USD. */
    currency?: string;
    /**
     * The date the first post-discharge billing statement for this care was
     * sent, as an ISO 8601 date: `2026-11-02`. Anything after the tenth
     * character is ignored. Both clocks start here. Required whenever the
     * timeline prints: the deriver throws when it is missing or not a real
     * calendar date.
     */
    firstPostDischargeStatementDate: string;
    /**
     * The date the written notice of intended extraordinary collection actions
     * was sent, ISO 8601. Missing means none has been sent: the statement then
     * says collection also waits for 30 days' written notice. Throws when
     * present but not a real calendar date.
     */
    ecaNoticeDate?: string;
    /**
     * The patient's financial assistance application. Missing means no
     * application (`status: "none"`).
     */
    application?: {
      /**
       * `none` and `denied` print the timeline. `pending` prints the
       * suspension. `approved` prints by `discount`. Missing means `none`;
       * any other value throws.
       */
      status?: "none" | "pending" | "approved" | "denied";
      /**
       * What an approval gives. Read only when `status` is `approved`, where
       * it is required: missing or any other value throws, because a wrong
       * amount on this notice is worse than no statement.
       */
      discount?: "free" | "partial";
    };
    /**
     * Gross charges for this care, before any discount, in `currency`.
     * Required for an approved partial discount (it is what AGB is a
     * percentage of): missing, negative or not finite throws there. Read
     * nowhere else.
     */
    grossCharges?: number;
    /**
     * The AGB percentage, `37.85` for 37.85%, to at most two decimal places
     * (more are rounded to two). Required for an approved partial discount:
     * missing, not finite, 0 or below, or above 100 throws there.
     */
    agbPercent?: number;
    /**
     * What the statement would otherwise ask for, after the policy's own
     * discount. Read only for an approved partial discount. Missing means the
     * patient owes the AGB figure. Given, the patient owes the lower of this
     * and AGB. Negative or not finite throws: a credit is not a balance this
     * notice caps.
     */
    balanceDue?: number;
  };
}

/**
 * The financial assistance notice and, by the patient's position, the
 * collections and application dates, the suspension, or the AGB-capped balance.
 */
export const FinancialAssistanceNotice: Nodes = async () => {
  const [{ facility, billing }] = useState((d: FinancialAssistanceData) => ({
    facility: d.facility,
    billing: d.billing,
  }));
  const fa = await useDeriver(financialAssistanceNotice, [facility, billing]);

  return (
    <>
      {fa.notice && <Paragraph id="financial-assistance-notice">{fa.notice}</Paragraph>}
      {fa.noticeIncomplete && (
        <Paragraph id="financial-assistance-notice-incomplete">{fa.noticeIncomplete}</Paragraph>
      )}
      {fa.suspended && <Paragraph id="financial-assistance-suspended">{fa.suspended}</Paragraph>}
      {fa.free && <Paragraph id="financial-assistance-free">{fa.free}</Paragraph>}
      {fa.agbBasis && <Paragraph id="financial-assistance-agb-basis">{fa.agbBasis}</Paragraph>}
      {fa.owed && <Paragraph id="financial-assistance-owed">{fa.owed}</Paragraph>}
      {fa.collections && <Paragraph id="financial-assistance-collections">{fa.collections}</Paragraph>}
      {fa.apply && <Paragraph id="financial-assistance-apply">{fa.apply}</Paragraph>}
    </>
  );
};
