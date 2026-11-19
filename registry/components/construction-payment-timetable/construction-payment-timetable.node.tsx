import { arrange, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { ConstructionPaymentTimetableLines } from "../arrangements/construction-payment-timetable/lines.ts";
import { stacked } from "../arrangements/construction-payment-timetable/stacked.arrangement.tsx";
import { table } from "../arrangements/construction-payment-timetable/table.arrangement.tsx";
import constructionPaymentTimetable from "../derivers/construction-payment-timetable.deriver.ts";

/**
 * The statutory payment timetable for one interim payment under a UK
 * construction contract: when payment falls due, the last day for the payer's
 * payment notice, the final date for payment and the last day for a pay less
 * notice. Each mistake below produces a deadline an adjudicator would reject.
 *
 * Counting: a period starts the day after its start date. Christmas Day, Good
 * Friday and the bank holidays of the contract's own region are not counted
 * (section 116). Weekends are, and a deadline that lands on a Saturday stays on
 * the Saturday rather than rolling to Monday.
 *
 * Region: the bank holidays come from the contract's region, so a Scottish job
 * never uses England's list. Northern Ireland is governed by its own Order.
 *
 * Pay less: counted back from the final date for payment as it stands after any
 * postponement, never forward from the due date.
 *
 * Notified sum: when the payer gives no valid payment notice, the payee's claim
 * stands if it arrived before the payment notice deadline (s110B(4)). Otherwise
 * the payee's own default notice stands, and the final date moves back by the
 * days it came after the deadline (s110B(3)). A payment notice after the
 * deadline is printed as late and treated as not given.
 *
 * Residential occupier: the Act does not apply, so one sentence says so and
 * nothing else is printed.
 *
 * All of it is decided in one deriver, because date arithmetic on request data
 * cannot run at publish time. The node only prints what it returns. The s110B
 * readings are interpretations, so have a lawyer confirm them.
 *
 * Installed by `dxcl add construction-payment-timetable`, which also installs
 * its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface ConstructionPaymentTimetableData {
  contract: {
    /**
     * Whose law and whose bank holidays apply: `EW` England and Wales, `SC`
     * Scotland, `NI` Northern Ireland. Required, even for a residential
     * occupier, because it names the Act. Anything else throws.
     */
    region: "EW" | "SC" | "NI";
    /**
     * Whether the contract is with a residential occupier (s106). `true`
     * prints only the "does not apply" sentence, and nothing else is read or
     * checked. Missing or anything but `true` means not residential.
     */
    residentialOccupier?: boolean;
    /**
     * Days after the valuation period ends that payment falls due, under the
     * contract. A whole number, 0 or more (0 is due on the day the period
     * ends). Missing (`undefined` or `null`) means the Scheme applies: due on
     * the later of 7 days after the period ends and receipt of the claim.
     * Negative, fractional or not a number throws.
     */
    dueDateDays?: number;
    /**
     * Days after the due date of the final date for payment. A fixed period,
     * a whole number, 1 or more. Missing means the Scheme's 17. Anything else
     * throws.
     */
    finalDateDays?: number;
    /**
     * The prescribed period for a pay less notice: days before the final date
     * for payment. A whole number, 0 or more (0 is the final date itself).
     * Missing means the Scheme's 7. Anything else throws.
     */
    payLessDays?: number;
  };
  payment: {
    /** The last day of the valuation period, as an ISO 8601 date (`2026-12-04`). Required. */
    valuationPeriodEnd: string;
    /**
     * When the payee's payment claim (application) was received, ISO 8601. A
     * claim states the sum the payee considers due and how it was worked out.
     * Leave the field out for anything that does not. Required when
     * `contract.dueDateDays` is missing, because under the Scheme it can set
     * the due date. When it was received before the payment notice deadline,
     * it stands as the notified sum if no payment notice is given. Missing or
     * `""` means no claim.
     */
    claimReceived?: string;
    /**
     * When the payer (or specified person) gave the payment notice, ISO 8601.
     * It counts only if given on or before the payment notice deadline. A
     * later one prints as late and is treated as not given. Missing or `""`
     * means none was given.
     */
    paymentNoticeIssued?: string;
    /**
     * When the payee gave its own notice in default of the payer's (s110B(2)),
     * ISO 8601. Read only when no valid payment notice was given and the claim
     * does not stand. It is ignored otherwise, because the Act then gives no
     * right to one. When read, it must be after the payment notice deadline,
     * or the deriver throws. Missing or `""` means none yet.
     */
    defaultNoticeIssued?: string;
  };
  /**
   * Bank holidays by region, each an array of ISO 8601 dates. This is the
   * shape of https://www.gov.uk/bank-holidays.json once its keys are mapped
   * to EW/SC/NI. Only `bankHolidays[contract.region]` is read. It is required
   * unless residential, and missing throws. It must contain at least one date
   * in every calendar year the timetable runs through, or the deriver throws,
   * because a year with none means the list is stale. Christmas Day and Good
   * Friday are excluded whether listed or not.
   */
  bankHolidays: { EW?: string[]; SC?: string[]; NI?: string[] };
}

/**
 * Where the dates sit.
 *
 * - `table`: the four dates in a label and date table. The default.
 * - `stacked`: one line per date, for a letter that reads as prose.
 */
export type ConstructionPaymentTimetableArrangement = "table" | "stacked";

/** Props for {@linkcode ConstructionPaymentTimetable}. */
export interface ConstructionPaymentTimetableProps {
  /** Where the dates sit. Defaults to `"table"`. */
  arrangement?: ConstructionPaymentTimetableArrangement;
}

export const ConstructionPaymentTimetable: SectionComponent<ConstructionPaymentTimetableProps> = async (
  { arrangement },
) => {
  const [contract] = useState((data: ConstructionPaymentTimetableData) => data.contract);
  const [payment] = useState((data: ConstructionPaymentTimetableData) => data.payment);
  const [bankHolidays] = useState((data: ConstructionPaymentTimetableData) => data.bankHolidays);
  const t = await useDeriver(constructionPaymentTimetable, [contract, payment, bankHolidays]);

  const lines: ConstructionPaymentTimetableLines = {
    applies: t.applies,
    notApplicable: t.notApplicable,
    basis: t.basis,
    dueDate: t.dueDate,
    paymentNoticeBy: t.paymentNoticeBy,
    finalDate: t.finalDate,
    payLessBy: t.payLessBy,
    noticeLate: t.noticeLate,
    noticeGiven: t.noticeGiven,
    applicationStands: t.applicationStands,
    defaultNotice: t.defaultNotice,
    defaultOpen: t.defaultOpen,
    reckoning: t.reckoning,
  };

  return (
    <Section id="payment-timetable" title="Payment timetable">
      {arrange(arrangement, {
        table: () => table(lines),
        stacked: () => stacked(lines),
      })}
    </Section>
  );
};
