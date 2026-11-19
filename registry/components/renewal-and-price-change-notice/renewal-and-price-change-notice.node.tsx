import { arrange, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { RenewalAndPriceChangeNoticeLines } from "../arrangements/renewal-and-price-change-notice/lines.ts";
import { stacked } from "../arrangements/renewal-and-price-change-notice/stacked.arrangement.tsx";
import { table } from "../arrangements/renewal-and-price-change-notice/table.arrangement.tsx";
import renewalAndPriceChangeNotice from "../derivers/renewal-and-price-change-notice.deriver.ts";

/**
 * The body of a renewal reminder or price-change notice: when the contract
 * renews and at what price, the last day to cancel, how to cancel, and what
 * the governing law adds. Every date is measured from the one that law uses.
 *
 * Cancel-by: the renewal date less the notice period, with month arithmetic
 * clamped to the end of the month (31 March less a month is the end of
 * February). Printed as "on or before". With no notice period the deadline is
 * the day before the renewal, because "on or before the renewal date" would
 * let a cancellation land after the charge.
 *
 * California (2025 Automatic Renewal Law): three windows, each against its own
 * anchor: the renewal date, the end of a trial or introductory price over 31
 * days, and the date a price change takes effect.
 *
 * New York (GOL section 5-903, consumers and businesses): the window is
 * measured from the cancellation deadline, not the renewal. A late notice
 * makes the automatic renewal unenforceable, and the notice says so.
 *
 * German consumer contracts past their minimum term run until cancelled, with
 * at most a month's notice, so they get no fixed date. A longer contractual
 * notice period is flagged and capped.
 *
 * UK subscription rules apply from a commencement date that is a parameter,
 * because it has moved twice. Anything else gets contractual wording only.
 *
 * A late notice drops the fixed deadline and prints what the law gives the
 * customer instead. A notice sent too early prints a "Check before sending"
 * warning for whoever issues it.
 *
 * All of this is decided once, in one deriver: every step is date or money
 * arithmetic on request data, which publishing refuses, and written as node
 * arms the jurisdictions would spend most of the default `branchLimit` of 32.
 * The node publishes 6 branches per arrangement.
 *
 * Arrangement: where the lines sit is a separate choice from what they say,
 * made with the `arrangement` prop. This file decides every line; each
 * arrangement, in `arrangements/renewal-and-price-change-notice/`, only
 * places them.
 *
 * Installed by `dxcl add renewal-and-price-change-notice`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface RenewalAndPriceChangeNoticeData {
  renewal: {
    /**
     * Whose rules apply. `"other"` means no statutory renewal rule here:
     * contractual wording only. Required: anything else throws.
     */
    jurisdiction: "US-CA" | "US-NY" | "DE" | "UK" | "other";
    /**
     * Who the customer is. California, Germany and the UK apply their rules to
     * consumers only; New York §5-903 applies to both. Required: anything else throws.
     */
    customerType: "consumer" | "business";
    /**
     * When the contract was entered, ISO 8601 (`2025-09-30`; anything after the
     * tenth character is ignored). Decides whether the 2025 California rules,
     * the 2022 German rule and the UK rules apply. It also starts the initial
     * term and any trial. Required: throws when missing or not a real date.
     */
    contractStart: string;
    /**
     * The latest date the contract was amended or extended (a renewal counts),
     * ISO 8601. Read for US-CA only: the 2025 rules apply when this or
     * `contractStart` is on or after 2025-07-01. Missing means never amended.
     * Throws when present but not a date.
     */
    amendedOn?: string;
    /**
     * Length of the first term in whole months. Decides whether California's
     * renewal window applies (12 or more) and when a German contract's minimum
     * term ends. Required: throws unless a positive integer.
     */
    initialTermMonths: number;
    /**
     * Length of each renewal term (the billing interval once renewed) in whole
     * months. Printed as the frequency ("every year", "every 3 months"), and
     * decides whether New York §5-903 applies (more than 1). Required: throws
     * unless a positive integer.
     */
    renewalTermMonths: number;
    /**
     * The renewal this notice is about: the day the next term starts and is
     * charged, ISO 8601. Ignored for a German consumer contract already past
     * its minimum term. Required: throws when missing, not a date, before
     * `contractStart`, or before `sendDate` (except in that German indefinite phase).
     */
    renewalDate: string;
    /**
     * How much notice the contract requires the customer to give before the
     * renewal date. `{ value: 0 }` means none: the deadline is then the day
     * before the renewal, because "on or before the renewal date" would let a
     * cancellation land after the charge. Required: throws unless `value` is a
     * non-negative integer and `unit` is `"days"` or `"months"`.
     */
    noticePeriod: { value: number; unit: "days" | "months" };
    /**
     * The day this notice is sent, ISO 8601. Every window is checked against
     * it. Required: throws when missing or not a date. It is not defaulted to
     * today, because a deriver has to give the same answer on every run.
     */
    sendDate: string;
    /** ISO 4217 code: `"USD"`, `"EUR"`, `"GBP"`. Required: throws when `Intl` does not know it. */
    currency: string;
    /**
     * The price charged for each renewal term now, in major units
     * (`1284750.00`). Required: throws unless a finite number of zero or more.
     */
    currentAmount: number;
    /**
     * The last payment the customer made, in major units. Printed when given,
     * with the increase when this renewal costs more. Missing omits the line.
     * Under the UK rules it also prints a check warning, because a UK reminder
     * must show it. Throws when present but not a finite number of zero or more.
     */
    previousAmount?: number;
    /**
     * A price change. Missing means none. A `newAmount` equal to
     * `currentAmount` once both are rounded to minor units is also treated as
     * no change. Throws when present and `newAmount` is not a finite number of
     * zero or more, or `effectiveDate` is not a date.
     */
    priceChange?: { newAmount: number; effectiveDate: string };
    /**
     * When a free trial or introductory price ends, ISO 8601. The trial
     * started on `contractStart`. Missing, or before `sendDate`, means no trial
     * line. Throws when present but not a date.
     */
    trialEnds?: string;
    /**
     * How to cancel, as it should be printed ("Sign in, open Billing, choose
     * Cancel renewal"). Every regime prints it. Required: throws when missing
     * or blank.
     */
    cancelInstructions: string;
    /**
     * When the UK subscription-contract rules (DMCC Act 2024, Part 4 Chapter 2)
     * take effect, ISO 8601. Missing means `"2027-01-01"`. It is a parameter
     * because the date has moved twice. Throws when present but not a date.
     */
    commencementDate?: string;
  };
}

/**
 * Where the lines sit.
 *
 * - `stacked`: one labelled line under another, down the text column. The default.
 * - `table`: labels in a narrow column with the values beside them, like a
 *   key-facts box.
 */
export type RenewalAndPriceChangeNoticeArrangement = "stacked" | "table";

/** Props for {@linkcode RenewalAndPriceChangeNotice}. */
export interface RenewalAndPriceChangeNoticeProps {
  /** Where the lines sit. Defaults to `"stacked"`. */
  arrangement?: RenewalAndPriceChangeNoticeArrangement;
}

/** Prints the renewal notice: what renews, when, the price change if any, and how to cancel. */
export const RenewalAndPriceChangeNotice: SectionComponent<RenewalAndPriceChangeNoticeProps> = async (
  { arrangement },
) => {
  const [renewal] = useState((data: RenewalAndPriceChangeNoticeData) => data.renewal);
  const r = await useDeriver(renewalAndPriceChangeNotice, [renewal]);

  const lines: RenewalAndPriceChangeNoticeLines = {
    statement: r.statement,
    renewal: r.renewal,
    trial: r.trial,
    priceChange: r.priceChange,
    previousPayment: r.previousPayment,
    cancelBy: r.cancelBy,
    howToCancel: renewal.cancelInstructions,
    consequence: r.consequence,
    warning: r.warning,
  };

  return (
    <Section id="renewal-notice" title={r.heading}>
      {arrange(arrangement, {
        stacked: () => stacked(lines),
        table: () => table(lines),
      })}
    </Section>
  );
};
