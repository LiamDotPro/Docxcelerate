import { Paragraph, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import invoiceBillingPeriods from "../derivers/invoice-billing-periods.deriver.ts";

/**
 * The periods a subscription invoice bills for. Calendar arithmetic is easy to
 * get nearly right and an invoice cannot be nearly right, so each decision is
 * made once and written down here.
 *
 * Month-end anchors are clamped, but the anchor day is never lost. An anchor of
 * the 31st bills on 31 Jan, 28 Feb, then 31 Mar, not 28 Mar: every period is
 * counted from the original anchor, never from the previous one. An annual
 * anchor of 29 Feb falls on 28 Feb in a common year. `Date.setMonth` turns 31
 * Jan into 3 Mar, and code that clamps once stays on the 28th for ever.
 *
 * End dates are inclusive. A period ends the day before the next anchor,
 * because under EN 16931 an end date is the last day of the period.
 *
 * Licensed items are billed in advance, for the period the invoice opens.
 * Metered items are billed in arrears, for the period before it. One invoice
 * can state both, so a usage line is never labelled with the subscription's
 * period. The header invoicing period runs from the earliest start to the
 * latest end, so it contains every line period, as Peppol R110 and R111
 * require; a validator rejects the invoice otherwise.
 *
 * A first period that starts off the anchor is labelled a part period.
 *
 * Dates are calendar dates in the seller's billing time zone, not UTC
 * instants: an invoice raised at 03:00 UTC on 1 Dec is still 30 Nov in New
 * York.
 *
 * None of this arithmetic can run on request data at publish time, so it is
 * all one deriver and the node prints what it returns. The deriver also
 * returns the ISO dates (`header`, `advance`, `arrears`) for a caller writing
 * UBL; the node does not print them.
 *
 * Installed by `dxcl add invoice-billing-periods`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface InvoiceBillingPeriodsData {
  billing: {
    /**
     * The original billing anchor: the date the cycle counts from, never a
     * date that has already been clamped. Its day of month, and its month for
     * an annual plan, are kept for every period. An anchor of 31 Jan bills on
     * 28 Feb and then 31 Mar. Pass `2026-01-31`, not the `2026-02-28` it last
     * fell on. It may come before or after the invoice date; periods are
     * counted from it both ways.
     *
     * Either a calendar date, `YYYY-MM-DD`, taken as it is, or an instant with
     * `Z` or an offset, such as `2026-02-01T00:00:00Z`, which is converted to
     * a calendar date in `timeZone`. Required. The deriver throws when it is
     * missing, is not a real date (for example `2026-02-30`), or is a date-time
     * with no offset.
     */
    anchor: string;
    /**
     * The unit the cycle repeats in. Required. Anything else, including
     * `"day"`, throws.
     */
    interval: "week" | "month" | "year";
    /**
     * How many units make up one period. 3 with `"month"` is quarterly.
     * Missing means 1. Throws when it is given but is not a whole number of at
     * least 1 (0, -1 and 1.5 all throw).
     */
    intervalCount?: number;
    /**
     * When the invoice is issued. This decides which period is billed in
     * advance (the one containing this date) and which in arrears (the one
     * before it). It accepts the same forms as `anchor`, and an instant is
     * read in `timeZone`. Required. It throws on the same inputs as `anchor`.
     */
    invoiceDate: string;
    /**
     * The seller's billing time zone, as an IANA name: `"America/New_York"`,
     * `"Europe/London"`. Every instant is turned into a calendar date here, so
     * a period never prints a day out for the customer. Required, even when
     * every date is a plain calendar date: the deriver throws when it is
     * missing or is not a zone the runtime knows.
     */
    timeZone: string;
    /**
     * When the subscription started, in the same forms as `anchor`. When it
     * falls strictly inside the period being billed, that period starts on
     * this date and is labelled a part period. Missing means the subscription
     * was running before every period stated, so nothing is a part period.
     * Throws when it is after `invoiceDate`. Throws when `hasArrearsItems` is
     * true and the subscription had not started by the end of the arrears
     * period, because there is no earlier period to bill usage for.
     */
    signupDate?: string;
    /**
     * Whether the invoice carries licensed items, which are billed in advance
     * for the coming period. Only `true` counts, so a missing value is false.
     * The deriver throws when neither this nor `hasArrearsItems` is true.
     */
    hasAdvanceItems: boolean;
    /**
     * Whether the invoice carries metered items, which are billed in arrears
     * for the previous period. Only `true` counts, so a missing value is false.
     */
    hasArrearsItems: boolean;
  };
}

export const InvoiceBillingPeriods: SectionComponent = async () => {
  const [billing] = useState((data: InvoiceBillingPeriodsData) => data.billing);
  const periods = await useDeriver(invoiceBillingPeriods, [billing]);

  return (
    <Section id="billing-periods" title="Billing period">
      <Paragraph id="billing-periods-statement">{periods.statement}</Paragraph>
      {periods.headerLine && <Paragraph id="billing-periods-header">{periods.headerLine}</Paragraph>}
    </Section>
  );
};
