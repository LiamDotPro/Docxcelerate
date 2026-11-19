import { arrange, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { DeMahnungFeeLines } from "../arrangements/de-mahnung-fee/lines.ts";
import { stacked } from "../arrangements/de-mahnung-fee/stacked.arrangement.tsx";
import { table } from "../arrangements/de-mahnung-fee/table.arrangement.tsx";
import deMahnungFee from "../derivers/de-mahnung-fee.deriver.ts";

/**
 * The part of a German Mahnung that says, per open invoice, since when the
 * debtor has been in Verzug and why, and what this letter may charge. The
 * component makes the legal decision so the reader does not have to.
 *
 * Verzug starts on the earliest of: a contractual payment date (§ 286 Abs. 2),
 * 30 days after due date and receipt of the invoice (§ 286 Abs. 3), or the
 * first Mahnung after the due date (§ 286 Abs. 1). The 30-day rule binds a
 * consumer only when the invoice warned of it, and a business whose invoice
 * date is unknown counts from receipt of the goods. A deadline ending on a
 * weekend or public holiday moves to the next working day (§ 193 BGB), the
 * nationwide holidays included, so the letter never claims Verzug a day early.
 *
 * The Mahnung that itself causes Verzug is free, and says so: its cost was not
 * caused by the delay. A later one carries material costs only, once per
 * letter, on the invoice that has been in Verzug longest. A consumer's fee is
 * held under a ceiling and the letter allows proof of lower damage; no
 * Pauschale is charged. A business owes the EUR 40 Pauschale once per invoice
 * and Mahnung fees are credited against it, so only the excess is charged.
 *
 * All of it is one deriver, because it is date arithmetic, sorting and money
 * on request data, which publishing refuses. The node prints what it returns.
 * Which arm prints is chosen once, outside the loops, so no loop body holds a
 * condition. The loops run over the deriver's rows, which the resolver reads
 * by scope, so any number of invoices is printed.
 *
 * Out of scope: Verzugszinsen, a Mahnung that sets its own payment deadline,
 * the Land holiday calendar (pass `holidays`), collection and lawyer fees, and
 * any totals. The wording is fixed German legal text.
 *
 * Arrangement: where the lines sit is a separate choice from what they say.
 * This file decides every line; each arrangement, in
 * `arrangements/de-mahnung-fee/`, only places them.
 *
 * Installed by `dxcl add de-mahnung-fee`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface DeMahnungFeeData {
  mahnung: {
    /**
     * Who owes the money. `consumer` is a Verbraucher (§ 13 BGB): no EUR 40
     * Pauschale, the 30-day rule only with a warning on the invoice, and a
     * fee of at most EUR 2.50. `business` is everyone else. Required:
     * anything else throws, because guessing either way puts an unlawful
     * charge on the letter.
     */
    debtorType: "consumer" | "business";
    /**
     * The date of this Mahnung, ISO 8601 (`2026-11-19`). It is taken as the
     * day the letter reaches the debtor. Anything after the tenth character is
     * ignored. Required: throws when missing or not YYYY-MM-DD.
     */
    date: string;
    /**
     * The material cost of this letter in euros (postage, print, registered
     * mail), never staff time. It is charged once for the whole letter, not once
     * per invoice. Missing means 0: nothing is charged. Throws when negative or
     * not finite. Throws above 2.50 when `debtorType` is `consumer`.
     */
    fee?: number;
    /**
     * Extra public holidays for the § 193 shift, ISO 8601. Use the holidays of
     * the debtor's Land, such as Reformationstag or Buß- und Bettag. The nine
     * nationwide holidays are always applied and need not be listed. Missing
     * means only weekends and the nationwide holidays. Throws when an entry is
     * not an ISO date.
     */
    holidays?: string[];
    /**
     * Every invoice this Mahnung covers, in any order. They are printed oldest
     * due date first. Required and non-empty: throws otherwise, because a
     * Mahnung with no invoice is not a letter.
     */
    invoices: Array<{
      /** The invoice number, printed as given. Required: throws when missing or blank. */
      number: string;
      /**
       * The amount still open on this invoice, in euros. It is printed and
       * never computed with. Required: throws when not finite or not above 0.
       */
      amount: number;
      /**
       * When payment fell due (Fälligkeit), ISO 8601. Required: throws when
       * missing or not a date. Also throws when `mahnung.date` is not after
       * this date shifted under § 193, because the invoice is not overdue yet.
       */
      dueDate: string;
      /**
       * When the debtor received the invoice, ISO 8601. Starts the § 286 Abs. 3
       * 30 days. Missing: a consumer gets no 30-day route at all, and a business
       * falls back to `goodsReceivedDate`. Throws when present but not a date.
       */
      invoiceReceivedDate?: string;
      /**
       * When the debtor received the goods or service, ISO 8601. Used only for a
       * business debtor, and only when `invoiceReceivedDate` is missing
       * (§ 286 Abs. 3 S. 2). Missing then means no 30-day route. Throws when
       * present but not a date.
       */
      goodsReceivedDate?: string;
      /**
       * A payment date fixed by contract, ISO 8601: either a calendar date
       * (§ 286 Abs. 2 Nr. 1), or a date the caller computed from an agreed event
       * such as "14 days after delivery" (Nr. 2; set `contractPaymentFromEvent`).
       * Missing means no Abs. 2 route. Throws when not a date or when it is
       * before `dueDate`.
       */
      contractPaymentDate?: string;
      /**
       * True when `contractPaymentDate` was computed from an event (Nr. 2)
       * rather than fixed in the calendar (Nr. 1). Changes only the wording.
       * Missing means false.
       */
      contractPaymentFromEvent?: boolean;
      /**
       * Whether the invoice warned that the debtor falls into Verzug 30 days
       * after due date and receipt. Read only for a consumer, for whom the
       * 30-day route applies only when this is true. Missing means false.
       */
      verzugWarningOnInvoice?: boolean;
      /**
       * The Mahnungen sent before this one for this invoice, in any order.
       * Missing or empty means this is the first.
       */
      mahnungen?: Array<{
        /**
         * The day it reached the debtor, ISO 8601. Throws when not a date, or
         * when it is not before `mahnung.date`. A Mahnung dated on or before the
         * shifted due date has no legal effect: it does not start Verzug, but
         * its fee still counts as charged.
         */
        date: string;
        /**
         * What it charged this invoice, in euros. Missing means 0. Throws when
         * negative or not finite. For a business debtor it is credited against
         * the Pauschale.
         */
        feeCharged?: number;
      }>;
    }>;
  };
}

/**
 * Where the lines sit.
 *
 * - `stacked`: each invoice as a heading with its lines beneath. The default.
 * - `table`: one row per invoice, for a letter covering many.
 */
export type DeMahnungFeeArrangement = "stacked" | "table";

/** Props for {@linkcode DeMahnungFee}. */
export interface DeMahnungFeeProps {
  /** Where the lines sit. Defaults to `"stacked"`. */
  arrangement?: DeMahnungFeeArrangement;
}

export const DeMahnungFee: SectionComponent<DeMahnungFeeProps> = async ({ arrangement }) => {
  const [mahnung] = useState((data: DeMahnungFeeData) => data.mahnung);
  const result = await useDeriver(deMahnungFee, [mahnung]);

  const lines: DeMahnungFeeLines = {
    business: result.business,
    proofLine: result.proofLine,
    invoices: result.invoices,
  };

  return (
    <Section id="de-mahnung-fee" title="Verzug und Mahnkosten">
      {arrange(arrangement, {
        stacked: () => stacked(lines),
        table: () => table(lines),
      })}
    </Section>
  );
};
