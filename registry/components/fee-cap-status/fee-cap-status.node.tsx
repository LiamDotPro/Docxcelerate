import { arrange, Paragraph, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { FeeCapStatusLines } from "../arrangements/fee-cap-status/lines.ts";
import { stacked } from "../arrangements/fee-cap-status/stacked.arrangement.tsx";
import { table } from "../arrangements/fee-cap-status/table.arrangement.tsx";
import feeCapStatus from "../derivers/fee-cap-status.deriver.ts";

/**
 * What an invoice under a fee limit may charge, and what it has to say about
 * it. Those are decisions about money and about what a client was told, not
 * about layout.
 *
 * Hard cap and ceiling: the invoice charges only what is left under the cap.
 * The excess is printed as written off, so the client sees the value given up
 * and nobody reads the bill as the whole value of the work. A cap already used
 * up bills the period at nil, never below it.
 *
 * Ceiling: once fees reach the notice threshold (85% unless told otherwise,
 * following FAR 52.232-7) the invoice prints the percentage used and the
 * notice. The percentage is floored, so a figure just under 100% never prints
 * as 100%.
 *
 * Estimate: not a limit, so fees are charged in full. Once they pass the
 * estimate the invoice must print the revised estimate and the date the client
 * was told. Without a revised estimate it prints an "incomplete" warning, at
 * the top, because an overrun the client was not told about can be challenged:
 * the SRA's position in England and Wales, and in Ireland LSRA 2015 s.150
 * requires a fresh notice when costs will be significantly greater. A draft
 * that cannot go out says so before anything else.
 *
 * Withholding: the agreed percentage of what the invoice charges is retained
 * until total retention reaches its limit, and after that nothing is.
 *
 * Four states, each with its own id: under, approaching, reached, exceeded.
 * The whole decision is one deriver, because every part of it is arithmetic on
 * request data, and that throws at publish time. The node prints what the
 * deriver returns, so nothing is computed on request data here.
 *
 * Arrangement: where the figures sit is chosen with the `arrangement` prop.
 * This file decides every line; each arrangement, in
 * `arrangements/fee-cap-status/`, only places them, so both print the same
 * rows under the same conditions.
 *
 * A publish build of this node emits 13 conditions against the default
 * `branchLimit` of 32: 8 in the node and 5 in the chosen arrangement. The
 * state arms are not collapsed into one paragraph, because the per-arm id is
 * what records which state a recipient got.
 *
 * Installed by `dxcl add fee-cap-status`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface FeeCapStatusData {
  feeCap: {
    /**
     * What kind of limit the engagement has.
     *
     * `hard` is a cap or not-to-exceed: nothing above it is charged.
     * `ceiling` is a hard cap with a notice duty as fees approach it (FAR
     * 52.232-7): it clips like `hard`, and past `noticeThreshold` prints the
     * notice. `estimate` is not a limit: fees are charged in full, and an
     * overrun has to have been told to the client. Required: anything else
     * throws, because a cap read as an estimate overcharges, and an estimate
     * read as a cap writes off fees that were owed.
     */
    kind: "hard" | "estimate" | "ceiling";
    /**
     * The cap, ceiling or original estimate, in major units of `currency`.
     * Required, finite and greater than 0: the deriver throws otherwise,
     * because every percentage divides by it.
     */
    amount: number;
    /**
     * Fees charged on earlier invoices under this limit, before any
     * withholding, not counting this invoice. Required, finite and 0 or more:
     * the deriver throws otherwise. A figure above `amount` on a hard cap or
     * ceiling is read as "cap used up", not as an error.
     */
    billedToDate: number;
    /**
     * The percentage of the limit at which the reader is warned that fees are
     * approaching it. A number greater than 0 and less than 100, for example `85`.
     * Missing means 85 for `ceiling` and no "approaching" state for `hard` and
     * `estimate`. Out of range or not finite throws.
     */
    noticeThreshold?: number;
    /**
     * For `estimate` only: the revised estimate and the day the client was told.
     * When given, it replaces `amount` as the estimate fees are measured
     * against, and the revision is printed. Missing means fees are measured
     * against `amount`, and an overrun prints the "incomplete" warning.
     * Ignored for `hard` and `ceiling`: a cap is not raised by notice.
     */
    revisedEstimate?: {
      /** The revised estimate, in major units. Finite and greater than 0, or it throws. */
      amount: number;
      /** When the client was told, as an ISO 8601 date (`2026-03-03`). Throws when not one. */
      notifiedOn: string;
    };
    /**
     * A retention the agreement allows: a percentage of each invoice held back
     * until the total held reaches a limit. Missing means nothing is withheld.
     */
    withholding?: {
      /** Percentage of what this invoice charges, greater than 0 and at most 100. Throws otherwise. */
      percent: number;
      /** The most that may be held in total, in major units. Greater than 0, or it throws. */
      limit: number;
      /**
       * Held back on earlier invoices. Required, 0 or more: throws when
       * missing, because guessing 0 would keep withholding after the limit.
       */
      withheldToDate: number;
    };
    /** ISO 4217 code. Defaults to GBP. A code Intl does not know throws. */
    currency?: string;
  };
  fees: {
    /**
     * Fees for the work in this invoice's period, before the cap and before
     * withholding, in major units. Required, finite and 0 or more: a credit
     * throws, because a credit against a capped engagement is a credit note,
     * not this.
     */
    thisPeriod: number;
  };
}

/**
 * Where the figures sit.
 *
 * - `table`: labels in a column, amounts right-aligned beside them, so the
 *   reader can check cap, billed and charged against each other. The default.
 * - `stacked`: one labelled figure under another, down the text column.
 */
export type FeeCapStatusArrangement = "table" | "stacked";

/** Props for {@linkcode FeeCapStatus}. */
export interface FeeCapStatusProps {
  /** Where the figures sit. Defaults to `"table"`. */
  arrangement?: FeeCapStatusArrangement;
}

export const FeeCapStatus: SectionComponent<FeeCapStatusProps> = async ({ arrangement }) => {
  const [feeCap] = useState((data: FeeCapStatusData) => data.feeCap);
  const [fees] = useState((data: FeeCapStatusData) => data.fees);
  const fee = await useDeriver(feeCapStatus, [feeCap, fees]);

  const lines: FeeCapStatusLines = {
    capLabel: fee.capLabel,
    capValue: fee.capValue,
    billedToDate: fee.billedToDate,
    thisPeriod: fee.thisPeriod,
    writtenOffLabel: fee.writtenOffLabel,
    writtenOff: fee.writtenOff,
    charged: fee.charged,
    withheldLabel: fee.withheldLabel,
    withheld: fee.withheld,
    payable: fee.payable,
    remainingLabel: fee.remainingLabel,
    remaining: fee.remaining,
  };

  return (
    <Section id="fee-cap" title={fee.heading}>
      {fee.incomplete && <Paragraph id="fee-cap-incomplete">{fee.incomplete}</Paragraph>}
      {fee.state === "under" && <Paragraph id="fee-cap-under">{fee.status}</Paragraph>}
      {fee.state === "approaching" && <Paragraph id="fee-cap-approaching">{fee.status}</Paragraph>}
      {fee.state === "reached" && <Paragraph id="fee-cap-reached">{fee.status}</Paragraph>}
      {fee.state === "exceeded" && <Paragraph id="fee-cap-exceeded">{fee.status}</Paragraph>}
      {arrange(arrangement, { table: () => table(lines), stacked: () => stacked(lines) })}
      {fee.revised && <Paragraph id="fee-cap-revised">{fee.revised}</Paragraph>}
      {fee.notice && <Paragraph id="fee-cap-notice">{fee.notice}</Paragraph>}
      {fee.retention && <Paragraph id="fee-cap-retention">{fee.retention}</Paragraph>}
    </Section>
  );
};
