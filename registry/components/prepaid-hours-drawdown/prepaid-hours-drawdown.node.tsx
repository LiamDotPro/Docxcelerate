import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import prepaidHoursDrawdown from "../derivers/prepaid-hours-drawdown.deriver.ts";

/**
 * A body section for a retainer invoice: it applies the period's time entries
 * to the client's prepaid hour blocks and prints a block ledger, the overage,
 * any split entry, what rolled and what lapsed, and a warning when the balance
 * is low or exhausted or hours are about to expire.
 *
 * Decisions already made for you:
 *
 * - Hours are drawn from the block that expires soonest, not the oldest
 *   purchase. Picking the oldest lets hours lapse that could have been used.
 * - A block is tested on the entry's date, never the invoice date. Time from
 *   31 March, invoiced in April, still draws on a block that ended on 31 March.
 * - Each entry is rounded up to the billing increment before it draws anything.
 *   An entry that runs past the last hours it may use is then split at the exact
 *   minute: the covered part and the overage always add up to the billed entry,
 *   so nothing is dropped or billed twice.
 * - Rollover is capped (in hours, or as a percentage of the block). Rolled hours
 *   keep a shorter expiry when the contract says they do not extend, and they
 *   never roll twice.
 * - Hours above the cap are shown as forfeited with their value at the block
 *   rate. That value is never credited. Where VAT was charged on the purchase,
 *   the section says it is not adjusted, because the prepayment was taxed when
 *   it was received.
 * - Role restrictions are honoured: partner time cannot draw on an
 *   associate-only block and goes to overage.
 *
 * All of it is one deriver, because every step is sorting, date maths or
 * arithmetic on request data, and publishing refuses all of those. The node
 * computes nothing: it prints the deriver's strings and branches on its flags,
 * with every test written inline so a publish build keeps each arm.
 *
 * The ledger is four fixed slots rather than a `.map()`. The resolver reads a
 * published loop's collection from request data and ignores the derived scope,
 * so a `.map()` over the deriver's rows would publish and then resolve to
 * nothing.
 *
 * The node publishes 13 branches against the default `branchLimit` of 32.
 *
 * Known limit: with increments such as 1 or 10 minutes, a displayed
 * two-decimal figure can differ by 0.01 from the sum of the displayed figures
 * above it. The arithmetic is done in minutes and is exact.
 *
 * Installed by `dxcl add prepaid-hours-drawdown`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface PrepaidHoursDrawdownData {
  retainer: {
    /**
     * ISO 4217 code every rate and amount is in: `"GBP"`, `"USD"`. It also sets
     * how many minor units make one major unit. Required: the deriver throws
     * when it is missing or not a code `Intl` knows.
     */
    currency: string;
    /**
     * The last day of time this invoice bills, ISO 8601, inclusive. A block
     * whose expiry is on or before this day is closed by this invoice (rolled
     * over or forfeited). Expiry warnings count from this day. There is no
     * invoice date on purpose: nothing here may be decided by it. Anything
     * after the tenth character is ignored. Required: throws when missing or
     * not YYYY-MM-DD.
     */
    periodEnd: string;
    /**
     * The billing increment in minutes: 6 for tenths of an hour, 15 for
     * quarters. Every entry is rounded UP to a multiple of it before it draws
     * hours. Required: throws when missing or not a positive integer. Use 1
     * for no rounding.
     */
    increment: number;
    /**
     * The engagement's overage rate per hour, in integer minor units. When it
     * is given, all overage is billed at this rate. Missing means each
     * overage entry is billed at its own `standardRate`. Throws when present
     * and not a non-negative integer.
     */
    overageRate?: number;
    /**
     * The low-balance threshold, in hours. If the hours left after this
     * period in blocks still open are more than 0 and no more than this, the
     * low-balance warning prints. Missing means no low-balance warning. The
     * exhausted warning prints regardless. Throws when present and not a
     * finite number of 0 or more.
     */
    lowBalanceHours?: number;
    /**
     * How many days after `periodEnd` a block's expiry is warned about,
     * inclusive. Missing means 30. 0 turns the warning off. Throws when
     * present and not a non-negative integer.
     */
    expiringDays?: number;
    /**
     * Whether VAT (or GST) was charged when the blocks were bought. True adds
     * the sentence that the VAT on forfeited hours is not adjusted. Missing
     * or false leaves VAT unmentioned, which is right where no VAT applies.
     */
    vatOnPurchase?: boolean;
    /**
     * Every block not already closed on an earlier invoice, in any order.
     * They are sorted into burn order. A block whose expiry is on or before
     * `periodEnd` is closed by this invoice, so leave out any block an earlier
     * invoice already closed, or it will be closed twice. Hours rolled in on
     * an earlier invoice that lapse on a different date from their block are
     * passed as a block of their own. Missing or empty means the component
     * prints nothing.
     */
    blocks?: Array<{
      /** Your id for the block. Referenced by `rolledInFrom` and echoed in the output. Required, unique: throws when missing, blank or duplicated. */
      id: string;
      /** What the block is called, as printed: `"Q2 2026 retainer"`. Missing prints `id`. */
      name?: string;
      /**
       * The block's size in hours, as bought. A percentage rollover cap is
       * taken from this. Required: throws when not a finite number of 0 or
       * more, or not a whole number of minutes (hours x 60 must be an integer).
       */
      hours: number;
      /**
       * Hours left going into this invoice, including any rolled in on an
       * earlier invoice, so it may be more than `hours`. Missing means `hours`
       * (a block new this period). Validated the same way as `hours`.
       */
      opening?: number;
      /**
       * What one prepaid hour cost, in integer minor units. Used only to value
       * forfeited hours. It is never billed. Required: throws when it is not a
       * non-negative integer.
       */
      rate: number;
      /** First day the block covers, ISO 8601, inclusive. Required: throws when missing or not a date. */
      start: string;
      /**
       * Last day the block covers, ISO 8601, inclusive: the end of its term.
       * Required: throws when missing, not a date, or before `start`.
       */
      expiry: string;
      /**
       * How many unused hours may roll into the block that names this one in
       * `rolledInFrom`: exactly one of `hours` (a fixed number) or `percent`
       * (0 to 100, of `hours`, rounded down to whole minutes). Missing means
       * nothing rolls and every unused hour is forfeited at expiry. Throws
       * when both or neither field is set, or a value is out of range. Throws
       * when this block closes this invoice with a cap and no block names it
       * in `rolledInFrom`: printing a forfeiture because the successor block
       * was left out of the data would be wrong.
       */
      rolloverCap?: { hours?: number; percent?: number };
      /**
       * For contracts where rolled hours do not extend: the day the hours
       * rolled OUT of this block lapse in the next term, ISO 8601, inclusive.
       * Missing means rolled hours run to the receiving block's `expiry`.
       * Throws when it is before the receiving block's `start` or after its
       * `expiry`.
       */
      rolledExpiry?: string;
      /**
       * The roles this block may cover, compared exactly with an entry's
       * `role`. Missing or empty covers every role. An entry with no role can
       * only draw on unrestricted blocks.
       */
      allowedRoles?: string[];
      /**
       * The id of the block whose unused hours roll into this one. Throws when
       * it names no block or itself, or when two blocks name the same one.
       */
      rolledInFrom?: string;
    }>;
    /**
     * The time recorded, in any order. Entries are sorted by date, and
     * entries on the same date keep their input order. Fields not listed
     * (a narrative, an id) are ignored. Missing or empty means nothing is
     * drawn.
     */
    entries?: Array<{
      /** The day the work was done, ISO 8601. Required: throws when missing, not a date, or after `periodEnd`. A late entry from an earlier period is allowed. */
      date: string;
      /** Who did the work, as printed in the split note. Required: throws when missing or blank. */
      timekeeper: string;
      /** The timekeeper's role, matched against `allowedRoles`: `"partner"`. Missing matches only unrestricted blocks. */
      role?: string;
      /** Time recorded, in minutes, before rounding. Required: throws when it is not a non-negative integer. 0 draws nothing. */
      minutes: number;
      /**
       * The timekeeper's standard rate per hour, in integer minor units. Read
       * only when this entry has overage and `overageRate` is missing; it
       * throws then if it is absent or not a non-negative integer.
       */
      standardRate?: number;
    }>;
  };
}

/**
 * Prints the block ledger, the overage, any split entry, what rolled and what
 * lapsed, and the balance and expiry warnings. Prints nothing when there are no
 * blocks.
 */
export const PrepaidHoursDrawdown: Nodes = async () => {
  const [retainer] = useState((data: PrepaidHoursDrawdownData) => data.retainer);
  const d = await useDeriver(prepaidHoursDrawdown, [retainer]);

  return (
    <>
      {d.show && (
        <Section id="prepaid-hours" title="Prepaid hours">
          <Paragraph id="prepaid-hours-intro">
            Time recorded up to {d.periodEnd} was applied to prepaid hours in date order, drawing first on the hours that
            expire soonest. Each entry was rounded up to the next {d.increment} before it was applied.
          </Paragraph>
          <Table
            id="prepaid-hours-ledger"
            columns={[{ width: "auto" }, { width: 19, align: "right" }, { width: 19, align: "right" }, { width: 19, align: "right" }, { width: 19, align: "right" }, { width: 19, align: "right" }, { width: 34 }]}
          >
            <Row id="prepaid-hours-ledger-header" header>
              <Cell>Block</Cell>
              <Cell>Opening</Cell>
              <Cell>Used</Cell>
              <Cell>Rolled over</Cell>
              <Cell>Forfeited</Cell>
              <Cell>Closing</Cell>
              <Cell>Expires</Cell>
            </Row>
            {d.blocks[0].name && (
              <Row id="prepaid-hours-block-1">
                <Cell>{d.blocks[0].name}</Cell>
                <Cell>{d.blocks[0].opening}</Cell>
                <Cell>{d.blocks[0].used}</Cell>
                <Cell>{d.blocks[0].rolled}</Cell>
                <Cell>{d.blocks[0].forfeited}</Cell>
                <Cell>{d.blocks[0].closing}</Cell>
                <Cell>{d.blocks[0].expires}</Cell>
              </Row>
            )}
            {d.blocks[1].name && (
              <Row id="prepaid-hours-block-2">
                <Cell>{d.blocks[1].name}</Cell>
                <Cell>{d.blocks[1].opening}</Cell>
                <Cell>{d.blocks[1].used}</Cell>
                <Cell>{d.blocks[1].rolled}</Cell>
                <Cell>{d.blocks[1].forfeited}</Cell>
                <Cell>{d.blocks[1].closing}</Cell>
                <Cell>{d.blocks[1].expires}</Cell>
              </Row>
            )}
            {d.blocks[2].name && (
              <Row id="prepaid-hours-block-3">
                <Cell>{d.blocks[2].name}</Cell>
                <Cell>{d.blocks[2].opening}</Cell>
                <Cell>{d.blocks[2].used}</Cell>
                <Cell>{d.blocks[2].rolled}</Cell>
                <Cell>{d.blocks[2].forfeited}</Cell>
                <Cell>{d.blocks[2].closing}</Cell>
                <Cell>{d.blocks[2].expires}</Cell>
              </Row>
            )}
            {d.blocks[3].name && (
              <Row id="prepaid-hours-block-4">
                <Cell>{d.blocks[3].name}</Cell>
                <Cell>{d.blocks[3].opening}</Cell>
                <Cell>{d.blocks[3].used}</Cell>
                <Cell>{d.blocks[3].rolled}</Cell>
                <Cell>{d.blocks[3].forfeited}</Cell>
                <Cell>{d.blocks[3].closing}</Cell>
                <Cell>{d.blocks[3].expires}</Cell>
              </Row>
            )}
            {d.showTotal && (
              <Row id="prepaid-hours-ledger-total">
                <Cell>Total</Cell>
                <Cell>{d.total.opening}</Cell>
                <Cell>{d.total.used}</Cell>
                <Cell>{d.total.rolled}</Cell>
                <Cell>{d.total.forfeited}</Cell>
                <Cell>{d.total.closing}</Cell>
                <Cell>{""}</Cell>
              </Row>
            )}
          </Table>
          {d.overageNote && <Paragraph id="prepaid-hours-overage">{d.overageNote}</Paragraph>}
          {d.splitNote && <Paragraph id="prepaid-hours-split">{d.splitNote}</Paragraph>}
          {d.rolledNote && <Paragraph id="prepaid-hours-rolled">{d.rolledNote}</Paragraph>}
          {d.forfeitedNote && <Paragraph id="prepaid-hours-forfeited">{d.forfeitedNote}</Paragraph>}
          {d.balanceArm === "exhausted" && <Paragraph id="prepaid-hours-exhausted">{d.balanceNote}</Paragraph>}
          {d.balanceArm === "low" && <Paragraph id="prepaid-hours-low">{d.balanceNote}</Paragraph>}
          {d.expiringNote && <Paragraph id="prepaid-hours-expiring">{d.expiringNote}</Paragraph>}
        </Section>
      )}
    </>
  );
};
