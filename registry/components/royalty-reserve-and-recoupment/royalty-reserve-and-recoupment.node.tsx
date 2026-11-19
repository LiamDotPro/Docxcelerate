import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import royaltyReserveAndRecoupment from "../derivers/royalty-reserve-and-recoupment.deriver.ts";

/**
 * A royalty statement's closing section: it takes the period's earned
 * royalties and works down, line by line, to the amount actually paid.
 *
 * Decisions already made for you:
 *
 * - A reserve against returns is held on physical formats only. E-books and
 *   digital audio cannot be returned to a warehouse, so they carry none. The
 *   reserve is rounded once per title, so the rows add up to the total.
 * - Returns are charged against the reserves held for that title, oldest first,
 *   not against this period's earnings. Only returns beyond them reduce this
 *   period's royalties.
 * - A reserve held for the contract's liquidation limit is released in full,
 *   after returns are charged against it.
 * - The statement shows its workings: what was withheld, and how each earlier
 *   reserve was used or released.
 * - Recoupment pools only the titles an advance cross-collateralises. Every
 *   other title recoups on its own, so one title's earnings never pay another's
 *   advance.
 * - An advance is still unrecouped, earned out this period or already earned
 *   out, each in its own words. An unrecouped advance is never printed as a
 *   negative balance or as a debt: it is recovered from royalties, never repaid.
 * - A total below the minimum payment is carried forward and the statement says
 *   so. It is never silently zeroed.
 * - Every amount is a whole number of minor units.
 *
 * All of it is decided in one deriver, because nearly every step is arithmetic,
 * sorting or filtering on request data, and publishing refuses all of those.
 * The node computes nothing: it prints the deriver's result into fixed slots
 * and branches on its outcome, with every test written inline so a publish
 * build keeps each arm.
 *
 * The rows are fixed slots rather than a `.map()`. The resolver reads a
 * published loop's collection from request data and ignores the derived scope,
 * so a `.map()` over the deriver's rows would publish and then resolve to none.
 *
 * The node publishes 18 branches against the default `branchLimit` of 32: three
 * outcomes, two settlement rows, the reserves table, six reserve rows, the
 * note and five advance rows. A document that also uses plan-change-proration
 * (14) needs `branchLimit` raised.
 *
 * Installed by `dxcl add royalty-reserve-and-recoupment`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface RoyaltyReserveAndRecoupmentData {
  royalty: {
    /**
     * ISO 4217 code every amount is in: `"USD"`, `"GBP"`, `"EUR"`. It also sets
     * how many minor units make one major unit. Required: the deriver throws
     * when it is missing or not a code `Intl` knows.
     */
    currency: string;
    /**
     * This statement's period number, counted the same way as `setInPeriod`
     * on every open reserve (1 for the first statement, 2 for the next...). It is
     * what a reserve's age is measured against. Required: throws when it is not
     * an integer of 1 or more.
     */
    periodIndex: number;
    /**
     * How the period is named in the opening line: `"1 January to 30 June 2026"`.
     * Missing means `"period <periodIndex>"`.
     */
    periodLabel?: string;
    /**
     * What each title earned this period, before any reserve or return, in
     * integer minor units. Two entries for the same title are added together.
     * Missing or empty means nothing was earned this period.
     */
    earned?: Array<{
      /** The title, as printed. Matched exactly against every other `title`. Required: throws when blank. */
      title: string;
      /**
       * Royalties on physical formats (hardback, paperback, audio CD): the only
       * royalties a reserve is held on. Missing means 0. Throws when not a
       * non-negative integer.
       */
      physicalMinor?: number;
      /**
       * Royalties on e-books and digital audio. No reserve is ever held on these.
       * Missing means 0. Throws when not a non-negative integer.
       */
      digitalMinor?: number;
    }>;
    /**
     * The share of this period's physical royalties held back against returns,
     * as a percentage: `25` for 25%. At most two decimals count (it is rounded to
     * basis points). Missing means 0: no new reserve. Throws when not a finite
     * number from 0 to 100.
     */
    reservePercent?: number;
    /**
     * How many periods a reserve may be held before it must be released:
     * `4` releases a reserve set in period 5 on the period-9 statement. Required
     * whenever any reserve exists (an open reserve, or a `reservePercent` above
     * 0 with physical royalties): throws then when it is missing or not an
     * integer of 1 or more, because a reserve with no release date is exactly
     * the mistake this component exists to prevent. Ignored when no reserve exists.
     */
    reserveLiquidationPeriods?: number;
    /**
     * Reserves held on earlier statements and not yet used or released, one
     * entry per title per period that set it. Missing or empty means none.
     */
    openReserves?: Array<{
      /** The title the reserve was held on. Required: throws when blank. */
      title: string;
      /**
       * The `periodIndex` of the statement that set it. Required: throws when
       * not an integer, or when it is not before `periodIndex` (this period's
       * reserve is the one this component computes, so it is never "open").
       */
      setInPeriod: number;
      /** What is still held, in integer minor units. Required: throws when not a non-negative integer. */
      amountMinor: number;
      /**
       * How the reserve's period is named in the note: `"the statement to 30 June 2024"`.
       * Missing means `"period <setInPeriod>"`.
       */
      label?: string;
    }>;
    /**
     * Royalties on copies returned this period, by title, in integer minor
     * units. Entries for the same title are added together. Missing or empty means no returns.
     */
    returnsCharged?: Array<{
      /** The title the returned copies belong to. Required: throws when blank. */
      title: string;
      /** Required: throws when not a non-negative integer. */
      amountMinor: number;
    }>;
    /**
     * One entry per advance: the titles it covers and what of it was unearned at
     * the start of this period. Titles in one entry are pooled (joint accounting,
     * cross-collateralised). A title in no entry has no advance and settles on
     * its own. Missing or empty means no title has an advance.
     */
    advances?: Array<{
      /**
       * The titles this advance covers, matched exactly. Required: throws when
       * empty, when an entry is blank, or when a title appears in two advances.
       */
      titles: string[];
      /**
       * The unearned part of the advance at the start of this period, in integer
       * minor units. 0 means already earned out. Required: throws when not a
       * non-negative integer.
       */
      unrecoupedOpeningMinor: number;
    }>;
    /**
     * The smallest amount the publisher pays out, in integer minor units. A total
     * due below it is carried forward instead of paid. Missing means no threshold:
     * any amount above 0 is paid. Throws when present and not a non-negative integer.
     */
    minimumPaymentMinor?: number;
    /**
     * An amount carried forward from the last statement because it was under the
     * minimum, in integer minor units. It is added to this period's total due.
     * Missing means 0. Throws when present and not a non-negative integer.
     */
    carriedForwardMinor?: number;
  };
}

/**
 * Prints the settlement from royalties earned to the amount paid, the reserve
 * ledger with its note, and each advance's position. The section always prints.
 */
export const RoyaltyReserveAndRecoupment: Nodes = async () => {
  const [royalty] = useState((data: RoyaltyReserveAndRecoupmentData) => data.royalty);
  const r = await useDeriver(royaltyReserveAndRecoupment, [royalty]);

  return (
    <Section id="royalty-reserve-and-recoupment" title="Royalties, reserves and advances">
      <Paragraph id="royalty-intro">{r.intro}</Paragraph>
      {r.outcome === "payable" && (
        <Paragraph id="royalty-payable">{r.payableNow} is payable to you for this period.</Paragraph>
      )}
      {r.outcome === "below-minimum" && (
        <Paragraph id="royalty-below-minimum">
          The {r.totalDue} due for this period is below the minimum payment of {r.minimum}, so it is carried forward to
          your next statement rather than paid now. It is not lost: it is added to what is due then.
        </Paragraph>
      )}
      {r.outcome === "nothing" && <Paragraph id="royalty-nothing-payable">Nothing is payable for this period.</Paragraph>}
      <Table id="royalty-settlement" columns={[{ width: "auto" }, { width: 34, align: "right" }]}>
        <Row id="royalty-settlement-header" header>
          <Cell>This period</Cell>
          <Cell>Amount</Cell>
        </Row>
        <Row id="royalty-settlement-physical">
          <Cell>Royalties earned on physical formats</Cell>
          <Cell>{r.physical}</Cell>
        </Row>
        <Row id="royalty-settlement-digital">
          <Cell>Royalties earned on e-books and digital audio</Cell>
          <Cell>{r.digital}</Cell>
        </Row>
        <Row id="royalty-settlement-reserve-held">
          <Cell>{r.heldLabel}</Cell>
          <Cell>{r.held}</Cell>
        </Row>
        <Row id="royalty-settlement-reserve-released">
          <Cell>Earlier reserves released</Cell>
          <Cell>{r.released}</Cell>
        </Row>
        <Row id="royalty-settlement-returns-beyond">
          <Cell>Returns beyond the reserves held for them</Cell>
          <Cell>{r.excess}</Cell>
        </Row>
        <Row id="royalty-settlement-net">
          <Cell>Net royalties this period</Cell>
          <Cell>{r.net}</Cell>
        </Row>
        <Row id="royalty-settlement-advances">
          <Cell>Set against unearned balances</Cell>
          <Cell>{r.advanceAdjustment}</Cell>
        </Row>
        {r.broughtForward && (
          <Row id="royalty-settlement-brought-forward">
            <Cell>Brought forward from your last statement</Cell>
            <Cell>{r.broughtForward}</Cell>
          </Row>
        )}
        <Row id="royalty-settlement-due">
          <Cell>Due for this period</Cell>
          <Cell>{r.totalDue}</Cell>
        </Row>
        {r.carriedForward && (
          <Row id="royalty-settlement-carried-forward">
            <Cell>Carried forward to your next statement</Cell>
            <Cell>{r.carriedForward}</Cell>
          </Row>
        )}
        <Row id="royalty-settlement-payable">
          <Cell>Payable now</Cell>
          <Cell>{r.payableNow}</Cell>
        </Row>
      </Table>
      {r.showReserves && (
        <Table
          id="royalty-reserves"
          columns={[
            { width: "auto" },
            { width: 24, align: "right" },
            { width: 24, align: "right" },
            { width: 24, align: "right" },
            { width: 24, align: "right" },
            { width: 24, align: "right" },
          ]}
        >
          <Row id="royalty-reserves-header" header>
            <Cell>Title</Cell>
            <Cell>Held at start</Cell>
            <Cell>Returns charged</Cell>
            <Cell>Released</Cell>
            <Cell>Held this period</Cell>
            <Cell>Held at end</Cell>
          </Row>
          {r.reserveRows[0].title && (
            <Row id="royalty-reserve-1">
              <Cell>{r.reserveRows[0].title}</Cell>
              <Cell>{r.reserveRows[0].start}</Cell>
              <Cell>{r.reserveRows[0].returns}</Cell>
              <Cell>{r.reserveRows[0].released}</Cell>
              <Cell>{r.reserveRows[0].held}</Cell>
              <Cell>{r.reserveRows[0].end}</Cell>
            </Row>
          )}
          {r.reserveRows[1].title && (
            <Row id="royalty-reserve-2">
              <Cell>{r.reserveRows[1].title}</Cell>
              <Cell>{r.reserveRows[1].start}</Cell>
              <Cell>{r.reserveRows[1].returns}</Cell>
              <Cell>{r.reserveRows[1].released}</Cell>
              <Cell>{r.reserveRows[1].held}</Cell>
              <Cell>{r.reserveRows[1].end}</Cell>
            </Row>
          )}
          {r.reserveRows[2].title && (
            <Row id="royalty-reserve-3">
              <Cell>{r.reserveRows[2].title}</Cell>
              <Cell>{r.reserveRows[2].start}</Cell>
              <Cell>{r.reserveRows[2].returns}</Cell>
              <Cell>{r.reserveRows[2].released}</Cell>
              <Cell>{r.reserveRows[2].held}</Cell>
              <Cell>{r.reserveRows[2].end}</Cell>
            </Row>
          )}
          {r.reserveRows[3].title && (
            <Row id="royalty-reserve-4">
              <Cell>{r.reserveRows[3].title}</Cell>
              <Cell>{r.reserveRows[3].start}</Cell>
              <Cell>{r.reserveRows[3].returns}</Cell>
              <Cell>{r.reserveRows[3].released}</Cell>
              <Cell>{r.reserveRows[3].held}</Cell>
              <Cell>{r.reserveRows[3].end}</Cell>
            </Row>
          )}
          {r.reserveRows[4].title && (
            <Row id="royalty-reserve-5">
              <Cell>{r.reserveRows[4].title}</Cell>
              <Cell>{r.reserveRows[4].start}</Cell>
              <Cell>{r.reserveRows[4].returns}</Cell>
              <Cell>{r.reserveRows[4].released}</Cell>
              <Cell>{r.reserveRows[4].held}</Cell>
              <Cell>{r.reserveRows[4].end}</Cell>
            </Row>
          )}
          {r.reserveRows[5].title && (
            <Row id="royalty-reserve-6">
              <Cell>{r.reserveRows[5].title}</Cell>
              <Cell>{r.reserveRows[5].start}</Cell>
              <Cell>{r.reserveRows[5].returns}</Cell>
              <Cell>{r.reserveRows[5].released}</Cell>
              <Cell>{r.reserveRows[5].held}</Cell>
              <Cell>{r.reserveRows[5].end}</Cell>
            </Row>
          )}
          <Row id="royalty-reserves-total">
            <Cell>All titles</Cell>
            <Cell>{r.reserveTotal.start}</Cell>
            <Cell>{r.reserveTotal.returns}</Cell>
            <Cell>{r.reserveTotal.released}</Cell>
            <Cell>{r.reserveTotal.held}</Cell>
            <Cell>{r.reserveTotal.end}</Cell>
          </Row>
        </Table>
      )}
      {r.reserveNote && <Paragraph id="royalty-reserves-note">{r.reserveNote}</Paragraph>}
      <Table
        id="royalty-advances"
        columns={[
          { width: "auto" },
          { width: 26, align: "right" },
          { width: 26, align: "right" },
          { width: 26, align: "right" },
          { width: 26, align: "right" },
        ]}
      >
        <Row id="royalty-advances-header" header>
          <Cell>Titles</Cell>
          <Cell>Unearned at start</Cell>
          <Cell>Net royalties</Cell>
          <Cell>Unearned at end</Cell>
          <Cell>Payable</Cell>
        </Row>
          <Row id="royalty-advance-1">
              <Cell>
                <Paragraph id="royalty-advance-1-titles">{r.units[0].titles}</Paragraph>
                <Paragraph id="royalty-advance-1-status">{r.units[0].status}</Paragraph>
              </Cell>
              <Cell>{r.units[0].opening}</Cell>
              <Cell>{r.units[0].net}</Cell>
              <Cell>{r.units[0].closing}</Cell>
              <Cell>{r.units[0].payable}</Cell>
            </Row>
          {r.units[1].titles && (
            <Row id="royalty-advance-2">
              <Cell>
                <Paragraph id="royalty-advance-2-titles">{r.units[1].titles}</Paragraph>
                <Paragraph id="royalty-advance-2-status">{r.units[1].status}</Paragraph>
              </Cell>
              <Cell>{r.units[1].opening}</Cell>
              <Cell>{r.units[1].net}</Cell>
              <Cell>{r.units[1].closing}</Cell>
              <Cell>{r.units[1].payable}</Cell>
            </Row>
          )}
          {r.units[2].titles && (
            <Row id="royalty-advance-3">
              <Cell>
                <Paragraph id="royalty-advance-3-titles">{r.units[2].titles}</Paragraph>
                <Paragraph id="royalty-advance-3-status">{r.units[2].status}</Paragraph>
              </Cell>
              <Cell>{r.units[2].opening}</Cell>
              <Cell>{r.units[2].net}</Cell>
              <Cell>{r.units[2].closing}</Cell>
              <Cell>{r.units[2].payable}</Cell>
            </Row>
          )}
          {r.units[3].titles && (
            <Row id="royalty-advance-4">
              <Cell>
                <Paragraph id="royalty-advance-4-titles">{r.units[3].titles}</Paragraph>
                <Paragraph id="royalty-advance-4-status">{r.units[3].status}</Paragraph>
              </Cell>
              <Cell>{r.units[3].opening}</Cell>
              <Cell>{r.units[3].net}</Cell>
              <Cell>{r.units[3].closing}</Cell>
              <Cell>{r.units[3].payable}</Cell>
            </Row>
          )}
          {r.units[4].titles && (
            <Row id="royalty-advance-5">
              <Cell>
                <Paragraph id="royalty-advance-5-titles">{r.units[4].titles}</Paragraph>
                <Paragraph id="royalty-advance-5-status">{r.units[4].status}</Paragraph>
              </Cell>
              <Cell>{r.units[4].opening}</Cell>
              <Cell>{r.units[4].net}</Cell>
              <Cell>{r.units[4].closing}</Cell>
              <Cell>{r.units[4].payable}</Cell>
            </Row>
          )}
          {r.units[5].titles && (
            <Row id="royalty-advance-6">
              <Cell>
                <Paragraph id="royalty-advance-6-titles">{r.units[5].titles}</Paragraph>
                <Paragraph id="royalty-advance-6-status">{r.units[5].status}</Paragraph>
              </Cell>
              <Cell>{r.units[5].opening}</Cell>
              <Cell>{r.units[5].net}</Cell>
              <Cell>{r.units[5].closing}</Cell>
              <Cell>{r.units[5].payable}</Cell>
            </Row>
          )}
      </Table>
    </Section>
  );
};
