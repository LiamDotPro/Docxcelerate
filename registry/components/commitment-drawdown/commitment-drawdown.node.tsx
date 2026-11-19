import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import commitmentDrawdown from "../derivers/commitment-drawdown.deriver.ts";

/**
 * A usage invoice section that draws prepaid commitment and credit grants down
 * line by line, before tax, and says what is left, what is about to lapse, and
 * whether the term's minimum was met.
 *
 * Decisions already made for you:
 *
 * - Covered and overage are split per line, not on the total. Applying credit
 *   to the total loses which usage was charged, and puts tax on the wrong base.
 * - Whether a grant may be used is tested on the usage date, not the invoice
 *   date. June usage invoiced on 2 July still draws on a grant that expired on
 *   1 July.
 * - Burn order is by type (rollover, then paid commitment, then promotional),
 *   and within a type by priority (lower first), then earliest expiry, then
 *   earliest effective date, then earliest creation. The credit closest to
 *   expiring is used before it lapses, instead of being picked first-in-first-out.
 * - Promotional credit never counts toward a minimum commitment.
 * - The true-up runs only on the invoice that ends the term, and subtracts
 *   true-ups already billed.
 * - Credits about to expire are named, with the amount projected to go unused
 *   when a burn rate is known.
 * - The tax basis is decided from how the credit was sold. Credit taxed when
 *   bought (a prepayment or a single-purpose voucher): only the overage carries
 *   VAT here. A multi-purpose voucher: VAT falls on the full usage, and the
 *   credit is deducted after tax like a payment. This is the one choice that
 *   changes what the customer owes.
 *
 * All of it is one deriver, because every step is arithmetic, sorting or date
 * maths on request data, and publishing refuses all of those. The node computes
 * nothing: it prints the deriver's strings and branches on its flags, with
 * every test written inline so a publish build keeps each arm.
 *
 * The lines are six fixed slots and the grants four, rather than a `.map()`.
 * The resolver reads a published loop's collection from request data and
 * ignores the derived scope, so a `.map()` over the deriver's rows would
 * publish and then resolve to nothing.
 *
 * The node publishes 18 branches against the default `branchLimit` of 32.
 *
 * Installed by `dxcl add commitment-drawdown`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface CommitmentDrawdownData {
  drawdown: {
    /**
     * ISO 4217 code every money amount is in: `"GBP"`, `"USD"`, `"JPY"`. It also
     * sets how many minor units make one major unit. Required: the deriver throws
     * when it is missing or not a code `Intl` knows.
     */
    currency: string;
    /**
     * The invoice's issue date, ISO 8601. Used only for the expiry warnings and
     * to mark a grant as expired; never to decide whether a grant could be used.
     * Anything after the tenth character is ignored. Required: throws when
     * missing or not YYYY-MM-DD.
     */
    invoiceDate: string;
    /**
     * The usage period this invoice bills, ISO 8601, `end` inclusive. `end` is
     * the usage date of any line without one, and is compared with `term.end`
     * to decide whether this invoice closes the term. Required: throws when
     * either end is missing or not a date, or when `end` is before `start`.
     */
    period: { start: string; end: string };
    /**
     * The commitment term, when the contract has a minimum spend. Missing means
     * no minimum: no true-up, and no true-up line on any invoice.
     */
    term?: {
      /** First day of the term, ISO 8601. Required when `term` is given; printed in the true-up line. */
      start: string;
      /** Last day of the term, ISO 8601. Required when `term` is given. */
      end: string;
      /**
       * The minimum spend over the whole term, in integer minor units. Throws
       * when it is not a non-negative integer. Zero means no true-up.
       */
      commitment: number;
      /**
       * Eligible spend already billed on EARLIER invoices in this term, minor
       * units: usage charged plus usage covered by paid commitment or rollover,
       * never usage covered by promotional credit. Not including this invoice.
       * Missing means 0. Throws when present and not a non-negative integer.
       */
      eligibleSpendToDate?: number;
      /**
       * True-ups already billed against this term, minor units. Missing means 0.
       * Throws when present and not a non-negative integer.
       */
      trueUpsBilled?: number;
    };
    /**
     * The priced usage lines, in invoice order. Grants are drawn in this order.
     * Every line counts toward the minimum commitment, so leave out charges that
     * do not. Missing or empty means nothing to draw against.
     */
    lines?: Array<{
      /** Your reference for the line, echoed in the deriver's `allocation`. Missing means its index. */
      key?: string;
      /** What the line is, as printed. Required: throws when missing or blank. */
      description: string;
      /**
       * The day the usage happened, ISO 8601. For a line covering the whole
       * period, the last day of usage. Missing means `period.end`, so a grant
       * has to still be live on the last day to cover it. Never the invoice date.
       */
      usageDate?: string;
      /** How much was used, in the line's unit. Throws when it is not a finite number of 0 or more. */
      quantity: number;
      /**
       * Decimal places the quantity is counted in (0 for tokens, 1 for minutes
       * to the tenth). An overage quantity is rounded to this. Missing means the
       * number of decimals `quantity` was given with, at most 6.
       */
      quantityDecimals?: number;
      /** The unit, printed after an overage quantity: `"hours"`, `"tokens"`. Missing prints the bare number. */
      unit?: string;
      /**
       * Price of one unit in minor units, fractions allowed (0.0003 for $3 per
       * million tokens). Read only when a unit-denominated grant covers the
       * line, to price its overage. Throws then if it is missing or not a finite
       * number of 0 or more.
       */
      unitPrice?: number;
      /**
       * What the line costs before any credit, in integer minor units. This is
       * authoritative: it is not recomputed from quantity × unitPrice, so tiered
       * pricing is honoured. Zero draws nothing and is not printed. Throws when
       * it is not a non-negative integer: a refund or an adjustment is not usage.
       */
      amount: number;
      /** The price the line was billed at, matched against a grant's `priceIds`. Missing means only unscoped grants cover it. */
      priceId?: string;
    }>;
    /**
     * Every grant with a balance going into this invoice, in any order: they are
     * sorted into burn order. Missing or empty means nothing is drawn, and the
     * component prints nothing unless a true-up is due.
     */
    grants?: Array<{
      /** Your id for the grant, echoed in `balances`. Required. */
      id: string;
      /** What the grant is called, as printed: `"Annual prepaid commitment 2025–26"`. Required. */
      name: string;
      /**
       * `rollover`: unused paid commitment carried from an earlier term.
       * `commit`: paid commitment for this term. `promotional`: free credit.
       * Rollover burns first and promotional last. Promotional credit is a
       * discount, so it never counts toward the minimum and is always taken off
       * before tax. Required: anything else throws.
       */
      kind: "rollover" | "commit" | "promotional";
      /**
       * `money`: `balance` is integer minor units. `units`: `balance` is a count
       * of the lines' units (tokens, minutes). Required: anything else throws.
       */
      denomination: "money" | "units";
      /** Printed after a units balance: `"minutes"`. Missing prints `"units"`. Ignored for money. */
      unit?: string;
      /**
       * What is left going into this invoice. Money: a non-negative integer.
       * Units: a finite number of 0 or more. Throws otherwise. Zero is skipped.
       */
      balance: number;
      /** Burn order within its kind, lower first. Missing sorts after every numbered grant of its kind. */
      priority?: number;
      /** First day the grant may be used, ISO 8601, inclusive. Required: throws when missing or not a date. */
      effective: string;
      /**
       * The day the grant stops covering usage, ISO 8601. It covers usage dated
       * strictly before this day. Missing means it never expires: it sorts after
       * dated grants and never warns.
       */
      expires?: string;
      /** When the grant was created, ISO 8601. The last tie-breaker. Missing falls back to input order. */
      created?: string;
      /**
       * The prices this grant may cover. Missing or empty covers every line.
       * Required (non-empty) for a `units` grant: throws without it, because
       * minutes must not be spent on tokens.
       */
      priceIds?: string[];
      /**
       * Trailing average draw per day, in the grant's denomination (minor units
       * or units). Used only to project how much of an expiring grant goes
       * unused. Missing means the warning names the expiry without a projection.
       */
      dailyBurn?: number;
    }>;
    /**
     * How the paid credit (commit and rollover) was taxed.
     * `taxed-at-purchase`: VAT was charged when it was bought (a prepayment or
     * a single-purpose voucher), so only overage carries VAT here.
     * `multi-purpose`: a multi-purpose voucher, so VAT falls on the full usage
     * now and the credit is deducted after VAT like a payment. Required: throws
     * when missing or anything else. Guessing either way changes what the
     * customer owes.
     */
    creditTaxBasis: "taxed-at-purchase" | "multi-purpose";
    /**
     * How many days after `invoiceDate` an expiring grant is warned about,
     * inclusive. Missing means 30. Throws when present and not a non-negative
     * integer. 0 turns the warnings off.
     */
    warnDays?: number;
  };
}

/**
 * Prints the usage lines with what credit covered and what was charged, the
 * state of each grant, the tax basis and any true-up. Prints nothing when there
 * is nothing to say.
 */
export const CommitmentDrawdown: Nodes = async () => {
  const [drawdown] = useState((data: CommitmentDrawdownData) => data.drawdown);
  const d = await useDeriver(commitmentDrawdown, [drawdown]);

  return (
    <>
      {d.show && (
        <Section id="commitment-drawdown" title="Prepaid credit applied">
          <Paragraph id="commitment-drawdown-intro">
            Prepaid credit was applied to the usage for {d.periodLabel} line by line and before tax: rollover first, then
            prepaid commitment, then promotional credit.
          </Paragraph>
          {d.showLines && (
            <Table
              id="commitment-drawdown-lines"
              columns={[{ width: "auto" }, { width: 30, align: "right" }, { width: 30, align: "right" }, { width: 36, align: "right" }]}
            >
              <Row id="commitment-drawdown-lines-header" header>
                <Cell>Usage</Cell>
                <Cell>Amount</Cell>
                <Cell>Covered by credit</Cell>
                <Cell>Charged</Cell>
              </Row>
              {d.lines[0].description && (
                <Row id="commitment-drawdown-line-1">
                  <Cell>{d.lines[0].description}</Cell>
                  <Cell>{d.lines[0].usage}</Cell>
                  <Cell>{d.lines[0].covered}</Cell>
                  <Cell>{d.lines[0].charged}</Cell>
                </Row>
              )}
              {d.lines[1].description && (
                <Row id="commitment-drawdown-line-2">
                  <Cell>{d.lines[1].description}</Cell>
                  <Cell>{d.lines[1].usage}</Cell>
                  <Cell>{d.lines[1].covered}</Cell>
                  <Cell>{d.lines[1].charged}</Cell>
                </Row>
              )}
              {d.lines[2].description && (
                <Row id="commitment-drawdown-line-3">
                  <Cell>{d.lines[2].description}</Cell>
                  <Cell>{d.lines[2].usage}</Cell>
                  <Cell>{d.lines[2].covered}</Cell>
                  <Cell>{d.lines[2].charged}</Cell>
                </Row>
              )}
              {d.lines[3].description && (
                <Row id="commitment-drawdown-line-4">
                  <Cell>{d.lines[3].description}</Cell>
                  <Cell>{d.lines[3].usage}</Cell>
                  <Cell>{d.lines[3].covered}</Cell>
                  <Cell>{d.lines[3].charged}</Cell>
                </Row>
              )}
              {d.lines[4].description && (
                <Row id="commitment-drawdown-line-5">
                  <Cell>{d.lines[4].description}</Cell>
                  <Cell>{d.lines[4].usage}</Cell>
                  <Cell>{d.lines[4].covered}</Cell>
                  <Cell>{d.lines[4].charged}</Cell>
                </Row>
              )}
              {d.lines[5].description && (
                <Row id="commitment-drawdown-line-6">
                  <Cell>{d.lines[5].description}</Cell>
                  <Cell>{d.lines[5].usage}</Cell>
                  <Cell>{d.lines[5].covered}</Cell>
                  <Cell>{d.lines[5].charged}</Cell>
                </Row>
              )}
              <Row id="commitment-drawdown-lines-total">
                <Cell>Total</Cell>
                <Cell>{d.totals.usage}</Cell>
                <Cell>{d.totals.covered}</Cell>
                <Cell>{d.totals.charged}</Cell>
              </Row>
            </Table>
          )}
          {d.showGrants && (
            <Table
              id="commitment-drawdown-grants"
              columns={[{ width: "auto" }, { width: 34, align: "right" }, { width: 30, align: "right" }, { width: 50 }]}
            >
              <Row id="commitment-drawdown-grants-header" header>
                <Cell>Credit</Cell>
                <Cell>Applied</Cell>
                <Cell>Remaining</Cell>
                <Cell>Notes</Cell>
              </Row>
              {d.grants[0].name && (
                <Row id="commitment-drawdown-grant-1">
                  <Cell>{d.grants[0].name}</Cell>
                  <Cell>{d.grants[0].applied}</Cell>
                  <Cell>{d.grants[0].remaining}</Cell>
                  <Cell>{d.grants[0].note}</Cell>
                </Row>
              )}
              {d.grants[1].name && (
                <Row id="commitment-drawdown-grant-2">
                  <Cell>{d.grants[1].name}</Cell>
                  <Cell>{d.grants[1].applied}</Cell>
                  <Cell>{d.grants[1].remaining}</Cell>
                  <Cell>{d.grants[1].note}</Cell>
                </Row>
              )}
              {d.grants[2].name && (
                <Row id="commitment-drawdown-grant-3">
                  <Cell>{d.grants[2].name}</Cell>
                  <Cell>{d.grants[2].applied}</Cell>
                  <Cell>{d.grants[2].remaining}</Cell>
                  <Cell>{d.grants[2].note}</Cell>
                </Row>
              )}
              {d.grants[3].name && (
                <Row id="commitment-drawdown-grant-4">
                  <Cell>{d.grants[3].name}</Cell>
                  <Cell>{d.grants[3].applied}</Cell>
                  <Cell>{d.grants[3].remaining}</Cell>
                  <Cell>{d.grants[3].note}</Cell>
                </Row>
              )}
            </Table>
          )}
          {d.expiring && <Paragraph id="commitment-drawdown-expiring">{d.expiring}</Paragraph>}
          {d.taxArm === "purchase" && <Paragraph id="commitment-drawdown-tax-purchase">{d.taxNote}</Paragraph>}
          {d.taxArm === "redemption" && <Paragraph id="commitment-drawdown-tax-redemption">{d.taxNote}</Paragraph>}
          {d.trueUpArm === "shortfall" && <Paragraph id="commitment-drawdown-true-up">{d.trueUpNote}</Paragraph>}
          {d.trueUpArm === "met" && <Paragraph id="commitment-drawdown-commitment-met">{d.trueUpNote}</Paragraph>}
        </Section>
      )}
    </>
  );
};
