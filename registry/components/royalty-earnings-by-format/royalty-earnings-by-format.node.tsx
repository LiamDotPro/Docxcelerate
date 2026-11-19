import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import royaltyEarningsByFormat from "../derivers/royalty-earnings-by-format.deriver.ts";

/**
 * An author's royalty statement section: one period's unit sales, returns and
 * subsidiary-rights income as royalty lines, one per format, escalator tier and
 * discount band, then the period's total.
 *
 * Decisions already made for you:
 *
 * - Escalators count life-to-date copies, not the period's. A period that
 *   crosses a threshold is split into two lines at the copy where the rate
 *   changes, never paid at one rate.
 * - Each format has its own basis: list price or net receipts.
 * - High discount is the contract's choice, not ours. A list-price sale at or
 *   above the threshold earns a fraction of the tier rate, or the full rate on
 *   net receipts. It never earns a fraction of a list-price royalty when the
 *   contract switches the base.
 * - Returns are charged at the highest rate reached so far, and never pull the
 *   life-to-date count back below a tier.
 * - Subsidiary income is shared on the money actually received, after foreign
 *   withholding tax. The tax gets a row of its own and is not deducted again.
 * - Each line is rounded once to the minor unit, and the total is the sum of the
 *   rounded lines, the way a publisher's ledger adds up. That can differ by a
 *   penny from rounding the exact sum.
 *
 * All of it is decided in one deriver, because every step does arithmetic or
 * sorting on request data, and publishing refuses both. The node computes
 * nothing: it reads the deriver's result and branches on it, with every test
 * written inline so a publish build keeps each arm.
 *
 * Installed by `dxcl add royalty-earnings-by-format`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface RoyaltyEarningsData {
  royalty: {
    /**
     * ISO 4217 code every amount is in: `"GBP"`, `"USD"`, `"JPY"`. It also sets
     * how many minor units make one major unit. Required: the deriver throws
     * when it is missing or not a code `Intl` knows.
     */
    currency: string;
    /**
     * One entry per format the contract pays on, in the order they should be
     * printed. Two entries may share a `format` (a home and an export edition);
     * give each a `label` so their lines can be told apart. Missing or empty
     * means no format lines.
     */
    formats?: Array<{
      /**
       * Which format this is. It names the line when `label` is missing:
       * "Hardback", "Paperback", "Ebook", "Audiobook". Required: anything else
       * throws.
       */
      format: "hardback" | "paperback" | "ebook" | "audio";
      /** The format's name as printed: `"Trade paperback (export)"`. Missing means the name above. */
      label?: string;
      /**
       * What the royalty percentage is taken of. `"list"` is the recommended
       * retail price per copy. `"net"` is what the publisher actually received.
       * Required: a format with no basis throws, because guessing one gets
       * every line wrong.
       */
      basis: "list" | "net";
      /**
       * List price of one copy, in minor units (`2500` for £25.00). Required
       * when `basis` is `"list"`, and throws without it. When `basis` is
       * `"net"`, it is read only to work out net receipts for a sale that does
       * not give them.
       */
      listPriceMinor?: number;
      /**
       * The escalator. Each tier starts at a life-to-date copy number:
       * `[{ fromUnit: 1, percent: 10 }, { fromUnit: 2501, percent: 12.5 }]`
       * pays 10% on copies 1 to 2,500 and 12.5% from copy 2,501. The first tier
       * must start at 1, and `fromUnit` must strictly ascend. The percent is
       * 0 to 100 with at most two decimals. Anything else throws. Required and
       * non-empty.
       */
      tiers: Array<{ fromUnit: number; percent: number }>;
      /**
       * Copies sold on all earlier statements, before this period. Returns are
       * never subtracted from it. Required, even when it is 0: a missing count
       * would silently pay the first tier's rate, which is the exact mistake
       * this component exists to prevent. Throws when it is not a non-negative
       * integer.
       */
      lifeToDateUnitsBefore: number;
      /**
       * This period's sales, in ledger order. When a threshold is crossed, the
       * earlier rows fill the lower tier first. Missing or empty means none.
       */
      sales?: Array<{
        /** Copies sold. Non-negative integer, else throws. 0 is skipped. */
        units: number;
        /**
         * Trade discount off list, 0 to 100. Missing means 0. It only decides
         * the high-discount band, and only on a `"list"` format.
         */
        discountPercent?: number;
        /**
         * What the publisher received for this whole row, in minor units. Used
         * by a `"net"` format, and by a list-price sale that high discount has
         * switched to net receipts. Missing means it is worked out as
         * units × listPriceMinor × (100 − discountPercent) / 100. When both
         * this and `listPriceMinor` are missing where it is needed, the deriver
         * throws.
         */
        netReceiptsMinor?: number;
      }>;
      /** Copies returned this period. Missing means none. */
      returns?: {
        /**
         * A positive count: 120 returned copies is `120`. A negative value
         * throws, because a sign applied by both the data and the component
         * turns a debit into a credit. 0 prints nothing.
         */
        units: number;
        /**
         * Net receipts credited back for those copies, in minor units.
         * Required when `basis` is `"net"` and `units` is above 0, and throws
         * without it. Ignored on a `"list"` format.
         */
        netReceiptsMinor?: number;
      };
    }>;
    /**
     * The contract's high-discount clause, applied to `"list"` formats only.
     * A net-receipts royalty already reflects the discount. Missing means the
     * contract has no clause, and every sale earns the full tier rate.
     */
    highDiscount?: {
      /** The discount at or above which a sale is high-discount. Missing means 55. */
      thresholdPercent?: number;
      /**
       * `"fraction"`: the tier rate times `fraction`, on list price.
       * `"net"`: the full tier rate, on net receipts instead of list price.
       * Missing means `"fraction"`.
       */
      treatment?: "fraction" | "net";
      /**
       * For `"fraction"`. Missing means 4/5. `numerator` and `denominator`
       * must be positive integers with numerator ≤ denominator, else throws.
       */
      fraction?: { numerator: number; denominator: number };
      /**
       * Deeper steps for `"fraction"`, such as
       * `[{ thresholdPercent: 62.5, numerator: 3, denominator: 5 }]`. The
       * thresholds must strictly ascend from `thresholdPercent`, else throws.
       * The deepest step a sale reaches applies. Ignored for `"net"`.
       */
      deeper?: Array<{ thresholdPercent: number; numerator: number; denominator: number }>;
    };
    /**
     * Income from rights the publisher sold on (translation, audio, serial),
     * in the order they should be printed. Missing or empty means none.
     */
    subRights?: Array<{
      /** What was licensed, as printed: `"Spanish translation"`. Required, non-blank. */
      right: string;
      /** Who licensed it, as printed. Required, non-blank. */
      licensee: string;
      /**
       * The licence fee before any withholding, in minor units. This is the
       * amount the contract names, not the amount that reached the bank.
       * Required, non-negative integer.
       */
      grossMinor: number;
      /**
       * Foreign tax the licensee withheld from `grossMinor`, in minor units.
       * Missing means 0. It throws when it exceeds `grossMinor`.
       */
      foreignTaxWithheldMinor?: number;
      /** The author's contractual share, 0 to 100 with at most two decimals: `50`, `80`. Required. */
      authorSharePercent: number;
    }>;
  };
}

/**
 * Prints the royalty lines for each format and sold right, then the total, and
 * says so plainly when returns made it a debit. Always returns the section.
 */
export const RoyaltyEarningsByFormat: Nodes = async () => {
  const [royalty] = useState((data: RoyaltyEarningsData) => data.royalty);
  const p = await useDeriver(royaltyEarningsByFormat, [royalty]);

  return (
    <Section id="royalty-earnings" title="Royalties earned this period">
      {p.outcome === "none" && (
        <Paragraph id="royalty-earnings-none">
          No copies were sold or returned, and no subsidiary income was received, in this period.
        </Paragraph>
      )}
      {p.outcome === "lines" && (
        <Table
          id="royalty-earnings-lines"
          columns={[
            { width: "auto" },
            { width: 22, align: "right" },
            { width: 34, align: "right" },
            { width: 26, align: "right" },
            { width: 28, align: "right" },
          ]}
        >
          <Row id="royalty-earnings-lines-header" header>
            <Cell>Description</Cell>
            <Cell>Copies</Cell>
            <Cell>Basis</Cell>
            <Cell>Rate</Cell>
            <Cell>Royalty</Cell>
          </Row>
          {p.lines.map((line) => (
            <Row id="royalty-earnings-line">
              <Cell>{line.description}</Cell>
              <Cell>{line.units}</Cell>
              <Cell>{line.basis}</Cell>
              <Cell>{line.rate}</Cell>
              <Cell>{line.amount}</Cell>
            </Row>
          ))}
          <Row id="royalty-earnings-total">
            <Cell>Total royalties this period</Cell>
            <Cell></Cell>
            <Cell></Cell>
            <Cell></Cell>
            <Cell>{p.total}</Cell>
          </Row>
        </Table>
      )}
      {p.hasMemo && (
        <Paragraph id="royalty-earnings-tax-note">
          Foreign tax rows are shown for your records. The tax was withheld before your share was calculated, so it is
          not deducted again.
        </Paragraph>
      )}
      {p.negative && (
        <Paragraph id="royalty-earnings-debit">
          Returns this period exceeded earnings, so this statement shows a debit of {p.totalAbs}.
        </Paragraph>
      )}
      {p.lifeToDate && (
        <Paragraph id="royalty-earnings-life-to-date">Copies sold to date, before returns: {p.lifeToDate}.</Paragraph>
      )}
    </Section>
  );
};
