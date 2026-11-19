import {
  arrange,
  Paragraph,
  Section,
  type Section as SectionComponent,
  useDeriver,
  useState,
} from "docxcelerate/template";
import { columns } from "../arrangements/marketplace-payout-statement/columns.arrangement.tsx";
import { signed } from "../arrangements/marketplace-payout-statement/signed.arrangement.tsx";
import marketplacePayoutStatement from "../derivers/marketplace-payout-statement.deriver.ts";

/**
 * A marketplace seller's statement: one reconciled row per step, in a fixed
 * order, from what happened in the period to what reaches their bank, closing
 * on a balance proved against the opening balance. It makes five decisions for
 * the reader:
 *
 * - Facilitator sales tax is not revenue. Sales tax the marketplace collected
 *   as a marketplace facilitator is shown, then taken straight back out.
 * - A refund returns only part of the commission. The commission on a refunded
 *   order comes back less the platform's refund administration fee, and
 *   fulfilment fees never come back.
 * - A won chargeback is not income. It is the reversal of an earlier
 *   deduction, and its dispute fee stays deducted.
 * - The reserve is on the statement. New holds come off, and matured holds
 *   come back.
 * - There is never a negative "payout". A negative closing balance is either
 *   carried forward or debited from the card or bank account on file.
 *
 * Every amount is an integer of minor units. The only rounding in the
 * statement is VAT on each fee line. When the platform supplies its own
 * closing figure, the deriver refuses a statement that does not tie to it.
 *
 * All of this lives in one deriver, because every row is a sum over request
 * data, which publishing refuses. The node and its arrangements only print
 * what the deriver returns.
 *
 * The rows are fixed named slots rather than a `.map()`. The resolver reads a
 * published loop's collection from request data and ignores the derived scope,
 * so a `.map()` over the deriver's rows would publish and then resolve to no
 * rows; conditions and `{{derived.…}}` tokens do honour the derived scope.
 *
 * The node publishes 18 branches against the default `branchLimit` of 32: 13
 * conditional rows, the reverse-charge note and four outcomes.
 *
 * Installed by `dxcl add marketplace-payout-statement`, which also installs its
 * deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface MarketplacePayoutStatementData {
  payout: {
    /**
     * ISO 4217 code every amount is in: "GBP", "USD", "EUR", "JPY". It also sets
     * how many minor units make one major unit. Required: the deriver throws
     * when it is missing or not a code `Intl` knows.
     */
    currency: string;
    /**
     * The balance the previous statement closed on and could not pay out, in
     * integer minor units. Positive: funds that were unavailable then. Negative:
     * a negative balance carried forward. Required, and 0 for a first statement:
     * the deriver throws when it is missing or not an integer, because a
     * statement that silently drops last period's negative balance does not tie
     * to the bank.
     */
    openingBalance: number;
    /**
     * Every order settled in the period. Missing or empty means no sales, and
     * the sales row prints the currency's zero.
     */
    orders?: Array<{
      /** The marketplace's order id. Read only in error messages. */
      id: string;
      /**
       * The item price paid by the buyer, in integer minor units, excluding
       * facilitator-collected sales tax. Required: throws when it is not a
       * non-negative integer.
       */
      itemTotal: number;
      /** Shipping charged to the buyer, minor units, excluding tax. Missing means 0. */
      shipping?: number;
      /**
       * Sales tax the marketplace collected and remits as marketplace
       * facilitator (or deemed supplier), in minor units. Missing means 0. It is
       * never part of sales.
       */
      facilitatorTax?: number;
    }>;
    /** Refunds made in the period. Missing or empty means none. */
    refunds?: Array<{
      /** The refunded order. Read only in error messages. */
      orderId: string;
      /**
       * What went back to the buyer, in minor units, excluding facilitator tax
       * (that was never the seller's). Required, a non-negative integer.
       */
      amount: number;
      /**
       * The commission on the refunded amount that the marketplace credits
       * back, before its administration fee. Required, a non-negative integer.
       * Fulfilment fees are never credited back, so there is no field for them.
       */
      commissionRefunded: number;
      /**
       * The refund administration fee the marketplace keeps out of that
       * commission, e.g. the lesser of 20% of the referral fee or 5.00 per item.
       * Required, and 0 where the platform keeps none: a missing fee would
       * silently refund the full commission. Throws when it is negative, not an
       * integer, or larger than `commissionRefunded`.
       */
      adminFee: number;
    }>;
    /**
     * Commission and every other fee charged in the period, one entry per line
     * of the platform's fee invoice. Missing or empty means none.
     */
    fees?: Array<{
      /** What the fee is called. Required, non-blank. Read only in error messages. */
      label: string;
      /**
       * The fee before VAT, in integer minor units. Negative is a fee credit or
       * adjustment. Required: throws when it is not an integer.
       */
      amount: number;
      /**
       * The VAT rate charged on this fee, as a percentage: 20, 8.1. Missing
       * means no VAT on this fee. Throws when it is not finite or outside 0 to 100.
       */
      vatPercent?: number;
      /**
       * True when this fee is supplied under the reverse charge (the seller is
       * VAT-registered in another country). No VAT is charged on it, whatever
       * `vatPercent` says, and the statement prints the reverse-charge note.
       * Missing means false.
       */
      reverseCharge?: boolean;
    }>;
    /** Chargebacks raised or settled in the period. Missing or empty means none. */
    chargebacks?: Array<{
      /** The disputed order. Read only in error messages. */
      orderId: string;
      /** The disputed amount, in minor units. Required, a non-negative integer. */
      amount: number;
      /**
       * The dispute fee for this chargeback, in minor units. Missing means 0.
       * Deducted with the chargeback and never returned, even on a win.
       */
      disputeFee?: number;
      /** Where the dispute stands. Required: anything else throws. */
      outcome: "open" | "won" | "lost";
      /**
       * True when the chargeback and its fee were already deducted on an
       * earlier statement, so this period only sees its outcome. Missing means
       * it was raised in this period.
       */
      raisedEarlier?: boolean;
    }>;
    /** The rolling reserve. Missing means no reserve: both rows print nothing. */
    reserve?: {
      /** Newly held this period, in minor units. Missing means 0. */
      held?: number;
      /** Holds that matured and were released this period, in minor units. Missing means 0. */
      released?: number;
      /** How long a hold lasts, in days, printed in the held row's label. Missing leaves it out. */
      days?: number;
    };
    /**
     * The payout account. Only the last four digits are ever printed: non-digits
     * are stripped and the last four are kept, so passing a full account number
     * prints no more than four digits. Fewer than four digits, or missing, prints
     * "your bank account".
     */
    accountLast4?: string;
    /**
     * What happens to a negative closing balance. "carry": nothing is paid and
     * it is deducted from the next payout. "debit": it is collected from the
     * card or bank account on file. Missing means "carry", because promising a
     * debit nobody makes is worse than saying it carries. Anything else throws.
     * Ignored unless the closing balance is negative.
     */
    negativeBalanceMethod?: "carry" | "debit";
    /**
     * The platform's own closing figure for the period, in minor units. When it
     * is given and the rows do not sum to it, the deriver throws and names both
     * figures: a statement that does not tie must not be sent. Missing skips
     * the check.
     */
    closingBalance?: number;
  };
}

/**
 * Where the rows sit.
 *
 * - `signed`: one amount column, each row signed. The default.
 * - `columns`: money in and money out in two columns, unsigned, the way a bank
 *   statement reads.
 */
export type MarketplacePayoutStatementArrangement = "signed" | "columns";

/** Props for {@linkcode MarketplacePayoutStatement}. */
export interface MarketplacePayoutStatementProps {
  /** Where the rows sit. Defaults to `"signed"`. */
  arrangement?: MarketplacePayoutStatementArrangement;
}

/**
 * Prints the reconciled rows from the period's opening balance to its closing
 * balance, then the one sentence that says what happens to the money.
 */
export const MarketplacePayoutStatement: SectionComponent<MarketplacePayoutStatementProps> = async (
  { arrangement },
) => {
  const [payout] = useState((data: MarketplacePayoutStatementData) => data.payout);
  const s = await useDeriver(marketplacePayoutStatement, [payout]);

  return (
    <Section id="marketplace-payout-statement" title="Your payout statement">
      {arrange(arrangement, { signed: () => signed(s.rows), columns: () => columns(s.rows) })}
      {s.reverseChargeNote && <Paragraph id="payout-reverse-charge">{s.reverseChargeNote}</Paragraph>}
      {s.outcome === "payout" && <Paragraph id="payout-paid">{s.outcomeAmount} will be paid to {s.payTo}.</Paragraph>}
      {s.outcome === "nil" && <Paragraph id="payout-nil">There is nothing to pay this period.</Paragraph>}
      {s.outcome === "carry" && (
        <Paragraph id="payout-carry">
          Nothing is paid this period. The negative balance of {s.outcomeAmount} is carried forward and will be
          deducted from your next payout.
        </Paragraph>
      )}
      {s.outcome === "debit" && (
        <Paragraph id="payout-debit">
          Nothing is paid this period. The negative balance of {s.outcomeAmount} will be collected from the card or
          bank account on file.
        </Paragraph>
      )}
    </Section>
  );
};
