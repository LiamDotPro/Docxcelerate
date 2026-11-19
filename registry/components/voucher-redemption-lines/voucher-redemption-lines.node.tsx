import { Cell, type Nodes, Paragraph, Row, Table, useDeriver, useState } from "docxcelerate/template";
import voucherRedemptionLines from "../derivers/voucher-redemption-lines.deriver.ts";

/**
 * The totals block under an invoice's or receipt's line items, with every
 * voucher, coupon and gift card where its tax treatment requires, so the VAT or
 * sales tax shown is the tax actually due. No heading, so it can sit directly
 * under the line table.
 *
 * Decisions already made for you:
 *
 * - A voucher is not always a discount. A retailer's coupon, and in the US a
 *   free promotional card, reduce the taxable amount and print above the tax. A
 *   multi-purpose voucher, a gift card and (unless the seller's state treats it
 *   as a discount) a US manufacturer's coupon are payments: tax is charged on
 *   the full price and they print below the total.
 * - A single-purpose voucher had its VAT accounted for when it was sold. The
 *   part it covers is shown as paid by voucher and left out of the VAT
 *   analysis, so VAT is not charged twice.
 * - A coupon is spread across the VAT bands of the lines it applies to by
 *   largest remainder, so the parts add up to the coupon exactly and no line
 *   goes below zero.
 * - Value left on a voucher or card is a remaining balance. It is never change
 *   and the amount due is never negative.
 *
 * The node computes nothing. One deriver validates, allocates, taxes, sorts,
 * caps and formats, because all of that is arithmetic or filtering on request
 * data and publishing refuses it. The node reads the result and branches on it.
 *
 * Each group is four fixed slots rather than a `.map()`: the resolver reads a
 * published loop's collection from request data and ignores the derived scope,
 * so a `.map()` over the deriver's rows would publish and then resolve to no
 * rows. Every test is written inline, never hoisted into a `const`.
 *
 * A publish build emits 15 branches against the default `branchLimit` of 32:
 * 12 slot rows and 3 notes.
 *
 * Installed by `dxcl add voucher-redemption-lines`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface VoucherRedemptionData {
  sale: {
    /**
     * ISO 4217 code every amount is in: `"GBP"`, `"EUR"`, `"USD"`. It also sets
     * how many minor units make one major unit. Required: the deriver throws
     * when it is missing or not a code `Intl` knows.
     */
    currency: string;
    /**
     * Whose rules apply. `UK` and `EU` share the VAT voucher rules (Directive
     * 2016/1065; VATA 1994 Sch 10A) and print "VAT". `US` prints "Sales tax".
     * It also decides which redemption kinds are allowed. Required: anything
     * else throws.
     */
    jurisdiction: "UK" | "EU" | "US";
    /**
     * US only. True when the seller's state treats a manufacturer's coupon as
     * a discount that reduces the taxable price (Texas, for one). Missing or
     * false means the majority rule: sales tax is charged on the pre-coupon
     * price, and the coupon reduces only the amount due. Ignored outside the US.
     */
    manufacturerCouponsReduceBase?: boolean;
    /**
     * The sale's lines, used only for their amounts and rates. This component
     * does not print them; the line-item table is a separate component.
     * Required: a missing or empty list throws.
     */
    lines: Array<{
      /**
       * Identifies the line for `appliesToLineKeys`. It is never printed.
       * Required, non-empty and unique: a duplicate throws.
       */
      key: string;
      /**
       * The line's price before tax, in integer minor units (12999 for 129.99),
       * after any line-level discount. Required: throws unless it is an integer
       * from 0 to 90,000,000,000. Zero is allowed: the line takes no part in any
       * allocation.
       */
      netMinor: number;
      /**
       * The VAT rate (UK/EU) or the combined sales-tax rate (US) for this line,
       * as a percentage: `20`, `5.5`, `8.875`. `0` is a zero-rated or untaxed
       * line. Lines that share a rate share a band. Required: throws unless it
       * is a finite number from 0 to 100 with no more than three decimal places.
       */
      ratePercent: number;
    }>;
    /**
     * Every voucher, coupon or card used on the sale, in the order they were
     * presented. That order breaks ties within a group, but it does not decide
     * the group: discounts always come before tax and payments after.
     * Missing or empty prints the totals with no redemption rows.
     */
    redemptions?: Array<{
      /**
       * What it is. UK/EU: `spv` (single-purpose voucher: place of supply and
       * VAT rate known when it was sold), `mpv` (multi-purpose voucher),
       * `discount-coupon` (retailer's money-off coupon). US: `retailer-coupon`,
       * `manufacturer-coupon` (the manufacturer reimburses the seller),
       * `gift-card` (bought for money), `promo-card` (a card given away free).
       * Required: a kind from the other jurisdiction, or an unknown kind, throws.
       * A gift card in the UK is an `mpv`.
       */
      kind:
        | "spv"
        | "mpv"
        | "discount-coupon"
        | "retailer-coupon"
        | "manufacturer-coupon"
        | "gift-card"
        | "promo-card";
      /**
       * Face value in integer minor units. For coupons and promo cards it is on
       * the same tax-exclusive basis as `netMinor`. For `spv`, `mpv` and
       * `gift-card` it is money the customer already paid, so it is set against
       * the tax-inclusive total. Required: throws unless it is a non-negative
       * integer up to 90,000,000,000. `0` is skipped and prints no row.
       */
      amountMinor: number;
      /**
       * The lines it may be used against, by `key`. Missing means every line.
       * Read for coupons, promo cards and `spv`. Ignored for `mpv` and
       * `gift-card`, which pay the whole total. Throws when it is an empty array
       * or names a key that is not on the sale.
       */
      appliesToLineKeys?: string[];
      /**
       * The code printed after the description: "Coupon SAVE10". Missing prints
       * the description alone.
       */
      code?: string;
      /**
       * `spv` only: the VAT rate fixed when the voucher was sold. When given,
       * every line the voucher applies to must be at this rate. When missing,
       * those lines must all share one rate, and that rate is used. Either
       * mismatch throws, because a voucher that spans rates is multi-purpose by
       * definition. Ignored on other kinds.
       */
      spvRatePercent?: number;
    }>;
  };
}

/**
 * Prints the subtotal, the discounts, the tax per band, the total, the payments
 * and the amount due, then a note for each treatment that was used.
 */
export const VoucherRedemptionLines: Nodes = async () => {
  const [sale] = useState((data: VoucherRedemptionData) => data.sale);
  const r = await useDeriver(voucherRedemptionLines, [sale]);

  return (
    <>
      <Table id="voucher-redemption-lines" columns={[{ width: "auto" }, { width: 35, align: "right" }]}>
        <Row id="voucher-redemption-subtotal">
          <Cell>Subtotal</Cell>
          <Cell>{r.subtotal}</Cell>
        </Row>
        {r.discounts[0].description && (
          <Row id="voucher-redemption-discount-1">
            <Cell>{r.discounts[0].description}</Cell>
            <Cell>{r.discounts[0].amount}</Cell>
          </Row>
        )}
        {r.discounts[1].description && (
          <Row id="voucher-redemption-discount-2">
            <Cell>{r.discounts[1].description}</Cell>
            <Cell>{r.discounts[1].amount}</Cell>
          </Row>
        )}
        {r.discounts[2].description && (
          <Row id="voucher-redemption-discount-3">
            <Cell>{r.discounts[2].description}</Cell>
            <Cell>{r.discounts[2].amount}</Cell>
          </Row>
        )}
        {r.discounts[3].description && (
          <Row id="voucher-redemption-discount-4">
            <Cell>{r.discounts[3].description}</Cell>
            <Cell>{r.discounts[3].amount}</Cell>
          </Row>
        )}
        {r.bands[0].description && (
          <Row id="voucher-redemption-tax-1">
            <Cell>{r.bands[0].description}</Cell>
            <Cell>{r.bands[0].amount}</Cell>
          </Row>
        )}
        {r.bands[1].description && (
          <Row id="voucher-redemption-tax-2">
            <Cell>{r.bands[1].description}</Cell>
            <Cell>{r.bands[1].amount}</Cell>
          </Row>
        )}
        {r.bands[2].description && (
          <Row id="voucher-redemption-tax-3">
            <Cell>{r.bands[2].description}</Cell>
            <Cell>{r.bands[2].amount}</Cell>
          </Row>
        )}
        {r.bands[3].description && (
          <Row id="voucher-redemption-tax-4">
            <Cell>{r.bands[3].description}</Cell>
            <Cell>{r.bands[3].amount}</Cell>
          </Row>
        )}
        <Row id="voucher-redemption-total">
          <Cell>Total</Cell>
          <Cell>{r.total}</Cell>
        </Row>
        {r.tenders[0].description && (
          <Row id="voucher-redemption-tender-1">
            <Cell>{r.tenders[0].description}</Cell>
            <Cell>{r.tenders[0].amount}</Cell>
          </Row>
        )}
        {r.tenders[1].description && (
          <Row id="voucher-redemption-tender-2">
            <Cell>{r.tenders[1].description}</Cell>
            <Cell>{r.tenders[1].amount}</Cell>
          </Row>
        )}
        {r.tenders[2].description && (
          <Row id="voucher-redemption-tender-3">
            <Cell>{r.tenders[2].description}</Cell>
            <Cell>{r.tenders[2].amount}</Cell>
          </Row>
        )}
        {r.tenders[3].description && (
          <Row id="voucher-redemption-tender-4">
            <Cell>{r.tenders[3].description}</Cell>
            <Cell>{r.tenders[3].amount}</Cell>
          </Row>
        )}
        <Row id="voucher-redemption-amount-due">
          <Cell>Amount due</Cell>
          <Cell>{r.amountDue}</Cell>
        </Row>
      </Table>
        {r.spvNote && (
          <Paragraph id="voucher-redemption-note-spv">
            VAT on the part paid by voucher was accounted for when the voucher was sold, so it is not charged again here.
          </Paragraph>
        )}
        {r.mpvNote && (
          <Paragraph id="voucher-redemption-note-mpv">
            VAT is charged on the full price. A voucher that can be spent across different VAT rates is a form of payment,
            so it reduces the amount to pay, not the VAT.
          </Paragraph>
        )}
        {r.manufacturerNote && (
          <Paragraph id="voucher-redemption-note-manufacturer">
            Sales tax is charged on the price before the manufacturer's coupon, because the manufacturer reimburses its
            value to us.
          </Paragraph>
        )}
    </>
  );
};
