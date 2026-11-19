import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import promptPaymentDiscount from "../derivers/prompt-payment-discount.deriver.ts";

/**
 * The early-payment terms block of an invoice, and the figures behind it,
 * worked per VAT band under the rule of the invoice's jurisdiction.
 *
 * The decision already made for you: VAT is never simply recomputed on the
 * discounted price, and the discount is never taken off the gross total.
 *
 * - UK, from 1 April 2015: VAT is charged on the full price. Per band the
 *   block shows the discounted net, the VAT on it (rounded once per band) and
 *   the amount payable if paid early. It then prints HMRC's wording when no
 *   credit note will be issued. That sentence matters: without it the
 *   supplier has to issue a credit note. `issuesCreditNote` defaults to false
 *   because printing it carries no obligation the seller can forget.
 * - Germany (Skonto, UStG s.14(4) no.7): the reduction and a fixed date are
 *   stated and the invoice's VAT stays on the full price. On a mixed-rate
 *   invoice the reduction is split per rate by largest remainder, because
 *   s.17 corrects VAT per rate. EU and FR follow the same rule.
 * - Zero-rated, exempt and reverse-charge bands: the discount reduces the net
 *   only. Their VAT stays 0.
 * - US: the terms are printed in the usual shorthand ("2/10, net 30"). Whether
 *   sales tax on the early-payment amount is on the full or the discounted
 *   price is looked up per state in a table in the deriver (NY, MA, CA, TX),
 *   never assumed. An unknown state throws until the owner adds it.
 *
 * No discount, no deadline or a deadline before the issue date prints
 * nothing. A deadline after the due date fails the document.
 *
 * All of it is one deriver, because every figure is arithmetic on request
 * data, and publishing refuses that. The node computes nothing: it reads the
 * deriver's result and branches on it, with every test written inline so a
 * publish build keeps each arm.
 *
 * The band rows are six fixed slots rather than a `.map()`. The resolver
 * reads a published loop's collection from request data and ignores the
 * derived scope, so a `.map()` over the deriver's rows would publish and then
 * resolve to no rows; conditions and `{{derived.…}}` tokens do honour the
 * derived scope.
 *
 * The node publishes 9 branches against the default `branchLimit` of 32: the
 * section, the statement, the table and six rows.
 *
 * Installed by `dxcl add prompt-payment-discount`, which also installs its
 * deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface PromptPaymentDiscountData {
  invoice: {
    /**
     * The date of issue, ISO 8601 (`2026-11-30`). The discount deadline is
     * counted from it. Anything after the tenth character is ignored.
     * Required when a discount applies: the deriver throws when it is missing
     * or does not start with YYYY-MM-DD. Not read when there is no discount.
     */
    issueDate: string;
    /**
     * ISO 4217 code the amounts are in: "GBP", "EUR", "USD". It also sets the
     * minor units (two for GBP, none for JPY). Required when a discount
     * applies: missing, or a code Intl rejects, throws.
     */
    currency: string;
    /**
     * Whose rules apply. `UK`: VAT stays on the full price, per-band figures
     * and HMRC's wording. `DE`: Skonto line in German, per-rate split when
     * there is more than one band. `EU` (any member state without rules of
     * its own here) and `FR`: the same rule as DE, in English and in French.
     * `US`: shorthand terms, sales tax per `usState`. Required when a
     * discount applies: anything else throws.
     */
    jurisdiction: "UK" | "EU" | "DE" | "FR" | "US";
    /**
     * US only: the two-letter state code ("NY"), trimmed and upper-cased
     * before lookup. It decides whether sales tax on the early-payment amount
     * is on the full or the discounted price. Required for US: a missing code,
     * or a state not in the deriver's table, throws. The table is meant to be
     * extended by the owner. Ignored outside the US.
     */
    usState?: string;
    /**
     * The invoice's payment term in days ("net 30" is 30). Used to reject a
     * discount deadline after the due date, and printed in the US shorthand.
     * Missing means no check, and the shorthand prints "2/10" with no ", net N".
     * Present but not a positive integer throws.
     */
    dueDays?: number;
    /**
     * The invoice's VAT analysis, one entry per rate and treatment, in the
     * order the invoice prints them. Rows print in this order. Required and
     * non-empty when a discount applies: missing or empty throws. More than
     * six bands with a non-zero net throws (six fixed slots).
     */
    bands: Array<{
      /**
       * The rate as a percentage: 20, 5, 8.875. At most 4 decimal places.
       * Required for `standard` and `reduced`: missing, ≤ 0, ≥ 100, not
       * finite or more than 4 decimal places throws. Ignored for the other
       * treatments, which are labelled by treatment.
       */
      rate: number | null;
      /**
       * How the band is taxed. `zero`, `exempt` and `reverse-charge` have
       * their net reduced and their VAT left at 0. On a US invoice only
       * `standard`/`reduced` bands carry sales tax. Anything outside the
       * union throws. Two bands with the same treatment (and, for
       * standard/reduced, the same rate) throw.
       */
      treatment: "standard" | "reduced" | "zero" | "exempt" | "reverse-charge";
      /**
       * The band's net at full price, in major units, as the invoice prints
       * it. No more decimals than the currency allows: more throws. Negative
       * or not finite throws. 0 is allowed and prints no row.
       */
      net: number;
      /**
       * The VAT (on a US invoice: the sales tax) charged on the full-price
       * net, in major units, as printed. Read by DE/EU/FR, where the band's
       * gross is net + vat, and by US states that tax the full price. UK
       * recomputes from `rate` instead. Same number rules as `net`. Non-zero
       * on a zero, exempt or reverse-charge band throws.
       */
      vat: number;
    }>;
    /**
     * The early-payment terms. Missing means no discount: the component
     * prints nothing and checks nothing else.
     */
    promptPaymentDiscount?: {
      /**
       * The discount as a percentage of the full price: 2, 2.5. Missing,
       * null or 0 prints nothing. Negative, ≥ 100, not finite or more than
       * 2 decimal places throws.
       */
      percent: number;
      /**
       * Days after `issueDate` within which payment earns the discount. The
       * deadline is `issueDate` + `days`, printed as a date. Missing or null
       * prints nothing (no deadline). 0 or negative prints nothing (a
       * deadline on or before the issue date is not an early-payment term).
       * Not an integer throws. Greater than `dueDays` throws.
       */
      days: number;
      /**
       * UK only: whether the seller issues a credit note when the discount is
       * taken. Missing or false prints HMRC's "No credit note will be
       * issued…" wording, which makes the invoice enough on its own. True
       * prints a sentence promising a credit note instead. Ignored outside
       * the UK.
       */
      issuesCreditNote?: boolean;
    };
  };
}

/**
 * Prints the early-payment terms and, where the jurisdiction needs them, the
 * figures per band. Prints nothing when there is no discount to state.
 */
export const PromptPaymentDiscount: Nodes = async () => {
  const [invoice] = useState((d: PromptPaymentDiscountData) => d.invoice);
  const d = await useDeriver(promptPaymentDiscount, [invoice]);

  return (
    <>
      {d.show && (
        <Section id="prompt-payment-discount" title={d.heading}>
          <Paragraph id="prompt-payment-discount-terms">{d.terms}</Paragraph>
          {d.statement && <Paragraph id="prompt-payment-discount-statement">{d.statement}</Paragraph>}
          {d.showTable && (
            <Table
              id="prompt-payment-discount-figures"
              columns={[{ width: "auto" }, { width: 34, align: "right" }, { width: 34, align: "right" }, { width: 34, align: "right" }]}
            >
              <Row id="prompt-payment-discount-figures-header" header>
                <Cell>{d.columns.rate}</Cell>
                <Cell>{d.columns.first}</Cell>
                <Cell>{d.columns.second}</Cell>
                <Cell>{d.columns.third}</Cell>
              </Row>
              {d.rows[0].label && (
                <Row id="prompt-payment-discount-band-1">
                  <Cell>{d.rows[0].label}</Cell>
                  <Cell>{d.rows[0].first}</Cell>
                  <Cell>{d.rows[0].second}</Cell>
                  <Cell>{d.rows[0].third}</Cell>
                </Row>
              )}
              {d.rows[1].label && (
                <Row id="prompt-payment-discount-band-2">
                  <Cell>{d.rows[1].label}</Cell>
                  <Cell>{d.rows[1].first}</Cell>
                  <Cell>{d.rows[1].second}</Cell>
                  <Cell>{d.rows[1].third}</Cell>
                </Row>
              )}
              {d.rows[2].label && (
                <Row id="prompt-payment-discount-band-3">
                  <Cell>{d.rows[2].label}</Cell>
                  <Cell>{d.rows[2].first}</Cell>
                  <Cell>{d.rows[2].second}</Cell>
                  <Cell>{d.rows[2].third}</Cell>
                </Row>
              )}
              {d.rows[3].label && (
                <Row id="prompt-payment-discount-band-4">
                  <Cell>{d.rows[3].label}</Cell>
                  <Cell>{d.rows[3].first}</Cell>
                  <Cell>{d.rows[3].second}</Cell>
                  <Cell>{d.rows[3].third}</Cell>
                </Row>
              )}
              {d.rows[4].label && (
                <Row id="prompt-payment-discount-band-5">
                  <Cell>{d.rows[4].label}</Cell>
                  <Cell>{d.rows[4].first}</Cell>
                  <Cell>{d.rows[4].second}</Cell>
                  <Cell>{d.rows[4].third}</Cell>
                </Row>
              )}
              {d.rows[5].label && (
                <Row id="prompt-payment-discount-band-6">
                  <Cell>{d.rows[5].label}</Cell>
                  <Cell>{d.rows[5].first}</Cell>
                  <Cell>{d.rows[5].second}</Cell>
                  <Cell>{d.rows[5].third}</Cell>
                </Row>
              )}
            </Table>
          )}
          <Paragraph id="prompt-payment-discount-payable">{d.payable}</Paragraph>
        </Section>
      )}
    </>
  );
};
