import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import disbursementsAndRecharges from "../derivers/disbursements-and-recharges.deriver.ts";

/**
 * A professional-services invoice section that splits the firm's expenses into
 * the two VAT treatments the law recognises and prints each as its own block
 * with its own total.
 *
 * A recharge is a cost the firm incurred in making its own supply. It takes the
 * main supply's VAT rate on what the firm is out of pocket: net of input VAT it
 * reclaimed, including tax it could not. A disbursement is a payment made as the
 * client's agent at exact cost (HMRC Notice 700 s.25; Directive 2006/112/EC
 * Art 79(c)): outside the scope of VAT, printed with no VAT, and listed after the
 * recharges' VAT total.
 *
 * Why agents get this wrong: they put every expense in one list, then add VAT to
 * everything (overcharging genuine disbursements) or to nothing (undercharging
 * recharges, the error HMRC assesses four years back). They pass the receipt's
 * own VAT rate through, keep a markup on a "disbursement", and convert foreign
 * receipts at the invoice-date rate.
 *
 * Decisions already made for you:
 *
 * - An item is a disbursement only when it is marked `paidAsAgent` and nothing
 *   contradicts that. A markup, input VAT the firm reclaimed, or a category HMRC
 *   treats as the firm's own cost (travel, subsistence, postage, electronic
 *   search fees since Brabners) moves it to the recharges with a printed note, or
 *   makes `strict` throw.
 * - Recharges never inherit the receipt's rate; they always take the main
 *   supply's. A zero-rated fare recharged becomes standard-rated.
 * - Foreign receipts are converted at the rate given for the expense date, and
 *   the original amount, the rate and the converted figure are all printed.
 * - VAT is worked once on the recharges' net, half away from zero, not per line.
 * - US invoices skip VAT and split on each line's `taxable` flag.
 * - Disbursements are never folded together: each must be itemised.
 *
 * All of it is decided in one deriver, because every step does arithmetic,
 * sorting or date formatting on request data, and publishing refuses all of
 * those. The node computes nothing: it branches on the deriver's result with
 * every test written inline so a publish build keeps each arm.
 *
 * The lines are fixed slots (8 recharges, 6 disbursements) rather than a
 * `.map()`. The resolver reads a published loop's collection from request data
 * and ignores the derived scope, so a `.map()` over the deriver's rows would
 * publish and then resolve to no rows.
 *
 * The node publishes 20 branches: the section, the two blocks, eight recharge
 * rows, the two VAT rows, six disbursement rows and the notes. A document that
 * also carries tax-treatment-notice (15) will need `branchLimit` raised above the
 * default 32.
 *
 * Installed by `dxcl add disbursements-and-recharges`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface DisbursementsAndRechargesData {
  invoice: {
    /**
     * Whose rules apply. `UK` and `EU` apply the disbursement test and charge VAT on
     * recharges. `DE` and `FR` are read as `EU`, so a document that also uses
     * invoice-document-type shares one field. `US` skips all VAT logic and splits on
     * each expense's `taxable` flag. Required whenever there are expenses: anything
     * else throws.
     */
    jurisdiction: "UK" | "EU" | "DE" | "FR" | "US";
    /**
     * The invoice's ISO 4217 currency. Every figure is printed in it, and it sets the
     * minor units. Compared upper-cased. Required whenever there are expenses: throws
     * when missing or not a code Intl knows.
     */
    currency: string;
    /**
     * VAT rate of the main supply (the firm's fees), as a percentage: 20, 5, 0, 8.5.
     * At most two decimal places. Every recharge takes this rate. `null` means the
     * fees carry no VAT on this invoice (exempt, or reverse-charged to a business
     * customer abroad). Recharges then follow the same treatment: no VAT row, and the
     * block says why. Required (number or null) for UK/EU; throws when undefined,
     * negative, non-finite or more than 2 dp. Ignored for US.
     */
    mainSupplyVatRate: number | null;
    /**
     * Every expense to pass on, in any order. Each block prints them by date, with
     * ties in input order. Missing or empty prints nothing, and then nothing else on
     * `invoice` is read or validated.
     */
    expenses?: Array<{
      /** The day the cost was incurred, ISO 8601; a time after the date is ignored.
       * Required: throws when missing or not YYYY-MM-DD. Also the day `fxRate` must be for. */
      date: string;
      /** What it was, printed as given. Required: throws when missing or blank. */
      description: string;
      /**
       * Recognised values, compared exactly: "travel", "subsistence", "accommodation",
       * "postage", "telephone", "photocopying", "printing", "electronic-search".
       * Under UK/EU these are the firm's own costs, so an item marked `paidAsAgent`
       * with one of them is reclassified as a recharge. Any other value, or missing,
       * has no effect. Never printed.
       */
      category?: string;
      /**
       * The receipt amount before any VAT or sales tax on it, in the expense's
       * currency, in major units (12.50). Negative is a refund of an expense and is
       * allowed. Required: throws when not finite, or when it has more decimals than
       * the currency has minor units.
       */
      net: number;
      /**
       * VAT on the receipt that the firm reclaimed as input tax. It is never charged
       * on and never printed. Under UK/EU, a value above 0 on an item marked
       * `paidAsAgent` reclassifies it: a firm that reclaimed the VAT received the
       * supply itself. Missing means 0. Throws when negative or when it has too many
       * decimals.
       */
      inputVatReclaimed?: number;
      /**
       * Tax on the receipt the firm paid and could not reclaim (foreign VAT, US sales
       * tax). It is part of the cost: added to `net` before conversion, markup and
       * VAT. Missing means 0. Throws when negative or when it has too many decimals.
       */
      irrecoverableTax?: number;
      /** ISO 4217 code of the receipt. Missing means `invoice.currency`. Compared upper-cased. */
      currency?: string;
      /**
       * Units of `invoice.currency` per one unit of the expense currency, on `date`
       * (1 EUR = 0.8571 GBP is 0.8571). At most 6 decimal places. It is printed as
       * given. Required when the currencies differ: throws when missing, 0 or less,
       * non-finite or more than 6 dp. Ignored when they match.
       */
      fxRate?: number;
      /**
       * Handling markup, as a percentage of the converted cost. At most 2 dp. Above 0
       * makes the item a recharge in every jurisdiction: a disbursement is passed on
       * at exact cost. Missing means 0. Throws when negative.
       */
      markupPercent?: number;
      /**
       * Whether the firm paid this as the client's agent: the client received the
       * supply, the firm recorded it outside its own costs (a suspense account), and
       * it is passed on at exact cost. Missing means false, so the item is a
       * recharge, because undercharging VAT is the error that gets assessed. Ignored
       * for US.
       */
      paidAsAgent?: boolean;
      /** Printed in brackets after the description when given and not blank. */
      supplierName?: string;
      /**
       * US only: whether the state taxes this reimbursed expense. Missing means not
       * taxable, and a note says so (strict throws). Ignored outside the US.
       */
      taxable?: boolean;
    }>;
  };
}

/** Props for {@linkcode DisbursementsAndRecharges}. */
export interface DisbursementsAndRechargesProps {
  /**
   * Refuse to reclassify. When true, the deriver throws on the first item that would
   * get a note (a reclassified disbursement, or a US item with no `taxable`), naming
   * its index and reason. Defaults to false.
   */
  strict?: boolean;
}

/**
 * Prints the recharges, their VAT, and the disbursements after them. Prints
 * nothing when there are no expenses, or none worth a line.
 */
export const DisbursementsAndRecharges: Nodes<DisbursementsAndRechargesProps> = async ({ strict }) => {
  const [invoice] = useState((d: DisbursementsAndRechargesData) => d.invoice);
  const p = await useDeriver(disbursementsAndRecharges, [invoice, strict === true]);

  return (
    <>
      {p.any && (
        <Section id="disbursements-and-recharges" title="Expenses">
          {p.hasRecharges && (
            <Section id="disbursements-and-recharges-recharges" title={p.rechargeTitle}>
              <Paragraph id="disbursements-and-recharges-recharges-statement">{p.rechargeStatement}</Paragraph>
              <Table
                id="disbursements-and-recharges-recharges-lines"
                columns={[{ width: 26 }, { width: "auto" }, { width: 32, align: "right" }]}
              >
                <Row id="disbursements-and-recharges-recharges-header" header>
                  <Cell>Date</Cell>
                  <Cell>Description</Cell>
                  <Cell>{p.rechargeAmountHeading}</Cell>
                </Row>
                {p.recharges[0].description && (
                  <Row id="disbursements-and-recharges-recharge-1">
                    <Cell>{p.recharges[0].date}</Cell>
                    <Cell>{p.recharges[0].description}</Cell>
                    <Cell>{p.recharges[0].amount}</Cell>
                  </Row>
                )}
                {p.recharges[1].description && (
                  <Row id="disbursements-and-recharges-recharge-2">
                    <Cell>{p.recharges[1].date}</Cell>
                    <Cell>{p.recharges[1].description}</Cell>
                    <Cell>{p.recharges[1].amount}</Cell>
                  </Row>
                )}
                {p.recharges[2].description && (
                  <Row id="disbursements-and-recharges-recharge-3">
                    <Cell>{p.recharges[2].date}</Cell>
                    <Cell>{p.recharges[2].description}</Cell>
                    <Cell>{p.recharges[2].amount}</Cell>
                  </Row>
                )}
                {p.recharges[3].description && (
                  <Row id="disbursements-and-recharges-recharge-4">
                    <Cell>{p.recharges[3].date}</Cell>
                    <Cell>{p.recharges[3].description}</Cell>
                    <Cell>{p.recharges[3].amount}</Cell>
                  </Row>
                )}
                {p.recharges[4].description && (
                  <Row id="disbursements-and-recharges-recharge-5">
                    <Cell>{p.recharges[4].date}</Cell>
                    <Cell>{p.recharges[4].description}</Cell>
                    <Cell>{p.recharges[4].amount}</Cell>
                  </Row>
                )}
                {p.recharges[5].description && (
                  <Row id="disbursements-and-recharges-recharge-6">
                    <Cell>{p.recharges[5].date}</Cell>
                    <Cell>{p.recharges[5].description}</Cell>
                    <Cell>{p.recharges[5].amount}</Cell>
                  </Row>
                )}
                {p.recharges[6].description && (
                  <Row id="disbursements-and-recharges-recharge-7">
                    <Cell>{p.recharges[6].date}</Cell>
                    <Cell>{p.recharges[6].description}</Cell>
                    <Cell>{p.recharges[6].amount}</Cell>
                  </Row>
                )}
                {p.recharges[7].description && (
                  <Row id="disbursements-and-recharges-recharge-8">
                    <Cell>{p.recharges[7].date}</Cell>
                    <Cell>{p.recharges[7].description}</Cell>
                    <Cell>{p.recharges[7].amount}</Cell>
                  </Row>
                )}
                {p.showVat && (
                  <Row id="disbursements-and-recharges-recharges-net">
                    <Cell></Cell>
                    <Cell>Net total</Cell>
                    <Cell>{p.rechargeNet}</Cell>
                  </Row>
                )}
                {p.showVat && (
                  <Row id="disbursements-and-recharges-recharges-vat">
                    <Cell></Cell>
                    <Cell>{p.vatLabel}</Cell>
                    <Cell>{p.vat}</Cell>
                  </Row>
                )}
                <Row id="disbursements-and-recharges-recharges-total">
                  <Cell></Cell>
                  <Cell>{p.rechargeTotalLabel}</Cell>
                  <Cell>{p.rechargeTotal}</Cell>
                </Row>
              </Table>
            </Section>
          )}
          {p.hasDisbursements && (
            <Section id="disbursements-and-recharges-disbursements" title={p.disbursementTitle}>
              <Paragraph id="disbursements-and-recharges-disbursements-statement">{p.disbursementStatement}</Paragraph>
              <Table
                id="disbursements-and-recharges-disbursements-lines"
                columns={[{ width: 26 }, { width: "auto" }, { width: 32, align: "right" }]}
              >
                <Row id="disbursements-and-recharges-disbursements-header" header>
                  <Cell>Date</Cell>
                  <Cell>Description</Cell>
                  <Cell>Amount</Cell>
                </Row>
                {p.disbursements[0].description && (
                  <Row id="disbursements-and-recharges-disbursement-1">
                    <Cell>{p.disbursements[0].date}</Cell>
                    <Cell>{p.disbursements[0].description}</Cell>
                    <Cell>{p.disbursements[0].amount}</Cell>
                  </Row>
                )}
                {p.disbursements[1].description && (
                  <Row id="disbursements-and-recharges-disbursement-2">
                    <Cell>{p.disbursements[1].date}</Cell>
                    <Cell>{p.disbursements[1].description}</Cell>
                    <Cell>{p.disbursements[1].amount}</Cell>
                  </Row>
                )}
                {p.disbursements[2].description && (
                  <Row id="disbursements-and-recharges-disbursement-3">
                    <Cell>{p.disbursements[2].date}</Cell>
                    <Cell>{p.disbursements[2].description}</Cell>
                    <Cell>{p.disbursements[2].amount}</Cell>
                  </Row>
                )}
                {p.disbursements[3].description && (
                  <Row id="disbursements-and-recharges-disbursement-4">
                    <Cell>{p.disbursements[3].date}</Cell>
                    <Cell>{p.disbursements[3].description}</Cell>
                    <Cell>{p.disbursements[3].amount}</Cell>
                  </Row>
                )}
                {p.disbursements[4].description && (
                  <Row id="disbursements-and-recharges-disbursement-5">
                    <Cell>{p.disbursements[4].date}</Cell>
                    <Cell>{p.disbursements[4].description}</Cell>
                    <Cell>{p.disbursements[4].amount}</Cell>
                  </Row>
                )}
                {p.disbursements[5].description && (
                  <Row id="disbursements-and-recharges-disbursement-6">
                    <Cell>{p.disbursements[5].date}</Cell>
                    <Cell>{p.disbursements[5].description}</Cell>
                    <Cell>{p.disbursements[5].amount}</Cell>
                  </Row>
                )}
                <Row id="disbursements-and-recharges-disbursements-total">
                  <Cell></Cell>
                  <Cell>{p.disbursementTotalLabel}</Cell>
                  <Cell>{p.disbursementTotal}</Cell>
                </Row>
              </Table>
            </Section>
          )}
          {p.notes && <Paragraph id="disbursements-and-recharges-notes">{p.notes}</Paragraph>}
        </Section>
      )}
    </>
  );
};
