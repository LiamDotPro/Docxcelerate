import { arrange, Paragraph, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { DepositCancellationOutcomeLines } from "../arrangements/deposit-cancellation-outcome/lines.ts";
import { stacked } from "../arrangements/deposit-cancellation-outcome/stacked.arrangement.tsx";
import { table } from "../arrangements/deposit-cancellation-outcome/table.arrangement.tsx";
import depositCancellationOutcome from "../derivers/deposit-cancellation-outcome.deriver.ts";

/**
 * The body section for a cancelled order: what happens to each kind of money
 * paid in advance, how much is refunded, how much is kept and why, and the VAT
 * or sales-tax position of each part.
 *
 * Three kinds of money that look alike are taxed differently. A security
 * deposit is held against damage and carries no VAT; it comes back less
 * itemised deductions. A refundable advance is part-payment of the price, and a
 * contractual cancellation fee may be retained from it. A non-refundable
 * advance is retained in full.
 *
 * UK, since 1 March 2019 (Revenue and Customs Brief 13/2018): VAT already
 * accounted for on a retained advance stays due and no credit note is issued
 * for it; the refunded part gets a credit note at the advance's original rate.
 * The older advice that a forfeited deposit is compensation outside the scope
 * of VAT is deliberately not followed.
 *
 * Germany: VAT on a refunded advance is corrected when the money is actually
 * paid back, not on cancellation (UStG s.17(2) no.2), and an agreed
 * cancellation fee is consideration with VAT backed out at the advance's rate
 * (CJEU C-622/23). Damages not agreed in the contract are not covered.
 *
 * US: whether a forfeited deposit bears sales tax depends on the state, so the
 * data carries an explicit flag rather than the component guessing.
 *
 * A partly retained advance is split per VAT band, VAT is worked on the
 * retained part, and the refunded VAT is whatever is left of the VAT charged,
 * so the two halves always add up to it. No amount is printed negative; the
 * labels say which way the money goes.
 *
 * The statute references have not been checked against primary sources. An
 * adviser needs to review them before this goes in front of a customer.
 *
 * Everything is decided once, in one deriver, and the node prints what it
 * returns. Deductions are fixed slots, not a loop: the resolver reads a
 * published loop's collection from request data and ignores derived scope, so a
 * loop over derived rows would print nothing. Published, the section is 19
 * branches (the nothing-paid arm, the arrangement's wrapper, 11 rows and 6
 * notes) against the default `branchLimit` of 32.
 *
 * Arrangement: where the amounts sit is the `arrangement` prop. This file
 * decides every line; each arrangement, in
 * `arrangements/deposit-cancellation-outcome/`, only places them.
 *
 * Installed by `dxcl add deposit-cancellation-outcome`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface DepositCancellationOutcomeData {
  order: {
    /**
     * The order, booking or contract that was cancelled, printed as given
     * (trimmed). Required: the deriver throws when it is missing or blank.
     */
    reference: string;
    /**
     * ISO 4217 code: "GBP", "EUR", "USD". Required, no default. Its minor-unit
     * exponent comes from Intl; a code Intl rejects throws, naming
     * `order.currency`.
     */
    currency: string;
    /**
     * Whose rules apply. Required: anything else throws. There is no generic
     * "EU" arm, because the German repayment rule is not the Directive's default.
     */
    jurisdiction: "UK" | "DE" | "US";
    /**
     * For US only: the state whose sales-tax rule decides a forfeited deposit,
     * printed as given ("California" or "CA"). Required when jurisdiction is US
     * and anything is retained from an advance; throws otherwise. Ignored
     * outside the US.
     */
    usState?: string;
    /**
     * For US only: whether `usState` charges sales tax on a forfeited deposit.
     * Required when jurisdiction is US and anything is retained from an
     * advance: missing throws, because guessing is wrong in roughly half the
     * states. Ignored outside the US.
     */
    usForfeitTaxable?: boolean;
    /**
     * The day the order was cancelled, ISO 8601 (`2026-11-03`). Anything after
     * the tenth character is ignored. Required: throws when missing or not
     * YYYY-MM-DD. For UK, a date before 2019-03-01 throws: the
     * pre-Brief-13/2018 rule is out of scope.
     */
    cancelledOn: string;
    /**
     * Every payment taken in advance, in the order they should be read. Missing
     * or empty means nothing was paid: the section prints the intro and the
     * "nothing paid" line.
     */
    payments?: Array<{
      /** Which kind of money it is. Required: anything else throws, naming `order.payments[i].kind`. */
      kind: "security" | "refundable-advance" | "non-refundable-advance";
      /**
       * What was paid, VAT-inclusive, in major units (1250.5 for 1,250.50).
       * Required: not a finite number, negative, or more decimals than the
       * currency allows → throws. Zero is allowed and contributes nothing.
       */
      gross: number;
      /**
       * The VAT analysis of this payment as invoiced, one entry per rate. Each
       * net + vat must sum, across bands, to `gross` exactly in minor units, or
       * it throws. Missing on an advance means it carried no VAT (one 0% band
       * of the whole gross). On a security deposit, and on any US payment, a
       * band with vat > 0 throws: a security deposit carries no VAT, and
       * refunding collected US sales tax is out of scope.
       */
      bands?: Array<{
        /** The VAT rate in percent, 0–100, at most two decimals: 20, 19, 7, 5.5. */
        rate: number;
        /** Net in major units, ≥ 0. */
        net: number;
        /** VAT in major units, ≥ 0, as charged on the original invoice. */
        vat: number;
      }>;
      /**
       * The invoice this advance was billed on. Named in the credit-note line
       * when part of it is refunded with VAT. Missing: the line names no invoice.
       */
      invoiceNumber?: string;
    }>;
    /**
     * The cancellation fee kept from the refundable advances, as the contract
     * sets it. Missing means none is kept: refundable advances come back in full.
     */
    retention?: {
      /**
       * A share of the pooled refundable-advance gross, 0–100, at most two
       * decimals. Give exactly one of `percent` and `amount`; both or neither throws.
       */
      percent?: number;
      /**
       * A fixed sum in major units, ≥ 0. More than the pooled refundable-advance
       * gross throws (the section never prints a negative refund).
       */
      amount?: number;
      /**
       * Why it is kept, printed in the "Retained: <reason>" line. Required when
       * `retention` is given; missing or blank throws.
       */
      reason: string;
    };
    /**
     * Itemised deductions from the security deposit, in print order. At most 6;
     * more throws. Missing or empty means the deposit comes back whole. Given
     * when no security deposit was paid, or summing to more than it, throws.
     */
    deductions?: Array<{
      /** What it is for, printed as given (trimmed). Required, non-blank. */
      label: string;
      /** VAT-inclusive, major units, > 0. Zero or negative throws. */
      amount: number;
      /**
       * True when the deduction pays for a supply (cleaning, a service), so VAT
       * is due on it. False means compensation, outside the scope of VAT.
       * Required for UK and DE: missing throws, because either default
       * misstates VAT. Ignored for US.
       */
      isSupply?: boolean;
      /**
       * The rate VAT is backed out at, percent. Required when `isSupply` is
       * true for UK and DE; missing throws. Ignored otherwise.
       */
      vatRate?: number;
    }>;
    /**
     * The number of the credit note issued for the refunded part of the
     * advances. Printed in the credit-note line. Missing (UK or DE, refunded
     * VAT > 0) prints "will follow" instead of "issued".
     */
    creditNoteNumber?: string;
    /**
     * When the refund was actually paid, ISO 8601. Missing means not yet paid.
     * In DE it also decides the VAT-correction sentence. Throws when present
     * but not a date.
     */
    refundPaidOn?: string;
  };
}

/**
 * Where the amounts sit.
 *
 * - `table`: label and amount in two columns, amounts right-aligned. The default.
 * - `stacked`: one "Label: amount" line under another.
 */
export type DepositCancellationOutcomeArrangement = "table" | "stacked";

/** Props for {@linkcode DepositCancellationOutcome}. */
export interface DepositCancellationOutcomeProps {
  /** Where the amounts sit. Defaults to `"table"`. */
  arrangement?: DepositCancellationOutcomeArrangement;
}

export const DepositCancellationOutcome: SectionComponent<DepositCancellationOutcomeProps> = async (
  { arrangement },
) => {
  const [order] = useState((data: DepositCancellationOutcomeData) => data.order);
  const c = await useDeriver(depositCancellationOutcome, [order]);

  const lines: DepositCancellationOutcomeLines = c.lines;

  return (
    <Section id="deposit-cancellation" title="Your payments after cancellation">
      <Paragraph id="deposit-cancellation-intro">Order {c.reference} was cancelled on {c.cancelledOn}.</Paragraph>
      {c.nothingPaid && (
        <Paragraph id="deposit-cancellation-nothing-paid">
          Nothing was paid in advance on this order, so nothing is refunded or retained.
        </Paragraph>
      )}
      {arrange(arrangement, {
        table: () => table(lines),
        stacked: () => stacked(lines),
      })}
      {c.refundPaid && <Paragraph id="deposit-cancellation-refund-paid">{c.refundPaid}</Paragraph>}
      {c.creditNote && <Paragraph id="deposit-cancellation-credit-note">{c.creditNote}</Paragraph>}
      {c.vatCorrection && <Paragraph id="deposit-cancellation-vat-correction">{c.vatCorrection}</Paragraph>}
      {c.retainedVat && <Paragraph id="deposit-cancellation-retained-vat">{c.retainedVat}</Paragraph>}
      {c.securityVat && <Paragraph id="deposit-cancellation-security-vat">{c.securityVat}</Paragraph>}
      {c.salesTax && <Paragraph id="deposit-cancellation-sales-tax">{c.salesTax}</Paragraph>}
    </Section>
  );
};
