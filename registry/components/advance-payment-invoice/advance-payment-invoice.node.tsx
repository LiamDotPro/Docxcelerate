import { arrange, Paragraph, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { AdvancePaymentInvoiceLines } from "../arrangements/advance-payment-invoice/lines.ts";
import { stacked } from "../arrangements/advance-payment-invoice/stacked.arrangement.tsx";
import { table } from "../arrangements/advance-payment-invoice/table.arrangement.tsx";
import advancePaymentInvoice from "../derivers/advance-payment-invoice.deriver.ts";

/**
 * The body of a deposit or prepayment invoice raised against an identified
 * order: the order total, each deposit already requested, the share requested
 * now and what is left to invoice, with a per-rate VAT analysis of this
 * deposit when VAT applies.
 *
 * A percentage is taken of the order gross, never of net plus VAT. The
 * customer expects the figure on the quote, and the two routes differ by a
 * cent (30% of 100.05 gross at 20% is 30.02; of net plus VAT it is 30.01).
 *
 * A mixed-rate order is spread across its rates by largest remainder in minor
 * units, so the band shares add up to the deposit exactly and previous
 * deposits plus this one plus what is left add up to the order total. No
 * remainder cent goes missing. Previous deposits are not split by rate, only
 * this one: largest remainder is not monotone, so a per-rate balance could go
 * a cent negative.
 *
 * Previous deposits (at most 8) and rate bands (at most 4) are fixed slots
 * rather than a `.map()`: the resolver reads a published loop's collection
 * from request data and ignores the derived scope, so a loop over the
 * deriver's lists would publish and then print no rows. More than that throws.
 *
 * Requests are capped at the order total: more than 100% fails the document
 * rather than printing a balance that is negative.
 *
 * Each kind of deposit gets its own tax wording. A UK or EU prepayment is
 * charged VAT now, because the payment is a tax point, and the date of payment
 * prints when it differs from the issue date. Germany says it is an
 * Anzahlungsrechnung. A US deposit carries no sales tax until the final
 * invoice. A refundable security deposit is not a payment for the supply, so
 * it carries no VAT and says so. Which wording applies is an ordinary
 * conditional with an id per arm, so a resolved document records which one a
 * recipient got; every test is written inline against `deposit.kind` and
 * `invoice.jurisdiction`, never hoisted into a `const`, because while
 * publishing a hoisted `===` is settled to a plain boolean and every recipient
 * gets the same arm.
 *
 * Every figure comes from one deriver, because arithmetic on request data
 * throws at publish. The invoice number, issue date and heading belong to
 * `invoice-document-type`, which reads the same `invoice.jurisdiction` and
 * `invoice.issueDate`.
 *
 * Arrangement: this file decides every line; each arrangement, in
 * `arrangements/advance-payment-invoice/`, only places them.
 *
 * Not covered: the final invoice that deducts these deposits, exempt order
 * lines, non-refundable US deposits that some states tax on receipt, and
 * supplies where an advance payment does not make VAT chargeable (reverse
 * charge, intra-Community supplies of goods, margin scheme). The statute
 * references are not checked against primary sources: have an adviser look.
 *
 * Installed by `dxcl add advance-payment-invoice`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface AdvancePaymentInvoiceData {
  order: {
    /**
     * The order, quote or contract this deposit is raised against, printed as
     * given (trimmed) so the customer can match it. Required: the deriver
     * throws when it is missing or blank, because a deposit that names no
     * order cannot be reconciled.
     */
    reference: string;
    /**
     * When the order was placed, ISO 8601 (`2026-10-14`). Anything after the
     * tenth character is ignored. Missing or empty means the line is not
     * printed. Throws when present but not an ISO date.
     */
    date?: string;
    /**
     * ISO 4217 code: "GBP", "EUR", "USD", "JPY". Required, with no default.
     * Its exponent (number of decimal places) comes from Intl. A code Intl
     * rejects throws, naming `order.currency`.
     */
    currency: string;
    /**
     * The order's lines, enough to give the total at each rate. Required and
     * non-empty: missing or empty throws, because there is no order total to
     * take a deposit of. Descriptions are not needed and not printed.
     */
    lines: Array<{
      /**
       * The line's VAT-inclusive amount in major units (pounds, not pence). Give
       * exactly one of `gross` or `net`: both, or neither, throws. It is
       * rounded once to minor units, half away from zero. May be negative (a
       * discount), as long as its rate band stays at zero or above.
       */
      gross?: number;
      /**
       * The line's amount excluding VAT, in major units. VAT is added once per
       * band on the sum of that band's net lines, not once per line. Same
       * rules as `gross`. In the US, give the pre-tax amount here.
       */
      net?: number;
      /**
       * The VAT rate as a percentage: 20, 5, 0, 7.7. 0 is zero-rated and
       * prints as "0%". Required outside the US: missing, negative or not
       * finite throws. Not read in the US, where no tax is charged on a
       * deposit. Exempt (no-rate) lines are not supported.
       */
      vatRate?: number;
    }>;
  };
  deposit: {
    /**
     * The share of the order gross requested on this invoice, as a percentage
     * (30, 12.5). Must be greater than 0 and no more than 100. Give exactly one
     * of `percent` or `amount`: both, or neither, throws.
     */
    percent?: number;
    /**
     * A fixed amount requested on this invoice, gross, in major units. Must be
     * greater than 0. Rounded to minor units, half away from zero.
     */
    amount?: number;
    /**
     * `prepayment` is part of the price, paid in advance and deducted from the
     * final invoice. `security` is a refundable security deposit, returned at
     * the end of the hire: no VAT, no VAT analysis, and it does not reduce
     * what is left to invoice. Required: anything else throws. Read as a
     * prepayment by mistake, a security deposit would be charged VAT. Read as
     * a security deposit by mistake, a prepayment would go without the VAT
     * that is due.
     */
    kind: "prepayment" | "security";
    /**
     * When the payment was received (or will be), ISO 8601. Read only for a
     * prepayment outside the US. Printed when given and on a different day
     * from `invoice.issueDate`. Missing or empty means no line. Throws when
     * present but not an ISO date.
     */
    receivedDate?: string;
  };
  /**
   * Prepayments already requested against this order, in the order they were
   * issued. Missing or empty means none. Each is printed on its own line and
   * counts towards the 100% cap.
   */
  previousDeposits?: Array<{
    /** The earlier deposit invoice's number, printed as given (trimmed). Missing or blank throws. */
    number: string;
    /** What it requested, gross, in major units. Must be greater than 0. */
    amount: number;
  }>;
  invoice: {
    /**
     * Whose rules apply. `EU` means a member state with no rules of its own
     * here, which gets the Directive's defaults. `FR` gets the EU wording.
     * `US` has no VAT: no rate is read, no VAT analysis is printed. Required:
     * anything else throws.
     */
    jurisdiction: "UK" | "EU" | "DE" | "FR" | "US";
    /**
     * This invoice's date of issue, ISO 8601. Required: throws when missing or
     * not YYYY-MM-DD. It is only compared against `deposit.receivedDate` and
     * is not printed here; `invoice-document-type` prints it.
     */
    issueDate: string;
  };
}

/**
 * Where the lines sit.
 *
 * - `table`: labels in a column with the figures right-aligned beside them, so
 *   the subtraction can be checked down the page. The default.
 * - `stacked`: one labelled line under another, down the text column.
 */
export type AdvancePaymentInvoiceArrangement = "table" | "stacked";

/** Props for {@linkcode AdvancePaymentInvoice}. */
export interface AdvancePaymentInvoiceProps {
  /** Where the lines sit. Defaults to `"table"`. */
  arrangement?: AdvancePaymentInvoiceArrangement;
}

export const AdvancePaymentInvoice: SectionComponent<AdvancePaymentInvoiceProps> = async ({ arrangement }) => {
  const [order] = useState((data: AdvancePaymentInvoiceData) => data.order);
  const [deposit] = useState((data: AdvancePaymentInvoiceData) => data.deposit);
  const [previousDeposits] = useState((data: AdvancePaymentInvoiceData) => data.previousDeposits);
  const [invoice] = useState((data: AdvancePaymentInvoiceData) => data.invoice);
  const figures = await useDeriver(advancePaymentInvoice, [order, deposit, previousDeposits, invoice]);

  const lines: AdvancePaymentInvoiceLines = {
    orderReference: figures.orderReference,
    orderDate: figures.orderDate,
    orderTotal: figures.orderTotal,
    previous: figures.previousSlots,
    depositLabel: figures.depositLabel,
    deposit: figures.deposit,
    leftToInvoice: figures.leftToInvoice,
    paymentDateLabel: figures.paymentDateLabel,
    paymentDate: figures.paymentDate,
    vatApplies: figures.vatApplies,
    bands: figures.bandSlots,
    netTotal: figures.netTotal,
    vatTotal: figures.vatTotal,
  };

  return (
    <Section id="advance-payment" title={figures.title}>
      {arrange(arrangement, {
        table: () => table(lines),
        stacked: () => stacked(lines),
      })}
      {deposit.kind === "security" && (
        <Paragraph id="advance-payment-tax-security">
          This is a refundable security deposit, returned at the end of the hire. It is not a payment for the supply
          and is not deducted from the order, so no VAT or sales tax is charged on it.
        </Paragraph>
      )}
      {deposit.kind !== "security" && invoice.jurisdiction === "UK" && (
        <Paragraph id="advance-payment-tax-uk">
          VAT is charged on this payment now. A payment received before the supply is made creates a tax point at the
          earlier of the date of this invoice and the date the payment is received (VATA 1994 s6(4)).
        </Paragraph>
      )}
      {deposit.kind !== "security" && (invoice.jurisdiction === "EU" || invoice.jurisdiction === "FR") && (
        <Paragraph id="advance-payment-tax-eu">
          VAT is charged on this payment on account now: it becomes chargeable when the payment is received (Art. 65
          Directive 2006/112/EC).
        </Paragraph>
      )}
      {deposit.kind !== "security" && invoice.jurisdiction === "DE" && (
        <Paragraph id="advance-payment-tax-de">
          Anzahlungsrechnung über eine noch nicht ausgeführte Leistung. Die Umsatzsteuer entsteht mit Vereinnahmung der
          Anzahlung (§ 13 Abs. 1 Nr. 1 Buchst. a Satz 4 UStG).
        </Paragraph>
      )}
      {deposit.kind !== "security" && invoice.jurisdiction === "US" && (
        <Paragraph id="advance-payment-tax-us">
          This deposit is refundable and no sales tax is charged on it. Sales tax will be charged on the final invoice.
        </Paragraph>
      )}
    </Section>
  );
};
