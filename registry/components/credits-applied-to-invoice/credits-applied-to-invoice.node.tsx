import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import creditsAppliedToInvoice from "../derivers/credits-applied-to-invoice.deriver.ts";

/**
 * An invoice section that sets a customer's open credits (credit notes and
 * payments on account) against the invoice and prints the reconciliation from
 * the invoice total to the amount due.
 *
 * Decisions already made for you:
 *
 * - Credits reduce the amount due, never the invoice total or its VAT. A credit
 *   note adjusted VAT when it was issued and a payment on account is only a
 *   means of payment, so applying either before VAT would reduce the tax base a
 *   second time. The section belongs below the VAT analysis and its first line
 *   repeats the VAT-inclusive total unchanged.
 * - Credits are taken oldest first and each is capped at what is still owed, or
 *   in the order of an explicit `credits.allocation` when the caller gives one.
 * - Whatever is left of each credit is printed as carried forward.
 * - Only credits in the invoice's own currency are applied. A foreign credit is
 *   listed as not applied, in its own currency, never converted.
 * - Exclusions are judged in a fixed order: expired first (an expired credit is
 *   gone whatever else is true of it), then disputed, then a different currency.
 *   Each excluded credit is named with its one reason.
 * - All arithmetic is on integer minor units, so total minus the sum applied is
 *   the amount due exactly, and it can never be negative.
 *
 * All of it is one deriver, because sorting, `min` and subtraction on request
 * data all throw at publish. The node computes nothing: it interpolates, loops
 * over the rows the deriver returns and branches on what it says.
 *
 * The loops run over derived rows, which needs a resolver that resolves a
 * repeat's source by scope; the framework commit that adds it must be present.
 *
 * Installed by `dxcl add credits-applied-to-invoice`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface CreditsAppliedToInvoiceData {
  invoice: {
    /**
     * ISO 4217 code of the invoice: "GBP", "EUR", "JPY". Sets the exponent
     * (two for GBP, none for JPY) from Intl. Only credits in this currency are
     * applied. Required: the deriver throws, naming `invoice.currency`, when it
     * is missing or a code Intl rejects.
     */
    currency: string;
    /**
     * The invoice's date of issue, ISO 8601 (`2026-11-30`). Anything after the
     * tenth character is ignored. Expiry is judged against this day, never
     * against "today", so the same data always gives the same document.
     * Required: throws when missing or not YYYY-MM-DD.
     */
    issueDate: string;
    /**
     * The invoice total including VAT, in integer minor units (128475000 for
     * £1,284,750.00), exactly as the VAT analysis above printed it. This
     * component never recalculates or reduces it. Required: throws when it is
     * not a safe integer ≥ 0. A negative total is a credit note, not an invoice.
     */
    totalMinor: number;
    /**
     * The VAT included in `totalMinor`, in minor units. When given, a sentence
     * says it is unchanged by the credits. Missing: no VAT sentence. Throws when
     * present but not a safe integer between 0 and `totalMinor`.
     */
    vatTotalMinor?: number;
  };
  credits: {
    /**
     * Every open credit on the customer's account, in any order. Missing,
     * empty, or every entry at `remainingMinor` 0 means there is nothing to
     * reconcile, and the component prints nothing at all.
     */
    open?: Array<{
      /**
       * The credit's reference, printed as given ("CN-2026-000418"), and what
       * an allocation names. Required, non-blank and unique: a missing or
       * duplicate id throws.
       */
      id: string;
      /**
       * `credit-note`: an issued credit note, which already adjusted VAT.
       * `payment-on-account`: money received and not yet applied. The two are
       * printed in separate groups. Anything else throws.
       */
      kind: "credit-note" | "payment-on-account";
      /**
       * When it was issued or received, ISO 8601. Sets the oldest-first order
       * and is printed. Required: throws when missing or not a date.
       */
      issuedOn: string;
      /**
       * ISO 4217 code of the credit. Compared with `invoice.currency` trimmed
       * and upper-cased. A different code is listed "Not applied – different
       * currency" and its amount is shown in its own currency, never converted.
       * Required: throws when missing or a code Intl rejects.
       */
      currency: string;
      /**
       * What is left of the credit, in integer minor units of its own currency.
       * 0 means used up: it is omitted entirely, not listed. Throws when it is
       * not a safe integer ≥ 0.
       */
      remainingMinor: number;
      /**
       * The last day the credit can be used, ISO 8601. A credit expiring on the
       * issue date is still applied; one expiring before it is excluded and
       * listed as expired. Missing: it never expires. Throws when present but
       * not a date.
       */
      expiresOn?: string;
      /**
       * True when the credit is blocked by a dispute. It is excluded and listed
       * "On hold – disputed". Missing or anything but `true` means not disputed.
       */
      disputed?: boolean;
    }>;
    /**
     * An explicit allocation from your ledger, applied in this order in place
     * of oldest first. Each amount is capped at the credit's remainder and at
     * what is still owed. Eligible credits not named are not applied and carry
     * forward in full. Missing or empty: oldest first. An entry naming an
     * unknown id, naming an id twice, or with an amount that is not a safe
     * integer ≥ 0 throws. One naming an excluded credit (expired, disputed,
     * other currency) applies nothing; the credit stays listed with its reason.
     */
    allocation?: Array<{ creditId: string; amountMinor: number }>;
  };
}

/**
 * Prints the credits applied to an invoice and the reconciliation from total to
 * amount due. Prints nothing when the customer has no open credit.
 */
export const CreditsAppliedToInvoice: Nodes = async () => {
  const [invoice] = useState((data: CreditsAppliedToInvoiceData) => data.invoice);
  const [credits] = useState((data: CreditsAppliedToInvoiceData) => data.credits);
  const p = await useDeriver(creditsAppliedToInvoice, [invoice, credits]);

  return (
    <>
      {p.show && (
        <Section id="credits-applied" title="Credits applied">
          <Paragraph id="credits-applied-intro">{p.intro}</Paragraph>
          {p.hasCreditNotes && (
            <Table
              id="credits-applied-credit-notes"
              columns={[{ width: "auto" }, { width: 34 }, { width: 30, align: "right" }, { width: 34, align: "right" }]}
            >
              <Row id="credits-applied-credit-notes-header" header>
                <Cell>Credit note</Cell>
                <Cell>Issued</Cell>
                <Cell>Applied</Cell>
                <Cell>Carried forward</Cell>
              </Row>
              {p.creditNotes.map((row) => (
                <Row id="credits-applied-credit-note">
                  <Cell>{row.reference}</Cell>
                  <Cell>{row.issued}</Cell>
                  <Cell>{row.applied}</Cell>
                  <Cell>{row.carriedForward}</Cell>
                </Row>
              ))}
            </Table>
          )}
          {p.hasPayments && (
            <Table
              id="credits-applied-payments"
              columns={[{ width: "auto" }, { width: 34 }, { width: 30, align: "right" }, { width: 34, align: "right" }]}
            >
              <Row id="credits-applied-payments-header" header>
                <Cell>Payment on account</Cell>
                <Cell>Received</Cell>
                <Cell>Applied</Cell>
                <Cell>Carried forward</Cell>
              </Row>
              {p.payments.map((row) => (
                <Row id="credits-applied-payment">
                  <Cell>{row.reference}</Cell>
                  <Cell>{row.issued}</Cell>
                  <Cell>{row.applied}</Cell>
                  <Cell>{row.carriedForward}</Cell>
                </Row>
              ))}
            </Table>
          )}
          {p.hasExcluded && (
            <Table
              id="credits-applied-excluded"
              columns={[{ width: "auto" }, { width: 34 }, { width: 30, align: "right" }, { width: 34 }]}
            >
              <Row id="credits-applied-excluded-header" header>
                <Cell>Not applied</Cell>
                <Cell>Type</Cell>
                <Cell>Amount</Cell>
                <Cell>Reason</Cell>
              </Row>
              {p.excluded.map((row) => (
                <Row id="credits-applied-excluded-credit">
                  <Cell>{row.reference}</Cell>
                  <Cell>{row.type}</Cell>
                  <Cell>{row.amount}</Cell>
                  <Cell>{row.reason}</Cell>
                </Row>
              ))}
            </Table>
          )}
          <Table id="credits-applied-reconciliation" columns={[{ width: "auto" }, { width: 40, align: "right" }]}>
            <Row id="credits-applied-total">
              <Cell>Invoice total</Cell>
              <Cell>{p.total}</Cell>
            </Row>
            {p.lessCreditNotes && (
              <Row id="credits-applied-less-credit-notes">
                <Cell>Less credit notes applied</Cell>
                <Cell>{p.lessCreditNotes}</Cell>
              </Row>
            )}
            {p.lessPayments && (
              <Row id="credits-applied-less-payments">
                <Cell>Less payments on account applied</Cell>
                <Cell>{p.lessPayments}</Cell>
              </Row>
            )}
            <Row id="credits-applied-due">
              <Cell>Amount due</Cell>
              <Cell>{p.due}</Cell>
            </Row>
          </Table>
          {p.vat && (
            <Paragraph id="credits-applied-vat">
              The VAT on this invoice, {p.vat}, is unchanged: credit notes adjusted VAT when they were issued, and a
              payment on account is a means of payment.
            </Paragraph>
          )}
          {p.outcome === "covered" && (
            <Paragraph id="credits-applied-covered">Nothing to pay – covered by credits.</Paragraph>
          )}
          {p.outcome === "part" && (
            <Paragraph id="credits-applied-part">The amount due after credits is {p.due}.</Paragraph>
          )}
          {p.carriedForward && (
            <Paragraph id="credits-applied-carried-forward">
              Credit of {p.carriedForward} remains on your account and is carried forward to your next invoice.
            </Paragraph>
          )}
        </Section>
      )}
    </>
  );
};
