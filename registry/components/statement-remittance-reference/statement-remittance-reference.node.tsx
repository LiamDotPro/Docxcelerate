import {
  Cell,
  Paragraph,
  Row,
  Section,
  type Section as SectionComponent,
  Table,
  useDeriver,
  useState,
} from "docxcelerate/template";
import statementRemittanceReference from "../derivers/statement-remittance-reference.deriver.ts";

/**
 * The reference a payer quotes when one payment settles several open items,
 * and the tear-off advice that goes with it.
 *
 * Decisions already made for the reader:
 *
 * - Credits net oldest due first. Never newest first, and never against the
 *   largest invoice. The invoice the credits only partly cover is part-paid.
 * - Disputed items never reach the amount to pay. They are listed, unticked.
 * - The reference is chosen in four steps: the one item's own reference, else
 *   a compacted invoice list, else a statement-level RF where the rail takes
 *   one, else the account number with a statement tag. It is folded to the
 *   rail's case and characters before it is measured, not after.
 * - Nothing is truncated: a reference that fits nowhere throws, because a bank
 *   that silently cuts one leaves a payment nobody can match.
 * - KID and OCR rails carry one reference each, so the advice asks for one
 *   payment per item.
 * - Amounts are rounded once, to whole minor units, and the rows are checked
 *   to add up to the net.
 *
 * It is all one deriver, because every step sorts, sums, measures or compares
 * request data, and publishing refuses all of those. Rows and reference lines
 * are `.map()` calls over the deriver's arrays. Since #83 (`resolveRepeat`
 * reads its source by scope) a published loop over `derived` resolves, which is
 * why this component does not need the fixed slots `plan-change-proration` uses.
 *
 * Installed by `dxcl add statement-remittance-reference`, which also installs
 * its deriver.
 */

/** A payment rail the payer's money travels on. Each sets the reference's length, characters and shape. */
export type StatementRail =
  | "bacs"
  | "fps"
  | "sepa"
  | "sepa-rf"
  | "ach-ccd"
  | "ach-ctx"
  | "swift"
  | "nz"
  | "kid"
  | "se-ocr";

/** What this component reads. Add these fields to your document data type. */
export interface StatementRemittanceReferenceData {
  statement: {
    /**
     * ISO 4217 code the amounts are in. It sets the minor units (2 for GBP, 0 for JPY, 3 for BHD).
     * Required: throws when it is missing or `Intl` does not know it. It must match the rail:
     * GBP for bacs and fps; EUR for sepa and sepa-rf; USD for ach-ccd and ach-ctx; NZD for nz;
     * NOK for kid; SEK for se-ocr. Any currency is allowed for swift. A mismatch throws,
     * because a statement on the wrong rail gives a reference the payer's bank refuses.
     */
    currency: string;
    /**
     * The statement date, ISO 8601 (`2026-11-30`). Anything after the tenth character is ignored.
     * It is printed on the advice, and its YYMMDD form is part of the RF and account-tag references.
     * Required: throws when it is missing or not YYYY-MM-DD.
     */
    date: string;
    /**
     * The customer's account number as you print it: `"C-004821"`. It is printed on the advice
     * as given (trimmed), and folded per rail wherever it goes into a reference.
     * Required: throws when it is missing or blank.
     */
    customerNumber: string;
    /**
     * The rail the payer is asked to pay on. It decides the reference's shape.
     * Required: anything else throws.
     */
    rail: StatementRail;
    /**
     * Whether the payee's bank matches an ISO 11649 RF reference quoted on this rail.
     * Read only on bacs, fps, sepa and swift. On sepa-rf it is always treated as true.
     * It is ignored on the other rails. Missing means false.
     */
    acceptsRf?: boolean;
  };
  /**
   * Every open item on the statement, in any order. The deriver sorts them.
   * Missing or empty means there is nothing to pay.
   */
  items: Array<{
    /**
     * What the item is. A `credit-note` and a `payment` (one received but not yet allocated)
     * reduce what is owed. Required: anything else throws.
     */
    type: "invoice" | "credit-note" | "payment";
    /**
     * The document number, printed as given. On list rails, and in every ACH segment,
     * this is what is quoted. Required: throws when it is missing or blank. Two items of the
     * same type with the same number throw, because a list with a duplicate cannot be matched.
     */
    number: string;
    /**
     * The document date, ISO 8601. It is the tie-break after the due date, and the order the
     * credits are applied in. Required: throws when it is missing or not a date.
     */
    date: string;
    /**
     * When an invoice falls due, ISO 8601. Missing means it was due on its `date`.
     * Ignored on credits. Throws when it is present but not a date.
     */
    dueDate?: string;
    /**
     * The amount still open on the item, in major units, never signed. The `type` says which
     * way it goes. It is rounded half away from zero to the currency's minor unit. An amount that
     * rounds to zero drops the item entirely. Throws when it is negative or not a finite number,
     * rather than guessing what a negative credit note means.
     */
    amount: number;
    /**
     * The item's own payment reference, when it has one other than its number. Use it for an
     * RF, a KID or an OCR printed on the invoice.
     * Rule 1 (one item left) quotes it in place of `number`.
     * On sepa-rf it is used only when it is a valid RF.
     * On kid and se-ocr it is required for every item left to pay: there it must be 2 to 25 digits
     * once spaces are removed, or the deriver throws.
     * Not read on the ACH rails or by the list.
     */
    reference?: string;
    /**
     * True for an item under dispute. It is listed on the advice with an unticked box and its
     * amount, but takes no part in netting and is kept out of the amount to pay.
     * Missing means false. Applies to credits too: a disputed credit is not set off.
     */
    disputed?: boolean;
  }>;
}

/**
 * Prints what the payer owes after credits, the reference to quote, and the
 * tear-off remittance advice.
 */
export const StatementRemittanceReference: SectionComponent = async () => {
  const [statement] = useState((data: StatementRemittanceReferenceData) => data.statement);
  const [items] = useState((data: StatementRemittanceReferenceData) => data.items);
  const p = await useDeriver(statementRemittanceReference, [statement, items]);

  return (
    <Section id="statement-remittance" title="Paying this statement">
      {p.outcome === "credit" && (
        <Paragraph id="statement-remittance-credit">Your account is {p.net} in credit. There is nothing to pay.</Paragraph>
      )}
      {p.outcome === "clear" && (
        <Paragraph id="statement-remittance-clear">There is nothing to pay on this statement.</Paragraph>
      )}
      {p.creditsNote && <Paragraph id="statement-remittance-credits">{p.creditsNote}</Paragraph>}
      {p.disputedNote && <Paragraph id="statement-remittance-disputed">{p.disputedNote}</Paragraph>}
      {p.outcome === "pay" && (
        <Paragraph id="statement-remittance-pay">The amount to pay on this statement is {p.net}.</Paragraph>
      )}
      {p.referenceLabel && (
        <Paragraph id="statement-remittance-reference-label">
          Quote this {p.referenceLabel} with your payment, exactly as written:
        </Paragraph>
      )}
      {p.referenceLines.map((line) => <Paragraph id="statement-remittance-reference-line">{line}</Paragraph>)}
      {p.perItemNote && <Paragraph id="statement-remittance-per-item">{p.perItemNote}</Paragraph>}
      {p.outcome === "pay" && (
        <Paragraph id="statement-remittance-cut" align="center">
          - - - - - - - - - - - - - - - - - - - -  Detach here  - - - - - - - - - - - - - - - - - - - -
        </Paragraph>
      )}
      {p.outcome === "pay" && <Paragraph id="statement-remittance-advice-heading">Remittance advice</Paragraph>}
      {p.outcome === "pay" && (
        <Paragraph id="statement-remittance-advice-account">Account number: {p.accountNumber}</Paragraph>
      )}
      {p.outcome === "pay" && (
        <Paragraph id="statement-remittance-advice-date">Statement date: {p.statementDate}</Paragraph>
      )}
      {p.tickInstruction && (
        <Paragraph id="statement-remittance-tick">
          If you are not paying the total, tick the items you are paying and write the amount in the box, so we can
          apply your payment to the right items.
        </Paragraph>
      )}
      {p.outcome === "pay" && (
        <Table
          id="statement-remittance-advice"
          columns={[{ width: 10 }, { width: "auto" }, { width: 28 }, { width: 32, align: "right" }, {
            width: 34,
            align: "right",
          }]}
        >
          <Row id="statement-remittance-advice-header" header>
            <Cell>Pay</Cell>
            <Cell>Item</Cell>
            <Cell>Due</Cell>
            <Cell>Amount</Cell>
            <Cell>Running total</Cell>
          </Row>
          {p.rows.map((row) => (
            <Row id="statement-remittance-advice-row">
              <Cell>☐</Cell>
              <Cell>{row.item}</Cell>
              <Cell>{row.due}</Cell>
              <Cell>{row.amount}</Cell>
              <Cell>{row.running}</Cell>
            </Row>
          ))}
          <Row id="statement-remittance-advice-total">
            <Cell></Cell>
            <Cell>Total to pay</Cell>
            <Cell></Cell>
            <Cell></Cell>
            <Cell>{p.net}</Cell>
          </Row>
          <Row id="statement-remittance-advice-paid">
            <Cell></Cell>
            <Cell>Amount paid</Cell>
            <Cell></Cell>
            <Cell></Cell>
            <Cell></Cell>
          </Row>
        </Table>
      )}
    </Section>
  );
};
