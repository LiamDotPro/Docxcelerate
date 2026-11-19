import { arrange, Paragraph, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { DirectDebitCollectionNoticeLines } from "../arrangements/direct-debit-collection-notice/lines.ts";
import { stacked } from "../arrangements/direct-debit-collection-notice/stacked.arrangement.tsx";
import { table } from "../arrangements/direct-debit-collection-notice/table.arrangement.tsx";
import directDebitCollectionNotice from "../derivers/direct-debit-collection-notice.deriver.ts";

/**
 * The "how you will pay" block for an invoice collected by direct debit: the
 * advance notice the payer is owed, which says how much is collected, when, by
 * whom and under which mandate. It is worded for the scheme the debit runs on,
 * SEPA Core, UK Bacs or US ACH. Use it instead of `payment-summary`, not
 * alongside it; the document decides which of the two to place.
 *
 * The decision it makes for the reader: the notice never names a collection
 * date the scheme does not allow. A debit that goes through is still
 * non-compliant if its notice was too short. SEPA counts the notice in
 * calendar days (14, or a shorter period agreed with the payer) and moves the
 * date to the next TARGET2 day, with the Easter closing days worked out for the
 * year. Bacs counts working days (10, or an agreed period of at least 3) and
 * skips the bank holidays the data supplies. ACH gives 10 days' notice of
 * amount and date only when the amount differs from the last debit (Reg E
 * 1005.10(d)). When the notice is issued inside the lead time, the requested
 * date is replaced by "on or after" the earliest lawful date. A zero or
 * negative amount says nothing will be collected.
 *
 * Scheme text is never paraphrased. The SEPA notice states the 8-week refund
 * right. A Bacs notice prints the Direct Debit Guarantee only from the approved
 * text the data supplies; if that is missing it prints a visible "incomplete"
 * warning rather than inventing the wording. The SEPA creditor identifier is
 * checked with its check digits, computed without the business code, and one
 * that fails stops the document rather than reaching a payer.
 *
 * Every scheme and outcome test is written inline against `debit.scheme` or
 * the deriver's result, never hoisted into a `const`: while publishing, a
 * hoisted `===` is settled to a plain boolean at build time and every
 * recipient gets the same arm.
 *
 * All calculation is in one deriver, because date arithmetic, rounding and
 * formatting on request data cannot be done at publish time. It runs for every
 * notice, whatever the outcome, so a bad identifier stops a zero-amount notice
 * too.
 *
 * A publish build of this node emits 13 branches against the default
 * `branchLimit` of 32.
 *
 * Arrangement: where the mandate details sit is a separate choice from what
 * they say, made with the `arrangement` prop. This file decides every line;
 * each arrangement, in `arrangements/direct-debit-collection-notice/`, only
 * places them.
 *
 * Installed by `dxcl add direct-debit-collection-notice`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface DirectDebitCollectionNoticeData {
  debit: {
    /**
     * Which scheme collects the payment. `sepa-core` is the SEPA Core scheme
     * (euro), `bacs` is UK Direct Debit (sterling), `ach` is a US ACH debit
     * (dollars). The currency follows the scheme. Required: anything else
     * throws, because each scheme has its own lead time and its own wording.
     */
    scheme: "sepa-core" | "bacs" | "ach";
    /**
     * Who collects, as the payer will see it on their statement:
     * `"Northwind Facilities Ltd"`. Printed as given (trimmed). Required:
     * missing or blank throws.
     */
    creditorName: string;
    /**
     * SEPA only: the creditor identifier, e.g. `"DE98ZZZ09999999999"`. Spaces
     * are removed and letters upper-cased before it is checked and printed.
     * Required for `sepa-core`. If it is missing, has the wrong shape, or
     * fails the check digits, the deriver throws. Ignored for other schemes.
     */
    creditorId?: string;
    /**
     * Bacs only: the six-digit Service User Number. Spaces are removed.
     * Required for `bacs`: missing or anything other than six digits throws.
     * Ignored for other schemes.
     */
    sun?: string;
    /**
     * The mandate reference (SEPA), Direct Debit reference (Bacs) or
     * authorization reference (ACH). Printed as given (trimmed). Required:
     * blank throws. Longer than 35 characters throws for `sepa-core`, and
     * longer than 18 for `bacs`. ACH has no limit here.
     */
    mandateReference: string;
    /**
     * The payer's IBAN. Only its last four characters are ever printed, as
     * "IBAN ending 3456". Spaces are removed and letters upper-cased. Wins
     * over `payerAccountLast4` when both are given. Fewer than four
     * alphanumeric characters throws. Not checksum-validated.
     */
    payerIban?: string;
    /**
     * The last four characters of the payer's account, printed as "Account
     * ending 1234" (or "IBAN ending 1234" under SEPA). Anything other than
     * exactly four letters or digits throws. When neither this nor
     * `payerIban` is given, the account line is left out: the notice is
     * lawful without it.
     */
    payerAccountLast4?: string;
    /**
     * The date the creditor wants to collect, ISO 8601: `"2027-03-26"`.
     * Anything after the tenth character is ignored. Required: missing, not
     * YYYY-MM-DD, or not a real date (`2027-02-30`) throws. It may fall on a
     * closed day or before the notice period ends. The printed date is moved
     * to the earliest lawful open day.
     */
    requestedDate: string;
    /**
     * What is collected, in major units of the scheme's currency: `1284.5`.
     * Rounded to minor units before any comparison. Zero (after rounding)
     * prints "nothing will be collected". Negative means the account is in
     * credit, and also collects nothing. Required: not a finite number throws.
     */
    amount: number;
    /**
     * ACH only: the amount of the last debit under this authorization, in
     * dollars. When it differs from `amount` (compared in cents), Reg E's
     * 10-day notice of amount and date applies and the notice says the amount
     * changed. Missing means "not known to differ": no lead time applies.
     * Not a finite number when present throws. Ignored for other schemes.
     */
    previousAmount?: number;
    /**
     * True when this is the first collection under the mandate. Changes one
     * sentence only. Missing means a recurring collection.
     */
    first?: boolean;
    /**
     * A notice period the payer agreed to in place of the scheme default:
     * calendar days under SEPA (any value replaces 14), working days under
     * Bacs (values below 3 are raised to 3, the Bacs floor, so the notice
     * stays lawful). Ignored for ACH, whose 10 days cannot be agreed away.
     * Missing means the default. Not a non-negative integer throws.
     */
    agreedNoticeDays?: number;
    /**
     * Bank holidays as ISO 8601 dates, on top of weekends. Bacs uses UK bank
     * holidays for the payer's bank's nation. ACH uses Federal Reserve
     * holidays. Required for `bacs` (an empty array means "none in range"):
     * missing throws, because without it the notice period can end a working
     * day early. Optional for `ach`, where missing means weekends only: the
     * lead time there is in calendar days, so only the settlement date shifts.
     * Ignored for SEPA, whose TARGET2 calendar is computed. Any entry that is
     * not an ISO date throws.
     */
    bankHolidays?: string[];
    /**
     * Bacs only: the Direct Debit Guarantee, exactly as approved by your
     * sponsoring bank, one entry per paragraph (include the heading as the
     * first entry if your approved text has one). Printed verbatim, one
     * paragraph per entry. Missing, empty, or containing a blank or
     * non-string entry prints the "Guarantee missing" warning in its place.
     * Ignored for other schemes.
     */
    guaranteeText?: string[];
  };
  document: {
    /**
     * When the notice is issued, ISO 8601. The notice period is counted from
     * this day (the day itself is not counted). Required: missing or not a
     * date throws.
     */
    issuedAt: string;
  };
}

/**
 * Where the mandate details sit.
 *
 * - `stacked`: one labelled line under another. The default.
 * - `table`: labels in a narrow column with the values beside them.
 */
export type DirectDebitCollectionNoticeArrangement = "stacked" | "table";

/** Props for {@linkcode DirectDebitCollectionNotice}. */
export interface DirectDebitCollectionNoticeProps {
  /** Where the mandate details sit. Defaults to `"stacked"`. */
  arrangement?: DirectDebitCollectionNoticeArrangement;
}

export const DirectDebitCollectionNotice: SectionComponent<DirectDebitCollectionNoticeProps> = async (
  { arrangement },
) => {
  const [debit] = useState((data: DirectDebitCollectionNoticeData) => data.debit);
  const [doc] = useState((data: DirectDebitCollectionNoticeData) => data.document);
  const d = await useDeriver(directDebitCollectionNotice, [debit, doc.issuedAt]);

  const lines: DirectDebitCollectionNoticeLines = {
    creditorName: d.creditorName,
    creditorIdLabel: d.creditorIdLabel,
    creditorId: d.creditorId,
    mandateLabel: d.mandateLabel,
    mandateReference: d.mandateReference,
    account: d.account,
  };

  return (
    <Section id="direct-debit-notice" title={d.heading}>
      {d.outcome === "clear" && (
        <Paragraph id="dd-notice-nothing-clear">
          There is nothing to pay on this invoice, so nothing will be collected by {d.schemeName}.
        </Paragraph>
      )}
      {d.outcome === "credit" && (
        <Paragraph id="dd-notice-nothing-credit">
          Your account is {d.amount} in credit, so nothing will be collected by {d.schemeName} for this invoice.
        </Paragraph>
      )}
      {d.outcome === "collect" && !d.late && (
        <Paragraph id="dd-notice-collect-on">
          We will collect {d.amount} from your account by {d.schemeName} on {d.date}.
        </Paragraph>
      )}
      {d.outcome === "collect" && d.late && (
        <Paragraph id="dd-notice-collect-after">
          This notice is issued less than {d.noticeLabel} before the date we planned, so we will collect {d.amount}{" "}
          from your account by {d.schemeName} on or after {d.date}, the earliest date the notice period allows.
        </Paragraph>
      )}
      {d.outcome === "collect" && debit.scheme === "ach" && d.amountChanged && (
        <Paragraph id="dd-notice-amount-changed">
          This amount differs from your last payment of {d.previousAmount}.
        </Paragraph>
      )}
      {d.outcome === "collect" && debit.first && (
        <Paragraph id="dd-notice-first">
          This is the first payment we will collect under this {d.mandateNoun}.
        </Paragraph>
      )}
      {d.outcome === "collect" && !debit.first && (
        <Paragraph id="dd-notice-recurring">
          This is a regular payment under your existing {d.mandateNoun}.
        </Paragraph>
      )}
      {d.outcome === "collect" && (
        <>
          {arrange(arrangement, {
            stacked: () => stacked(lines),
            table: () => table(lines),
          })}
        </>
      )}
      {d.outcome === "collect" && debit.scheme === "sepa-core" && (
        <Paragraph id="dd-notice-sepa-refund">
          You are entitled to a refund from your bank under the terms and conditions of your agreement with your
          bank. A refund must be claimed within 8 weeks starting from the date on which your account was debited.
        </Paragraph>
      )}
      {d.outcome === "collect" && debit.scheme === "bacs" && !d.guaranteeMissing &&
        (debit.guaranteeText ?? []).map((entry) => <Paragraph id="dd-notice-guarantee">{entry}</Paragraph>)}
      {d.outcome === "collect" && debit.scheme === "bacs" && d.guaranteeMissing && (
        <Paragraph id="dd-notice-guarantee-missing">
          Direct Debit Guarantee missing: a Direct Debit advance notice must carry the Guarantee in the wording your
          sponsoring bank approved, and none was supplied. Add the approved text before this notice is issued.
        </Paragraph>
      )}
    </Section>
  );
};
