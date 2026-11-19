import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import statementOfAccount from "../derivers/statement-of-account.deriver.ts";

/**
 * A balance-forward statement for one period: the opening balance, every
 * transaction dated inside the period with a running balance, the totals, the
 * closing balance, and one sentence saying what that closing balance means.
 *
 * Decisions already made for you:
 *
 * - The opening balance is what you state. It is never rebuilt from older
 *   rows: anything before the period start is assumed to be in it already, and
 *   anything after the statement date belongs to the next statement.
 * - Credits are credits. Invoices and refunds are debits; payments and credit
 *   notes are credits. A credit note is never a negative invoice, so both
 *   columns add up and the totals reconcile.
 * - Same-day entries print in a fixed order: date, then invoice, credit note,
 *   payment, refund, adjustment, then document number. A payment and the
 *   invoice it settles on one day never show the balance dipping into credit.
 * - Rounding happens only when a figure is printed, half away from zero to the
 *   currency's minor unit. The running balance is carried unrounded, so the
 *   last running balance always equals the closing balance.
 * - A supplied closing balance is checked against opening + debits − credits,
 *   and a disagreement of more than half a minor unit throws.
 * - The closing sentence follows the rounded figure: an amount due, "nothing
 *   to pay", or a balance in the customer's favour. A credit prints as an
 *   amount with "CR", never a minus sign.
 * - Mixed currencies throw, and an empty period still prints the opening
 *   balance and says there were no transactions.
 *
 * All of it is one deriver, because sorting, filtering and arithmetic on
 * request data are refused at publish. The node only branches on the result
 * and maps over its rows; a `.map()` over derived rows publishes as a loop the
 * engine walks.
 *
 * Installed by `dxcl add statement-of-account`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface StatementOfAccountData {
  statement: {
    /**
     * ISO 4217 code every amount is in: "GBP", "JPY", "KWD". It also sets how
     * many decimals a printed figure has. Required: the deriver throws when it
     * is missing or not a code `Intl` knows.
     */
    currency: string;
    /**
     * The first day the statement covers, ISO 8601 (YYYY-MM-DD; anything after
     * the tenth character is ignored). Inclusive. Required: throws when missing
     * or not a date.
     */
    periodStart: string;
    /**
     * The last day the statement covers, and the date it is drawn up, ISO 8601.
     * Inclusive. Required: throws when missing, not a date, or before
     * `periodStart`.
     */
    statementDate: string;
    /**
     * The balance brought forward at the start of `periodStart`, in major units
     * (1234.56). Positive means the customer owes it; negative means the account
     * was in credit. Used exactly as given, never rebuilt from entries.
     * Required: throws when it is not a finite number.
     */
    openingBalance: number;
    /**
     * The closing balance your ledger holds, in major units. When given, it is
     * checked against opening + debits − credits, and the deriver throws if
     * they differ by more than half a minor unit. Missing means no check; the
     * computed figure is printed.
     */
    closingBalance?: number;
    /**
     * When an amount owed is due, ISO 8601. Printed on the amount-due line
     * ("payable by 30 November 2026"). Ignored when the account is clear or in
     * credit. Missing means the line gives no date.
     */
    dueBy?: string;
    /**
     * Every document on the account, in any order. Those dated before
     * `periodStart` or after `statementDate` are left out, but every entry is
     * validated first, so a bad one throws even when it is out of the period.
     * Missing or empty means no transactions this period.
     */
    entries?: Array<{
      /**
       * What the document is. Invoices and refunds are debits; payments and
       * credit notes are credits; an adjustment goes by the sign of its amount.
       * Required: throws on any other value.
       */
      type: "invoice" | "credit-note" | "payment" | "refund" | "adjustment";
      /** The document's date, ISO 8601. Required: throws when not a date. */
      date: string;
      /** The document number as printed: "INV-2026-104518". Required: throws when empty. */
      number: string;
      /**
       * The amount in major units. Zero or more for every type except
       * "adjustment", which may be negative (a credit adjustment). A negative
       * amount on any other type throws: a credit note is its own type, not a
       * negative invoice. Required: throws when not a finite number.
       */
      amount: number;
      /**
       * The currency this entry is in. Missing means the statement's currency.
       * Any other value throws: one statement, one currency.
       */
      currency?: string;
      /** Your reference or the customer's: a PO, a remittance reference. Printed in brackets. Missing prints nothing. */
      reference?: string;
      /** What it was for. Printed after the number. Missing prints nothing. */
      description?: string;
    }>;
  };
}

/**
 * Prints the ledger for the period, or the opening balance and a "no
 * transactions" line when there is none, then what the closing balance means.
 */
export const StatementOfAccount: Nodes = async () => {
  const [statement] = useState((data: StatementOfAccountData) => data.statement);
  const s = await useDeriver(statementOfAccount, [statement]);

  return (
    <Section id="statement-of-account" title="Statement of account">
      <Paragraph id="statement-of-account-period">Transactions on your account from {s.periodLabel}.</Paragraph>
      {s.hasEntries && (
        <Table
          id="statement-of-account-ledger"
          // 34mm fits "£12,345,678.90 CR" at body size.
          columns={[{ width: 24 }, { width: "auto" }, { width: 28, align: "right" }, { width: 28, align: "right" }, { width: 34, align: "right" }]}
        >
          <Row id="statement-of-account-ledger-header" header>
            <Cell>Date</Cell>
            <Cell>Details</Cell>
            <Cell>Debit</Cell>
            <Cell>Credit</Cell>
            <Cell>Balance</Cell>
          </Row>
          <Row id="statement-of-account-opening">
            <Cell>{s.openingDate}</Cell>
            <Cell>Balance brought forward</Cell>
            <Cell></Cell>
            <Cell></Cell>
            <Cell>{s.openingBalance}</Cell>
          </Row>
          {s.rows.map((row) => (
            <Row id="statement-of-account-entry">
              <Cell>{row.date}</Cell>
              <Cell>{row.details}</Cell>
              <Cell>{row.debit}</Cell>
              <Cell>{row.credit}</Cell>
              <Cell>{row.balance}</Cell>
            </Row>
          ))}
          <Row id="statement-of-account-totals">
            <Cell></Cell>
            <Cell>Totals</Cell>
            <Cell>{s.totalDebits}</Cell>
            <Cell>{s.totalCredits}</Cell>
            <Cell></Cell>
          </Row>
          <Row id="statement-of-account-closing">
            <Cell>{s.closingDate}</Cell>
            <Cell>Closing balance</Cell>
            <Cell></Cell>
            <Cell></Cell>
            <Cell>{s.closingBalance}</Cell>
          </Row>
        </Table>
      )}
      {!s.hasEntries && (
        <Paragraph id="statement-of-account-empty-opening">
          Balance brought forward on {s.openingDateLong}: {s.openingBalance}.
        </Paragraph>
      )}
      {!s.hasEntries && (
        <Paragraph id="statement-of-account-empty">There were no transactions on your account this period.</Paragraph>
      )}
      {s.outcome === "owing" && (
        <Paragraph id="statement-of-account-owing">Amount due: {s.closingAbs}{s.dueSuffix}.</Paragraph>
      )}
      {s.outcome === "clear" && (
        <Paragraph id="statement-of-account-clear">Your account is clear. There is nothing to pay.</Paragraph>
      )}
      {s.outcome === "credit" && (
        <Paragraph id="statement-of-account-credit">
          Your account is {s.closingAbs} in credit. This balance is in your favour, and there is nothing to pay.
        </Paragraph>
      )}
    </Section>
  );
};
