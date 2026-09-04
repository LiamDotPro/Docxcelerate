import { Paragraph, type Section as SectionComponent, Section, useFormat, useState } from "docxcelerate/template";

/**
 * What is owed, by when, and what happens next.
 *
 * Three outcomes, three ids: in credit, clear, or owing. A recipient who owes
 * nothing should not be shown a payment deadline.
 *
 * Installed by `dxcl add payment-summary`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface PaymentData {
  account: {
    /** Your reference for the account, printed so a caller can quote it. */
    reference: string;
    /** What is owed. Negative means the account is in credit. */
    balanceDue: number;
    /** When it is due. Anything `Date` can parse. Ignored when nothing is owed. */
    dueBy?: string | number | Date;
    /** ISO 4217 code. Defaults to GBP. */
    currency?: string;
  };
}

export const PaymentSummary: SectionComponent = () => {
  const format = useFormat();
  const [account] = useState((data: PaymentData) => data.account);
  const amount = format.currency(Math.abs(account.balanceDue), account.currency ?? "GBP");

  return (
    <Section id="payment-summary" title="Your balance">
      <Paragraph id="payment-reference">
        Please quote {account.reference} on anything you send us about this account.
      </Paragraph>
      {account.balanceDue < 0 && (
        <Paragraph id="payment-in-credit">
          Your account is {amount} in credit. There is nothing to pay, and the
          credit carries over to your next statement.
        </Paragraph>
      )}
      {account.balanceDue === 0 && (
        <Paragraph id="payment-clear">
          Your account is clear. There is nothing to pay.
        </Paragraph>
      )}
      {account.balanceDue > 0 && (
        <Paragraph id="payment-due">
          The balance on your account is {amount}
          {account.dueBy ? `, payable by ${format.date(account.dueBy)}` : ""}.
        </Paragraph>
      )}
    </Section>
  );
};
