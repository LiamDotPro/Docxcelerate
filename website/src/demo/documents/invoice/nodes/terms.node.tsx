import { Paragraph, Section, type Section as SectionComponent, useState } from "docxcelerate/template";
import type { InvoiceData } from "../types.ts";

/** The terms, and who to ask about them. */
export const Terms: SectionComponent = () => {
  const [state] = useState((data: InvoiceData) => ({
    email: data.sender.email,
    lead: data.deliveryLead,
  }));

  return (
    <Section id="terms" title="Terms and notes">
      <Paragraph id="terms-payment">
        Payment within 14 days of the invoice date. Accounts unpaid at 30 days accrue interest
        at 8% above the Bank of England base rate, per the Late Payment of Commercial Debts Act.
      </Paragraph>
      <Paragraph id="terms-contact">
        Send remittance advice to {state.email}. Queries about this invoice go to your delivery
        lead, {state.lead}.
      </Paragraph>
    </Section>
  );
};
