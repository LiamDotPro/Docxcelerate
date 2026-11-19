import {
  defineMessages,
  Paragraph,
  type Section as SectionComponent,
  Section,
  useFormat,
  useMessages,
  useState,
} from "docxcelerate/template";

/**
 * What is owed, by when, and what happens next.
 *
 * Three outcomes, three ids: in credit, clear, or owing. Branching on the
 * amount rather than papering over it with one sentence that reads oddly at
 * zero is the whole point — a recipient who owes nothing should not be told
 * about a payment deadline, and a recipient in credit should be told plainly
 * that they are.
 *
 * Written in English, with French and German translations: the document's
 * locale picks one, and the amounts and dates follow it. Add a language by
 * adding an entry to `messages` — the type-checker says what it has to cover.
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

const messages = defineMessages("en", {
  title: "Your balance",
  reference: (reference: string) =>
    `Please quote ${reference} on anything you send us about this account.`,
  inCredit: (amount: string) =>
    `Your account is ${amount} in credit. There is nothing to pay, and the ` +
    "credit carries over to your next statement.",
  clear: "Your account is clear. There is nothing to pay.",
  due: (amount: string, dueBy?: string) =>
    `The balance on your account is ${amount}${dueBy ? `, payable by ${dueBy}` : ""}.`,
}, {
  fr: {
    title: "Votre solde",
    reference: (reference) =>
      `Merci de rappeler la référence ${reference} dans toute correspondance concernant ce compte.`,
    inCredit: (amount) =>
      `Votre compte est créditeur de ${amount}. Vous n’avez rien à payer, et ce ` +
      "crédit sera reporté sur votre prochain relevé.",
    clear: "Votre compte est à jour. Vous n’avez rien à payer.",
    due: (amount, dueBy) =>
      `Le solde de votre compte s’élève à ${amount}${dueBy ? `, à régler au plus tard le ${dueBy}` : ""}.`,
  },
  de: {
    title: "Ihr Kontostand",
    reference: (reference) =>
      `Bitte geben Sie bei allen Zuschriften zu diesem Konto die Referenz ${reference} an.`,
    inCredit: (amount) =>
      `Ihr Konto weist ein Guthaben von ${amount} auf. Es ist nichts zu zahlen, und ` +
      "das Guthaben wird auf Ihre nächste Abrechnung übertragen.",
    clear: "Ihr Konto ist ausgeglichen. Es ist nichts zu zahlen.",
    due: (amount, dueBy) =>
      `Der offene Betrag auf Ihrem Konto beträgt ${amount}${dueBy ? `, zahlbar bis zum ${dueBy}` : ""}.`,
  },
});

export const PaymentSummary: SectionComponent = () => {
  const format = useFormat();
  const t = useMessages(messages);
  const [account] = useState((data: PaymentData) => data.account);
  const amount = format.currency(Math.abs(account.balanceDue), account.currency ?? "GBP");

  return (
    <Section id="payment-summary" title={t.title}>
      <Paragraph id="payment-reference">{t.reference(account.reference)}</Paragraph>
      {account.balanceDue < 0 && (
        <Paragraph id="payment-in-credit">{t.inCredit(amount)}</Paragraph>
      )}
      {account.balanceDue === 0 && (
        <Paragraph id="payment-clear">{t.clear}</Paragraph>
      )}
      {account.balanceDue > 0 && (
        <Paragraph id="payment-due">
          {t.due(amount, account.dueBy ? format.date(account.dueBy) : undefined)}
        </Paragraph>
      )}
    </Section>
  );
};
