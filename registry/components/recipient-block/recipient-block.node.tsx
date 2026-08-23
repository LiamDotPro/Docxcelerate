import { type Nodes, Paragraph, useState } from "docxcelerate/template";

/**
 * Who the document is addressed to, and how it greets them.
 *
 * The salutation is the interesting half. A document generated in bulk meets
 * recipients whose names it does not have — a joint tenancy, a company, a
 * record where the field was never filled in — and "Dear ," is the kind of
 * mistake that gets screenshotted. So the greeting branches, and the fallback
 * is deliberately formal rather than clever.
 *
 * Installed by `dxcl add recipient-block`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface RecipientData {
  recipient: {
    /** The addressee. Leave empty where you genuinely do not have a name. */
    name?: string;
    /** Street, town, postcode — one entry per line. */
    addressLines: string[];
    /** How to greet them, when a first name would be too familiar. */
    formalName?: string;
  };
}

export const RecipientBlock: Nodes = () => {
  const [recipient] = useState((data: RecipientData) => data.recipient);
  const greeting = recipient.formalName ?? recipient.name;

  return (
    <>
      {recipient.name && <Paragraph id="recipient-name">{recipient.name}</Paragraph>}
      <Paragraph id="recipient-address">{recipient.addressLines.join(", ")}</Paragraph>
      {greeting
        ? <Paragraph id="salutation">Dear {greeting},</Paragraph>
        : <Paragraph id="salutation-unnamed">Dear Sir or Madam,</Paragraph>}
    </>
  );
};
