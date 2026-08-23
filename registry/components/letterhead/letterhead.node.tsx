import { type Nodes, Paragraph, useFormat, useState } from "docxcelerate/template";

/**
 * The block at the top of the page saying who sent this and when.
 *
 * Written as loose paragraphs rather than a section, because a letterhead has
 * no heading — it is the thing a heading would sit under. The address is one
 * paragraph with the lines joined, so a sender with two address lines and one
 * with five both come out as a block rather than as a ragged run of nodes.
 *
 * Installed by `dxcl add letterhead`. It is your copy: rename it, re-order it,
 * or drop the date once it is in the project.
 */

/** What this component reads. Add these fields to your document data type. */
export interface LetterheadData {
  sender: {
    /** The organisation, as it should be printed. */
    name: string;
    /** Street, town, postcode — one entry per line. */
    addressLines: string[];
  };
  /** When the document was sent. Anything `Date` can parse. */
  sentOn: string | number | Date;
}

export const Letterhead: Nodes = () => {
  const format = useFormat();
  const [sender] = useState((data: LetterheadData) => ({
    name: data.sender.name,
    address: data.sender.addressLines.join(", "),
    sentOn: format.date(data.sentOn),
  }));

  return (
    <>
      <Paragraph id="letterhead-sender">{sender.name}</Paragraph>
      <Paragraph id="letterhead-address">{sender.address}</Paragraph>
      <Paragraph id="letterhead-date">{sender.sentOn}</Paragraph>
    </>
  );
};
