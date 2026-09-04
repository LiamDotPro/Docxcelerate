import { Image, type Nodes, Paragraph, useState } from "docxcelerate/template";

/**
 * How a letter ends: a closing line, a signature, and who signed it.
 *
 * The image is optional and the name is not — a letter signed by nobody is a
 * letter nobody owns, whereas a missing image is a rendering detail.
 *
 * Installed by `dxcl add signature-block`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface SignatoryData {
  signatory: {
    /** The person signing, as they should be printed. */
    name: string;
    /** Their role, printed beneath the name. */
    role: string;
    /** Path to a signature image, relative to the document project. */
    signatureImage?: string;
    /** How the letter closes. Defaults to "Yours sincerely". */
    closing?: string;
  };
}

export const SignatureBlock: Nodes = () => {
  const [signatory] = useState((data: SignatoryData) => data.signatory);

  return (
    <>
      <Paragraph id="closing">{signatory.closing ?? "Yours sincerely"},</Paragraph>
      {signatory.signatureImage && (
        <Image
          id="signature-image"
          src={signatory.signatureImage}
          alt={`Signature of ${signatory.name}`}
          width={160}
          height={54}
        />
      )}
      <Paragraph id="signatory-name">{signatory.name}</Paragraph>
      <Paragraph id="signatory-role">{signatory.role}</Paragraph>
    </>
  );
};
