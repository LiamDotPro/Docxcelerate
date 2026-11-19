import { type Nodes, Paragraph, useDeriver, useState } from "docxcelerate/template";
import structuredPaymentReference from "../derivers/structured-payment-reference.deriver.ts";

/**
 * The structured payment reference a bank matches a payment against, with
 * its check digits worked out, printed the way that scheme groups it, or an
 * instruction to quote the invoice number when no correct reference can be given.
 *
 * Bases are strings, and must stay strings. A number drops the leading zeros
 * a Belgian OGM and a Swiss QR reference count, and a 26-digit QR reference
 * is past 2^53, so arithmetic on it gives check digits that look right and
 * are not. Every check is worked digit by digit on the string.
 *
 * A base too long for its scheme is refused, never truncated: a truncated
 * reference still validates, and books the payment against the wrong invoice.
 * A reference the ERP already issued is validated rather than trusted.
 *
 * A Norwegian MOD11 KID whose check digit comes out as "-" is treated as
 * unissuable: many banks reject it, so the invoice number is printed instead.
 * Every failure lands on the same fallback, which always gets a payment
 * booked, even if by hand. Each outcome has its own id, so a resolved document
 * records which one this recipient got.
 *
 * All of it is one deriver, because computing on request data cannot be
 * published any other way. The node only chooses between three lines.
 *
 * Not checked: whether the reference is registered with the bank, or that a
 * KID or OCR length matches the creditor's agreement.
 *
 * Installed by `dxcl add structured-payment-reference`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface StructuredPaymentReferenceData {
  /**
   * How this invoice is to be referenced when paid. Missing (undefined or null)
   * means no structured reference: the fallback line is printed.
   */
  payment?: {
    /**
     * Which scheme the reference follows.
     *
     * - `rf`: ISO 11649 creditor reference, any country.
     * - `no-kid-mod10` / `no-kid-mod11`: Norwegian KID, Luhn or MOD11 check.
     * - `fi`: Finnish national reference number (viitenumero).
     * - `fi-rf`: the Finnish number in its international RF form.
     * - `be-ogm`: Belgian structured communication (OGM/VCS).
     * - `ch-qrr`: Swiss QR reference, for a QR-IBAN.
     * - `se-ocr` / `se-ocr-length`: Swedish Bankgiro OCR, without or with the length digit.
     *
     * Required when `payment` is given. Anything else throws, because a
     * mistyped scheme would otherwise fall back silently on every invoice.
     */
    scheme:
      | "rf"
      | "no-kid-mod10"
      | "no-kid-mod11"
      | "fi"
      | "fi-rf"
      | "be-ogm"
      | "ch-qrr"
      | "se-ocr"
      | "se-ocr-length";
    /**
     * What the check digits are generated from: an invoice or customer number,
     * as a string. Whitespace is ignored. Digits only, except for `rf`, which
     * also takes letters (case does not matter). A string that breaks the
     * scheme's rules, or is too long for it, prints the fallback; it is never
     * truncated. Missing or empty, with no `reference` either, prints the
     * fallback. A number throws: it has already lost its leading zeros.
     * Ignored when `reference` is given.
     */
    base?: string;
    /**
     * A full reference the ERP already issued, check digits included. It is
     * validated and printed in the scheme's grouping, not recalculated.
     * Whitespace is ignored, and for `be-ogm` so are `+`, `*` and `/`. A
     * reference that fails its check prints the fallback. Takes precedence
     * over `base`. A number throws.
     */
    reference?: string;
  };
  invoice: {
    /**
     * The invoice number, printed as given in the fallback line. Required:
     * the deriver throws when it is missing or blank, whatever the outcome,
     * so a document that needs the fallback can never be missing it.
     */
    number: string;
  };
}

/**
 * Prints one line: the payment reference in its scheme's grouping, or the
 * quote-the-invoice-number fallback. It has no heading, so it can sit inside
 * the host's payment block.
 */
export const StructuredPaymentReference: Nodes = async () => {
  const [payment] = useState((data: StructuredPaymentReferenceData) => data.payment);
  const [invoice] = useState((data: StructuredPaymentReferenceData) => data.invoice);
  const ref = await useDeriver(structuredPaymentReference, [payment, invoice]);

  return (
    <>
      {ref.issued && <Paragraph id="structured-reference-issued">{ref.issued}</Paragraph>}
      {ref.invalid && <Paragraph id="structured-reference-invalid">{ref.invalid}</Paragraph>}
      {ref.unissuable && <Paragraph id="structured-reference-unissuable">{ref.unissuable}</Paragraph>}
    </>
  );
};
