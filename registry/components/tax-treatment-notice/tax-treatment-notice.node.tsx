import { derive } from "docxcelerate/document";
import { type Nodes, Paragraph, useState } from "docxcelerate/template";

/**
 * The one mention an invoice's tax treatment requires, or a warning when it
 * cannot be given. Nothing at all for `standard`, and no heading, so it can
 * sit under the totals on every invoice.
 *
 * Decisions already made for you:
 *
 * - A reverse-charge or exempt supply never prints "0% VAT". That reads as
 *   zero-rating, a different legal treatment, so each treatment prints its own
 *   statutory mention.
 * - Where an exemption only holds with an identifier (the customer's VAT
 *   number, both VAT numbers, a resale certificate), a missing identifier
 *   prints an "incomplete" warning instead of the mention. An invoice that
 *   claims an exemption it cannot support is worse than one that visibly says
 *   it is unfinished. An unrecognised treatment warns too: a missing mention
 *   is how an exemption is lost.
 * - The UK construction reverse charge states the VAT that would have been
 *   due, for information; it is not charged and not part of the amount payable.
 * - The French and German small-business wording is printed exactly as
 *   written, never paraphrased.
 *
 * Every test is written inline against `invoice.taxTreatment` and friends,
 * not hoisted into a `const`: while publishing, a hoisted `===` is settled to
 * a plain boolean at build time and every recipient gets the same arm.
 *
 * Nothing is computed here except the CIS amount, which is formatted by a
 * deriver, installed at `derivers/tax-treatment-notice.deriver.ts`. The
 * installer registers only node exports, so add it to `derivers/index.ts` by
 * hand. Derivers run before a node's `when` is checked, so a document using
 * this node fails to build naming it whatever its treatment, not only on the
 * reverse-charge arm:
 * `import taxTreatmentNoticeVat from "./tax-treatment-notice.deriver.ts";`
 * and `taxTreatmentNoticeVat: taxTreatmentNoticeVat.run,` in the object.
 *
 * A publish build of this node emits 15 branches against the default
 * `branchLimit` of 32, so a document that already carries many decisions may
 * need `branchLimit` raised. The arms are not collapsed to save branches: the
 * per-arm ids are what let a resolved document record which mention a
 * recipient got.
 *
 * Not checked: that a VAT number is well-formed or live, or that a certificate
 * is valid. Presence is the only test, so `""` and `0` count as missing.
 *
 * Installed by `dxcl add tax-treatment-notice`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface TaxTreatmentData {
  invoice: {
    /**
     * Which VAT or sales-tax treatment the invoice is issued under.
     *
     * `standard` prints nothing. A value outside the union, such as a typo or
     * a treatment added upstream and not here, prints the "unrecognised"
     * warning rather than nothing, because a missing mention is how an
     * exemption is lost.
     */
    taxTreatment:
      | "standard"
      | "zero-rated-export"
      | "intra-community-supply"
      | "reverse-charge-services"
      | "uk-domestic-reverse-charge"
      | "small-business-exempt"
      | "us-resale-exempt"
      | "margin-scheme";
    /**
     * Where the seller is established, as an upper-case ISO 3166-1 alpha-2
     * code: `"FR"`, `"DE"`. Only read for `small-business-exempt`. It is
     * compared exactly, so `"fr"` does not match. Missing, or any value other
     * than FR or DE, prints the small-business "incomplete" warning.
     */
    jurisdiction?: string;
    /**
     * The export evidence reference for `zero-rated-export`, typically the
     * customs MRN. Printed after the mention when given. Missing prints the
     * mention alone: the evidence has to be kept on record, but the invoice
     * does not have to carry it.
     */
    exportEvidence?: string;
    /**
     * For `uk-domestic-reverse-charge` only: the VAT that would have been
     * due, in pounds sterling, already rounded the way the seller rounds VAT.
     * It is printed and never added to anything. Missing or 0 prints the
     * "incomplete" warning, because HMRC wants the amount stated, and the
     * domestic reverse charge never applies to a zero-rated supply. A negative
     * figure (a credit note) is printed with its sign.
     */
    notionalVat?: number;
  };
  seller: {
    /**
     * The seller's VAT number, printed as given. Required for
     * `reverse-charge-services`. Missing there prints the "incomplete" warning.
     */
    vatNumber?: string;
  };
  customer: {
    /**
     * The customer's VAT number, printed as given. Required for
     * `intra-community-supply` and `reverse-charge-services`. Missing or empty
     * there prints the "incomplete" warning.
     */
    vatNumber?: string;
    /**
     * The customer's resale exemption certificate, for `us-resale-exempt`.
     * Both fields are required. If either is missing or empty, the
     * "incomplete" warning prints instead.
     */
    resaleCertificate?: {
      /** The certificate or permit number, printed as given. */
      number: string;
      /** The state that issued it, printed as given: `"California"` or `"CA"`. */
      state: string;
    };
  };
}

/**
 * The legal tax-treatment notice for an invoice: one paragraph chosen by
 * `invoice.taxTreatment`, or an "incomplete" warning when the data that
 * treatment requires is missing.
 */
export const TaxTreatmentNotice: Nodes = () => {
  const [{ invoice, seller, customer }] = useState((data: TaxTreatmentData) => ({
    invoice: data.invoice,
    seller: data.seller,
    customer: data.customer,
  }));
  const t = invoice.taxTreatment;
  const cv = customer.vatNumber;
  const sv = seller.vatNumber;
  const rc = customer.resaleCertificate;
  // A comparison cannot carry `undefined`, so a missing jurisdiction is read as
  // "" and falls to the incomplete arm. A publish stand-in is never nullish.
  const jurisdiction = invoice.jurisdiction ?? "";

  return (
    <>
      {t === "zero-rated-export" && invoice.exportEvidence && (
        <Paragraph id="tax-notice-export-evidenced">
          Zero-rated export of goods. Export evidence: {invoice.exportEvidence}.
        </Paragraph>
      )}
      {t === "zero-rated-export" && !invoice.exportEvidence && (
        <Paragraph id="tax-notice-export">Zero-rated export of goods.</Paragraph>
      )}
      {t === "intra-community-supply" && cv && (
        <Paragraph id="tax-notice-intra-community">
          Exempt intra-Community supply (Art 138 Directive 2006/112/EC). Customer VAT number: {cv}.
        </Paragraph>
      )}
      {t === "intra-community-supply" && !cv && (
        <Paragraph id="tax-notice-intra-community-incomplete">
          Tax treatment incomplete: an exempt intra-Community supply must show the customer's VAT number, and none was given. Without it the exemption does not apply. Add the number before this invoice is issued.
        </Paragraph>
      )}
      {t === "reverse-charge-services" && sv && cv && (
        <Paragraph id="tax-notice-reverse-charge">
          Reverse charge (Art 196 Directive 2006/112/EC). Supplier VAT number: {sv}. Customer VAT number: {cv}.
        </Paragraph>
      )}
      {t === "reverse-charge-services" && !(sv && cv) && (
        <Paragraph id="tax-notice-reverse-charge-incomplete">
          Tax treatment incomplete: a reverse-charge supply must show both the supplier's and the customer's VAT numbers, and at least one is missing. Add them before this invoice is issued.
        </Paragraph>
      )}
      {t === "uk-domestic-reverse-charge" && invoice.notionalVat && (
        <Paragraph
          id="tax-notice-uk-reverse-charge"
          derivers={[
            derive("taxTreatmentNoticeVat", {
              output: "taxTreatmentNoticeVatLabel",
              inputs: [invoice.notionalVat],
            }),
          ]}
        >
          Reverse charge: customer to account to HMRC for the VAT (VATA 1994 s55A). VAT that would have been due: {"{{derived.taxTreatmentNoticeVatLabel}}"}. It is shown for information and is not included in the amount payable.
        </Paragraph>
      )}
      {t === "uk-domestic-reverse-charge" && !invoice.notionalVat && (
        <Paragraph id="tax-notice-uk-reverse-charge-incomplete">
          Tax treatment incomplete: the UK domestic reverse charge must state the VAT that would have been due, and no amount was given. Add it before this invoice is issued.
        </Paragraph>
      )}
      {t === "small-business-exempt" && jurisdiction === "FR" && (
        <Paragraph id="tax-notice-small-business-fr">TVA non applicable, art. 293 B du CGI</Paragraph>
      )}
      {t === "small-business-exempt" && jurisdiction === "DE" && (
        <Paragraph id="tax-notice-small-business-de">Gemäß § 19 UStG wird keine Umsatzsteuer berechnet</Paragraph>
      )}
      {t === "small-business-exempt" && jurisdiction !== "FR" && jurisdiction !== "DE" && (
        <Paragraph id="tax-notice-small-business-incomplete">
          Tax treatment incomplete: this invoice is marked small-business exempt, but its jurisdiction is not FR or DE, the two this mention is written for. Add the wording that jurisdiction requires before this invoice is issued.
        </Paragraph>
      )}
      {t === "us-resale-exempt" && rc?.number && rc?.state && (
        <Paragraph id="tax-notice-resale">
          Exempt: sale for resale. Resale certificate {rc.number}, issued in {rc.state}.
        </Paragraph>
      )}
      {t === "us-resale-exempt" && !(rc?.number && rc?.state) && (
        <Paragraph id="tax-notice-resale-incomplete">
          Tax treatment incomplete: a sale for resale must show the customer's resale certificate number and the state that issued it, and at least one is missing. Add them before this invoice is issued.
        </Paragraph>
      )}
      {t === "margin-scheme" && (
        <Paragraph id="tax-notice-margin-scheme">Margin scheme — Second-hand goods</Paragraph>
      )}
      {t !== "standard" &&
        t !== "zero-rated-export" &&
        t !== "intra-community-supply" &&
        t !== "reverse-charge-services" &&
        t !== "uk-domestic-reverse-charge" &&
        t !== "small-business-exempt" &&
        t !== "us-resale-exempt" &&
        t !== "margin-scheme" && (
        <Paragraph id="tax-notice-unrecognised">
          Tax treatment incomplete: this invoice's tax treatment is not one this notice recognises, so no mention has been printed.
        </Paragraph>
      )}
    </>
  );
};
