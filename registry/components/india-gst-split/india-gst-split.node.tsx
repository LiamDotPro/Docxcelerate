import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import indiaGstSplit from "../derivers/india-gst-split.deriver.ts";

/**
 * The GST block of an Indian tax invoice: place of supply, the reverse-charge
 * declaration, the HSN/SAC-wise tax summary under the right tax heads, the
 * round-off, the total payable and that total in words.
 *
 * Decisions already made for you:
 *
 * - The tax head comes from comparing two state codes: the supplier's, read
 *   from the first two digits of its GSTIN, and the place of supply. The
 *   billing address and any typed state name play no part.
 * - The same code gets CGST plus SGST at half the rate each, or CGST plus UTGST
 *   in a Union Territory without a legislature (Delhi, Jammu and Kashmir and
 *   Puducherry have one, so they get SGST). A different code gets IGST at the
 *   full rate.
 * - A supply to an SEZ, or an export (place-of-supply code 96), is always
 *   inter-state. Under a letter of undertaking it carries IGST at nil and the
 *   Rule 46 "without payment of integrated tax" mention; without one it carries
 *   IGST at the full rate and the "on payment" mention.
 * - Each half of an intra-state tax is rounded to the paisa on its own, so an
 *   odd paisa makes CGST plus SGST differ by one paisa from IGST on the same
 *   base. That is correct; the base is never split after rounding.
 * - "Reverse charge: Yes/No" always prints. Under reverse charge the tax is
 *   shown but left out of the amount payable.
 * - The total is rounded to the nearest rupee on its own visible line, and is
 *   written in lakh and crore, never million and billion.
 *
 * Everything is decided in one deriver, because GSTIN prefixes, grouping, sums
 * and rounding cannot be computed from request data at publish time. The node
 * only picks a table and interpolates finished strings.
 *
 * Not checked: the GSTIN check digit, whether a GSTIN is live, whether the rate
 * is the right one for the HSN code, and whether the LUT is current.
 *
 * Installed by `dxcl add india-gst-split`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface IndiaGstSplitData {
  seller: {
    /**
     * The supplier's GSTIN, 15 characters. Trimmed and upper-cased, then
     * printed by nothing here: its first two digits are the supplier's state
     * code, and that is the only state this component believes. Required:
     * the deriver throws when it is missing, not 15 characters of [0-9A-Z],
     * or does not start with a known state code (01 to 38, or 97). An
     * unregistered supplier cannot charge GST, so there is no fallback.
     */
    gstin: string;
    /**
     * The supplier's state code as your system holds it, two digits: "27".
     * Optional, and only a cross-check: when given and it differs from the
     * GSTIN's first two digits, the deriver throws rather than pick one.
     */
    stateCode?: string;
  };
  placeOfSupply: {
    /**
     * Two-digit GST state code of the place of supply: "27", "07", or "96"
     * for a supply outside India (an export). Required: throws when missing
     * or not one of 01 to 38, 96 or 97. A number such as 7 is not accepted;
     * pass "07".
     */
    stateCode: string;
    /**
     * The state's name as it should print: "Maharashtra". Printed as
     * "Maharashtra (27)". Trimmed. Required: throws when missing or blank,
     * because Rule 46 makes the place of supply mandatory. Never used to
     * decide anything.
     */
    stateName: string;
  };
  /** Missing entirely means "not an SEZ". */
  customer?: {
    /**
     * Whether the recipient is an SEZ unit or developer receiving for
     * authorised operations. True makes the supply inter-state whatever the
     * state codes say. Missing or anything but `true` means false. Throws
     * with place of supply "96": an SEZ is in India.
     */
    isSez?: boolean;
  };
  supply: {
    /**
     * Whether this zero-rated supply (an export or a supply to an SEZ) is
     * made under a bond or letter of undertaking without payment of IGST.
     * True prints IGST at nil with the LUT mention. Missing means false.
     * Throws when true on a supply that is neither an export nor to an SEZ,
     * because that would print nil tax on a taxable domestic supply.
     */
    exportUnderLut?: boolean;
    /**
     * Whether tax is payable by the recipient on reverse charge. Printed as
     * "Reverse charge: Yes" or "No" on every invoice. True shows the tax but
     * leaves it out of the amount payable. Required, and must be a boolean:
     * the deriver throws when it is missing, because the field is mandatory
     * and guessing "No" is how it goes wrong.
     */
    reverseCharge: boolean;
  };
  /**
   * The invoice's lines, as far as tax is concerned. At least one is
   * required: an empty or missing list throws. Lines sharing an HSN/SAC
   * code and a rate are summed into one summary row.
   */
  lines: Array<{
    /**
     * The HSN (goods) or SAC (services) code, digits only, 4, 6 or 8 long:
     * "84714190", "998314". Trimmed. Throws otherwise, including when a code
     * was stored as a number and lost its leading zero.
     */
    hsnSac: string;
    /**
     * The line's taxable value in paise, an integer: 1234550 is
     * ₹12,345.50. Negative is allowed (a discount line, or a credit note) and
     * is taxed as the mirror image of its positive. Throws when it is not a
     * safe integer.
     */
    taxableMinor: number;
    /**
     * The full GST rate in percent, 0 or above, at most two decimal places:
     * 18, 5, 0.25, 0, 40. Not checked against a list of slabs, because the
     * slabs change. Throws when negative, not finite, or with a third
     * decimal place.
     */
    rate: number;
  }>;
}

/** The GST block of an Indian tax invoice, decided from the place of supply. */
export const IndiaGstSplit: Nodes = async () => {
  const [seller] = useState((data: IndiaGstSplitData) => data.seller);
  const [placeOfSupply] = useState((data: IndiaGstSplitData) => data.placeOfSupply);
  const [customer] = useState((data: IndiaGstSplitData) => data.customer);
  const [supply] = useState((data: IndiaGstSplitData) => data.supply);
  const [lines] = useState((data: IndiaGstSplitData) => data.lines);
  // After every other hook; it runs for every invoice and has no placeholder.
  const split = await useDeriver(indiaGstSplit, [seller, placeOfSupply, customer, supply, lines]);

  return (
    <Section id="india-gst-split" title="GST summary">
      <Paragraph id="gst-split-place-of-supply">Place of supply: {split.placeOfSupply}</Paragraph>
      <Paragraph id="gst-split-reverse-charge">Reverse charge: {split.reverseCharge}</Paragraph>
      {split.intraState
        ? (
          <Table
            id="gst-split-intra"
            columns={[
              { width: 20 },
              { width: "auto", align: "right" },
              { width: 13, align: "right" },
              { width: "auto", align: "right" },
              { width: 13, align: "right" },
              { width: "auto", align: "right" },
              { width: "auto", align: "right" },
            ]}
          >
            <Row id="gst-split-intra-header" header>
              <Cell>HSN/SAC</Cell>
              <Cell>Taxable value</Cell>
              <Cell>CGST rate</Cell>
              <Cell>CGST</Cell>
              <Cell>{split.stateTaxLabel} rate</Cell>
              <Cell>{split.stateTaxLabel}</Cell>
              <Cell>Total tax</Cell>
            </Row>
            {split.rows.map((row) => (
              <Row id="gst-split-intra-row">
                <Cell>{row.hsnSac}</Cell>
                <Cell>{row.taxable}</Cell>
                <Cell>{row.halfRate}</Cell>
                <Cell>{row.centralTax}</Cell>
                <Cell>{row.halfRate}</Cell>
                <Cell>{row.stateTax}</Cell>
                <Cell>{row.totalTax}</Cell>
              </Row>
            ))}
            <Row id="gst-split-intra-total">
              <Cell>Total</Cell>
              <Cell>{split.taxableTotal}</Cell>
              <Cell>{""}</Cell>
              <Cell>{split.centralTaxTotal}</Cell>
              <Cell>{""}</Cell>
              <Cell>{split.stateTaxTotal}</Cell>
              <Cell>{split.taxTotal}</Cell>
            </Row>
          </Table>
        )
        : (
          <Table
            id="gst-split-inter"
            columns={[{ width: 20 }, { width: "auto", align: "right" }, { width: 16, align: "right" }, {
              width: "auto",
              align: "right",
            }]}
          >
            <Row id="gst-split-inter-header" header>
              <Cell>HSN/SAC</Cell>
              <Cell>Taxable value</Cell>
              <Cell>IGST rate</Cell>
              <Cell>IGST</Cell>
            </Row>
            {split.rows.map((row) => (
              <Row id="gst-split-inter-row">
                <Cell>{row.hsnSac}</Cell>
                <Cell>{row.taxable}</Cell>
                <Cell>{row.rate}</Cell>
                <Cell>{row.integratedTax}</Cell>
              </Row>
            ))}
            <Row id="gst-split-inter-total">
              <Cell>Total</Cell>
              <Cell>{split.taxableTotal}</Cell>
              <Cell>{""}</Cell>
              <Cell>{split.integratedTaxTotal}</Cell>
            </Row>
          </Table>
        )}
      {split.zeroRatedMention && <Paragraph id="gst-split-zero-rated">{split.zeroRatedMention}</Paragraph>}
      {split.reverseChargeNote && (
        <Paragraph id="gst-split-reverse-charge-note">{split.reverseChargeNote}</Paragraph>
      )}
      <Paragraph id="gst-split-round-off">Round off: {split.roundOff}</Paragraph>
      <Paragraph id="gst-split-grand-total">Total amount payable: {split.grandTotal}</Paragraph>
      <Paragraph id="gst-split-in-words">Amount in words: {split.grandTotalInWords}</Paragraph>
    </Section>
  );
};
