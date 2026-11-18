import {
  arrange,
  Cell,
  Paragraph,
  Row,
  Section,
  type Section as SectionComponent,
  Table,
  useDeriver,
  useState,
} from "docxcelerate/template";
import cisInvoiceDeduction from "../derivers/cis-invoice-deduction.deriver.ts";

/**
 * What a construction subcontractor is actually paid on one invoice, and the
 * working that shows it. Every usual mistake here gives a payable figure that
 * HMRC or the contractor will dispute, so the decisions are made once.
 *
 * Decisions already made for you:
 *
 * - CIS comes off the labour only: the ex-VAT gross less materials at cost,
 *   plant hire and consumables. Never off materials, and never off VAT.
 * - VAT is worked out on the whole gross and added before CIS comes off. CIS
 *   never reduces it, and it never reduces CIS.
 * - Silence from the customer means the reverse charge applies. Only written
 *   end-user or intermediary status takes the invoice out of it.
 * - The 5% disregard: when reverse-charge services are 5% or less of the
 *   invoice the whole invoice is standard-rated. Zero-rated work is never
 *   reverse-charged.
 * - The deduction is floored to the penny, so it never over-deducts.
 * - An unverified subcontractor is deducted at 30% and the document prints a
 *   warning to verify them.
 * - Unclear data fails the document rather than guessing. A guessed treatment
 *   gives a disputed figure either way.
 * - Everything is settled in one deriver, because every figure is arithmetic
 *   on request data and so cannot run at publish.
 *
 * Under the reverse charge the notional VAT is returned by the deriver for
 * `tax-treatment-notice`, which reads it from request data: the calling
 * application passes the same figure.
 *
 * Arrangement: where the lines sit is a separate choice from what they say.
 * Every arrangement prints the same lines under the same conditions.
 *
 * Installed by `dxcl add cis-invoice-deduction`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface CisInvoiceDeductionData {
  invoice: {
    /**
     * The invoice value before VAT, materials included: what VAT is worked out
     * on. Required: the deriver throws when it is missing, not a finite number,
     * or negative. A credit note is not handled here.
     */
    grossExVat: number;
    /**
     * What the subcontractor paid for materials, at cost and net of any VAT
     * they can reclaim. A subcontractor not registered for VAT includes the VAT
     * they paid. Any markup belongs in labour, not here. Required, and 0 when
     * there are none: the deriver throws when it is missing or negative, and
     * when materials, plant hire and consumables together exceed `grossExVat`.
     */
    materialsCost: number;
    /** Plant hired in for the job, at cost. Missing means 0, and the line is not printed. Throws when negative. */
    plantHire?: number;
    /** Consumable stores and fuel used on the job (not travel), at cost. Missing means 0, and the line is not printed. Throws when negative. */
    consumables?: number;
    /**
     * The VAT rate the supply carries: 0.2 standard, 0.05 reduced, 0 zero-rated.
     * Required when the supplier is VAT-registered; ignored when not. Any other
     * value throws.
     */
    vatRate: 0.2 | 0.05 | 0;
    /**
     * How much of the invoice value is reverse-charge construction services,
     * as a fraction from 0 to 1. At 0.05 or less the 5% disregard applies and
     * the whole invoice is standard-rated; above it the whole invoice is
     * reverse-charged. Missing means 1: the whole invoice is construction
     * services. Outside 0 to 1, or not a number, throws.
     */
    reverseChargeShare?: number;
  };
  supplier: {
    /**
     * Whether the subcontractor is registered for VAT. Required: anything but a
     * boolean throws. A guessed answer either drops the VAT from the payable or
     * charges VAT that cannot lawfully be charged. Not registered means no VAT
     * line and no reverse charge.
     */
    vatRegistered: boolean;
    /**
     * Their CIS status as HMRC verified it: `gross` (gross payment status, 0%),
     * `net` (registered, 20%), `unverified` (unmatched or not verified, 30%,
     * with a warning). Required: anything else throws.
     */
    cisStatus: "gross" | "net" | "unverified";
    /** The verification number HMRC gave the contractor. Printed as given when present; nothing is printed when it is missing or empty. */
    verificationNumber?: string;
  };
  customer: {
    /** Whether the contractor is VAT-registered. Required as a boolean when the supplier is VAT-registered and `vatRate` is not 0; otherwise not read. Throws when required and not a boolean. */
    vatRegistered: boolean;
    /** Whether the contractor is registered for CIS. The same requirement as `vatRegistered`. */
    cisRegistered: boolean;
    /**
     * Whether the customer has told the supplier in writing that it is an end
     * user or an intermediary supplier. Either one takes the invoice out of the
     * reverse charge. Missing, `null` or `""` is silence, and silence means the
     * reverse charge applies. Any other value throws.
     */
    endUserNotified?: "end-user" | "intermediary";
  };
}

/**
 * Where the lines sit.
 *
 * - `table`: labels in a column, figures right-aligned beside them, so the
 *   sums can be checked down the page. The default.
 * - `stacked`: one "Label: figure" line under another, for a narrow column.
 */
export type CisInvoiceDeductionArrangement = "table" | "stacked";

/** Props for {@linkcode CisInvoiceDeduction}. */
export interface CisInvoiceDeductionProps {
  /** Where the lines sit. Defaults to `"table"`. */
  arrangement?: CisInvoiceDeductionArrangement;
}

export const CisInvoiceDeduction: SectionComponent<CisInvoiceDeductionProps> = async ({ arrangement }) => {
  const [invoice] = useState((data: CisInvoiceDeductionData) => data.invoice);
  const [supplier] = useState((data: CisInvoiceDeductionData) => data.supplier);
  const [customer] = useState((data: CisInvoiceDeductionData) => data.customer);
  const cis = await useDeriver(cisInvoiceDeduction, [invoice, supplier, customer]);

  return (
    <Section id="cis-deduction" title="CIS deduction">
      {cis.warning && <Paragraph id="cis-deduction-warning">{cis.warning}</Paragraph>}
      {arrange(arrangement, {
        table: () => (
          <Table id="cis-deduction-lines" columns={[{ width: "auto" }, { width: 45, align: "right" }]}>
            <Row id="cis-deduction-gross">
              <Cell>Gross amount, excluding VAT</Cell>
              <Cell>{cis.gross}</Cell>
            </Row>
            <Row id="cis-deduction-materials">
              <Cell>Less materials at cost</Cell>
              <Cell>{cis.materials}</Cell>
            </Row>
            {cis.plantHire && (
              <Row id="cis-deduction-plant-hire">
                <Cell>Less plant hire</Cell>
                <Cell>{cis.plantHire}</Cell>
              </Row>
            )}
            {cis.consumables && (
              <Row id="cis-deduction-consumables">
                <Cell>Less consumables</Cell>
                <Cell>{cis.consumables}</Cell>
              </Row>
            )}
            <Row id="cis-deduction-labour">
              <Cell>Labour subject to CIS</Cell>
              <Cell>{cis.labour}</Cell>
            </Row>
            {cis.vatLabel && (
              <Row id="cis-deduction-vat">
                <Cell>{cis.vatLabel}</Cell>
                <Cell>{cis.vatValue}</Cell>
              </Row>
            )}
            <Row id="cis-deduction-cis">
              <Cell>Less {cis.cisLabel}</Cell>
              <Cell>{cis.cisDeduction}</Cell>
            </Row>
            <Row id="cis-deduction-payable">
              <Cell>Amount payable</Cell>
              <Cell>{cis.payable}</Cell>
            </Row>
          </Table>
        ),
        stacked: () => (
          <>
            <Paragraph id="cis-deduction-gross">Gross amount, excluding VAT: {cis.gross}</Paragraph>
            <Paragraph id="cis-deduction-materials">Less materials at cost: {cis.materials}</Paragraph>
            {cis.plantHire && <Paragraph id="cis-deduction-plant-hire">Less plant hire: {cis.plantHire}</Paragraph>}
            {cis.consumables && (
              <Paragraph id="cis-deduction-consumables">Less consumables: {cis.consumables}</Paragraph>
            )}
            <Paragraph id="cis-deduction-labour">Labour subject to CIS: {cis.labour}</Paragraph>
            {cis.vatLabel && <Paragraph id="cis-deduction-vat">{cis.vatLabel}: {cis.vatValue}</Paragraph>}
            <Paragraph id="cis-deduction-cis">Less {cis.cisLabel}: {cis.cisDeduction}</Paragraph>
            <Paragraph id="cis-deduction-payable">Amount payable: {cis.payable}</Paragraph>
          </>
        ),
      })}
      {supplier.verificationNumber && (
        <Paragraph id="cis-deduction-verification">
          CIS verification number: {supplier.verificationNumber}
        </Paragraph>
      )}
    </Section>
  );
};
