import {
  Cell,
  Paragraph,
  Row,
  Section,
  type Section as SectionComponent,
  Table,
  useDeriver,
  useSetPlaceholders,
  useSetPrompts,
  useState,
} from "docxcelerate/template";
import payLessNotice from "../derivers/pay-less-notice.deriver.ts";

/**
 * A UK construction pay less notice under section 111 of the Housing Grants,
 * Construction and Regeneration Act 1996 (as amended in 2009): the sum the payer
 * considers due, every deduction itemised, and whether the notice is in time.
 *
 * Decisions already made for you:
 *
 * - The basis is kept. An agent nets everything into one figure and loses the
 *   basis section 111(4)(b) requires. Here the itemised schedule is the basis;
 *   the generated paragraph explains the schedule but does not replace it.
 * - The LD cap is cumulative. Liquidated damages are capped at the cap less what
 *   earlier notices already deducted, and the line says when the cap bit.
 * - CIS is rejected. A CIS deduction is tax withholding, not a pay less ground,
 *   and belongs on the CIS statement. Listing one throws.
 * - No negative sum is printed. Deductions above the notified sum give a sum due
 *   of nil, and a separate line says the excess is not recovered through this
 *   notice.
 * - Lateness is called out. A late notice has no effect, so the banner says the
 *   notified sum remains payable. An agent never checks this.
 *
 * All of it is one deriver, because arithmetic, sorting by kind, date maths and
 * formatting on request data all throw at publish. The node computes nothing:
 * it reads the deriver's result and branches on it, each test written inline.
 * `excludedDays` exists because section 116 leaves Christmas Day, Good Friday
 * and bank holidays out of the count back from the final date for payment.
 *
 * The schedule is 12 fixed slots rather than a `.map()`: a published loop
 * resolves its collection from request data, not from derived output, so a
 * `.map()` over the deriver's lines would publish and resolve to no rows.
 *
 * The node publishes 15 branches against the default `branchLimit` of 32: two
 * banners, twelve schedule rows and the excess paragraph.
 *
 * Nothing in the runtime falls back to the placeholder when generation fails;
 * the placeholder is what previews show. The legal basis survives a bad
 * generation only because the schedule is static.
 *
 * Installed by `dxcl add pay-less-notice`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface PayLessNoticeData {
  payLess: {
    /** The contract (or subcontract) the notice is given under, printed as given. Required; it is interpolated, not checked. */
    contractRef: string;
    /** Who gives the notice: the party paying. Printed as given. Required. */
    payer: string;
    /** Who the notice is given to: the party being paid. Printed as given. Required. */
    payee: string;
    /** The notified sum this notice reduces: the sum in the notice s.111 says it is measured against. */
    notifiedSum: {
      /**
       * The notified sum in integer pence (487231645 for £4,872,316.45). Required:
       * the deriver throws when it is missing, negative or not a safe integer. Zero is allowed.
       */
      amountMinor: number;
      /**
       * Which notice fixed it. `payment-notice`: the payer's payment notice.
       * `application`: the payee's payment application, standing as the payee's
       * notice under s.110B(4). `default-notice`: the payee's default payment notice
       * under s.110B(2), given because no valid payment notice was. Required:
       * anything else throws.
       */
      source: "payment-notice" | "application" | "default-notice";
      /** Which notice it is ("Interim Application No. 14"), printed after the source in brackets. Missing or blank: omitted. */
      reference?: string;
    };
    /** The final date for payment of the notified sum, ISO 8601 (YYYY-MM-DD; anything after the tenth character is ignored). Required: throws when missing or malformed. */
    finalDateForPayment: string;
    /**
     * The last day this notice may be served, ISO 8601. When given, it is used as
     * is, and `prescribedPeriodDays` and `excludedDays` are ignored. Throws when
     * malformed or after `finalDateForPayment`. One of this or
     * `prescribedPeriodDays` is required; with neither, the deriver throws
     * rather than guess at validity.
     */
    payLessDeadline?: string;
    /**
     * The prescribed period: the days before the final date for payment by which
     * a pay less notice must be served (7 under the Scheme; the contract may say
     * otherwise). Counted back from `finalDateForPayment`, skipping `excludedDays`.
     * Throws unless it is a positive integer. Ignored when `payLessDeadline` is given.
     */
    prescribedPeriodDays?: number;
    /**
     * Days s.116 leaves out of the count: Christmas Day, Good Friday and the bank
     * holidays where the work is (gov.uk/bank-holidays.json lists all three), as
     * ISO 8601 dates. Only days inside the counting window matter. Missing means
     * every day counts, which makes the deadline later than the law allows when
     * the window crosses one. Throws on a malformed entry.
     */
    excludedDays?: string[];
    /** The date the notice is served, ISO 8601. Required: throws when missing or malformed. */
    servedDate: string;
    /** Liquidated damages for delay. Missing: no LD line. */
    ld?: {
      /** The LD rate in integer pence per `per`. Throws unless it is a positive safe integer. */
      rateMinor: number;
      /** What the rate is per. Throws on anything else. */
      per: "day" | "week";
      /** Days of culpable delay. Throws unless it is a non-negative integer. 0: no LD line. */
      delayDays: number;
      /** The contract's total LD cap in integer pence. Missing: uncapped. Throws when it is negative or not an integer. */
      capMinor?: number;
      /** LDs already deducted under earlier notices, in integer pence. Missing means 0. Throws when it is negative or not an integer. */
      previouslyDeductedMinor?: number;
    };
    /**
     * Every other deduction, in any order; each kind is printed together, in the
     * order given. Missing or empty: no deductions besides LDs.
     */
    deductions?: Array<{
      /**
       * `contra`: a contra-charge (work done or cost incurred by the payer for the
       * payee). `abatement`: work valued lower than the notified sum assumed.
       * `setoff`: a cross-claim. `cis`: always throws, because CIS withholding is
       * not a pay-less ground and goes on the CIS statement. Anything else throws.
       */
      kind: "contra" | "abatement" | "setoff" | "cis";
      /** What the deduction is for, printed as given. Required: throws when missing or blank, because it is the basis. */
      description: string;
      /** The amount deducted in integer pence. Throws when it is negative or not an integer. A deduction of 0 is left out. */
      amountMinor: number;
      /** A supporting reference (a contra-charge number, a notice date), printed after the description. Missing or blank: omitted. */
      reference?: string;
    }>;
  };
}

/**
 * Prints the validity banner, the particulars, the sum due, the itemised
 * schedule, any excess, and a generated explanation of the basis.
 */
export const PayLessNotice: SectionComponent = async () => {
  const [payLess] = useState((data: PayLessNoticeData) => data.payLess);
  const [contractRef] = useState((data: PayLessNoticeData) => data.payLess.contractRef);
  const [payer] = useState((data: PayLessNoticeData) => data.payLess.payer);
  const [payee] = useState((data: PayLessNoticeData) => data.payLess.payee);
  useSetPrompts({
    systemPrompt:
      "You write the basis of calculation in a UK construction pay less notice under section 111 of the Housing Grants, Construction and Regeneration Act 1996, on behalf of the payer. Precise, neutral and factual; no argument, no persuasion.",
    generalPrompt:
      "In one paragraph, state the basis on which the sum due was calculated: start from the notified sum, take each deduction in the order given with its amount and what it is for, and end with the sum due on the date of service.",
    negativePrompt:
      "Do not change, round, total differently or add any figure. Do not introduce a deduction, ground, date, clause number or reference that is not in the figures. Do not mention CIS, tax or VAT. Do not state any sum below nil or describe a negative payment. Do not threaten proceedings or comment on the payee's conduct.",
  });
  useSetPlaceholders(
    "The basis of calculation: each deduction in the schedule above set against the notified sum, written per notice from those figures.",
  );
  const p = await useDeriver(payLessNotice, [payLess]);

  return (
    <Section id="pay-less-notice" title="Pay less notice">
      {p.timing === "late" && (
        <Paragraph id="pay-less-notice-late">
          Served late. This notice was served on {p.served}, after {p.deadline}, the last day a pay less notice could be
          given for the final date for payment of {p.finalDate}. Under section 111 of the Housing Grants, Construction
          and Regeneration Act 1996 a late pay less notice has no effect: the notified sum of {p.notified} remains
          payable in full by {p.finalDate}.
        </Paragraph>
      )}
      {p.timing === "in-time" && (
        <Paragraph id="pay-less-notice-in-time">
          Served in time. This notice is served on {p.served}, on or before {p.deadline}, the last day a pay less notice
          can be given for the final date for payment of {p.finalDate}.
        </Paragraph>
      )}
      <Paragraph id="pay-less-notice-particulars">
        From {payer} to {payee}, under {contractRef}. This is a pay less notice under section 111 of the Housing Grants,
        Construction and Regeneration Act 1996, given against {p.notifiedSource}, which states a notified sum of{" "}
        {p.notified} with a final date for payment of {p.finalDate}.
      </Paragraph>
      <Paragraph id="pay-less-notice-sum-due">
        {payer} considers the sum due on the date this notice is served to be {p.sumDue}.
      </Paragraph>
      <Table id="pay-less-notice-schedule" columns={[{ width: 38 }, { width: "auto" }, { width: 32, align: "right" }]}>
        <Row id="pay-less-notice-schedule-header" header>
          <Cell>Kind</Cell>
          <Cell>Basis</Cell>
          <Cell>Amount</Cell>
        </Row>
        <Row id="pay-less-notice-notified">
          <Cell>Notified sum</Cell>
          <Cell>{p.notifiedSource}</Cell>
          <Cell>{p.notified}</Cell>
        </Row>
          {p.lines[0].description && (
            <Row id="pay-less-notice-line-1">
              <Cell>{p.lines[0].kind}</Cell>
              <Cell>{p.lines[0].description}</Cell>
              <Cell>{p.lines[0].amount}</Cell>
            </Row>
          )}
          {p.lines[1].description && (
            <Row id="pay-less-notice-line-2">
              <Cell>{p.lines[1].kind}</Cell>
              <Cell>{p.lines[1].description}</Cell>
              <Cell>{p.lines[1].amount}</Cell>
            </Row>
          )}
          {p.lines[2].description && (
            <Row id="pay-less-notice-line-3">
              <Cell>{p.lines[2].kind}</Cell>
              <Cell>{p.lines[2].description}</Cell>
              <Cell>{p.lines[2].amount}</Cell>
            </Row>
          )}
          {p.lines[3].description && (
            <Row id="pay-less-notice-line-4">
              <Cell>{p.lines[3].kind}</Cell>
              <Cell>{p.lines[3].description}</Cell>
              <Cell>{p.lines[3].amount}</Cell>
            </Row>
          )}
          {p.lines[4].description && (
            <Row id="pay-less-notice-line-5">
              <Cell>{p.lines[4].kind}</Cell>
              <Cell>{p.lines[4].description}</Cell>
              <Cell>{p.lines[4].amount}</Cell>
            </Row>
          )}
          {p.lines[5].description && (
            <Row id="pay-less-notice-line-6">
              <Cell>{p.lines[5].kind}</Cell>
              <Cell>{p.lines[5].description}</Cell>
              <Cell>{p.lines[5].amount}</Cell>
            </Row>
          )}
          {p.lines[6].description && (
            <Row id="pay-less-notice-line-7">
              <Cell>{p.lines[6].kind}</Cell>
              <Cell>{p.lines[6].description}</Cell>
              <Cell>{p.lines[6].amount}</Cell>
            </Row>
          )}
          {p.lines[7].description && (
            <Row id="pay-less-notice-line-8">
              <Cell>{p.lines[7].kind}</Cell>
              <Cell>{p.lines[7].description}</Cell>
              <Cell>{p.lines[7].amount}</Cell>
            </Row>
          )}
          {p.lines[8].description && (
            <Row id="pay-less-notice-line-9">
              <Cell>{p.lines[8].kind}</Cell>
              <Cell>{p.lines[8].description}</Cell>
              <Cell>{p.lines[8].amount}</Cell>
            </Row>
          )}
          {p.lines[9].description && (
            <Row id="pay-less-notice-line-10">
              <Cell>{p.lines[9].kind}</Cell>
              <Cell>{p.lines[9].description}</Cell>
              <Cell>{p.lines[9].amount}</Cell>
            </Row>
          )}
          {p.lines[10].description && (
            <Row id="pay-less-notice-line-11">
              <Cell>{p.lines[10].kind}</Cell>
              <Cell>{p.lines[10].description}</Cell>
              <Cell>{p.lines[10].amount}</Cell>
            </Row>
          )}
          {p.lines[11].description && (
            <Row id="pay-less-notice-line-12">
              <Cell>{p.lines[11].kind}</Cell>
              <Cell>{p.lines[11].description}</Cell>
              <Cell>{p.lines[11].amount}</Cell>
            </Row>
          )}
        <Row id="pay-less-notice-total-deductions">
          <Cell>Total deductions</Cell>
          <Cell></Cell>
          <Cell>{p.totalDeductions}</Cell>
        </Row>
        <Row id="pay-less-notice-due">
          <Cell>Sum due</Cell>
          <Cell>On the date of service</Cell>
          <Cell>{p.sumDue}</Cell>
        </Row>
      </Table>
      {p.excess && (
        <Paragraph id="pay-less-notice-excess">
          The deductions exceed the notified sum by {p.excess}. Nothing is payable under this notice, and the excess is
          not claimed or recovered through it.
        </Paragraph>
      )}
      <Paragraph id="pay-less-notice-basis" infoPrompt={`The figures, which are final: ${p.basisFacts}`} />
    </Section>
  );
};
