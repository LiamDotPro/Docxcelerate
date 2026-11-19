import { Cell, type Nodes, Paragraph, Row, Section, Table, useDeriver, useState } from "docxcelerate/template";
import goodFaithEstimateComparison from "../derivers/good-faith-estimate-comparison.deriver.ts";

/**
 * For a bill sent to an uninsured or self-pay patient: each provider or
 * facility's bill against its own line on the good faith estimate (GFE), and,
 * where one billed at least $400 more, the patient's right to use the No
 * Surprises Act patient-provider dispute process.
 *
 * Decisions already made for you:
 *
 * - Each provider is compared with its own estimate. The $400 test is never
 *   applied to the estimate total or the bill total, so one statement can offer
 *   a dispute for the anesthesiologist and not for the surgeon.
 * - Everything billed counts, including items that were not on the estimate.
 * - A provider missing from the estimate is compared with $0.
 * - "At least $400" is inclusive, and the sums are done in integer cents.
 * - Only for uninsured or self-pay patients (self-pay includes an insured
 *   patient who chose not to bill their plan). Nothing prints for a patient
 *   billing their plan or enrolled in a federal programme, or when no estimate
 *   was given.
 * - The 120 days run from the date on the initial bill, not the date of service
 *   or of receipt. The deadline uses date-only arithmetic, so a timezone cannot
 *   move it by a day.
 * - The collections promise prints whenever dispute rights do: the patient
 *   needs it to dispute safely.
 * - A state process replaces the federal one when the state runs its own.
 *
 * All of it is one deriver, because summing, comparing, counting, joining names,
 * date arithmetic and formatting money are all things publishing refuses on
 * request data. The node computes nothing; every test is written inline on
 * `g.*` so a publish build keeps each arm. The wording is regulated, so no
 * prose is generated.
 *
 * The six provider rows are fixed slots rather than a `.map()`. The resolver
 * reads a published loop's collection from request data and ignores the derived
 * scope, so a `.map()` over the deriver's rows would publish and then resolve to
 * no rows.
 *
 * The node publishes 12 branches against the default `branchLimit` of 32: the
 * section, six rows, three outcomes, and the state or federal contact.
 *
 * The legal points (a provider missing from the estimate compared with $0, the
 * CMS number, 120 days from the bill date) should be checked against
 * 45 CFR 149.620 and the current CMS notice before use.
 *
 * Installed by `dxcl add good-faith-estimate-comparison`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface GoodFaithEstimateComparisonData {
  patientBill: {
    /**
     * How the patient is paying for this care. Required: the deriver throws
     * when it is missing or not one of these four, because printing nothing
     * for a misspelt value would silently withhold dispute rights.
     *
     * - `uninsured`: no coverage. Entitled to a GFE.
     * - `self-pay`: has coverage but chose not to bill the plan. Entitled to a GFE.
     * - `insured`: billing the plan. Prints nothing.
     * - `federal-program`: enrolled in Medicare, Medicaid or another federal
     *   health programme. Not entitled to a GFE. Prints nothing.
     */
    coverage: "uninsured" | "self-pay" | "insured" | "federal-program";
    /**
     * The date printed on the initial bill for this care, ISO 8601
     * (`2026-11-19`). Anything after the tenth character is ignored. The
     * dispute deadline is this date plus 120 calendar days. Required whenever
     * the component would print (eligible coverage, an estimate, at least one
     * provider): the deriver throws when it is missing or not YYYY-MM-DD.
     * Not read when nothing prints.
     */
    initialBillDate: string;
    /**
     * The good faith estimate the patient was given. Missing means no estimate
     * exists, and the component prints nothing.
     */
    estimate?: {
      /**
       * The date on the estimate, ISO 8601. Printed in the intro. Required when
       * `estimate` is given: throws when it is missing or not a date.
       */
      date: string;
    };
    /**
     * The state's own patient-provider dispute process, given only where the
     * state runs one in place of the federal process. Missing means the
     * federal process: the CMS contact is printed. When given, both fields
     * are required and are printed as given; the deriver throws if either is
     * missing or blank.
     */
    stateProcess?: {
      /** Who runs it, as printed: "the Example State Department of Insurance". */
      name: string;
      /** How to reach it, as printed: a phone number, a URL, or both. */
      contact: string;
    };
    /**
     * One entry per provider or facility billing on this statement, in the
     * order they should be printed. At most 6: the deriver throws above 6.
     * Missing or empty prints nothing. Two entries with the same name (after
     * trimming) throw, because splitting one provider across entries would
     * split its comparison.
     */
    providers: Array<{
      /** The provider's or facility's name, as printed. Required, non-blank. */
      name: string;
      /**
       * The total expected charges for this provider on the GFE, in integer
       * cents (125000 is $1,250.00). Missing means this provider was not on
       * the estimate: it is compared with $0 and printed as "Not on your
       * estimate". Throws when present and not a non-negative integer.
       */
      estimateCents?: number;
      /**
       * Every item this provider billed, whether or not it was on the
       * estimate. Each line has its amount in integer cents, net of any
       * discount. Other fields on a line are ignored. Required (it may be
       * empty, which bills $0): throws when it is missing or not an array, or
       * when an `amountCents` is not a non-negative integer.
       */
      billedLines: Array<{ amountCents: number }>;
    }>;
  };
}

/**
 * Prints the comparison table and, where a provider billed at least $400 over
 * its own estimate, the dispute rights. Prints nothing when the patient is not
 * eligible, no estimate was given, or there are no providers.
 */
export const GoodFaithEstimateComparison: Nodes = async () => {
  const [bill] = useState((data: GoodFaithEstimateComparisonData) => data.patientBill);
  const g = await useDeriver(goodFaithEstimateComparison, [bill]);

  return (
    <>
      {g.outcome !== "none" && (
        <Section id="gfe-comparison" title="Your bill and your good faith estimate">
          <Paragraph id="gfe-comparison-intro">
            This statement compares what each provider and facility billed with the good faith estimate you were given
            on {g.estimateDate}. Each is compared with its own estimate, and everything it billed counts, including
            items that were not on the estimate.
          </Paragraph>
          <Table
            id="gfe-comparison-table"
            columns={[{ width: "auto" }, { width: 30, align: "right" }, { width: 30, align: "right" }, { width: 34, align: "right" }]}
          >
            <Row id="gfe-comparison-header" header>
              <Cell>Provider or facility</Cell>
              <Cell>Estimate</Cell>
              <Cell>Billed</Cell>
              <Cell>Difference</Cell>
            </Row>
            {g.rows[0].name && (
              <Row id="gfe-comparison-provider-1">
                <Cell>{g.rows[0].name}</Cell>
                <Cell>{g.rows[0].estimate}</Cell>
                <Cell>{g.rows[0].billed}</Cell>
                <Cell>{g.rows[0].difference}</Cell>
              </Row>
            )}
            {g.rows[1].name && (
              <Row id="gfe-comparison-provider-2">
                <Cell>{g.rows[1].name}</Cell>
                <Cell>{g.rows[1].estimate}</Cell>
                <Cell>{g.rows[1].billed}</Cell>
                <Cell>{g.rows[1].difference}</Cell>
              </Row>
            )}
            {g.rows[2].name && (
              <Row id="gfe-comparison-provider-3">
                <Cell>{g.rows[2].name}</Cell>
                <Cell>{g.rows[2].estimate}</Cell>
                <Cell>{g.rows[2].billed}</Cell>
                <Cell>{g.rows[2].difference}</Cell>
              </Row>
            )}
            {g.rows[3].name && (
              <Row id="gfe-comparison-provider-4">
                <Cell>{g.rows[3].name}</Cell>
                <Cell>{g.rows[3].estimate}</Cell>
                <Cell>{g.rows[3].billed}</Cell>
                <Cell>{g.rows[3].difference}</Cell>
              </Row>
            )}
            {g.rows[4].name && (
              <Row id="gfe-comparison-provider-5">
                <Cell>{g.rows[4].name}</Cell>
                <Cell>{g.rows[4].estimate}</Cell>
                <Cell>{g.rows[4].billed}</Cell>
                <Cell>{g.rows[4].difference}</Cell>
              </Row>
            )}
            {g.rows[5].name && (
              <Row id="gfe-comparison-provider-6">
                <Cell>{g.rows[5].name}</Cell>
                <Cell>{g.rows[5].estimate}</Cell>
                <Cell>{g.rows[5].billed}</Cell>
                <Cell>{g.rows[5].difference}</Cell>
              </Row>
            )}
          </Table>
          {g.outcome === "within" && (
            <Paragraph id="gfe-comparison-within">
              Every provider and facility on this statement billed no more than its good faith estimate.
            </Paragraph>
          )}
          {g.outcome === "over" && (
            <Paragraph id="gfe-comparison-over">
              At least one provider or facility billed more than its good faith estimate, but each by less than $400.
              The table shows the difference for each.
            </Paragraph>
          )}
          {g.outcome === "dispute" && (
            <Section id="gfe-dispute" title="Your right to dispute this bill">
              <Paragraph id="gfe-dispute-eligible">
                You can dispute {g.disputeSubject}, because it is at least $400 more than that provider's own good faith
                estimate.
              </Paragraph>
              <Paragraph id="gfe-dispute-deadline">
                To dispute, start the process by {g.deadline}. That is 120 calendar days from {g.initialBillDate}, the
                date on your initial bill.
              </Paragraph>
              <Paragraph id="gfe-dispute-collections">
                While a dispute is open, the provider or facility it concerns will not send the disputed bill to
                collections or threaten to, will pause any collection already started, and will not charge late fees on
                the disputed amount.
              </Paragraph>
              {g.stateName && (
                <Paragraph id="gfe-dispute-how-state">
                  To start a dispute, contact {g.stateName}: {g.stateContact}.
                </Paragraph>
              )}
              {!g.stateName && (
                <Paragraph id="gfe-dispute-how-federal">
                  To start a dispute, contact the Centers for Medicare & Medicaid Services at 1-800-985-3059 or visit
                  www.cms.gov/nosurprises.
                </Paragraph>
              )}
            </Section>
          )}
        </Section>
      )}
    </>
  );
};
