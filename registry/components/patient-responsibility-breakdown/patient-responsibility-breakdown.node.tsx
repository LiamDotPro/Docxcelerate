import { arrange, type Nodes, Paragraph, Section, useDeriver, useState } from "docxcelerate/template";
import type { PatientResponsibilityBreakdownLines } from "../arrangements/patient-responsibility-breakdown/lines.ts";
import { stacked } from "../arrangements/patient-responsibility-breakdown/stacked.arrangement.tsx";
import { table } from "../arrangements/patient-responsibility-breakdown/table.arrangement.tsx";
import patientResponsibilityBreakdown from "../derivers/patient-responsibility-breakdown.deriver.ts";

/**
 * A statement section that takes each claim line from billed charge to allowed
 * amount to what the patient owes, and shows the deductible, copay and
 * coinsurance on each line and the balances left after the last one.
 *
 * Decisions already made for you:
 *
 * - Write-offs are not owed. The gap between billed and allowed is the
 *   provider's write-off and is never shown as owed. The one exception is an
 *   out-of-network line the No Surprises Act does not protect: there the gap is
 *   a balance bill.
 * - Plan order is followed. Copay comes before or after the deductible, and
 *   counts toward it or not, as the plan says.
 * - Coinsurance is on the remainder: what is left after the deductible and
 *   copay, not the allowed amount.
 * - One cap across the statement. Deductible and out-of-pocket balances carry
 *   from line to line in service-date order. Cost share stops at the
 *   out-of-pocket maximum, and later covered lines say the plan pays 100%.
 * - Some charges sit outside the cap. Non-covered charges and balance bills are
 *   added after it and never count toward the deductible or the maximum.
 * - Totals add up exactly. All arithmetic is in integer cents and each line's
 *   coinsurance is rounded once, half up, so the total is the sum of the
 *   printed lines and never drifts by a cent from the rows.
 * - Missing fields fall on the side that never overstates what the patient
 *   owes: a missing `covered` means covered, a missing `network` means in
 *   network, a missing `nsaProtected` means protected.
 * - Denials are not billed. A covered line the plan allowed nothing for is not
 *   billed to the patient.
 *
 * All of it is decided in one deriver, because every step is arithmetic,
 * sorting or comparison on request data, and publishing refuses all of those.
 * The node computes nothing: it reads the result and branches on it, with every
 * test written inline so a publish build keeps each arm.
 *
 * Arrangement: where the lines sit is a separate choice from what they say,
 * made with the `arrangement` prop. This file decides every line; each
 * arrangement, in `arrangements/patient-responsibility-breakdown/`, only
 * places them, so both print the same lines under the same conditions.
 *
 * The rows are a published loop over the deriver's rows (`.map()` over derived
 * data resolves by scope), so there is no row limit.
 *
 * Installed by `dxcl add patient-responsibility-breakdown`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface PatientResponsibilityData {
  patientResponsibility: {
    /**
     * ISO 4217 code the amounts are in. It also sets how many minor units
     * make one major unit. Missing means "USD". Throws when `Intl` does not
     * know the code.
     */
    currency?: string;
    /**
     * Where the patient's plan stood before the first line on this statement:
     * the accumulators from the payer's latest EOB, for the plan year these
     * services fall in.
     */
    plan: {
      /**
       * Deductible still to meet before the first line, in integer minor
       * units (50000 is $500.00). 0 means already met. Required: throws when
       * it is missing, negative or not an integer.
       */
      deductibleRemainingMinor: number;
      /**
       * Out-of-pocket maximum still to meet before the first line, in integer
       * minor units. 0 means already met: every covered line prints "plan
       * pays 100%". Required: throws when it is missing, negative or not an
       * integer.
       */
      oopRemainingMinor: number;
      /**
       * The patient's coinsurance as a fraction: 0.2 is 20%. It is rounded to
       * the nearest basis point (0.01%) before use, so the arithmetic stays in
       * integers. Required: throws when it is missing, not finite, below 0 or
       * above 1.
       */
      coinsuranceRate: number;
      /**
       * True: the copay is taken first and the deductible applies to what is
       * left. False: the deductible is taken first and the copay applies to
       * what is left after it. Missing means true, the common design.
       * Ignored on `copayOnly` lines.
       */
      copayAppliesBeforeDeductible?: boolean;
      /**
       * True: a copay paid also reduces the deductible remaining, from the
       * moment it is taken. With copay first, that lowers the same line's
       * deductible portion. With deductible first, it affects only later
       * lines. Missing means false, the ACA-typical design.
       */
      copayCountsTowardDeductible?: boolean;
    };
    /**
     * One entry per claim line, in any order. They are sorted by
     * `serviceDate`, with ties kept in input order, and printed in that order,
     * because the running balances are only true in that order. Missing or
     * empty means there is nothing to show, and the component prints nothing.
     */
    lines?: Array<{
      /**
       * Date of service, ISO 8601. Anything after the tenth character is
       * ignored. Required: throws when it is missing or not YYYY-MM-DD.
       */
      serviceDate: string;
      /**
       * What the service was, as it should be printed. Required: throws when
       * it is missing or blank.
       */
      description: string;
      /**
       * Procedure or revenue code (CPT, HCPCS, revenue code), printed beside
       * the date. Missing or empty prints the date alone.
       */
      code?: string;
      /**
       * The provider's charge, in integer minor units. Required: throws when
       * it is missing, negative or not an integer.
       */
      billedMinor: number;
      /**
       * What the plan allowed, in integer minor units. Required on a covered
       * line: throws when it is missing, negative, not an integer, or above
       * `billedMinor`. 0 on a covered line with a non-zero charge is a denial.
       * Ignored when `covered` is false.
       */
      allowedMinor?: number;
      /**
       * The plan's copay for this service, in integer minor units. Missing
       * means 0. Never charged above what is left of the allowed amount.
       * Throws when it is negative or not an integer.
       */
      copayMinor?: number;
      /**
       * True when the plan pays everything after the copay on this service
       * (the usual office-visit design): no deductible and no coinsurance on
       * this line. Missing means false.
       */
      copayOnly?: boolean;
      /**
       * Whether the plan covers this service at all. False: the whole billed
       * charge is the patient's, outside the deductible and the out-of-pocket
       * maximum. Missing means true, because billing a full charge on a guess
       * is the worse mistake.
       */
      covered?: boolean;
      /**
       * Whether the provider is in the plan's network. Missing means "in".
       * Throws on any other value.
       */
      network?: "in" | "out";
      /**
       * Out-of-network lines only: whether the No Surprises Act protects this
       * service (emergency care, an out-of-network clinician at an in-network
       * facility without consent, air ambulance). Protected: in-network cost
       * sharing and no balance bill. Missing means true, because balance
       * billing a protected service is unlawful. Ignored on in-network lines.
       */
      nsaProtected?: boolean;
    }>;
  };
}

/**
 * Where the lines sit.
 *
 * - `table`: one row per claim line, the amounts in columns. The default.
 * - `stacked`: one labelled block per claim line, down the text column.
 */
export type PatientResponsibilityBreakdownArrangement = "table" | "stacked";

/** Props for {@linkcode PatientResponsibilityBreakdown}. */
export interface PatientResponsibilityBreakdownProps {
  /** Where the lines sit. Defaults to `"table"`. */
  arrangement?: PatientResponsibilityBreakdownArrangement;
}

/**
 * Prints what the patient owes on each claim line, then what is written off,
 * what sits outside the cap and what is still to meet. Prints nothing when
 * there are no lines.
 */
export const PatientResponsibilityBreakdown: Nodes<PatientResponsibilityBreakdownProps> = async (
  { arrangement },
) => {
  const [patientResponsibility] = useState((data: PatientResponsibilityData) => data.patientResponsibility);
  const p = await useDeriver(patientResponsibilityBreakdown, [patientResponsibility]);

  const lines: PatientResponsibilityBreakdownLines = { rows: p.rows, total: p.total };

  return (
    <>
      {p.outcome !== "none" && (
        <Section id="patient-responsibility" title="Your share of the cost">
          {p.outcome === "owes" && (
            <Paragraph id="patient-responsibility-owes">You owe {p.total.youOwe} for the services below.</Paragraph>
          )}
          {p.outcome === "nothing-owed" && (
            <Paragraph id="patient-responsibility-nothing-owed">You owe nothing for the services below.</Paragraph>
          )}
          {arrange(arrangement, { table: () => table(lines), stacked: () => stacked(lines) })}
          {p.hasAdjustment && (
            <Paragraph id="patient-responsibility-adjustment">
              {p.adjustment} of the billed charges is written off under the provider's agreement with your plan. You do
              not owe it.
            </Paragraph>
          )}
          {p.hasExcluded && (
            <Paragraph id="patient-responsibility-excluded">
              {p.excluded} of what you owe is for services your plan does not cover or for out-of-network charges above
              the allowed amount. It does not count toward your deductible or out-of-pocket maximum.
            </Paragraph>
          )}
          {p.oopMet && (
            <Paragraph id="patient-responsibility-oop-met">
              You have reached your plan's out-of-pocket maximum. Covered services are paid in full for the rest of the
              plan year.
            </Paragraph>
          )}
          {!p.oopMet && (
            <Paragraph id="patient-responsibility-remaining">
              Still to meet after these services: {p.deductibleRemaining} of your deductible and {p.oopRemaining} of
              your out-of-pocket maximum.
            </Paragraph>
          )}
        </Section>
      )}
    </>
  );
};
