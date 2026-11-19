import { arrange, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { ExemptionCertificateCheckLines } from "../arrangements/exemption-certificate-check/lines.ts";
import { stacked } from "../arrangements/exemption-certificate-check/stacked.arrangement.tsx";
import { table } from "../arrangements/exemption-certificate-check/table.arrangement.tsx";
import exemptionCertificateCheck from "../derivers/exemption-certificate-check.deriver.ts";

/**
 * An invoice section that decides, for each line of a US sale, whether it may
 * go out exempt from sales tax, which of the buyer's certificates allows that,
 * and for a line that stays taxable, which rule failed. The reader gets the
 * lines to bill tax on, each with a reason they can defend in an audit.
 *
 * Decisions already made for you:
 *
 * - A certificate is never a yes or no for the whole customer. It is tested
 *   line by line against the line's category, the ship-to state (never the
 *   billing state), whether that state accepts the form, its expiry on the
 *   invoice date, and when it was received.
 * - Streamlined Form F0003 counts only in the full-member states of a list you
 *   supply. The MTC uniform certificate counts only in the states you list as
 *   accepting it, and for resale only where you say so. Both lists print
 *   their as-of dates, so a reader sees what the check was judged against.
 * - Resale, agricultural, industrial and direct-mail reasons cover only lines
 *   of their own category, so a resale certificate never exempts own-use
 *   items. Government, tribal, diplomat, charity, religious and education
 *   reasons cover the whole order. Direct pay covers the whole order and moves
 *   the tax to the buyer. "Other" is never accepted automatically.
 * - In a Streamlined member state a certificate the buyer has promised keeps a
 *   line exempt until 90 days after the invoice date, printed with that date.
 *   Once it passes the line is taxable: the 90 days is a deadline, not
 *   permission never to collect.
 * - A missing MTC `resaleOnly` is read as true: wrongly allowing a non-resale
 *   exemption is the costly direction.
 *
 * All of it is decided in one deriver, because every step is arithmetic,
 * comparison or date handling on request data, and publishing refuses all of
 * those. The node computes nothing: it reads the deriver's result and branches
 * on it. The deriver's `lines` match `us-sales-tax-breakdown`'s
 * `salesTax.lines`, so an application can call its `run` upstream; a deriver's
 * output cannot reach another node's data.
 *
 * Arrangement: where the lines sit is chosen with the `arrangement` prop. This
 * file decides every line; each arrangement, in
 * `arrangements/exemption-certificate-check/`, only places them.
 *
 * The lines are sixteen fixed slots rather than a `.map()`. The resolver reads
 * a published loop's collection from request data and ignores the derived
 * scope, so a `.map()` over the deriver's rows would publish and then resolve
 * to nothing. More than 16 lines throws.
 *
 * The node publishes 17 branches against the default `branchLimit` of 32: the
 * section and sixteen rows.
 *
 * Installed by `dxcl add exemption-certificate-check`, which also installs its
 * deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface ExemptionCertificateCheckData {
  exemption: {
    /**
     * The invoice (sale) date, ISO 8601: `2026-11-30`. Anything after the tenth
     * character is ignored. Expiry is judged on this date, and the 90-day
     * window for a pending certificate runs from it. Required: the deriver
     * throws when it is missing or does not start with YYYY-MM-DD.
     */
    invoiceDate: string;
    /**
     * The day the check is made, ISO 8601. Used only to decide whether a
     * pending certificate's 90 days have run out. Missing means `invoiceDate`,
     * so an invoice printed on its own date always honours a pending
     * certificate. Pass it when the document is printed or reissued later.
     * Throws when it is present but not a date, or when it is before
     * `invoiceDate`.
     */
    checkedOn?: string;
    /**
     * The state the goods ship to, as a two-letter USPS code. It is trimmed and
     * upper-cased before use. Every certificate is judged against this state,
     * never the billing address. Required: throws when, after trimming and
     * upper-casing, it is not two letters A to Z.
     */
    shipToState: string;
    /** ISO 4217 code the amounts are in. Missing means `"USD"`. Throws when `Intl` does not know the code. */
    currency?: string;
    /**
     * The SST full-member states, and the date the list was correct as of.
     * Never hardcoded: membership changes, and the as-of date is printed.
     * Associate members (Tennessee) are not full members. Leave them out unless
     * you have confirmed they accept F0003. Required: throws when it is
     * missing, when `asOf` is not an ISO date, or when `states` is not an
     * array. An empty array means no state accepts F0003.
     */
    sstMembers: {
      /** ISO 8601 date the list is correct as of. Printed. */
      asOf: string;
      /** Two-letter codes. Trimmed and upper-cased before use. */
      states: string[];
    };
    /**
     * The states that accept the MTC uniform certificate, and the date the list
     * was correct as of. Missing means no MTC certificate is accepted anywhere:
     * every MTC certificate fails "No list of states accepting the MTC uniform
     * certificate was supplied". Throws when it is present and `asOf` is not an
     * ISO date or `states` is not an array.
     */
    mtcAcceptingStates?: {
      /** ISO 8601 date the list is correct as of. Printed. */
      asOf: string;
      states: Array<{
        /** Two-letter code. Trimmed and upper-cased before use. */
        code: string;
        /**
         * True when the state takes the MTC certificate only for resale
         * (reason G). Missing is read as `true`, because wrongly allowing a
         * non-resale exemption is the costly direction.
         */
        resaleOnly: boolean;
      }>;
    };
    /**
     * Every exemption certificate the buyer holds or has promised, in any
     * order. Missing or empty means none: every line is taxable with "No
     * exemption certificate on file".
     */
    certificates?: Array<{
      /**
       * Which form. `SST` is Streamlined Form F0003, `MTC` is the uniform
       * certificate, and `state` is a state's own form. Anything else throws.
       */
      form: "SST" | "MTC" | "state";
      /**
       * The state the certificate is for. On a `state` form it is the only
       * state that form covers. On SST and MTC forms it is read only when
       * `states` is missing. Trimmed and upper-cased.
       */
      state?: string;
      /**
       * The states a multi-state SST or MTC certificate was completed for.
       * Ignored on a `state` form. When both `states` and `state` are missing,
       * the certificate fails "Certificate names no state".
       */
      states?: string[];
      /**
       * The exemption reason, using the F0003 letters (MTC and state forms are
       * mapped onto them by the caller): A federal government, B state or local
       * government, C tribal government, D foreign diplomat, E charitable
       * organization, F religious organization, G resale, H agricultural
       * production, I industrial production or manufacturing, J direct pay
       * permit, K direct mail, L other, M educational organization. Anything
       * else throws.
       */
      reason: "A" | "B" | "C" | "D" | "E" | "F" | "G" | "H" | "I" | "J" | "K" | "L" | "M";
      /** The certificate or permit number, printed in the exempt reason when given. Missing prints the form alone. */
      number?: string;
      /**
       * Last day the certificate is valid, ISO 8601, inclusive. Expiry rules
       * differ by state, so the caller works the date out. Missing means it
       * does not expire. Throws when it is present but not a date.
       */
      expires?: string;
      /**
       * When the seller received it, ISO 8601. Missing means it was held at
       * the time of sale. Throws when it is present but not a date. Ignored
       * when `pending` is true.
       */
      receivedOn?: string;
      /**
       * True when the buyer has claimed the exemption but not yet sent the
       * certificate. Honoured only when the ship-to state is an SST member, and
       * only until `invoiceDate` + 90 days. Missing means false.
       */
      pending?: boolean;
    }>;
    /**
     * The invoice lines, in invoice order. Missing or empty means nothing to
     * check, and the component prints nothing. More than 16 throws.
     */
    lines?: Array<{
      /** What the line is, as printed. Missing or blank throws. */
      description: string;
      /**
       * Net price in minor units (cents), before tax. Must be a safe integer.
       * Zero and negative (a return or discount) are decided like any other
       * line and printed with their sign. Anything else throws.
       */
      amountMinor: number;
      /**
       * What the buyer will do with the line, matched exactly against the
       * reasons that depend on use: `"resale"` (G), `"agricultural"` (H),
       * `"industrial"` (I), `"direct-mail"` (K). Any other value, such as
       * `"own-use"`, is covered only by reasons that depend on the buyer
       * (A–F, J, M). Missing means the same as `"own-use"`, and the failure
       * says "this line has no category". Copied through to the output line.
       */
      category?: string;
    }>;
  };
}

/**
 * Where the lines sit.
 *
 * - `table`: one row per line with columns for amount, status and reason. The default.
 * - `stacked`: one paragraph per line, for a narrow column or long failure reasons.
 */
export type ExemptionCertificateCheckArrangement = "table" | "stacked";

/** Props for {@linkcode ExemptionCertificateCheck}. */
export interface ExemptionCertificateCheckProps {
  /** Where the lines sit. Defaults to `"table"`. */
  arrangement?: ExemptionCertificateCheckArrangement;
}

/**
 * Prints each invoice line with whether it is exempt and why, then the totals
 * and the dates the lists were judged against. Prints nothing when there are
 * no lines.
 */
export const ExemptionCertificateCheck: SectionComponent<ExemptionCertificateCheckProps> = async (
  { arrangement },
) => {
  const [exemption] = useState((d: ExemptionCertificateCheckData) => d.exemption);
  const c = await useDeriver(exemptionCertificateCheck, [exemption]);

  const lines: ExemptionCertificateCheckLines = { rows: c.rows, summary: c.summary, basis: c.basis };

  return (
    <>
      {c.hasLines && (
        <Section id="exemption-check" title="Sales tax exemptions">
          {arrange(arrangement, {
            table: () => table(lines),
            stacked: () => stacked(lines),
          })}
        </Section>
      )}
    </>
  );
};
