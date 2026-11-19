import { type Nodes, Paragraph, Section, useDeriver, useState } from "docxcelerate/template";
import preferentialOriginStatement from "../derivers/preferential-origin-statement.deriver.ts";

/**
 * The statement on origin a trade agreement prescribes, printed on an invoice
 * so the importer can claim preferential duty, or a warning in its place when
 * it cannot be made truthfully. Nothing at all when no line is originating, so
 * it can sit on every export invoice.
 *
 * Decisions already made for you:
 *
 * - The statement is either true and complete, or it is not printed. A
 *   statement the importer relies on and that turns out to be incomplete costs
 *   them a post-clearance duty demand, which is worse than no statement.
 * - The EUR 6,000 threshold is tested on the originating lines only, converted
 *   at the rate you supply. It is never tested against the invoice total, and
 *   never against a fixed sterling figure: the law is in euro.
 * - Above the threshold the exporter's reference (REX, or EORI for a UK
 *   exporter) is mandatory. Without it the warning prints instead.
 * - Every non-originating line is listed by its reference, because the
 *   statement says "except where otherwise clearly indicated". Without that
 *   list it would cover every line on the document.
 * - A multiple-shipment period over 12 months is refused the same way.
 * - Each agreement's wording is printed exactly, in the prescribed order: the
 *   period, the sentence, the exceptions, the criteria, then place and date
 *   over the exporter's name. There are no arrangements, because moving any of
 *   that would separate the exceptions from the sentence they qualify.
 *
 * All of it is decided in one deriver, because it sums, filters, joins and
 * formats dates on request data, and publishing refuses all of those. The node
 * computes nothing: it reads the deriver's result and branches on it, with
 * every test written inline so a publish build keeps each arm.
 *
 * Not checked: whether the goods originate, which agreement applies, or that a
 * REX or EORI number is well formed. Presence is the only test.
 *
 * Installed by `dxcl add preferential-origin-statement`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface PreferentialOriginData {
  origin: {
    /**
     * The agreement the goods are claimed under. It picks the wording, the
     * reference label and the origin criterion codes. Required: anything else
     * throws, because a statement in the wrong agreement's words is invalid
     * for all of it.
     */
    agreement: "eu-uk-tca" | "eu-gsp" | "ceta" | "eu-japan" | "uk-dcts";
    /**
     * Where the products originate, printed into "these products are of …
     * preferential origin". `eu-uk-tca` takes "UK" or "EU". `ceta` and
     * `eu-japan` take "EU" only. `eu-gsp` and `uk-dcts` take the beneficiary
     * country as it should be printed ("Bangladesh"), never "EU" or "UK".
     * Compared exactly after trimming. Missing, or not allowed for the
     * agreement, throws.
     */
    originCountry: "UK" | "EU" | string;
    /**
     * The exporter's reference, printed as given (trimmed). That is the EORI
     * for a UK exporter under the TCA, and the REX number for an EU exporter
     * or a GSP/DCTS exporter. Printed whenever given. Missing or blank is
     * allowed while the originating products come to EUR 6,000 or less: the
     * reference clause is then left out. Above that, missing means the
     * statement is withheld and a warning prints. Its format is not checked.
     */
    exporterReference?: string;
    /**
     * How many euros one unit of `consignment.currency` is worth (for GBP,
     * about 1.17). Used only to test the EUR 6,000 threshold; never printed
     * except inside the warning. Ignored when the currency is EUR. Otherwise
     * required: missing, zero, negative or not finite throws. Which rate the
     * agreement wants is the caller's decision.
     */
    eurRate?: number;
    /**
     * For a statement covering multiple shipments of identical products:
     * ISO 8601 dates at both ends, both inclusive. It prints as the "Period"
     * line. Missing means a single shipment, and no period line prints.
     * Either end missing or malformed throws, and so does `to` before
     * `from`. A period longer than 12 months withholds the statement.
     */
    period?: { from: string; to: string };
    /** Where the statement is made, printed as given. Required; missing or blank throws. */
    place: string;
    /** When it is made, ISO 8601 (`2026-11-30`). Required; missing or not YYYY-MM-DD throws. */
    date: string;
    /** The exporter's name, printed as given. Required; missing or blank throws. */
    exporterName: string;
  };
  consignment: {
    /**
     * ISO 4217 code of `valueMinor`, three upper-case letters. It sets the
     * minor units. Required: missing or malformed throws.
     */
    currency: string;
    /**
     * Every line of the document, in document order. Missing or empty prints
     * nothing.
     */
    lines: Array<{
      /**
       * The line's reference as the document shows it ("3", "SKU-1182"). It
       * is how exceptions and criteria point at the line. Required and
       * unique: blank or duplicated throws.
       */
      ref: string;
      /** The line's value in integer minor units. Missing, negative or not an integer throws. */
      valueMinor: number;
      /**
       * Whether the line is claimed as originating. Anything but `true`
       * (including missing) means not originating: a preference claim has
       * to be explicit.
       */
      originating?: boolean;
      /**
       * The origin criterion for an originating line. `eu-uk-tca` and
       * `eu-japan` take one or more of A–E. `eu-gsp` and `uk-dcts` take P
       * (wholly obtained) or W (sufficiently worked). `ceta` ignores it.
       * Missing or invalid on an originating line, where the agreement
       * uses criteria, withholds the statement. Ignored on lines that are
       * not originating.
       */
      criterion?: string;
      /**
       * The four-digit HS heading for a GSP/DCTS "W" line ("6109"). Required
       * there: missing or not four digits withholds the statement. Ignored
       * everywhere else.
       */
      hsHeading?: string;
    }>;
  };
}

/**
 * Prints the statement on origin the chosen agreement prescribes, or a warning
 * saying why it cannot be made. Prints nothing when no line is originating.
 */
export const PreferentialOriginStatement: Nodes = async () => {
  const [{ origin, consignment }] = useState((data: PreferentialOriginData) => ({
    origin: data.origin,
    consignment: data.consignment,
  }));
  const s = await useDeriver(preferentialOriginStatement, [origin, consignment]);

  return (
    <>
      {s.outcome === "statement" && (
        <Section id="origin-statement" title={s.title}>
          {s.period && <Paragraph id="origin-statement-period">{s.period}</Paragraph>}
          <Paragraph id="origin-statement-declaration">{s.declaration}</Paragraph>
          {s.exceptions && <Paragraph id="origin-statement-exceptions">{s.exceptions}</Paragraph>}
          {s.criteriaLine && <Paragraph id="origin-statement-criteria">{s.criteriaLine}</Paragraph>}
          <Paragraph id="origin-statement-place-date">{s.placeAndDate}</Paragraph>
          <Paragraph id="origin-statement-exporter">{s.exporterName}</Paragraph>
        </Section>
      )}
      {s.outcome === "withheld" && (
        <Section id="origin-statement-withheld" title={s.title}>
          <Paragraph id="origin-statement-withheld-warning">{s.warning}</Paragraph>
          <Paragraph id="origin-statement-withheld-place-date">{s.placeAndDate}</Paragraph>
          <Paragraph id="origin-statement-withheld-exporter">{s.exporterName}</Paragraph>
        </Section>
      )}
    </>
  );
};
