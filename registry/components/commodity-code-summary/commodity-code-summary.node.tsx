import { arrange, Paragraph, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { CommodityCodeSummaryLines } from "../arrangements/commodity-code-summary/lines.ts";
import { linesThenSummary } from "../arrangements/commodity-code-summary/lines-then-summary.arrangement.tsx";
import { summaryFirst } from "../arrangements/commodity-code-summary/summary-first.arrangement.tsx";
import commodityCodeSummary from "../derivers/commodity-code-summary.deriver.ts";

/**
 * The customs block of a commercial invoice: each goods line's tariff code,
 * origin, preference claim, quantity, supplementary unit, net mass and value,
 * then the summary a customs declaration is keyed from.
 *
 * Decisions already made for the reader:
 *
 * - Tariff codes are text. A code supplied as a number warns, because its
 *   leading zero is already gone (0901 coffee becomes 901).
 * - The digit count is checked against the destination: 6 is HS, 8 is EU CN,
 *   10 is the UK and US tariffs. A short or missing code warns on its own line,
 *   since a 6-digit code passes a generic template but cannot be declared there.
 * - Lines group on code, origin and preference together, never on code alone.
 *   Customs declares those combinations separately.
 * - Mass is rounded once: lines are converted to exact grams, summed per group
 *   and rounded to whole grams there. The total is the sum of the rounded
 *   groups, so the summary always adds up.
 * - Net mass above the declared gross mass warns, before a broker rejects it.
 *
 * All of it is one deriver, because every step does string checks, sorting,
 * grouping or arithmetic on request data, and publishing refuses all of those.
 * The tables loop over the deriver's rows with `.map()`, which the resolver
 * reads by scope. The node computes nothing: it builds the lines the
 * arrangements place and decides only whether there is anything to place.
 *
 * Arrangement: where the lines and the summary sit is a separate choice from
 * what they say, made with the `arrangement` prop. An arrangement only places,
 * so there is no summary-only one: it would drop the lines and their warnings.
 *
 * Installed by `dxcl add commodity-code-summary`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface CommodityCodeSummaryData {
  customs: {
    /**
     * ISO 4217 code every `valueMinor` is in: "GBP", "EUR", "JPY". It also sets
     * how many minor units make a major one. Required: the deriver throws when
     * it is missing or not a code `Intl` knows.
     */
    currency: string;
    /**
     * How many digits the destination's tariff needs: 6 (HS), 8 (EU Combined
     * Nomenclature) or 10 (UK Global Tariff, US HTS). Every line's code is
     * checked against it. Required: anything else throws, because a check
     * against the wrong length passes codes that cannot be declared.
     */
    destinationCodeLength: 6 | 8 | 10;
    /**
     * The declared gross mass of the whole shipment, in kilograms, packaging
     * included. Printed under the summary and compared with the summed net
     * mass. Missing means no gross line and no check. Throws when present but
     * not a finite number of 0 or more.
     */
    grossMassKg?: number;
    /**
     * The goods lines, in invoice order. They are printed in this order and
     * grouped for the summary. Missing or empty prints one sentence saying
     * there are no goods lines, and no tables.
     */
    lines: Array<{
      /**
       * The invoice's own line reference, printed so a broker can quote it.
       * Missing or empty prints the line's 1-based position ("Line 3").
       */
      ref?: string;
      /** The goods description as printed on the invoice. Missing prints nothing in its place. */
      description?: string;
      /**
       * The tariff (commodity) code, as text: "0901210000". Spaces and dots
       * are removed before checking ("0901 21 00 00" is fine). Missing or
       * empty warns on the line and the line groups under "None". A number
       * warns that leading zeros may have been lost, and is then checked as
       * its digits. Anything left that is not digits warns. Fewer digits
       * than `destinationCodeLength` warns. More digits is accepted and
       * printed as given.
       */
      tariffCode: string;
      /**
       * Country of origin, ISO 3166-1 alpha-2, trimmed and upper-cased: "CO".
       * Missing or empty warns on the line and groups under "None". Anything
       * that is not two letters after that warns, and groups as given.
       */
      origin: string;
      /**
       * Whether preferential origin is claimed for this line. Missing means
       * false. Claiming a preference without the proof behind it is the error
       * that costs money, so the default is to claim nothing.
       */
      preference?: boolean;
      /**
       * The invoice quantity, in the invoice's own units. Printed per line,
       * never summed: it is not a declaration figure. Throws when it is not a
       * finite number of 0 or more.
       */
      quantity: number;
      /**
       * The supplementary unit the tariff asks for on this code, and the
       * quantity in it: { code: "pa", quantity: 2400 } for pairs,
       * { code: "l", quantity: 3600 } for litres. Missing means the code has
       * none. Throws when present with an empty code, or a quantity that is
       * not a finite number of 0 or more.
       */
      supplementaryUnit?: { code: string; quantity: number };
      /**
       * Net mass of the line, without packaging. Required: throws when value
       * is not a finite number of 0 or more, or unit is not "kg", "g" or "lb".
       */
      netMass: { value: number; unit: "kg" | "g" | "lb" };
      /**
       * The line's value in integer minor units of `currency` (15600000 is
       * £156,000.00). Required: throws when it is not a non-negative integer.
       * A credit belongs in the invoice totals, not on a customs line.
       */
      valueMinor: number;
    }>;
  };
}

/**
 * Where the lines and the summary sit.
 *
 * - `lines-then-summary`: every line, then the summary under it. The default.
 * - `summary-first`: the summary and its warnings first, the lines after, so
 *   the figures a declaration needs sit on the first page of a long invoice.
 */
export type CommodityCodeSummaryArrangement = "lines-then-summary" | "summary-first";

/** Props for {@linkcode CommodityCodeSummary}. */
export interface CommodityCodeSummaryProps {
  /** Where the lines and the summary sit. Defaults to `"lines-then-summary"`. */
  arrangement?: CommodityCodeSummaryArrangement;
}

export const CommodityCodeSummary: SectionComponent<CommodityCodeSummaryProps> = async ({ arrangement }) => {
  const [customs] = useState((data: CommodityCodeSummaryData) => data.customs);
  const result = await useDeriver(commodityCodeSummary, [customs]);

  const lines: CommodityCodeSummaryLines = {
    basis: result.basis,
    attention: result.attention,
    lines: result.lines,
    groups: result.groups,
    totalNetMass: result.totalNetMass,
    totalValue: result.totalValue,
    grossMass: result.grossMass,
    grossWarning: result.grossWarning,
  };

  return (
    <Section id="commodity-code-summary" title="Commodity codes">
      {result.empty
        ? <Paragraph id="commodity-code-summary-none">There are no goods lines on this invoice.</Paragraph>
        : arrange(arrangement, {
          "lines-then-summary": () => linesThenSummary(lines),
          "summary-first": () => summaryFirst(lines),
        })}
    </Section>
  );
};
