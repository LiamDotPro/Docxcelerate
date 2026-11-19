import { arrange, Paragraph, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { GstInclusiveStatementLines } from "../arrangements/gst-inclusive-statement/lines.ts";
import { stacked } from "../arrangements/gst-inclusive-statement/stacked.arrangement.tsx";
import { table } from "../arrangements/gst-inclusive-statement/table.arrangement.tsx";
import gstInclusiveStatement from "../derivers/gst-inclusive-statement.deriver.ts";

/**
 * The items, the GST and the total of an Australian or New Zealand invoice:
 * how the GST is shown, and which identities the invoice has to carry. It
 * prints the item list itself because, on a mixed invoice, the taxable mark
 * has to sit on the item.
 *
 * Decisions already made for you:
 *
 * - "Total price includes GST." takes the place of a GST figure only when it
 *   is exactly true: every item taxable at the standard rate, prices GST
 *   inclusive, and the GST the rounding rule produced equal to 1/11 (AU) or
 *   3/23 (NZ) of the total. A GST-free, input-taxed, margin-scheme, zero-rated
 *   or exempt item, GST-exclusive prices, or a per-line rounding that drifts a
 *   cent all show the figure instead.
 * - A mixed invoice marks each taxable item with " *" and prints a legend, so
 *   the reader can see which items are taxable.
 * - GST is 1/11 or 3/23 of an inclusive price, never 10% or 15% of the gross.
 * - GST is worked in integer cents: one rounding on the taxable total (the
 *   ATO total-invoice rule, the default) or each taxable line rounded and then
 *   added (the taxable-sale rule, the issuer's choice).
 * - Thresholds use the GST-inclusive total, which is what the law uses.
 * - The AU buyer rule and the NZ buyer rule are different rules: AU asks for
 *   the buyer's name or ABN from A$1,000; NZ asks for the name and one
 *   identifier above NZ$1,000. The seller's ABN is always required in AU, the
 *   GST number above NZ$200 in NZ.
 * - ABNs, IRD/GST numbers (with the secondary weights) and NZBNs are
 *   checksummed. A bad checksum or a missing required identity throws: an
 *   invoice that cannot support a credit claim is worse than a failed build.
 * - A seller who is not registered gets "Invoice" and no GST wording at all.
 *   Claiming GST while unregistered is the unlawful direction, so a missing
 *   flag means not registered.
 * - The A$82.50 threshold changes nothing printed. A registered buyer may ask
 *   for a tax invoice above it, but issuing one below it is always lawful and
 *   whether the buyer asked is not in the data. Do not add it.
 * - All of it is one deriver, because totals, rounding, thresholds and
 *   checksums all compute on request data.
 *
 * Arrangement: where the items and totals sit is a separate choice from what
 * they say, made with the `arrangement` prop. The identity lines, statement,
 * legend and margin note are prose and stay here.
 *
 * Not checked: whether the buyer is registered, whether an ABN or GST number
 * is live (an ABR or IRD lookup), or whether a treatment is right for the goods.
 *
 * Installed by `dxcl add gst-inclusive-statement`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface GstInclusiveStatementData {
  gst: {
    /**
     * Whose GST rules apply. `AU`: 10%, A New Tax System (GST) Act 1999, amounts in AUD.
     * `NZ`: 15%, Goods and Services Tax Act 1985, amounts in NZD. Required: anything
     * else throws.
     */
    country: "AU" | "NZ";
    /**
     * Whether the seller is registered for GST. Missing, or anything but `true`, means
     * not registered: heading "Invoice", no GST rows, no marks, no legend, no statement,
     * and no threshold checks. Claiming GST while unregistered is the unlawful
     * direction, so that is not the default.
     */
    sellerRegistered: boolean;
    /**
     * AU: the seller's ABN (11 digits; spaces allowed). NZ: the seller's GST number,
     * which is its IRD number (8 or 9 digits; spaces and hyphens allowed).
     * AU: required when registered and any line is taxable. Printed as
     * "ABN 51 824 753 556" whenever given, even when not registered, because quoting an
     * ABN is not a GST claim and avoids no-ABN withholding.
     * NZ: required when registered, a line is taxable and the inclusive total is above
     * NZ$200. Printed as "GST number 49-091-850" when given and registered, and ignored
     * when not registered.
     * Throws when given but failing its checksum, or when required and missing or blank.
     */
    sellerTaxId?: string;
    /**
     * Whether line amounts include GST. Required when registered and any line is
     * taxable: missing throws, because guessing moves the GST by a tenth. Ignored
     * otherwise.
     */
    pricesIncludeGst?: boolean;
    /**
     * How GST is rounded. `total-invoice` (the default when missing): one rounding, on
     * the total of the taxable lines. `taxable-sale`: round each taxable line's GST,
     * then add. Any other value throws.
     */
    rounding?: "total-invoice" | "taxable-sale";
    /** The items, in printing order. Missing or empty throws. */
    lines: Array<{
      /** What was sold, printed as given (trimmed). Missing or blank throws. */
      description: string;
      /**
       * The line amount in cents, on the basis `pricesIncludeGst` names. Must be an
       * integer (`Number.isInteger`), or it throws. Negative is allowed for a discount
       * or rebate line. A negative total, or a negative total of the taxable lines,
       * throws: adjustment notes are out of scope.
       */
      amountMinor: number;
      /**
       * How GST treats the line. AU accepts `taxable | gst-free | input-taxed |
       * margin-scheme`. NZ accepts `taxable | zero-rated | exempt`. Anything else,
       * including the other country's term, throws, naming the allowed list. Only
       * `taxable` is standard-rated. Every other treatment makes the invoice mixed and
       * adds no GST.
       */
      treatment: "taxable" | "gst-free" | "input-taxed" | "margin-scheme" | "zero-rated" | "exempt";
    }>;
    /**
     * Who is buying. Missing is fine unless a threshold needs it. Every field
     * given is printed, whatever the threshold.
     */
    buyer?: {
      /** The buyer's legal or trading name, printed as "Buyer: …". Blank counts as missing. */
      name?: string;
      /**
       * AU only (ignored for NZ): the buyer's ABN, printed as "Buyer ABN 98 765 432 122".
       * Throws when given and failing the checksum.
       */
      abn?: string;
      /**
       * NZ only (ignored for AU): one identifier besides the name, printed as
       * "Buyer <label>: <value>". A kind outside the union throws. A blank value counts
       * as missing. An `nzbn` must be 13 digits passing the GS1 check digit, or it
       * throws.
       */
      identifier?: {
        kind: "address" | "phone" | "email" | "nzbn" | "url" | "trading-name";
        value: string;
      };
    };
  };
}

/** Where the items and totals sit. `table` (the default) or `stacked`. */
export type GstInclusiveStatementArrangement = "table" | "stacked";

/** Props for {@linkcode GstInclusiveStatement}. */
export interface GstInclusiveStatementProps {
  /** Where the items and totals sit. Defaults to `"table"`. */
  arrangement?: GstInclusiveStatementArrangement;
}

export const GstInclusiveStatement: SectionComponent<GstInclusiveStatementProps> = async ({ arrangement }) => {
  const [gst] = useState((d: GstInclusiveStatementData) => d.gst);
  const s = await useDeriver(gstInclusiveStatement, [gst]);

  const lines: GstInclusiveStatementLines = {
    items: s.items,
    subtotal: s.subtotal,
    gst: s.gst,
    gstLabel: s.gstLabel,
    total: s.total,
    totalLabel: s.totalLabel,
  };

  return (
    <Section id="gst-statement" title={s.heading}>
      {s.sellerTaxId && <Paragraph id="gst-seller-tax-id">{s.sellerTaxId}</Paragraph>}
      {s.buyerName && <Paragraph id="gst-buyer-name">{s.buyerName}</Paragraph>}
      {s.buyerTaxId && <Paragraph id="gst-buyer-tax-id">{s.buyerTaxId}</Paragraph>}
      {s.buyerIdentifier && <Paragraph id="gst-buyer-identifier">{s.buyerIdentifier}</Paragraph>}
      {arrange(arrangement, {
        table: () => table(lines),
        stacked: () => stacked(lines),
      })}
      {s.statement && <Paragraph id="gst-statement-text">{s.statement}</Paragraph>}
      {s.legend && <Paragraph id="gst-legend">{s.legend}</Paragraph>}
      {s.marginNote && <Paragraph id="gst-margin-note">{s.marginNote}</Paragraph>}
    </Section>
  );
};
