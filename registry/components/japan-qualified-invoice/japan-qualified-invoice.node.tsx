import { arrange, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { JapanQualifiedInvoiceLines } from "../arrangements/japan-qualified-invoice/lines.ts";
import { stacked } from "../arrangements/japan-qualified-invoice/stacked.arrangement.tsx";
import { table } from "../arrangements/japan-qualified-invoice/table.arrangement.tsx";
import japanQualifiedInvoice from "../derivers/japan-qualified-invoice.deriver.ts";

/**
 * The tax part of a Japanese invoice under the qualified invoice system
 * (インボイス制度, from 1 October 2023): the lines, the totals per tax rate, the
 * tax per rate, the grand total, the registration number and the transaction
 * date. The issuer's name is a required item too, but the letterhead prints it.
 *
 * Tax is worked once per rate per invoice, never per line: the NTA forbids
 * summing tax that was rounded line by line. The issuer chooses the rounding,
 * and there is no default, because a silent one could be one yen out from the
 * issuer's books.
 *
 * The basis is respected: tax-inclusive prices are taxed at 10/110 and 8/108,
 * never at 10% or 8%.
 *
 * Every reduced-rate line carries ※, and the legend 「※は軽減税率対象」 is
 * printed whenever one does, on a simplified invoice too.
 *
 * The registration number must be valid: T and 13 digits, with its check digit
 * verified. A wrong number fails the build instead of printing.
 *
 * A simplified qualified invoice (適格簡易請求書) only drops the buyer's name.
 * It keeps the T-number, the date, the marker, the rate and the tax. The
 * caller vouches that the business may issue one.
 *
 * A seller with no registration number gets a plain 「請求書」: no T-number, no
 * 適格 wording, no stated consumption tax. It does carry a note of the share
 * of tax the buyer may still deduct under the transitional measure. That share
 * comes from a date table, not a constant, because the FY2026 reform replaced
 * the old 80% then 50% schedule with 80, 70, 50 and 30.
 *
 * All of this is decided in one deriver, because it is arithmetic, grouping
 * and date work on request data, and a publish build refuses all three. The
 * node prints what the deriver returns. It is 18 branches: 10 slots plus
 * registration, buyer, standard total, standard tax, reduced total, reduced
 * tax, legend and note, against the default `branchLimit` of 32.
 *
 * Words: Japanese only, with no `defineMessages`. This is statutory wording,
 * and a translation would not be a qualified invoice.
 *
 * Not covered: qualified return invoices (a negative band throws), billing
 * periods (one transaction date only), lines that are not taxable at 10% or 8%
 * (非課税, 不課税 and 免税 throw), more than 10 lines (throws), whether a
 * simplified invoice is allowed for the business, and whether the T-number is
 * live on the NTA site. Only its format and check digit are verified.
 *
 * Arrangement: where the lines sit is a separate choice from what they say.
 * This file decides every line; each arrangement, in
 * `arrangements/japan-qualified-invoice/`, only places them.
 *
 * Installed by `dxcl add japan-qualified-invoice`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface JapanQualifiedInvoiceData {
  jpInvoice: {
    /**
     * The seller's qualified invoice issuer registration number (登録番号):
     * "T" and 13 digits, e.g. "T9234567890123". Spaces and hyphens are
     * ignored and a lower-case "t" is accepted. It is printed normalised.
     * The first digit is a check digit, which is verified. A wrong format or
     * a wrong check digit throws.
     * Missing, empty or blank means the seller is not registered. The
     * component then prints 「請求書」 with no T-number and no stated
     * consumption tax, plus the transitional-measure note.
     */
    registrationNumber?: string;
    /**
     * True for a simplified qualified invoice (適格簡易請求書). Only retail,
     * restaurants, taxis, parking and travel businesses may issue one. The
     * caller vouches for that, because the component cannot check it.
     * The only effect is that `buyerName` becomes optional. The T-number,
     * date, ※ marker, rate and tax are all still printed.
     * Missing means false. Throws when true and there is no registration
     * number: a simplified invoice without a T-number is not one.
     */
    simplified?: boolean;
    /**
     * How the per-rate tax is rounded to whole yen, applied once per rate on
     * this invoice: "floor" (切捨て), "ceil" (切上げ) or "half-up" (四捨五入,
     * .5 rounds up).
     * Required when registered: missing or anything else throws, because a
     * silent default could be ¥1 out from the issuer's books. Ignored when
     * not registered.
     */
    rounding: "floor" | "ceil" | "half-up";
    /**
     * Whether `amountYen` is before tax ("exclusive": tax = total × rate/100)
     * or includes it ("inclusive": tax = total × 10/110 or 8/108).
     * Required when registered: missing or anything else throws. Ignored when
     * not registered, where amounts are simply the prices charged.
     */
    basis: "exclusive" | "inclusive";
    /**
     * The transaction date (取引年月日), ISO 8601 "YYYY-MM-DD". Anything after
     * the tenth character is ignored. Printed as 「2026年11月30日」. When the
     * seller is not registered, it also selects the transitional deduction
     * share. Required: throws when missing, not a real calendar date, or
     * before 2023-10-01 (the start of the system).
     */
    transactionDate: string;
    /**
     * The buyer's name, printed exactly as given, honorific included:
     * "株式会社〇〇 御中", "山田太郎 様". The component never adds one,
     * because 御中 and 様 are not interchangeable.
     * Required on a full qualified invoice (registered, not simplified):
     * missing or blank throws. Otherwise it is printed when given and omitted
     * when missing.
     */
    buyerName?: string;
    /**
     * The items, in any order, 1 to 10 of them. They are printed 10% lines
     * first, then 8% lines, each group in input order.
     * Required: missing, empty, or more than 10 throws.
     */
    lines: Array<{
      /**
       * What was supplied, as printed. A reduced-rate line gets " ※"
       * appended. Required: a missing or blank description throws.
       */
      description: string;
      /**
       * The amount in whole yen, on the `basis` above. A negative value is a
       * discount line inside its rate band. Throws when it is not a safe
       * integer or when its magnitude exceeds 1,000,000,000,000.
       */
      amountYen: number;
      /**
       * "10" for the standard rate, "8" for the reduced rate (food and
       * drink, newspapers). Anything else throws, including 非課税 and 不課税
       * lines, which are out of scope.
       */
      rate: "10" | "8";
    }>;
  };
}

/** The placements `arrangement` chooses between. The first is the default. */
export type JapanQualifiedInvoiceArrangement = "table" | "stacked";

/** Props for {@linkcode JapanQualifiedInvoice}. */
export interface JapanQualifiedInvoiceProps {
  /** Where the lines sit. Missing means "table". */
  arrangement?: JapanQualifiedInvoiceArrangement;
}

/** Prints the heading as the Section title, then the invoice lines placed by `arrangement`. */
export const JapanQualifiedInvoice: SectionComponent<JapanQualifiedInvoiceProps> = async ({ arrangement }) => {
  const [jpInvoice] = useState((data: JapanQualifiedInvoiceData) => data.jpInvoice);
  const inv = await useDeriver(japanQualifiedInvoice, [jpInvoice]);

  // shortcut: ten fixed slots because a published loop ignores the derived scope;
  // switch to .map() once commit 70cafb6 (registry/india-gst-split) is on main.
  const lines: JapanQualifiedInvoiceLines = {
    buyer: jpInvoice.buyerName,
    date: inv.date,
    registration: inv.registration,
    amountHeader: inv.amountHeader,
    items: inv.items,
    standard: inv.standard,
    reduced: inv.reduced,
    totalLabel: inv.totalLabel,
    total: inv.total,
    legend: inv.legend,
    note: inv.note,
  };

  return (
    <Section id="jp-invoice" title={inv.heading}>
      {arrange(arrangement, {
        table: () => table(lines),
        stacked: () => stacked(lines),
      })}
    </Section>
  );
};
