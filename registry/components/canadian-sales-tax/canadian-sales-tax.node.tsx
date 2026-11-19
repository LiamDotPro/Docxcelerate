import { arrange, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { CanadianSalesTaxLines } from "../arrangements/canadian-sales-tax/lines.ts";
import { stacked } from "../arrangements/canadian-sales-tax/stacked.arrangement.tsx";
import { table } from "../arrangements/canadian-sales-tax/table.arrangement.tsx";
import canadianSalesTax from "../derivers/canadian-sales-tax.deriver.ts";

/**
 * The sales-tax block of a Canadian invoice: subtotal, each tax on its own
 * base, the tax-included total, the registration numbers, and what is missing
 * for the customer to claim it.
 *
 * Decisions already made for you:
 * - Tax follows the province of supply — where goods are delivered, where the
 *   customer is for services — never the seller's province. There is no
 *   seller-province field to get wrong.
 * - HST provinces get one line. BC and SK get GST plus PST, MB GST plus RST,
 *   QC GST plus QST. Alberta and the territories get GST alone. Never a
 *   combined 13% or 14.975% on one base.
 * - Each tax has its own base, because PST and QST exempt different things
 *   than GST. QST is charged on the price before GST, as it has been since
 *   2013, never tax on tax. Each tax is rounded once on its base, half away
 *   from zero, not line by line.
 * - Rates are effective-dated rows picked by the tax point date, because they
 *   change mid-year: NS HST went from 15% to 14% on 1 April 2025.
 * - A seller not registered for GST/HST charges none, and the block says so.
 * - CRA's invoice tiers are applied to the tax-included total (under $100,
 *   $100 to $499.99, $500 or more; not the old $30 and $150). Whatever the
 *   tier needs and the data lacks is printed as a warning.
 * - A GST/HST number is printed only if it passes the check digit. A wrong
 *   number is worse than none.
 *
 * Everything is decided in one deriver, because totals, rounding and date
 * comparisons on request data cannot be done at publish time. The node prints
 * what it returns.
 *
 * Installed by `dxcl add canadian-sales-tax`, which also installs its deriver.
 */

/** The thirteen provinces and territories, as Canada Post codes. */
export type CanadianProvince =
  | "AB" | "BC" | "MB" | "NB" | "NL" | "NS" | "NT" | "NU" | "ON" | "PE" | "QC" | "SK" | "YT";

/**
 * One effective-dated rate. Rows are picked by `salesTax.taxPointDate`;
 * exactly one row must cover the date for every tax the province charges.
 */
export interface CanadianTaxRate {
  /** Which tax. `GST` is federal; HST is the harmonized rate in one line. */
  tax: "GST" | "HST" | "PST" | "RST" | "QST";
  /** The province the rate belongs to. Required for HST, PST, RST and QST; ignored for GST. */
  province?: CanadianProvince;
  /**
   * The rate as a percentage: `5`, `9.975`. At most three decimal places;
   * more, a negative value, or a value that is not a finite number throws.
   */
  rate: number;
  /** First day the rate applies, ISO 8601 (`2025-04-01`). Inclusive. */
  from: string;
  /** Last day it applies, ISO 8601, inclusive. Missing means still in force. */
  to?: string;
}

/** What this component reads. Add these fields to your document data type. */
export interface CanadianSalesTaxData {
  salesTax: {
    /**
     * The province of supply: where goods are delivered, or where the
     * customer is for services. Not the seller's province. Required; anything
     * outside `CanadianProvince` throws.
     */
    supplyProvince: CanadianProvince;
    /**
     * The day tax became payable, ISO 8601: the earlier of the day the
     * consideration was paid and the day it fell due — usually the invoice
     * date. Picks the rate rows. Anything after the tenth character is
     * ignored. Required: missing, not YYYY-MM-DD, or before 2013-04-01 (when
     * the current federal/provincial structure began) throws.
     */
    taxPointDate: string;
    /**
     * Every line of the invoice, at its price before any tax, after line
     * discounts. Required (an empty array is allowed and prints zeros);
     * missing or not an array throws.
     */
    lines: Array<{
      /**
       * What was supplied. Not printed here; read for the $500 tier check and
       * named in the "No GST on" note. Missing prints as "line N" in that
       * note, and on a $500+ invoice adds a warning.
       */
      description?: string;
      /**
       * The pre-tax amount in Canadian dollars, to the cent. Negative is a
       * discount or credit line. Not a finite number throws.
       */
      amount: number;
      /**
       * GST/HST status. Missing means `taxable` — the Excise Tax Act's default.
       * Any other string throws.
       */
      gstCategory?: "taxable" | "zero-rated" | "exempt";
      /**
       * PST/RST/QST status, read only in BC, SK, MB and QC. Missing follows
       * the GST status: `taxable` if `gstCategory` is taxable, otherwise
       * `exempt`. Any other string throws.
       */
      provincialCategory?: "taxable" | "exempt";
    }>;
    /**
     * Rate rows. Missing or empty uses the table built into the deriver. When
     * given, they replace that table entirely; they are not merged with it.
     */
    rates?: CanadianTaxRate[];
    /** Payment terms. Not printed; its absence warns on a $500+ invoice. */
    paymentTerms?: string;
  };
  seller: {
    /** The seller's name. Not printed here (the letterhead does); missing warns at every tier. */
    name?: string;
    /**
     * Whether the seller is a GST/HST registrant. Missing means not registered
     * (a small supplier): no GST/HST line, and the small-supplier note prints.
     * Charging GST/HST without being registered is the worse error.
     */
    gstRegistered?: boolean;
    /**
     * The GST/HST number: 9-digit BN, `RT`, 4-digit account, as in
     * `843219767RT0001`. Spaces and hyphens are removed and letters upper-cased
     * before checking; printed in that normalised form. Read only when
     * `gstRegistered` is true. Missing: from $100 a warning prints. Present but
     * invalid: a warning prints at every tier and the number is not printed.
     */
    bn?: string;
    /**
     * Whether the seller is registered for the supply province's PST (BC, SK),
     * RST (MB) or QST (QC). False means no provincial line. Missing means the
     * same, plus a warning that it was not stated. Ignored elsewhere.
     */
    provincialRegistered?: boolean;
    /**
     * The QST number: 10 digits, `TQ`, 4 digits (`1219876543TQ0001`).
     * Normalised like `bn`. Read only when QST is charged. Missing: from $100 a
     * warning prints. Present but invalid: a warning at every tier, not printed.
     */
    qst?: string;
  };
  /** The customer. The whole object may be missing. */
  customer?: {
    /** The customer's name. Not printed here; its absence warns on a $500+ invoice. */
    name?: string;
  };
}

/**
 * Where the lines sit.
 *
 * - `table`: labels on the left, amounts right-aligned beside them. The default.
 * - `stacked`: one labelled line under another, for a narrow column or a letter.
 */
export type CanadianSalesTaxArrangement = "table" | "stacked";

/** Props for {@linkcode CanadianSalesTax}. */
export interface CanadianSalesTaxProps {
  /** Where the lines sit. Defaults to `"table"`. */
  arrangement?: CanadianSalesTaxArrangement;
}

export const CanadianSalesTax: SectionComponent<CanadianSalesTaxProps> = async ({ arrangement }) => {
  const [salesTax] = useState((data: CanadianSalesTaxData) => data.salesTax);
  const [seller] = useState((data: CanadianSalesTaxData) => data.seller);
  const [customer] = useState((data: CanadianSalesTaxData) => data.customer);
  const result = await useDeriver(canadianSalesTax, [salesTax, seller, customer]);

  const lines: CanadianSalesTaxLines = {
    subtotal: result.subtotal,
    federalLabel: result.federalLabel,
    federalAmount: result.federalAmount,
    provincialLabel: result.provincialLabel,
    provincialAmount: result.provincialAmount,
    total: result.total,
    bn: result.bn,
    qst: result.qst,
    smallSupplierNote: result.smallSupplierNote,
    statusNote: result.statusNote,
    warning: result.warning,
  };

  return (
    <Section id="canadian-sales-tax" title="Sales tax">
      {arrange(arrangement, {
        table: () => table(lines),
        stacked: () => stacked(lines),
      })}
    </Section>
  );
};
