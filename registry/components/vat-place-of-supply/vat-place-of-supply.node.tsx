import { type Nodes, Paragraph, useDeriver, useState } from "docxcelerate/template";
import vatPlaceOfSupply from "../derivers/vat-place-of-supply.deriver.ts";

/**
 * Whose VAT applies to a sale and at what rate, decided before anything is
 * printed, as one line saying what was chosen. No section and no heading, so
 * it can sit beside the totals on every invoice. The treatment and the rate are
 * also readable as `{{derived.vatPlaceOfSupply.treatment}}` and
 * `{{derived.vatPlaceOfSupply.rate}}`, so the totals can follow the decision.
 *
 * Decisions already made for you:
 *
 * - The EUR 10,000 distance-selling threshold is EU-wide, not per country. It
 *   is one figure: all of the seller's cross-border B2C sales of goods and
 *   digital services to other member states, added together. It is tested on
 *   the current year and on the preceding calendar year, and either one over
 *   means VAT is due at destination.
 * - Over the threshold, or opted in to the Union OSS, means the customer's
 *   country's rate. Under it, the seller's own rate applies.
 * - A VAT number counts only once it is marked validated. A string shaped like
 *   a VAT number is not proof of business status. A business claim with a
 *   missing or unvalidated number is charged VAT, never reverse charge, and
 *   the invoice says why.
 * - A consumer never gets reverse charge.
 * - Rates come from the document's data, never from a table in the code. A
 *   hard-coded table goes stale: Finland moved to 25.5% in September 2024,
 *   Slovakia to 23% in January 2025, Romania to 21% and Estonia to 24% in 2025.
 *   The component holds the rules and the caller holds the rates.
 * - Northern Ireland (`XI`) counts as the EU for goods and as the UK for
 *   services.
 *
 * Every comparison here is on request data, so the decision is made by one
 * deriver and the node only prints what it returns. Choosing the optional note
 * is an ordinary conditional.
 *
 * Not covered: reduced rates, the special place-of-supply rules for services
 * (property, events, transport, catering, hire), imports and the IOSS,
 * domestic reverse charges, checking a VAT number, keeping the sales total,
 * and calculating amounts. `outside-scope` is not a treatment
 * `tax-treatment-notice` knows.
 *
 * Installed by `dxcl add vat-place-of-supply`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface VatPlaceOfSupplyData {
  seller: {
    /**
     * Where the seller is established: an EU member state, `"GB"`, or `"XI"`
     * for Northern Ireland. An ISO 3166-1 alpha-2 code. It is trimmed and
     * upper-cased, and `"EL"` is read as Greece (`"GR"`). Required: the
     * deriver throws when it is missing, is not two letters, or is outside the
     * EU, GB and XI. Sellers established elsewhere are out of scope.
     */
    country: string;
    /**
     * True when the seller has opted to charge destination VAT on every EU
     * distance sale by registering for the Union OSS, even under the
     * threshold. Missing means false. When true, `euDistanceSales` is not read.
     */
    ossRegistered?: boolean;
    /**
     * The seller's EU-wide total, in euros, of cross-border B2C sales of goods
     * and digital services to other member states. It is one figure for the
     * whole EU, not one per country. Read only for a cross-border EU sale of
     * goods or digital services to a consumer, when `ossRegistered` is not
     * true. Missing there throws: the threshold cannot be guessed.
     */
    euDistanceSales?: {
      /** This calendar year so far, in EUR. Must be a finite number of 0 or more, or the deriver throws. */
      currentYear: number;
      /** The whole preceding calendar year, in EUR. Must be a finite number of 0 or more, or the deriver throws. */
      priorYear: number;
    };
  };
  customer: {
    /**
     * Where the customer is: who receives a service, or where a consumer
     * lives. ISO alpha-2, normalised like `seller.country`. Any country
     * outside the EU, GB and XI counts as "outside". Required: the deriver
     * throws when it is missing or not two letters.
     */
    country: string;
    /**
     * Whether the customer buys as a business. Missing means a consumer: B2B
     * is a claim that has to be made, and a missing claim must not turn into
     * a zero rate.
     */
    isBusiness: boolean;
    /**
     * The customer's VAT number. Needed, together with `vatNumberValidated`,
     * for reverse charge or an intra-Community supply. Missing, `""` or
     * whitespace on a cross-border EU business sale means VAT is charged and
     * the "unvalidated" note prints. The number itself is not printed here:
     * tax-treatment-notice prints it.
     */
    vatNumber?: string;
    /**
     * True only when the caller has checked the number, for example against
     * VIES, and it came back valid. Missing or false is treated like a
     * missing number. Only the literal `true` counts.
     */
    vatNumberValidated?: boolean;
  };
  supply: {
    /**
     * What is being sold. `digital` means telecommunications, broadcasting
     * and electronically supplied services (TBE). `services` means any other
     * service, under the general rules. Required: anything else throws.
     */
    kind: "goods" | "services" | "digital";
    /**
     * Goods only: the country the goods leave from. Missing means
     * `seller.country`. It must be in the EU, GB or XI, or the deriver throws
     * (imports are out of scope). If it differs from `seller.country`, the
     * distance-selling threshold does not apply.
     */
    shipsFrom?: string;
    /** Goods only: the country the goods arrive in. Missing means `customer.country`. */
    shipsTo?: string;
  };
  /**
   * The current standard VAT rate per country, as a percentage: `{ NL: 21,
   * FI: 25.5 }`. Keyed by upper-case ISO alpha-2. Greece is `GR`. Northern
   * Ireland uses `GB`. Only the rate the decision needs is read. Throws when
   * that rate is missing, not a finite number, below 0, or 100 or more.
   * Reduced rates are out of scope.
   */
  vatRates: Record<string, number>;
}

/**
 * The line saying which VAT applies and at what rate, and, when a business
 * claim was charged VAT, the note saying why.
 */
export const VatPlaceOfSupply: Nodes = async () => {
  const [{ seller, customer, supply, vatRates }] = useState((data: VatPlaceOfSupplyData) => ({
    seller: data.seller,
    customer: data.customer,
    supply: data.supply,
    vatRates: data.vatRates,
  }));
  // After every other hook.
  const pos = await useDeriver(vatPlaceOfSupply, [seller, customer, supply, vatRates]);

  return (
    <>
      <Paragraph id="vat-place-of-supply">{pos.line}</Paragraph>
      {pos.note && <Paragraph id="vat-place-of-supply-unvalidated">{pos.note}</Paragraph>}
    </>
  );
};
