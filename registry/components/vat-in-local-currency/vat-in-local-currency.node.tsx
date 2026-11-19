import { arrange, Paragraph, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import { block } from "../arrangements/vat-in-local-currency/block.arrangement.tsx";
import { columns } from "../arrangements/vat-in-local-currency/columns.arrangement.tsx";
import type { VatInLocalCurrencyLines } from "../arrangements/vat-in-local-currency/lines.ts";
import vatInLocalCurrency from "../derivers/vat-in-local-currency.deriver.ts";

/**
 * An invoice issued in one currency for VAT owed in another has to state the
 * VAT analysis in the currency the VAT is owed in. The rules are Art. 230 of
 * the VAT Directive and, in the UK, HMRC Notice 700 s.16: net per rate and VAT
 * per rate in sterling. This component restates the invoice's own VAT bands in
 * that currency.
 *
 * Decisions already made for you:
 *
 * - Every band is converted on its own. Net and VAT are converted separately
 *   and each is rounded once, to the minor units of the VAT currency. The
 *   local totals are the sums of those bands. They are never worked out by
 *   converting the gross total and backing VAT out of it, and never by adding
 *   up line-by-line conversions. Those methods drift a penny away from the
 *   figure an accounting system recomputes.
 * - The rate carries its direction. "1 EUR = 0.8571 GBP" multiplies and
 *   "1 GBP = 1.1667 EUR" divides, so a feed that quotes the inverse cannot
 *   turn the conversion upside down.
 * - The rate is checked against the tax point, not the invoice date, using
 *   the rule of the jurisdiction: in the UK the market selling rate, an HMRC
 *   period rate for the month of the tax point, or a rate agreed with HMRC; in
 *   the EU the latest selling rate or the ECB rate (Art. 91); in Germany the
 *   BMF monthly average; in Poland the NBP or ECB rate of the last working day
 *   before the tax point; in Bulgaria the fixed euro rate, used unrounded and
 *   never inverted.
 * - Advances keep their own rate. On a final invoice each advance is restated
 *   at the rate of its own chargeable event, so VAT already declared in local
 *   currency is not re-converted at today's rate.
 * - The rate, its direction, its source and its date are printed beside the
 *   figures. A wrong source or date is printed as a "Check before issuing"
 *   line. It is not silently accepted.
 *
 * All of this is decided in one deriver, because arithmetic on request data
 * cannot be published. This node only interpolates, loops and chooses between
 * the same-currency sentence and the figures.
 *
 * Installed by `dxcl add vat-in-local-currency`, which also installs its deriver and arrangements.
 */

/** One VAT rate band, in the currency of the document it belongs to. */
export interface VatBand {
  /**
   * The VAT rate as a percentage, a decimal string: "20", "5", "0", "7.7". A
   * number is accepted and read with String(n). At most 4 decimal places, and
   * it must not be negative. Required: missing or malformed throws, naming the
   * field (`vatConversion.bands[1].ratePercent`). Two bands with the same rate
   * ("20" and "20.0" count as the same) throw.
   */
  ratePercent: string | number;
  /**
   * The net at this rate, in major units of the invoice currency, a decimal
   * string: "1234.56". A number is accepted and read with String(n), so
   * 0.1 + 0.2 throws instead of being rounded. Negative for a credit note.
   * Zero is allowed. More decimal places than the currency has, or anything
   * that is not a plain decimal, throws, naming the field. Required.
   */
  net: string | number;
  /** The VAT at this rate, as the invoice states it. The same rules as `net`. Required. It is never recomputed from `net`. */
  vat: string | number;
}

/** An exchange rate with its direction stated: 1 `base` = `value` `quote`. */
export interface VatExchangeRate {
  /**
   * The rate as a decimal string, "0.8571". A number is accepted and read with
   * String(n). Must be greater than zero, with at most 10 decimal places.
   * Required. Printed exactly as given (trimmed).
   */
  value: string | number;
  /**
   * ISO 4217 code of the unit side. Together with `quote` it must be exactly
   * the invoice currency and the VAT currency, in either order. Anything else
   * throws, naming both currencies it expected. Base = invoice currency
   * multiplies. Base = VAT currency divides.
   */
  base: string;
  /** ISO 4217 code of the priced side. See `base`. Required. */
  quote: string;
  /**
   * Where the rate comes from. Required: anything outside the union throws. A
   * source the jurisdiction does not accept is still used, and a warning is
   * printed. `fixed-euro` is allowed only for the EUR/BGN pair (it throws
   * otherwise). That pair must use `fixed-euro`, with base EUR and value
   * exactly "1.95583". Anything else for that pair throws.
   */
  source: "market-selling" | "hmrc-period" | "ecb" | "nbp" | "bmf-monthly" | "fixed-euro" | "agreed";
  /**
   * The date the rate applies to, ISO 8601 (characters after the tenth are
   * ignored). For a period rate (`hmrc-period`, `bmf-monthly`) this is any day
   * in the month the rate was published for. Required except for `fixed-euro`,
   * where it is ignored. Missing or malformed throws.
   */
  rateDate?: string;
}

/** An advance invoiced earlier and deducted on this, the final invoice. */
export interface VatAdvance {
  /** The advance invoice's number, printed as given (trimmed). Required: missing or blank throws. */
  invoiceNumber: string;
  /**
   * When VAT on the advance became chargeable (the payment was received), ISO
   * 8601. Required. Missing or malformed throws, and so does a date after
   * `vatConversion.taxPointDate`. The advance's rate is checked against this
   * date, never against the final tax point.
   */
  chargeableDate: string;
  /**
   * The advance's VAT bands in the invoice currency, with the same rules as
   * `vatConversion.bands`. At least one entry. Every rate must also appear in
   * `vatConversion.bands`, or it throws.
   */
  bands: VatBand[];
  /** The rate the advance was declared at. The same rules as `vatConversion.rate`. Required. */
  rate: VatExchangeRate;
}

/** What this component reads. Add these fields to your document data type. */
export interface VatInLocalCurrencyData {
  vatConversion: {
    /**
     * Whose conversion rules apply. `EU` is a member state with no rules of
     * its own here, so the Directive's defaults (Art. 91) apply. Required:
     * anything else throws.
     */
    jurisdiction: "UK" | "EU" | "DE" | "PL" | "BG";
    /** ISO 4217 code the invoice is issued in. Required. A code Intl rejects throws. */
    invoiceCurrency: string;
    /**
     * ISO 4217 code the VAT is owed in. Required. It must fit the
     * jurisdiction, or it throws: UK is GBP, DE is EUR, PL is PLN, BG is BGN
     * for a tax point before 2026-01-01 and EUR from then on. EU accepts any
     * valid code. When it equals `invoiceCurrency`, nothing is converted and
     * the component prints one sentence saying so.
     */
    vatCurrency: string;
    /** The tax point (time of supply, chargeable event), ISO 8601. Required. Missing or malformed throws. */
    taxPointDate: string;
    /**
     * The date of issue, ISO 8601. Read only in PL: an invoice issued before
     * its tax point takes the rate from before the issue date (art. 31a(2)).
     * Required when jurisdiction is PL (missing throws). Elsewhere it is
     * optional, and validated when present.
     */
    issueDate?: string;
    /**
     * The invoice's VAT analysis, one band per rate, in the invoice currency,
     * printed in the order given. With `advances` these are the whole supply,
     * as a final invoice lists it. At least one band, or it throws.
     */
    bands: VatBand[];
    /** The rate for this invoice's tax point. Required. */
    rate: VatExchangeRate;
    /**
     * Advances already invoiced and deducted here. Missing or empty means
     * none. When given, every band must be zero or positive (it throws
     * otherwise). After the advances are subtracted, every band's remainder
     * must be zero or positive, or it throws, naming the rate.
     */
    advances?: VatAdvance[];
  };
}

/** Where the local-currency figures sit: beside each band, or in a statement of their own. */
export type VatInLocalCurrencyArrangement = "columns" | "block";

/** Props for {@linkcode VatInLocalCurrency}. */
export interface VatInLocalCurrencyProps {
  /** Where the local figures sit. Defaults to `"columns"`. */
  arrangement?: VatInLocalCurrencyArrangement;
}

export const VatInLocalCurrency: SectionComponent<VatInLocalCurrencyProps> = async ({ arrangement }) => {
  const [conversion] = useState((data: VatInLocalCurrencyData) => data.vatConversion);
  const local = await useDeriver(vatInLocalCurrency, [conversion]);

  // Built before the decision: the compiler publishes an early return only when
  // a lone `return` follows it.
  const lines: VatInLocalCurrencyLines = {
    invoiceCurrency: local.invoiceCurrency,
    vatCurrency: local.vatCurrency,
    bands: local.bands,
    total: local.total,
    hasAdvances: local.hasAdvances,
    advances: local.advances,
  };

  if (!local.converted) {
    // Its own id: a published document cannot carry two nodes with the same one.
    return (
      <Section id="vat-in-local-currency-same-currency" title={`VAT in ${local.vatCurrency}`}>
        <Paragraph id="vat-local-same-currency">
          VAT on this invoice is charged in {local.sameCurrency}, the currency it is issued in, so no conversion is
          needed.
        </Paragraph>
      </Section>
    );
  }

  return (
    <Section id="vat-in-local-currency" title={`VAT in ${local.vatCurrency}`}>
      <Paragraph id="vat-local-statement">{local.statement}</Paragraph>
      {arrange(arrangement, { columns: () => columns(lines), block: () => block(lines) })}
      {local.warnings.map((warning) => (
        <Paragraph id="vat-local-warning">Check before issuing: {warning.text}</Paragraph>
      ))}
    </Section>
  );
};
