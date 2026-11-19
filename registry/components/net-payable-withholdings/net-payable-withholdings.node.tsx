import { arrange, Paragraph, Section, type Section as SectionComponent, useDeriver, useState } from "docxcelerate/template";
import type { NetPayableWithholdingsLines } from "../arrangements/net-payable-withholdings/lines.ts";
import { stacked } from "../arrangements/net-payable-withholdings/stacked.arrangement.tsx";
import { table } from "../arrangements/net-payable-withholdings/table.arrangement.tsx";
import netPayableWithholdings from "../derivers/net-payable-withholdings.deriver.ts";

/**
 * How a Spanish or Italian invoice total becomes the amount the customer
 * actually transfers: each step its own row, labelled with its legal basis,
 * and the rows always add up. The decisions are already made here.
 *
 * Withholding comes off the base, never the VAT-inclusive total. IRPF is taken
 * from the base imponible; the ritenuta d'acconto from the fee plus the 4%
 * INPS rivalsa. A cassa contribution counts toward IVA but is not in the
 * ritenuta base.
 *
 * A consumer is never withheld from. In Spain every business, professional or
 * public-administration customer withholds; in Italy only a sostituto
 * d'imposta does, and never from a forfettario supplier.
 *
 * Split payment still shows the IVA, as a charged row, and then takes it off
 * the net payable. It is never shown as 0% IVA, and it cannot be combined with
 * a ritenuta.
 *
 * The 2.00 EUR bollo is due when the amount carrying no IVA is over 77.47 EUR.
 * It is added to the total only when re-charged; otherwise it is a note.
 *
 * Every amount is rounded to the cent on its own, half up, in integer cents,
 * in legal order, each step on the rows above it, so the total and the net are
 * figures a reader can check by adding the rows. Unclear or contradictory data
 * fails the document: a guessed withholding gives a wrong payable either way.
 *
 * The labels are in the country's language whatever the document locale,
 * because the mention requires the legal wording. All of it is one deriver,
 * since every figure is arithmetic on request data and publishing refuses
 * arithmetic. Rows are fixed slots, each printed by rule rather than by amount.
 *
 * Arrangement: where the rows sit is a separate choice, made with the
 * `arrangement` prop. This file decides every row and note; each arrangement
 * only places them.
 *
 * Installed by `dxcl add net-payable-withholdings`, which also installs its deriver.
 */

/** What this component reads. Add these fields to your document data type. */
export interface NetPayableWithholdingsData {
  invoice: {
    /** Whose rules apply: `ES` Spain, `IT` Italy. Required: anything else throws. */
    country: "ES" | "IT";
    /**
     * The price or fee before any contribution, VAT, recargo or withholding,
     * in euros: `"1000.00"`. Spain: the base imponible. Italy: the compenso or
     * corrispettivo, before rivalsa INPS and cassa. Required: throws when
     * missing, negative, or with more than two decimals. `"0"` is allowed.
     */
    base: string | number;
    /**
     * The VAT rate as a percentage. ES: `"21"`, `"10"`, `"4"` or `"0"`. IT:
     * `"22"`, `"10"`, `"5"`, `"4"` or `"0"`. `"0"` means no VAT is charged;
     * the mention saying why belongs to `tax-treatment-notice`. Required,
     * except for an Italian forfettario supplier, where it must be missing or
     * `"0"`. Any other value throws.
     */
    vatPercent?: string | number;
    /**
     * Expenses paid in the customer's name and re-charged at cost, outside VAT
     * and withholding: suplidos in Spain, anticipazioni in Italy. Added to the
     * total on its own row. Counts toward the Italian bollo trigger. Missing
     * or `"0"` means none.
     */
    exemptAmount?: string | number;
    /**
     * Italy only: the customer pays the IVA to the Treasury (art. 17-ter DPR
     * 633/1972). IVA is shown, then deducted from the net payable. Throws when
     * true and: the country is ES, the customer is a consumer, no IVA is
     * charged, or a ritenuta applies.
     */
    splitPayment?: boolean;
    /**
     * Italy only: the 2.00 EUR bollo is charged on and added to the total.
     * Otherwise, when due, it is printed as a note. Throws when true in ES.
     */
    stampRecharged?: boolean;
  };
  supplier: {
    /**
     * An individual doing professional work whose fees carry withholding.
     * Missing means false, and nothing is withheld.
     */
    professional?: boolean;
    /**
     * Spain only: withheld at 7% instead of 15%. The caller decides this.
     * Throws when true without `professional`, or in Italy.
     */
    newProfessional?: boolean;
    /** Italy only: regime forfettario. No IVA, never a ritenuta. Throws in Spain when true. */
    forfettario?: boolean;
    /**
     * Italy only: the supplier charges the 4% rivalsa INPS. Carries IVA and is
     * in the ritenuta base. Throws with `cassaPercent`, and in Spain.
     */
    rivalsaInps?: boolean;
    /**
     * Italy only: the cassa's contributo integrativo as a percentage, such as
     * `"4"`. Carries IVA, not in the ritenuta base. Throws unless above 0 and
     * at most 10 with up to two places, and in Spain.
     */
    cassaPercent?: string | number;
    /** The cassa's name as printed in its row. Ignored without `cassaPercent`. */
    cassaName?: string;
  };
  customer: {
    /** Who is paying. A `consumer` is never withheld from. Required. */
    type: "business" | "consumer" | "public-administration";
    /** Spain only: a retailer in the recargo de equivalencia buying goods. Throws in Italy when true. */
    reRegime?: boolean;
    /**
     * Italy only: the customer is a sostituto d'imposta. Must be a boolean
     * whenever a ritenuta could apply; missing then throws. Throws when true
     * for a consumer. Ignored in Spain.
     */
    withholdingAgent?: boolean;
  };
}

/**
 * Where the rows sit.
 *
 * - `table`: labels in a column with the figures right-aligned beside them.
 *   The default.
 * - `stacked`: one `Label: figure` paragraph per row, for a narrow column.
 */
export type NetPayableWithholdingsArrangement = "table" | "stacked";

/** Props for {@linkcode NetPayableWithholdings}. */
export interface NetPayableWithholdingsProps {
  /** Where the rows sit. Defaults to `"table"`. */
  arrangement?: NetPayableWithholdingsArrangement;
}

export const NetPayableWithholdings: SectionComponent<NetPayableWithholdingsProps> = async ({ arrangement }) => {
  const [invoice] = useState((data: NetPayableWithholdingsData) => data.invoice);
  const [supplier] = useState((data: NetPayableWithholdingsData) => data.supplier);
  const [customer] = useState((data: NetPayableWithholdingsData) => data.customer);
  const p = await useDeriver(netPayableWithholdings, [invoice, supplier, customer]);

  // Field by field: a spread of a derived stand-in is refused at publish.
  const lines: NetPayableWithholdingsLines = {
    baseLabel: p.baseLabel,
    base: p.base,
    rivalsaLabel: p.rivalsaLabel,
    rivalsa: p.rivalsa,
    cassaLabel: p.cassaLabel,
    cassa: p.cassa,
    vatLabel: p.vatLabel,
    vat: p.vat,
    reLabel: p.reLabel,
    re: p.re,
    exemptLabel: p.exemptLabel,
    exempt: p.exempt,
    stampLabel: p.stampLabel,
    stamp: p.stamp,
    totalLabel: p.totalLabel,
    total: p.total,
    withholdingLabel: p.withholdingLabel,
    withholding: p.withholding,
    splitPaymentLabel: p.splitPaymentLabel,
    splitPayment: p.splitPayment,
    netLabel: p.netLabel,
    net: p.net,
  };

  return (
    <Section id="net-payable" title={p.title}>
      {arrange(arrangement, {
        table: () => table(lines),
        stacked: () => stacked(lines),
      })}
      {p.withholdingNote && <Paragraph id="net-payable-withholding-note">{p.withholdingNote}</Paragraph>}
      {p.stampNote && <Paragraph id="net-payable-stamp-note">{p.stampNote}</Paragraph>}
    </Section>
  );
};
