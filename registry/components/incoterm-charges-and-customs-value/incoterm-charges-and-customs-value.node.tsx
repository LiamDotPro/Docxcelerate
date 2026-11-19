import { Cell, Paragraph, Row, Section, type Section as SectionComponent, Table, useDeriver, useState } from "docxcelerate/template";
import incotermChargesAndCustomsValue from "../derivers/incoterm-charges-and-customs-value.deriver.ts";

/**
 * A commercial-invoice section that states the Incoterms 2020 rule and lists
 * every cost leg of the shipment, from export packing to import VAT, then gives
 * the customs value on the basis the destination uses.
 *
 * Decisions already made for you:
 *
 * - No flat "shipping" line. Every leg is a named row with its amount, so
 *   customs can separate the goods price from the cost of delivering them. The
 *   US requires it by law (19 CFR 141.86(a)(8)), and the EU and UK need it
 *   before post-entry transport and import charges can come out of the value.
 * - Who pays each leg is fixed by the rule, not chosen by the author. A buyer's
 *   leg is printed and used for the customs value, but never enters the total.
 * - The customs value is not the invoice total. An EXW invoice is below the
 *   EU/UK CIF value, because the buyer's origin costs, freight and insurance to
 *   the border are added. A CIF invoice is above the US FOB value, because
 *   international freight and insurance come off. A DDP price includes duty,
 *   which comes back out. EU/UK air freight counts only at the Annex 23-01
 *   percentage.
 * - Four common mistakes are flagged in print: a sea-only rule used for air,
 *   container or land freight; a missing named place; DDP with the seller
 *   paying import VAT; and a CFR or CPT invoice that shows insurance.
 *
 * All of it is decided in one deriver, because every step does arithmetic on
 * request data, and publishing refuses that. The node computes nothing: it
 * reads the deriver's result and branches on it, with every test written inline
 * so a publish build keeps each arm.
 *
 * The legs are ten fixed slots rather than a `.map()`. The resolver reads a
 * published loop's collection from request data and ignores the derived scope,
 * so a `.map()` over the deriver's legs would publish and then resolve to no
 * rows. A three-column table is the only placement that keeps the total
 * checkable against the amounts above it, so there is no `arrangement` prop.
 *
 * The node publishes 15 branches against the default `branchLimit` of 32: four
 * warnings, ten legs and the customs note.
 *
 * Installed by `dxcl add incoterm-charges-and-customs-value`.
 */

/** The ten cost legs of a shipment, in the order the goods meet them. */
export type IncotermLegKind =
  | "export-packing"
  | "origin-inland"
  | "export-clearance"
  | "origin-terminal"
  | "main-carriage"
  | "insurance"
  | "destination-terminal"
  | "onward-delivery"
  | "import-duty"
  | "import-tax";

/** What this component reads. Add these fields to your document data type. */
export interface IncotermChargesAndCustomsValueData {
  shipment: {
    /**
     * The Incoterms 2020 rule the goods are sold under. Required: the deriver
     * throws on anything else, including a 2010-only rule such as DAT, because
     * the cost split is the 2020 one.
     */
    incoterm: "EXW" | "FCA" | "FAS" | "FOB" | "CFR" | "CIF" | "CPT" | "CIP" | "DAP" | "DPU" | "DDP";
    /**
     * The named place or port the rule applies at, printed as given (trimmed),
     * such as "Port of Felixstowe". Missing or blank prints the term without a
     * place and prints the named-place warning: a rule with no place does not
     * say where costs pass.
     */
    namedPlace?: string;
    /**
     * How the main carriage travels. `"sea"` means bulk or break-bulk on a
     * vessel; `"sea-container"` means containerised ocean freight. Used for the
     * sea-only warning and for EU/UK air-freight apportionment. Missing means
     * unknown: no mode warning, and air freight is counted in full. Throws on
     * any other string.
     */
    mode?: "sea" | "sea-container" | "air" | "road" | "rail" | "multimodal";
    /**
     * ISO 4217 code every amount is in. It also sets the minor-unit exponent
     * (two for EUR, none for JPY). Required: throws when missing or not a code
     * `Intl` knows.
     */
    currency: string;
    /**
     * The goods price in integer minor units. Legs marked `inPrice` are inside
     * this figure, not added to it. Required: throws unless it is a
     * non-negative integer. Zero is accepted and prints as zero.
     */
    goodsMinor: number;
    /**
     * Every cost leg there is a figure for, in any order. They print in the
     * journey order of `IncotermLegKind`. At most one leg per kind: a duplicate
     * throws, so add the amounts together first.
     *
     * A leg left out is treated as zero and prints no row, with one exception:
     * the customs value cannot be stated when a figure it needs is missing. For
     * a `cif-border` destination, supply the buyer's legs up to the border too,
     * or the value understates. Missing or empty means no legs.
     */
    legs?: Array<{
      /** Which leg. Throws on anything else. */
      kind: IncotermLegKind;
      /**
       * What it costs, in integer minor units. Required on every leg: throws
       * unless it is a non-negative integer. A buyer's leg's amount is printed
       * and used for the customs value, never added to the total.
       */
      amountMinor: number;
      /**
       * True when a leg the seller pays is already inside `goodsMinor`. It
       * prints "Included in price" with its amount and is not added again.
       * Missing means false: the leg is charged on top. Throws when true on a
       * leg the rule puts on the buyer, or when the included legs add up to
       * more than `goodsMinor`.
       */
      inPrice?: boolean;
    }>;
    /** How the importing country values goods for duty. Required: throws when missing. */
    destination: {
      /**
       * `"cif-border"`: the EU and UK, which value goods at the place of entry,
       * with freight and insurance to the border added. `"fob-export"`: the US,
       * Canada and Australia, which value goods at the place of export, with
       * international freight and insurance taken off. Required: throws on
       * anything else.
       */
      valuationBasis: "cif-border" | "fob-export";
      /**
       * The share of air freight that counts towards customs value, as a
       * whole-number percentage 0 to 100, from Annex 23-01 of Implementing
       * Regulation (EU) 2015/2447 or the UK's equivalent. Read only when
       * `valuationBasis` is `"cif-border"` and `mode` is `"air"`. Missing there
       * means 100%, and the customs note says so: over-declaring is the lawful
       * direction to be wrong in. Throws when present but not an integer from
       * 0 to 100.
       */
      airFreightDutiablePercent?: number;
    };
  };
}

/** Prints the delivery terms, every cost leg, the invoice total and the customs value. */
export const IncotermChargesAndCustomsValue: SectionComponent = async () => {
  const [shipment] = useState((d: IncotermChargesAndCustomsValueData) => d.shipment);
  const c = await useDeriver(incotermChargesAndCustomsValue, [shipment]);

  return (
    <Section id="incoterm-charges" title="Delivery terms and charges">
      <Paragraph id="incoterm-charges-term">Delivery terms: {c.term}</Paragraph>
      {c.warnSeaOnly && <Paragraph id="incoterm-charges-warning-sea-only">{c.warnSeaOnly}</Paragraph>}
      {c.warnNamedPlace && <Paragraph id="incoterm-charges-warning-named-place">{c.warnNamedPlace}</Paragraph>}
      {c.warnDdpVat && <Paragraph id="incoterm-charges-warning-ddp-vat">{c.warnDdpVat}</Paragraph>}
      {c.warnImpliedInsurance && (
        <Paragraph id="incoterm-charges-warning-implied-insurance">{c.warnImpliedInsurance}</Paragraph>
      )}
      <Table id="incoterm-charges-lines" columns={[{ width: "auto" }, { width: 45 }, { width: 32, align: "right" }]}>
        <Row id="incoterm-charges-header" header>
          <Cell>Charge</Cell>
          <Cell>Treatment</Cell>
          <Cell>Amount</Cell>
        </Row>
        <Row id="incoterm-charges-goods">
          <Cell>Goods</Cell>
          <Cell></Cell>
          <Cell>{c.goods}</Cell>
        </Row>
        {c.legs["export-packing"].label && (
          <Row id="incoterm-charges-export-packing">
            <Cell>{c.legs["export-packing"].label}</Cell>
            <Cell>{c.legs["export-packing"].treatment}</Cell>
            <Cell>{c.legs["export-packing"].amount}</Cell>
          </Row>
        )}
        {c.legs["origin-inland"].label && (
          <Row id="incoterm-charges-origin-inland">
            <Cell>{c.legs["origin-inland"].label}</Cell>
            <Cell>{c.legs["origin-inland"].treatment}</Cell>
            <Cell>{c.legs["origin-inland"].amount}</Cell>
          </Row>
        )}
        {c.legs["export-clearance"].label && (
          <Row id="incoterm-charges-export-clearance">
            <Cell>{c.legs["export-clearance"].label}</Cell>
            <Cell>{c.legs["export-clearance"].treatment}</Cell>
            <Cell>{c.legs["export-clearance"].amount}</Cell>
          </Row>
        )}
        {c.legs["origin-terminal"].label && (
          <Row id="incoterm-charges-origin-terminal">
            <Cell>{c.legs["origin-terminal"].label}</Cell>
            <Cell>{c.legs["origin-terminal"].treatment}</Cell>
            <Cell>{c.legs["origin-terminal"].amount}</Cell>
          </Row>
        )}
        {c.legs["main-carriage"].label && (
          <Row id="incoterm-charges-main-carriage">
            <Cell>{c.legs["main-carriage"].label}</Cell>
            <Cell>{c.legs["main-carriage"].treatment}</Cell>
            <Cell>{c.legs["main-carriage"].amount}</Cell>
          </Row>
        )}
        {c.legs["insurance"].label && (
          <Row id="incoterm-charges-insurance">
            <Cell>{c.legs["insurance"].label}</Cell>
            <Cell>{c.legs["insurance"].treatment}</Cell>
            <Cell>{c.legs["insurance"].amount}</Cell>
          </Row>
        )}
        {c.legs["destination-terminal"].label && (
          <Row id="incoterm-charges-destination-terminal">
            <Cell>{c.legs["destination-terminal"].label}</Cell>
            <Cell>{c.legs["destination-terminal"].treatment}</Cell>
            <Cell>{c.legs["destination-terminal"].amount}</Cell>
          </Row>
        )}
        {c.legs["onward-delivery"].label && (
          <Row id="incoterm-charges-onward-delivery">
            <Cell>{c.legs["onward-delivery"].label}</Cell>
            <Cell>{c.legs["onward-delivery"].treatment}</Cell>
            <Cell>{c.legs["onward-delivery"].amount}</Cell>
          </Row>
        )}
        {c.legs["import-duty"].label && (
          <Row id="incoterm-charges-import-duty">
            <Cell>{c.legs["import-duty"].label}</Cell>
            <Cell>{c.legs["import-duty"].treatment}</Cell>
            <Cell>{c.legs["import-duty"].amount}</Cell>
          </Row>
        )}
        {c.legs["import-tax"].label && (
          <Row id="incoterm-charges-import-tax">
            <Cell>{c.legs["import-tax"].label}</Cell>
            <Cell>{c.legs["import-tax"].treatment}</Cell>
            <Cell>{c.legs["import-tax"].amount}</Cell>
          </Row>
        )}
        <Row id="incoterm-charges-total">
          <Cell>Invoice total</Cell>
          <Cell></Cell>
          <Cell>{c.total}</Cell>
        </Row>
      </Table>
      <Paragraph id="incoterm-charges-customs-value">{c.customsLine}</Paragraph>
      {c.customsNote && <Paragraph id="incoterm-charges-customs-note">{c.customsNote}</Paragraph>}
    </Section>
  );
};
