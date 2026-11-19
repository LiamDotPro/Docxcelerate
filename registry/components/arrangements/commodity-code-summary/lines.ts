/**
 * One goods line, every cell already decided and formatted. "" means nothing
 * to print.
 *
 * Installed by `dxcl add commodity-code-summary`, with the arrangements.
 */
export interface CommodityCodeSummaryLine {
  /** The invoice's line reference and description. */
  item: string;
  /** The tariff code as text, or "None". */
  tariffCode: string;
  /** The two-letter country of origin, or "None". */
  origin: string;
  /** "Yes" or "No". */
  preference: string;
  /** The invoice quantity. */
  quantity: string;
  /** Supplementary quantity and unit ("2,400 pa"), or "". */
  supplementary: string;
  /** Net mass in kilograms, 3 decimals. */
  netMass: string;
  /** The line's value, formatted in the invoice currency. */
  value: string;
  /** What is wrong with this line, or "". */
  warning: string;
}

/** One summary row: lines sharing a tariff code, origin and preference claim. */
export interface CommodityCodeSummaryGroup {
  /** The tariff code as text, or "None". */
  tariffCode: string;
  /** The two-letter country of origin, or "None". */
  origin: string;
  /** "Yes" or "No". */
  preference: string;
  /** "1 line" or "n lines". */
  lineCount: string;
  /** Summed supplementary quantity and unit, or "". */
  supplementary: string;
  /** Net mass in kilograms, rounded once from the exact sum. */
  netMass: string;
  /** The summed value, formatted in the invoice currency. */
  value: string;
  /** Why the supplementary quantity was not totalled, or "". */
  warning: string;
}

/**
 * What the commodity code summary hands each of its arrangements: every part
 * of the block, already decided. An arrangement places these and decides
 * nothing, so both print the same lines under the same conditions.
 *
 * A "" field is a line this invoice does not carry. Test it in the JSX,
 * `{lines.attention && …}`, so the test is compiled into a condition the
 * engine settles per document.
 *
 * Installed by `dxcl add commodity-code-summary`, with the arrangements.
 */
export interface CommodityCodeSummaryLines {
  /** How the codes were checked and what the masses are. */
  basis: string;
  /** How many lines need attention, or "". */
  attention: string;
  /** The goods lines, in invoice order. */
  lines: CommodityCodeSummaryLine[];
  /** The summary rows, in tariff order. */
  groups: CommodityCodeSummaryGroup[];
  /** Net mass of the groups added together, in kilograms. */
  totalNetMass: string;
  /** Value of the groups added together. */
  totalValue: string;
  /** The declared gross mass in kilograms, or "". */
  grossMass: string;
  /** The net-over-gross warning, or "". */
  grossWarning: string;
}
