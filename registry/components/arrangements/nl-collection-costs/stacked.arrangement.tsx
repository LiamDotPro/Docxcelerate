import { Paragraph, type Yield } from "docxcelerate/template";
import type { NlCollectionCostsLines } from "./lines.ts";

/** One labelled figure under another, below the statement. The default: it reads in any width. */
export const stacked = (lines: NlCollectionCostsLines): Yield => (
  <>
    <Paragraph id="nl-collection-costs-invoices">Facturen: {lines.invoices}</Paragraph>
    <Paragraph id="nl-collection-costs-principal">Hoofdsom: {lines.principal}</Paragraph>
    {lines.interest && (
      <Paragraph id="nl-collection-costs-interest">Wettelijke rente bijgeteld: {lines.interest}</Paragraph>
    )}
    <Paragraph id="nl-collection-costs-costs">Incassokosten: {lines.costs}</Paragraph>
    {lines.vat && <Paragraph id="nl-collection-costs-vat">Btw (21%) over de incassokosten: {lines.vat}</Paragraph>}
    <Paragraph id="nl-collection-costs-total">{lines.totalLabel}: {lines.total}</Paragraph>
  </>
);
