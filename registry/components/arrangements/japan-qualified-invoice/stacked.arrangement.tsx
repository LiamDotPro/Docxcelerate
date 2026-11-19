import { Paragraph, type Yield } from "docxcelerate/template";
import type { JapanQualifiedInvoiceLines } from "./lines.ts";

/**
 * One Paragraph per item, band line and total down the text column. The
 * receipt layout simplified invoices from shops and restaurants are usually
 * printed in, and it reads at any width.
 */
export const stacked = (lines: JapanQualifiedInvoiceLines): Yield => (
  <>
    {lines.buyer && <Paragraph id="jp-invoice-buyer">{lines.buyer}</Paragraph>}
    <Paragraph id="jp-invoice-date">取引年月日：{lines.date}</Paragraph>
    {lines.registration && <Paragraph id="jp-invoice-registration">登録番号：{lines.registration}</Paragraph>}
    {lines.items[0].description && (
      <Paragraph id="jp-invoice-line-1">
        {lines.items[0].description}　{lines.items[0].amount}
      </Paragraph>
    )}
    {lines.items[1].description && (
      <Paragraph id="jp-invoice-line-2">
        {lines.items[1].description}　{lines.items[1].amount}
      </Paragraph>
    )}
    {lines.items[2].description && (
      <Paragraph id="jp-invoice-line-3">
        {lines.items[2].description}　{lines.items[2].amount}
      </Paragraph>
    )}
    {lines.items[3].description && (
      <Paragraph id="jp-invoice-line-4">
        {lines.items[3].description}　{lines.items[3].amount}
      </Paragraph>
    )}
    {lines.items[4].description && (
      <Paragraph id="jp-invoice-line-5">
        {lines.items[4].description}　{lines.items[4].amount}
      </Paragraph>
    )}
    {lines.items[5].description && (
      <Paragraph id="jp-invoice-line-6">
        {lines.items[5].description}　{lines.items[5].amount}
      </Paragraph>
    )}
    {lines.items[6].description && (
      <Paragraph id="jp-invoice-line-7">
        {lines.items[6].description}　{lines.items[6].amount}
      </Paragraph>
    )}
    {lines.items[7].description && (
      <Paragraph id="jp-invoice-line-8">
        {lines.items[7].description}　{lines.items[7].amount}
      </Paragraph>
    )}
    {lines.items[8].description && (
      <Paragraph id="jp-invoice-line-9">
        {lines.items[8].description}　{lines.items[8].amount}
      </Paragraph>
    )}
    {lines.items[9].description && (
      <Paragraph id="jp-invoice-line-10">
        {lines.items[9].description}　{lines.items[9].amount}
      </Paragraph>
    )}
    {lines.standard.total && (
      <Paragraph id="jp-invoice-standard-total">
        {lines.standard.label}　{lines.standard.total}
      </Paragraph>
    )}
    {lines.standard.tax && (
      <Paragraph id="jp-invoice-standard-tax">
        {lines.standard.taxLabel}　{lines.standard.tax}
      </Paragraph>
    )}
    {lines.reduced.total && (
      <Paragraph id="jp-invoice-reduced-total">
        {lines.reduced.label}　{lines.reduced.total}
      </Paragraph>
    )}
    {lines.reduced.tax && (
      <Paragraph id="jp-invoice-reduced-tax">
        {lines.reduced.taxLabel}　{lines.reduced.tax}
      </Paragraph>
    )}
    <Paragraph id="jp-invoice-total">
      {lines.totalLabel}　{lines.total}
    </Paragraph>
    {lines.legend && <Paragraph id="jp-invoice-legend">{lines.legend}</Paragraph>}
    {lines.note && <Paragraph id="jp-invoice-note">{lines.note}</Paragraph>}
  </>
);
