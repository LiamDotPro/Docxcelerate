import { Paragraph, type Yield } from "docxcelerate/template";
import type { AdvancePaymentInvoiceLines } from "./lines.ts";

/**
 * One labelled line under another, down the text column. For a narrow column,
 * where four figures side by side would not fit.
 */
export const stacked = (lines: AdvancePaymentInvoiceLines): Yield => (
  <>
    <Paragraph id="advance-payment-order-reference">Order reference: {lines.orderReference}</Paragraph>
    {lines.orderDate && <Paragraph id="advance-payment-order-date">Order date: {lines.orderDate}</Paragraph>}
    <Paragraph id="advance-payment-order-total">Order total: {lines.orderTotal}</Paragraph>
    {lines.previous.map((p) => (
      <Paragraph id="advance-payment-previous">Previously requested ({p.number}): {p.amount}</Paragraph>
    ))}
    <Paragraph id="advance-payment-deposit">{lines.depositLabel}: {lines.deposit}</Paragraph>
    <Paragraph id="advance-payment-left">Left to invoice: {lines.leftToInvoice}</Paragraph>
    {lines.paymentDate && (
      <Paragraph id="advance-payment-payment-date">{lines.paymentDateLabel}: {lines.paymentDate}</Paragraph>
    )}
    {lines.vatApplies && (
      <>
        {lines.bands.map((band) => (
          <Paragraph id="advance-payment-vat-band">
            VAT at {band.label}: {band.vat} on a net {band.net} (gross {band.gross}).
          </Paragraph>
        ))}
        <Paragraph id="advance-payment-vat-total">
          VAT on this deposit: {lines.vatTotal} on a net {lines.netTotal}.
        </Paragraph>
      </>
    )}
  </>
);
