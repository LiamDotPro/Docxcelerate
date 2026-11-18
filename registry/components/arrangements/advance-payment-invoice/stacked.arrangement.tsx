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
    {lines.previous[0].number && (
      <Paragraph id="advance-payment-previous-1">
        Previously requested ({lines.previous[0].number}): {lines.previous[0].amount}
      </Paragraph>
    )}
    {lines.previous[1].number && (
      <Paragraph id="advance-payment-previous-2">
        Previously requested ({lines.previous[1].number}): {lines.previous[1].amount}
      </Paragraph>
    )}
    {lines.previous[2].number && (
      <Paragraph id="advance-payment-previous-3">
        Previously requested ({lines.previous[2].number}): {lines.previous[2].amount}
      </Paragraph>
    )}
    {lines.previous[3].number && (
      <Paragraph id="advance-payment-previous-4">
        Previously requested ({lines.previous[3].number}): {lines.previous[3].amount}
      </Paragraph>
    )}
    {lines.previous[4].number && (
      <Paragraph id="advance-payment-previous-5">
        Previously requested ({lines.previous[4].number}): {lines.previous[4].amount}
      </Paragraph>
    )}
    {lines.previous[5].number && (
      <Paragraph id="advance-payment-previous-6">
        Previously requested ({lines.previous[5].number}): {lines.previous[5].amount}
      </Paragraph>
    )}
    {lines.previous[6].number && (
      <Paragraph id="advance-payment-previous-7">
        Previously requested ({lines.previous[6].number}): {lines.previous[6].amount}
      </Paragraph>
    )}
    {lines.previous[7].number && (
      <Paragraph id="advance-payment-previous-8">
        Previously requested ({lines.previous[7].number}): {lines.previous[7].amount}
      </Paragraph>
    )}
    <Paragraph id="advance-payment-deposit">{lines.depositLabel}: {lines.deposit}</Paragraph>
    <Paragraph id="advance-payment-left">Left to invoice: {lines.leftToInvoice}</Paragraph>
    {lines.paymentDate && (
      <Paragraph id="advance-payment-payment-date">{lines.paymentDateLabel}: {lines.paymentDate}</Paragraph>
    )}
    {lines.vatApplies && (
      <>
        {lines.bands[0].label && (
          <Paragraph id="advance-payment-vat-band-1">
            VAT at {lines.bands[0].label}: {lines.bands[0].vat} on a net {lines.bands[0].net} (gross{" "}
            {lines.bands[0].gross}).
          </Paragraph>
        )}
        {lines.bands[1].label && (
          <Paragraph id="advance-payment-vat-band-2">
            VAT at {lines.bands[1].label}: {lines.bands[1].vat} on a net {lines.bands[1].net} (gross{" "}
            {lines.bands[1].gross}).
          </Paragraph>
        )}
        {lines.bands[2].label && (
          <Paragraph id="advance-payment-vat-band-3">
            VAT at {lines.bands[2].label}: {lines.bands[2].vat} on a net {lines.bands[2].net} (gross{" "}
            {lines.bands[2].gross}).
          </Paragraph>
        )}
        {lines.bands[3].label && (
          <Paragraph id="advance-payment-vat-band-4">
            VAT at {lines.bands[3].label}: {lines.bands[3].vat} on a net {lines.bands[3].net} (gross{" "}
            {lines.bands[3].gross}).
          </Paragraph>
        )}
        <Paragraph id="advance-payment-vat-total">
          VAT on this deposit: {lines.vatTotal} on a net {lines.netTotal}.
        </Paragraph>
      </>
    )}
  </>
);
