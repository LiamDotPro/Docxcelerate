import { Paragraph, Text, type Yield } from "docxcelerate/template";
import type { NetPayableWithholdingsLines } from "./lines.ts";

/** One `Label: figure` paragraph per row, for a narrow column. */
export const stacked = (lines: NetPayableWithholdingsLines): Yield => (
  <>
    <Paragraph id="net-payable-base">{lines.baseLabel}: {lines.base}</Paragraph>
    {lines.rivalsaLabel && (
    <Paragraph id="net-payable-rivalsa">{lines.rivalsaLabel}: {lines.rivalsa}</Paragraph>
    )}
    {lines.cassaLabel && (
    <Paragraph id="net-payable-cassa">{lines.cassaLabel}: {lines.cassa}</Paragraph>
    )}
    {lines.vatLabel && (
    <Paragraph id="net-payable-vat">{lines.vatLabel}: {lines.vat}</Paragraph>
    )}
    {lines.reLabel && (
    <Paragraph id="net-payable-re">{lines.reLabel}: {lines.re}</Paragraph>
    )}
    {lines.exemptLabel && (
    <Paragraph id="net-payable-exempt">{lines.exemptLabel}: {lines.exempt}</Paragraph>
    )}
    {lines.stampLabel && (
    <Paragraph id="net-payable-stamp">{lines.stampLabel}: {lines.stamp}</Paragraph>
    )}
    <Paragraph id="net-payable-total">{lines.totalLabel}: {lines.total}</Paragraph>
    {lines.withholdingLabel && (
    <Paragraph id="net-payable-withholding">{lines.withholdingLabel}: {lines.withholding}</Paragraph>
    )}
    {lines.splitPaymentLabel && (
    <Paragraph id="net-payable-split-payment">{lines.splitPaymentLabel}: {lines.splitPayment}</Paragraph>
    )}
    <Paragraph id="net-payable-net">{lines.netLabel}: <Text bold>{lines.net}</Text></Paragraph>
  </>
);
