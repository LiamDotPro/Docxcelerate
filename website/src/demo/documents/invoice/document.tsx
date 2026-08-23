import { Document, PageBreak, Section, template } from "docxcelerate/template";
import {
  Charges,
  EngagementSummary,
  InvoiceMeta,
  Letterhead,
  Parties,
  Payment,
  RunningFooter,
  RunningHeader,
  ScanToPay,
  Terms,
  Totals,
} from "./nodes/index.ts";
import type { InvoiceData } from "./types.ts";

/**
 * Structure only: which nodes, in which order, and where the page turns.
 *
 * The break is part of what this document is, not a way of nudging a paragraph
 * off the bottom of a page. What is owed goes on one page and how to pay it on
 * the next, so that either can be handed to someone on its own.
 */
export const documentTemplate = template<InvoiceData>(
  <Document
    id="invoice"
    title="Invoice"
    header={<RunningHeader />}
    footer={<RunningFooter />}
  >
    <Letterhead />
    <InvoiceMeta />
    <Parties />
    <Section id="summary" title="Engagement summary">
      <EngagementSummary />
    </Section>
    <Charges />
    <Totals />

    <PageBreak id="to-payment" />

    <Payment />
    <Section id="scan" title="Scan to pay">
      <ScanToPay />
    </Section>
    <Terms />
  </Document>,
);
