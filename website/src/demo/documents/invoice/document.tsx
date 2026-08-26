import { Document, PageBreak, Section, template } from "docxcelerate/template";
import {
  Charges,
  Closer,
  EngagementSummary,
  InvoiceMeta,
  Letterhead,
  Parties,
  Payment,
  PaymentLetterhead,
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
 *
 * Page one carries no running header: the letterhead already is the top of the
 * page, and printing both names the sender twice. Page two needs one, because a
 * payment page that does not say which invoice it belongs to gets filed against
 * the wrong account — so the header runs everywhere except the first page.
 */
export const documentTemplate = template<InvoiceData>(
  <Document
    id="invoice"
    title="Invoice"
    header={<RunningHeader />}
    firstHeader={false}
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
    <Closer />

    <PageBreak id="to-payment" />

    <PaymentLetterhead />
    <Payment />
    <Section id="scan" title="Scan to pay">
      <ScanToPay />
    </Section>
    <Terms />
  </Document>,
);
