import { Document, PageBreak, template } from "docxcelerate/template";
import { Activity } from "./nodes/activity.node.tsx";
import { Allocation } from "./nodes/allocation.node.tsx";
import { Attribution } from "./nodes/attribution.node.tsx";
import { Costs } from "./nodes/costs.node.tsx";
import { Cover, CoverFooter } from "./nodes/cover.node.tsx";
import { Focus } from "./nodes/focus.node.tsx";
import { Holdings } from "./nodes/holdings.node.tsx";
import { Income } from "./nodes/income.node.tsx";
import { Letter } from "./nodes/letter.node.tsx";
import { Performance } from "./nodes/performance.node.tsx";
import { Plan } from "./nodes/plan.node.tsx";
import { RunningFooter, RunningHeader } from "./nodes/running.node.tsx";
import { Summary } from "./nodes/summary.node.tsx";
import type { PortfolioReviewData } from "./types.ts";

/**
 * Structure only: thirteen pages, each opening where the design opens it.
 *
 * Every section starts a page of its own. A review is read a section at a time
 * and photocopied a page at a time, and a section that began at the foot of
 * the page before is the one that gets left on the copier.
 *
 * The cover carries no running band — the photograph is the top of that page —
 * so the band runs on every page but the first, and the cover's foot names
 * who prepared the review.
 */
export const documentTemplate = template<PortfolioReviewData>(
  <Document
    id="portfolio-review"
    title="Quarterly portfolio review"
    header={<RunningHeader />}
    firstHeader={false}
    footer={<RunningFooter />}
    firstFooter={<CoverFooter />}
  >
    <Cover />
    <PageBreak id="to-letter" />
    <Letter />
    <PageBreak id="to-summary" />
    <Summary />
    <PageBreak id="to-performance" />
    <Performance />
    <PageBreak id="to-attribution" />
    <Attribution />
    <PageBreak id="to-allocation" />
    <Allocation />
    <PageBreak id="to-holdings" />
    <Holdings />
    <PageBreak id="to-focus" />
    <Focus />
    <PageBreak id="to-income" />
    <Income />
    <PageBreak id="to-activity" />
    <Activity />
    <PageBreak id="to-costs" />
    <Costs />
    <PageBreak id="to-plan" />
    <Plan />
  </Document>,
);
