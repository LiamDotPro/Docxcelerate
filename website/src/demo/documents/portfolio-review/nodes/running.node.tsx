import { Cell, type Nodes, PageNumber, Paragraph, Row, Table, useState } from "docxcelerate/template";
import type { PortfolioReviewData } from "../types.ts";

/**
 * The dark band at the top of every page after the cover: the firm, and whose
 * review this is.
 *
 * A page pulled out of the stack still says which client and which quarter it
 * belongs to, which is the whole reason the band is furniture rather than the
 * first line of each section.
 */
export const RunningHeader: Nodes = () => {
  const [head] = useState((data: PortfolioReviewData) => ({
    firm: data.firm.name,
    meta: `${data.client.shortName} · ${data.client.account} · Quarterly portfolio review · ${data.period.short}`,
  }));

  return (
    <>
      <Table id="head-band" variant="headBand" columns={[{ width: 70 }, { width: "auto", align: "right" }]}>
        <Row>
          <Cell variant="headName">{head.firm.toUpperCase()}</Cell>
          <Cell variant="headMeta">{head.meta}</Cell>
        </Row>
      </Table>
      <Paragraph id="head-rule" variant="headRule" />
    </>
  );
};

/** The regulator's line and the page count, over a hairline, on every page. */
export const RunningFooter: Nodes = () => {
  const [foot] = useState((data: PortfolioReviewData) => ({
    line: `${data.firm.name} Ltd · ${data.firm.regulation}`,
  }));

  return (
    <Table id="foot" variant="foot" columns={[{ width: "auto" }, { width: 9, align: "right" }, { width: 13, align: "right" }]}>
      <Row>
        <Cell>{foot.line}</Cell>
        <Cell variant="footPage">PAGE</Cell>
        <Cell variant="footPage">
          <PageNumber id="foot-page-number" format="currentOfTotal" separator=" OF " />
        </Cell>
      </Row>
    </Table>
  );
};
