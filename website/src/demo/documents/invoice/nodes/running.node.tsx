import {
  Cell,
  Image,
  type Nodes,
  PageNumber,
  Paragraph,
  Row,
  Table,
  useState,
} from "docxcelerate/template";
import { markPng, markSvg } from "../assets.ts";
import type { InvoiceData } from "../types.ts";

/**
 * The strip at the top of every page: who sent this, and which invoice.
 *
 * Running furniture rather than the first thing in the body, because page two
 * needs it as much as page one — a payment page that does not say which
 * invoice it belongs to is a page that gets filed against the wrong account.
 */
export const RunningHeader: Nodes = () => {
  const [state] = useState((data: InvoiceData) => ({
    name: data.sender.name,
    reference: data.reference,
  }));

  return (
    <>
      <Table id="running-head" columns={[{ width: "auto" }, { width: 60, align: "right" }]}>
        <Row>
          <Cell id="running-sender">{state.name}</Cell>
          <Cell id="running-reference">{state.reference}</Cell>
        </Row>
      </Table>
      <Paragraph id="head-rule" variant="rule" />
    </>
  );
};

/**
 * The dark strip at the foot of every page.
 *
 * The strip is the table's, not each cell's: a bar is one thing crossing the
 * page, and three cells that each paint themselves navy is three boxes that
 * happen to touch.
 *
 * The credit is the design's second decision, and it is written as an ordinary
 * `&&`: whether a given sender's invoice carries it is theirs to set, so the
 * build compiles it into a condition rather than settling it once for everyone.
 */
export const RunningFooter: Nodes = () => {
  const [state] = useState((data: InvoiceData) => ({
    registration: data.sender.registration,
    showCredit: data.showCredit,
  }));

  return (
    <Table
      id="running-foot"
      variant="footerBar"
      columns={[{ width: "auto" }, { width: 52 }, { width: 29, align: "right" }]}
    >
      <Row>
        <Cell id="foot-registration">{state.registration}</Cell>
        <Cell id="foot-credit">
          {state.showCredit && (
            <Paragraph id="credit-line">
              <Image
                id="credit-mark"
                src={markSvg}
                fallbackSrc={markPng}
                alt=""
                width={8}
                height={8}
              />
              {" Generated with Docxcelerate"}
            </Paragraph>
          )}
        </Cell>
        <Cell id="foot-page" variant="footerEdge">
          <PageNumber id="foot-page-number" />
        </Cell>
      </Row>
    </Table>
  );
};
