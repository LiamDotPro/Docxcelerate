import { Cell, Paragraph, Row, Table, type Yield } from "docxcelerate/template";
import type { RenewalAndPriceChangeNoticeLines } from "./lines.ts";

/**
 * Labels in a narrow column with the values beside them: the key-facts box
 * many reminders use. The consequence and warning stay full width below it,
 * because they are sentences, not facts.
 */
export const table = (lines: RenewalAndPriceChangeNoticeLines): Yield => (
  <>
    <Paragraph id="renewal-notice-statement">{lines.statement}</Paragraph>
    {/* 50mm holds the longest label, "Introductory period ends", on one line. */}
    <Table id="renewal-notice-details" columns={[{ width: 50 }, { width: "auto" }]}>
      <Row id="renewal-notice-renewal">
        <Cell>Renewal</Cell>
        <Cell>{lines.renewal}</Cell>
      </Row>
      {lines.trial && (
        <Row id="renewal-notice-trial">
          <Cell>Introductory period ends</Cell>
          <Cell>{lines.trial}</Cell>
        </Row>
      )}
      {lines.priceChange && (
        <Row id="renewal-notice-price-change">
          <Cell>Price change</Cell>
          <Cell>{lines.priceChange}</Cell>
        </Row>
      )}
      {lines.previousPayment && (
        <Row id="renewal-notice-previous-payment">
          <Cell>Previous payment</Cell>
          <Cell>{lines.previousPayment}</Cell>
        </Row>
      )}
      {lines.cancelBy && (
        <Row id="renewal-notice-cancel-by">
          <Cell>Cancellation deadline</Cell>
          <Cell>{lines.cancelBy}</Cell>
        </Row>
      )}
      <Row id="renewal-notice-how-to-cancel">
        <Cell>How to cancel</Cell>
        <Cell>{lines.howToCancel}</Cell>
      </Row>
    </Table>
    {lines.consequence && <Paragraph id="renewal-notice-consequence">{lines.consequence}</Paragraph>}
    {lines.warning && <Paragraph id="renewal-notice-warning">{lines.warning}</Paragraph>}
  </>
);
