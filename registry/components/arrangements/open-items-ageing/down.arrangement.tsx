import { Cell, Row, Table, type Yield } from "docxcelerate/template";
import type { OpenItemsAgeingLines } from "./lines.ts";

/**
 * One row per bucket with its amount and share beside it, for a narrow column
 * or seven-figure amounts.
 */
export const down = (lines: OpenItemsAgeingLines): Yield => (
  <>
    {lines.overdue && (
      <Table
        id="open-items-ageing-strip"
        columns={[{ width: "auto" }, { width: 35, align: "right" }, { width: 18, align: "right" }]}
      >
        <Row id="open-items-ageing-strip-header">
          <Cell>Days overdue</Cell>
          <Cell>Amount</Cell>
          <Cell>Share</Cell>
        </Row>
        <Row id="open-items-ageing-strip-not-due">
          <Cell>{lines.buckets[0].label}</Cell>
          <Cell>{lines.buckets[0].amount}</Cell>
          <Cell>{lines.buckets[0].share}</Cell>
        </Row>
        <Row id="open-items-ageing-strip-1-30">
          <Cell>{lines.buckets[1].label}</Cell>
          <Cell>{lines.buckets[1].amount}</Cell>
          <Cell>{lines.buckets[1].share}</Cell>
        </Row>
        <Row id="open-items-ageing-strip-31-60">
          <Cell>{lines.buckets[2].label}</Cell>
          <Cell>{lines.buckets[2].amount}</Cell>
          <Cell>{lines.buckets[2].share}</Cell>
        </Row>
        <Row id="open-items-ageing-strip-61-90">
          <Cell>{lines.buckets[3].label}</Cell>
          <Cell>{lines.buckets[3].amount}</Cell>
          <Cell>{lines.buckets[3].share}</Cell>
        </Row>
        <Row id="open-items-ageing-strip-over-90">
          <Cell>{lines.buckets[4].label}</Cell>
          <Cell>{lines.buckets[4].amount}</Cell>
          <Cell>{lines.buckets[4].share}</Cell>
        </Row>
        <Row id="open-items-ageing-strip-disputed">
          <Cell>{lines.buckets[5].label}</Cell>
          <Cell>{lines.buckets[5].amount}</Cell>
          <Cell>{lines.buckets[5].share}</Cell>
        </Row>
      </Table>
    )}
  </>
);
