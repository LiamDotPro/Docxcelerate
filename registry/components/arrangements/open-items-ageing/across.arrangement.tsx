import { Cell, Row, Table, type Yield } from "docxcelerate/template";
import type { OpenItemsAgeingLines } from "./lines.ts";

/**
 * The buckets side by side in three rows (name, amount, share), the way a
 * statement usually prints them. Six columns of seven-figure amounts only just
 * fit a text column; `down` exists for when they do not.
 */
export const across = (lines: OpenItemsAgeingLines): Yield => (
  <>
    {lines.overdue && (
      <Table
        id="open-items-ageing-strip"
        columns={[
          { width: "auto", align: "right" },
          { width: "auto", align: "right" },
          { width: "auto", align: "right" },
          { width: "auto", align: "right" },
          { width: "auto", align: "right" },
          { width: "auto", align: "right" },
        ]}
      >
        <Row id="open-items-ageing-strip-labels">
          <Cell>{lines.buckets[0].label}</Cell>
          <Cell>{lines.buckets[1].label}</Cell>
          <Cell>{lines.buckets[2].label}</Cell>
          <Cell>{lines.buckets[3].label}</Cell>
          <Cell>{lines.buckets[4].label}</Cell>
          <Cell>{lines.buckets[5].label}</Cell>
        </Row>
        <Row id="open-items-ageing-strip-amounts">
          <Cell>{lines.buckets[0].amount}</Cell>
          <Cell>{lines.buckets[1].amount}</Cell>
          <Cell>{lines.buckets[2].amount}</Cell>
          <Cell>{lines.buckets[3].amount}</Cell>
          <Cell>{lines.buckets[4].amount}</Cell>
          <Cell>{lines.buckets[5].amount}</Cell>
        </Row>
        <Row id="open-items-ageing-strip-shares">
          <Cell>{lines.buckets[0].share}</Cell>
          <Cell>{lines.buckets[1].share}</Cell>
          <Cell>{lines.buckets[2].share}</Cell>
          <Cell>{lines.buckets[3].share}</Cell>
          <Cell>{lines.buckets[4].share}</Cell>
          <Cell>{lines.buckets[5].share}</Cell>
        </Row>
      </Table>
    )}
  </>
);
