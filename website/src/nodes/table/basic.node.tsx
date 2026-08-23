import { Cell, Row, Table, useFormat, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * The columns are declared once, on the table, because every row shares them.
 * A row is an ordinary node, so a `.map()` produces one per entry and the
 * table needs to know nothing about loops.
 */
export const VisitLog: Table = () => {
  const { number } = useFormat("en-GB");
  const [state] = useState((data: SampleData) => ({ months: data.visitsByMonth }));

  return (
    <Table id="visit-log" columns={[{ width: "auto" }, { width: 28, align: "right" }]}>
      <Row header>
        <Cell>Month</Cell>
        <Cell>Visits</Cell>
      </Row>
      {state.months.map((entry) => (
        <Row>
          <Cell>{entry.month}</Cell>
          <Cell>{number(entry.visits)}</Cell>
        </Row>
      ))}
    </Table>
  );
};
