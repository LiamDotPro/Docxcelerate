import { Graph, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * Only `graphType` changes between the three static forms. Deriving the
 * running total inside the state initializer keeps the chart and the prose
 * from disagreeing, and keeps the work out of the returned element.
 */
export const CumulativeVisits: Graph = () => {
  const [state] = useState((data: SampleData) => {
    let running = 0;

    return {
      months: data.visitsByMonth.map((entry) => entry.month),
      toDate: data.visitsByMonth.map((entry) => (running += entry.visits)),
    };
  });

  return (
    <Graph
      id="cumulative-visits"
      title="Visits to date"
      graphType="line"
      data={{
        categories: state.months,
        series: [{ label: "Visits to date", values: state.toDate }],
      }}
      caption="Visits accumulated across the membership year"
    />
  );
};
