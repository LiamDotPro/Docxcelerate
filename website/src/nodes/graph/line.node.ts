import { graph } from "docxcelerate";
import type { SampleData } from "../sample-data.ts";

/**
 * Only `graphType` changes between the three static forms. Deriving the
 * running total here keeps the chart and the prose from disagreeing.
 */
export const CumulativeVisits = graph<SampleData>({
  id: "cumulative-visits",
  graphType: "line",
  data: (data) => {
    let running = 0;

    return {
      labels: data.visitsByMonth.map((entry) => entry.month),
      series: [
        {
          name: "Visits to date",
          values: data.visitsByMonth.map((entry) => (running += entry.visits)),
        },
      ],
    };
  },
  caption: () => "Visits accumulated across the membership year",
});
