import { Graph, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * `stacked="percent"` is the 100% stacked chart, and it is a different reading
 * from `stacked`: every column is drawn to the same height, so what is
 * compared is the mix within each month rather than the size of the months.
 *
 * The raw counts are what the file carries. Word works the shares out itself
 * and labels the axis 0% to 100%, so nothing here divides — pre-dividing would
 * hand Word shares of shares.
 */
export const VisitMix: Graph = () => {
  const [state] = useState((data: SampleData) => ({
    months: data.visitMix.map((entry) => entry.month),
    swim: data.visitMix.map((entry) => entry.swim),
    strength: data.visitMix.map((entry) => entry.strength),
    classes: data.visitMix.map((entry) => entry.classes),
  }));

  return (
    <Graph
      id="visit-mix"
      title="What you came for, month by month"
      graphType="bar"
      stacked="percent"
      data={{
        categories: state.months,
        series: [
          { label: "Swim", values: state.swim },
          { label: "Strength", values: state.strength },
          { label: "Classes", values: state.classes },
        ],
      }}
      caption="Share of each month's visits, by activity"
    />
  );
};
