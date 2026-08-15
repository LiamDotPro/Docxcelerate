import { Graph, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * Shares of a whole. `caption` is optional, but a chart in a letter is read
 * once and not returned to.
 */
export const ClassMix: Graph = () => {
  const [state] = useState((data: SampleData) => ({
    labels: data.classMix.map((entry) => entry.label),
    shares: data.classMix.map((entry) => entry.share),
  }));

  return (
    <Graph
      id="class-mix"
      graphType="pie"
      data={{
        labels: state.labels,
        series: [{ name: "Share of visits", values: state.shares }],
      }}
      caption="How you used the centre, by activity"
    />
  );
};
