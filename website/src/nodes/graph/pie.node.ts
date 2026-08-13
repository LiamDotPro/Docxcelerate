import { graph } from "docxcelerate";
import type { SampleData } from "../sample-data.ts";

/**
 * Shares of a whole. `caption` is optional, but a chart in a letter is read
 * once and not returned to.
 */
export const ClassMix = graph<SampleData>({
  id: "class-mix",
  graphType: "pie",
  data: (data) => ({
    labels: data.classMix.map((entry) => entry.label),
    series: [{ name: "Share of visits", values: data.classMix.map((entry) => entry.share) }],
  }),
  caption: () => "How you used the centre, by activity",
});
