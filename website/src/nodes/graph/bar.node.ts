import { graph } from "docxcelerate";
import type { SampleData } from "../sample-data.ts";

/**
 * A chart is declared, not drawn: `graphType` fixes the form, `data` returns
 * the payload. One node then serves the browser preview and the packed DOCX.
 */
export const VisitsByMonth = graph<SampleData>({
  id: "visits-by-month",
  graphType: "bar",
  data: (data) => ({
    labels: data.visitsByMonth.map((entry) => entry.month),
    series: [{ name: "Visits", values: data.visitsByMonth.map((entry) => entry.visits) }],
  }),
  caption: (data) => `Your visits to ${data.centreName}, last six months`,
});
