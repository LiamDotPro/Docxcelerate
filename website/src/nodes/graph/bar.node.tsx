import { Graph, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * A chart is declared, not drawn: `graphType` fixes the form, `data` carries
 * the payload. One node then serves the browser preview and the packed DOCX.
 */
export const VisitsByMonth: Graph = () => {
  const [state] = useState((data: SampleData) => ({
    centreName: data.centreName,
    months: data.visitsByMonth.map((entry) => entry.month),
    visits: data.visitsByMonth.map((entry) => entry.visits),
  }));

  return (
    <Graph
      id="visits-by-month"
      title={`Visits to ${state.centreName}`}
      graphType="bar"
      valueAxisTitle="Visits"
      data={{
        categories: state.months,
        series: [{ label: "Visits", values: state.visits }],
      }}
      caption={`Your visits to ${state.centreName}, last six months`}
    />
  );
};
