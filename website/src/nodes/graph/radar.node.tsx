import { Graph, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * Two shapes held against one another. A radar answers "strong where", not
 * "how much" — a reader cannot compare the areas of two rings, and is not
 * meant to. The axis is scored out of 100 so the spokes share a scale;
 * plotting raw counts of different things on one ring compares nothing.
 */
export const FacilityMix: Graph = () => {
  const [state] = useState((data: SampleData) => ({
    facilities: data.facilityUse.map((entry) => entry.facility),
    you: data.facilityUse.map((entry) => entry.you),
    average: data.facilityUse.map((entry) => entry.average),
  }));

  return (
    <Graph
      id="facility-mix"
      title="What you used, against the average"
      graphType="radar"
      data={{
        categories: state.facilities,
        series: [
          { label: "You", values: state.you },
          { label: "Centre average", values: state.average },
        ],
      }}
      caption="Use of each facility, scored out of 100"
    />
  );
};
