import { Graph, useState } from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const __NODE_COMPONENT__: Graph = () => {
  const [state] = useState((data: DocumentData) => ({
    name: data.recipientName,
  }));

  return (
    <Graph
      id="__NODE_ID__"
      title="__NODE_TITLE__"
      graphType="bar"
      data={{
        categories: ["Current"],
        series: [{ label: state.name, values: [1] }],
      }}
      caption={`__NODE_TITLE__ for ${state.name}.`}
    />
  );
};
