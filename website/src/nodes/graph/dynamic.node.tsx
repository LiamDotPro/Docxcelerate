/** @jsxImportSource docxcelerate/template */
import { Graph, useSetPlaceholders, useSetPrompts, useState } from "docxcelerate/template";
import type { SampleData } from "../sample-data.ts";

/**
 * For figures that need deriving rather than reading. `graphType` still fixes
 * the form locally, so the layout is known before the numbers are.
 */
export const PeakTimes: Graph = () => {
  const [state] = useState((data: SampleData) => ({
    name: data.memberName,
    ref: data.membershipRef,
    plan: data.plan,
  }));

  useSetPrompts({
    generalPrompt: `Plot when ${state.name} (${state.ref}) visits, bucketed by ` +
      `hour of the day across the week.`,
    infoPrompt: `Their plan is ${state.plan}, which allows entry at any hour.`,
  });

  useSetPlaceholders("Your busiest hours, Monday to Sunday");

  return <Graph id="peak-times" graphType="bar" />;
};
