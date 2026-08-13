import { graph } from "docxcelerate";
import type { SampleData } from "../sample-data.ts";

/**
 * For figures that need deriving rather than reading. `graphType` still fixes
 * the form locally, so the layout is known before the numbers are.
 */
export const PeakTimes = graph<SampleData>({
  id: "peak-times",
  graphType: "bar",
  placeholder: () => "Your busiest hours, Monday to Sunday",
  generalPrompt: (data) =>
    `Plot when ${data.memberName} (${data.membershipRef}) visits, bucketed by ` +
    `hour of the day across the week.`,
  infoPrompt: (data) => `Their plan is ${data.plan}, which allows entry at any hour.`,
});
