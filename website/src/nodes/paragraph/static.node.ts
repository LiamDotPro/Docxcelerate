import { paragraph } from "docxcelerate";
import type { SampleData } from "../sample-data.ts";

/** The smallest useful node: an id, and a render that turns data into a line. */
export const Greeting = paragraph<SampleData>({
  id: "greeting",
  render: (data) => `Dear ${data.memberName},`,
});
