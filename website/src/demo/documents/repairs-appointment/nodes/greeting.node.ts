import { paragraph } from "docxcelerate";
import type { RepairsData } from "../types.ts";

export const Greeting = paragraph<RepairsData>({
  id: "greeting",
  render: (data) => `Dear ${data.residentName},`,
});
