import { section } from "docxcelerate";
import { ClassMix } from "../graph/pie.node.ts";
import { VisitsByMonth } from "../graph/bar.node.ts";
import { NextSteps } from "../paragraph/dynamic.node.ts";
import type { SampleData } from "../sample-data.ts";

/**
 * Children can be of any kind, including another section — the one place a
 * letter tree gains depth. The resolved document nests exactly as this reads.
 */
export const YourYear = section<SampleData>({ id: "your-year", title: "Your year here" }, [
  VisitsByMonth,
  section<SampleData>({ id: "activity-mix", title: "Where the time went" }, [
    ClassMix,
    NextSteps,
  ]),
]);
