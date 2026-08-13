import { section } from "docxcelerate";
import { PriceChange } from "../paragraph/conditional.node.ts";
import { Greeting } from "../paragraph/static.node.ts";
import type { SampleData } from "../sample-data.ts";

/**
 * A section takes an id, a title and its children. The children are the same
 * node components you would place at the top level.
 */
export const Opening = section<SampleData>({ id: "opening", title: "Your renewal" }, [
  Greeting,
  PriceChange,
]);
