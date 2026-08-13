import type { NodeComponent } from "docxcelerate";
import type { SampleData } from "../sample-data.ts";

/**
 * No helper ships for this kind yet. A node component is only a function
 * returning a definition, so writing it out reaches the same place.
 *
 * The shipped renderers print the title and stop.
 */
export const Contents: NodeComponent<SampleData> = () => ({
  kind: "tableOfContents",
  id: "contents",
  title: "What is in this letter",
});
