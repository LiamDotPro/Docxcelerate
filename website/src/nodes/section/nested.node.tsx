import { Section } from "docxcelerate/template";
import { ClassMix } from "../graph/pie.node.tsx";
import { VisitsByMonth } from "../graph/bar.node.tsx";
import { NextSteps } from "../paragraph/dynamic.node.tsx";

/**
 * Children can be of any kind, including another section — the one place a
 * document tree gains depth. The resolved document nests exactly as this reads.
 */
export const YourYear: Section = () => (
  <Section id="your-year" title="Your year here">
    <VisitsByMonth />
    <Section id="activity-mix" title="Where the time went">
      <ClassMix />
      <NextSteps />
    </Section>
  </Section>
);
