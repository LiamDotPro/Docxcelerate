/** @jsxImportSource docxcelerate/template */
import { Section } from "docxcelerate/template";
import { PriceChange } from "../paragraph/conditional.node.tsx";
import { Greeting } from "../paragraph/static.node.tsx";

/**
 * A section takes an id, a title and its children. The children are the same
 * components you would place at the top level.
 */
export const Opening: Section = () => (
  <Section id="opening" title="Your renewal">
    <Greeting />
    <PriceChange />
  </Section>
);
