import { Document, Section, template } from "docxcelerate/template";
import { Greeting, PremiumChange, Renewal, Shopping } from "./nodes/index.ts";
import type { PolicyData } from "./types.ts";

export const documentTemplate = template<PolicyData>(
  <Document id="policy-renewal" title="Your renewal">
    <Section id="renewal" title="Your renewal">
      <Greeting />
      <Renewal />
      <PremiumChange />
    </Section>
    <Section id="your-options" title="Your options">
      <Shopping />
    </Section>
  </Document>,
);
