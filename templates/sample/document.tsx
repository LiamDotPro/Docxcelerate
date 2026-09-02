import { Document, Section, template } from "docxcelerate/template";
import * as Nodes from "./nodes/index.ts";
import type { DocumentData } from "./types.ts";

export const documentTemplate = template<DocumentData>(
  <Document id="welcome" title="Welcome">
    <Section id="opening" title="Opening">
      <Nodes.Greeting />
      <Nodes.BalanceSummary />
    </Section>
  </Document>,
);
