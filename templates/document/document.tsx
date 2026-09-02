import { Document, Section, template } from "docxcelerate/template";
import * as Nodes from "./nodes/index.ts";
import type { DocumentData } from "./types.ts";

export const documentTemplate = template<DocumentData>(
  <Document id="__DOCUMENT_ID__" title="__DOCUMENT_TITLE__">
    <Section id="opening" title="Opening">
      <Nodes.Greeting />
      <Nodes.Intro />
    </Section>
  </Document>,
);
