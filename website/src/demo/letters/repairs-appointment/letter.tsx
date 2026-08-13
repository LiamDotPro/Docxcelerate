/** @jsxImportSource docxcelerate/template */
import { Document, Section, template } from "docxcelerate/template";
import { Access, Appointment, Greeting, WhatToExpect } from "./nodes/index.ts";
import type { RepairsData } from "./types.ts";

export const letterTemplate = template<RepairsData>(
  <Document id="repairs-appointment" title="Your repair appointment">
    <Section id="appointment" title="Your appointment">
      <Greeting />
      <Appointment />
    </Section>
    <Section id="on-the-day" title="On the day">
      <WhatToExpect />
      <Access />
    </Section>
  </Document>,
);
