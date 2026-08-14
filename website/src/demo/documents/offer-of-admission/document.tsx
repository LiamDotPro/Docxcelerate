/** @jsxImportSource docxcelerate/template */
import { Document, Section, template } from "docxcelerate/template";
import {
  Conditions,
  Fees,
  Greeting,
  NextSteps,
  Offer,
  SignOff,
  TutorNote,
} from "./nodes/index.ts";
import type { OfferData } from "./types.ts";

export const documentTemplate = template<OfferData>(
  <Document id="offer-of-admission" title="Offer of Admission">
    <Section id="your-offer" title="Your offer">
      <Greeting />
      <Offer />
      <TutorNote />
    </Section>
    <Section id="conditions" title="Conditions">
      <Conditions />
    </Section>
    <Section id="fees-and-funding" title="Fees and funding">
      <Fees />
    </Section>
    <Section id="next-steps" title="Next steps">
      <NextSteps />
      <SignOff />
    </Section>
  </Document>,
);
