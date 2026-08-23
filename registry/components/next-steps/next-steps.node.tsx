import { type Paragraph as ParagraphComponent, Paragraph, useState } from "docxcelerate/template";

/**
 * The paragraph telling a reader what to do now — written per document.
 *
 * All four prompts are set, because this is the node most likely to invent
 * something. The info prompt hands the engine the facts it is allowed to use,
 * and the negative prompt closes off the two failures that matter in a document
 * somebody acts on: inventing a deadline, and inventing a way to contact you.
 *
 * Previews show the placeholder rather than calling anything, so the shape of
 * the finished document is readable before an engine exists.
 *
 * Installed by `dxcl add next-steps`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface NextStepsData {
  /** What the reader has to do, in your own words. One entry per step. */
  actions: string[];
  /** How to reach you. Printed nowhere; given to the engine as fact. */
  contact: string;
}

export const NextSteps: ParagraphComponent = () => {
  const [steps] = useState((data: NextStepsData) => ({
    actions: data.actions.join("; "),
    contact: data.contact,
  }));

  return (
    <Paragraph
      id="next-steps"
      systemPrompt="You are writing on behalf of the sender, plainly and without sales language."
      generalPrompt="Tell the reader what to do next, in one short paragraph of no more than three sentences."
      infoPrompt={`The steps are: ${steps.actions}. They can reach us at ${steps.contact}.`}
      negativePrompt="Do not invent deadlines, reference numbers, phone numbers or addresses. Do not apologise."
      placeholder="What the reader should do next, written per document from the steps above."
    />
  );
};
