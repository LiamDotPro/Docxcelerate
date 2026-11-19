import { type JsonObject, type Paragraph as ParagraphComponent, Paragraph, useAi } from "docxcelerate/template";

/** The adviser's voice, for every paragraph the engine writes in this review. */
export const VOICE =
  "You are Tomas Lindqvist, a chartered financial planner, writing to a client you know. " +
  "Plain British English, short sentences, no jargon, no sales language. Never forecast markets.";

/**
 * A paragraph the engine writes for this client, from the facts it is handed.
 *
 * One component for all of them because they differ only in what they ask and
 * what they are given: the voice and the refusal to forecast are the review's,
 * not any one paragraph's.
 */
export const Written: ParagraphComponent<{
  id: string;
  variant?: string;
  ask: string;
  placeholder: string;
  from?: JsonObject;
  /** Numbered, when what is written is a list of points, one per paragraph. */
  list?: "number";
}> = ({ id, variant, ask, placeholder, from, list }) => {
  useAi({
    ask,
    placeholder,
    voice: VOICE,
    from,
    avoid: "Do not forecast returns, give figures you were not handed, or promise outcomes.",
  });
  // Numbered points are numbered as the canvas numbers them: the bare figure.
  return <Paragraph id={id} variant={variant} list={list} listStyle={list === "number" ? "decimalPlain" : undefined} />;
};
