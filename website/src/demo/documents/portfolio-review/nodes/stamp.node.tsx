import { type JsonObject, type Nodes, Paragraph } from "docxcelerate/template";
import { Written } from "./written.node.tsx";

/**
 * What a decision stamp says: the decision in a word, and a line under it.
 *
 * Both are the engine's. The word is the decision itself, and the line says
 * what kind of decision it was — "recommended this quarter", "for retirement
 * in 2036" — which depends on the word. Placed in a cell drawn as `stamp`,
 * whose look in a proof is the outlined, unfilled one the canvas gives a
 * decision nobody has taken yet.
 */
export const StampWords: Nodes<{
  id: string;
  ask: string;
  placeholder: string;
  noteAsk: string;
  notePlaceholder: string;
  from: JsonObject;
}> = ({ id, ask, placeholder, noteAsk, notePlaceholder, from }) => (
  <>
    <Paragraph variant="stampLabel">Decision</Paragraph>
    <Written id={id} variant="stampWord" ask={ask} placeholder={placeholder} from={from} />
    <Written id={`${id}-note`} variant="stampNote" ask={noteAsk} placeholder={notePlaceholder} from={from} />
  </>
);
