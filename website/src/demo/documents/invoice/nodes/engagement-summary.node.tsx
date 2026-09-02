import {
  Paragraph,
  useSetPlaceholders,
  useSetPrompts,
  useState,
} from "docxcelerate/template";
import type { InvoiceData } from "../types.ts";

/**
 * The paragraph a human would otherwise write every month.
 *
 * The figures are already in the table below, so this is the one part of an
 * invoice that has to be composed rather than laid out — what the work was
 * for, in the client's terms. It is the engine's to write per recipient, which
 * is why the node carries prompts instead of text.
 *
 * The negative prompt is doing real work: an engine handed a table of charges
 * will restate the total unless told not to, and a total that disagrees with
 * the one printed below it is worse than no summary at all.
 */
export const EngagementSummary: Paragraph = () => {
  const [state] = useState((data: InvoiceData) => ({
    client: data.billedTo.name,
    lead: data.deliveryLead,
  }));

  useSetPrompts({
    systemPrompt:
      "You write the covering note on a consultancy invoice. Plain, specific, " +
      "and short. You are writing to a finance team who did not attend the work.",
    generalPrompt:
      `Write three sentences for ${state.client} summarising what this month's ` +
      "work delivered, leading with whatever the largest line paid for.",
    // The lines are named by pointing at them rather than by being pasted in.
    // Spelling them out here would mean walking the list while building, which
    // publishing refuses — there is no list until a request arrives — and the
    // engine is writing into a document that carries the table anyway.
    infoPrompt: "The work billed is itemised in the charges table below, each " +
      `line with its own note. The delivery lead is ${state.lead}.`,
    negativePrompt:
      "Do not restate any figure, total or rate — they are in the table below. " +
      "Do not thank the client, and do not mention the invoice itself.",
  });
  useSetPlaceholders(
    "Sprint 14 closed out the repairs booking endpoints and the account and " +
      "statements views in the tenant portal. Most of this invoice is the " +
      "eight-day API build; the document automation line covers the arrears " +
      "letter and statement templates now generating in production. September " +
      "continues under the support retainer.",
  );

  return <Paragraph id="engagement-summary" variant="summary" />;
};
