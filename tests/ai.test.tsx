import { test } from "node:test";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { buildDocument, EchoAiClient } from "docxcelerate";
import type { DocumentNode, ParagraphNode, PromptSpec } from "docxcelerate";
import { Document, Paragraph, Section, template, useAi, useState } from "docxcelerate/template";

/**
 * One hook for a node an engine writes.
 *
 * Underneath every field here is a single decision — this node is written per
 * document rather than at build time — and the point of collapsing the old
 * hooks into one call is that the decision looks like one decision. The
 * placeholder being required is the part that earns its keep: it is what a
 * reader sees in every preview and in every document whose generation was
 * skipped, and it was the easiest thing to forget when it lived in its own
 * hook.
 */
interface InvoiceData {
  reference: string;
  period: string;
  lines: Array<{ desc: string }>;
}

const invoice: InvoiceData = {
  reference: "INV-2026-0142",
  period: "August 2026",
  lines: [{ desc: "API build" }],
};

function buildWith(component: () => unknown, options = {}) {
  const Body = component as () => never;

  return buildDocument(
    template<InvoiceData>(
      <Document id="invoice" title="Invoice">
        <Section id="body" title="Body">
          <Body />
        </Section>
      </Document>,
    ),
    invoice,
    options,
  );
}

function paragraphOf(nodes: DocumentNode[]): ParagraphNode {
  const section = nodes[0];

  if (section?.kind !== "section" || section.children[0]?.kind !== "paragraph") {
    throw new Error("expected one paragraph in the body");
  }

  return section.children[0];
}

function promptOf(prompts: PromptSpec[] | undefined, kind: PromptSpec["kind"]): string | undefined {
  return prompts?.find((prompt) => prompt.kind === kind)?.text;
}

const summary = () => {
  useAi({
    ask: "Three sentences summarising what this invoice covers.",
    placeholder: "A short summary of the work this invoice covers.",
  });

  return <Paragraph id="summary" />;
};

// ---------------------------------------------------------------------------
// What the call does
// ---------------------------------------------------------------------------

test("calling it is what makes the node generated, with no mode declared", async () => {
  const built = await buildWith(summary, { dynamicMode: "placeholder" });

  assertEquals(paragraphOf(built.nodes).mode, "dynamic");
});

test("a node with no prompts at all stays static", async () => {
  const built = await buildWith(() => <Paragraph id="plain">Written here.</Paragraph>);

  assertEquals(paragraphOf(built.nodes).mode, "static");
});

test("the placeholder is what a preview shows", async () => {
  const built = await buildWith(summary, { dynamicMode: "placeholder" });

  assertEquals(
    paragraphOf(built.nodes).text,
    "A short summary of the work this invoice covers.",
  );
});

test("`ask` becomes the request the engine is given", async () => {
  const built = await buildWith(summary, {
    dynamicMode: "resolve",
    aiClient: new EchoAiClient(),
  });

  assertStringIncludes(
    paragraphOf(built.nodes).text ?? "",
    "Three sentences summarising what this invoice covers.",
  );
});

test("`voice` travels as the standing instruction", async () => {
  const built = await buildWith(() => {
    useAi({
      ask: "Summarise the invoice.",
      placeholder: "A summary.",
      voice: "A delivery lead writing to a finance contact.",
    });

    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "" } });

  assertEquals(
    promptOf(paragraphOf(built.nodes).prompts, "system"),
    "A delivery lead writing to a finance contact.",
  );
});

test("`avoid` travels as what the node must not say", async () => {
  const built = await buildWith(() => {
    useAi({
      ask: "Summarise the invoice.",
      placeholder: "A summary.",
      avoid: "Do not restate the totals.",
    });

    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "" } });

  assertEquals(
    promptOf(paragraphOf(built.nodes).prompts, "negative"),
    "Do not restate the totals.",
  );
});

// ---------------------------------------------------------------------------
// Facts
// ---------------------------------------------------------------------------

test("`from` takes the data itself rather than a sentence about it", async () => {
  const built = await buildWith(() => {
    const [data] = useState((input: InvoiceData) => input);

    useAi({
      ask: "Summarise the invoice.",
      placeholder: "A summary.",
      from: { period: data.period, reference: data.reference },
    });

    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "" } });

  const facts = promptOf(paragraphOf(built.nodes).prompts, "info") ?? "";

  assertStringIncludes(facts, "August 2026");
  assertStringIncludes(facts, "INV-2026-0142");
});

test("`from` still accepts a sentence, for facts that are already prose", async () => {
  const built = await buildWith(() => {
    useAi({
      ask: "Summarise the invoice.",
      placeholder: "A summary.",
      from: "The engagement ran from June to August.",
    });

    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "" } });

  assertEquals(
    promptOf(paragraphOf(built.nodes).prompts, "info"),
    "The engagement ran from June to August.",
  );
});

test("a fact nobody supplied is left out rather than written as null", async () => {
  const built = await buildWith(() => {
    useAi({
      ask: "Summarise the invoice.",
      placeholder: "A summary.",
      from: { period: "August 2026", chaser: undefined },
    });

    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "" } });

  const facts = promptOf(paragraphOf(built.nodes).prompts, "info") ?? "";

  assertEquals(facts.includes("chaser"), false);
});

// ---------------------------------------------------------------------------
// What it refuses
// ---------------------------------------------------------------------------

test("a missing placeholder is refused, because a blank preview reads as finished", async () => {
  await assertRejects(
    () =>
      buildWith(() => {
        (useAi as (config: unknown) => unknown)({ ask: "Summarise the invoice." });
        return <Paragraph id="summary" />;
      }),
    Error,
    "useAi needs a placeholder",
  );
});

test("a blank placeholder is refused the same way an absent one is", async () => {
  await assertRejects(
    () =>
      buildWith(() => {
        useAi({ ask: "Summarise.", placeholder: "   " });
        return <Paragraph id="summary" />;
      }),
    Error,
    "useAi needs a placeholder",
  );
});

test("the refusal says what a placeholder is actually for", async () => {
  let message = "";

  try {
    await buildWith(() => {
      useAi({ ask: "Summarise.", placeholder: "" });
      return <Paragraph id="summary" />;
    });
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }

  assertStringIncludes(message, "skipped or failed");
});

test("a voice with nothing to ask for is refused", async () => {
  await assertRejects(
    () =>
      buildWith(() => {
        (useAi as (config: unknown) => unknown)({
          placeholder: "A summary.",
          voice: "Plain English.",
        });
        return <Paragraph id="summary" />;
      }),
    Error,
    "needs an `ask`",
  );
});

test("an arm that supplies its own text stays written, and the prompts are dropped", async () => {
  // Hooks run in call order, so a component calls useAi before it knows which
  // arm it will take. An arm that says what it says is static, and the request
  // it never used goes nowhere — the alternative would be refusing a component
  // for a decision it had not made yet.
  const built = await buildWith(() => {
    useAi({ ask: "Summarise.", placeholder: "A summary." });
    return <Paragraph id="summary">Already written.</Paragraph>;
  });

  assertEquals(paragraphOf(built.nodes).mode, "static");
  assertEquals(paragraphOf(built.nodes).text, "Already written.");
  assertEquals(paragraphOf(built.nodes).prompts, undefined);
});

test("prompts written on the element beside its text are still a contradiction", async () => {
  // Here there is no decision to excuse it: both were written in the same place,
  // on the same tag, so one of them is a mistake.
  await assertRejects(
    () =>
      buildWith(() => (
        <Paragraph id="summary" generalPrompt="Summarise." placeholder="A summary.">
          Already written.
        </Paragraph>
      )),
    Error,
    "supplies both its content and a prompt",
  );
});

test("calling it outside a component says so rather than landing somewhere else", () => {
  let message = "";

  try {
    useAi({ ask: "Summarise.", placeholder: "A summary." });
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }

  assertStringIncludes(message, "useAi");
});

// ---------------------------------------------------------------------------
// Sharing house style
// ---------------------------------------------------------------------------

test("a shared hook can set the voice on a node it does not own", async () => {
  const useHouseVoice = () =>
    useAi({
      ask: "Summarise the invoice.",
      placeholder: "A summary.",
      voice: "You write for a housing association. Plain English, second person.",
    });

  const built = await buildWith(() => {
    useHouseVoice();
    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "" } });

  assertStringIncludes(
    promptOf(paragraphOf(built.nodes).prompts, "system") ?? "",
    "housing association",
  );
});

test("what it returns is the prompts now standing on the node", async () => {
  let returned: unknown;

  await buildWith(() => {
    returned = useAi({ ask: "Summarise.", placeholder: "A summary." });
    return <Paragraph id="summary" />;
  }, { dynamicMode: "placeholder" });

  assertEquals(returned, {
    generalPrompt: "Summarise.",
    placeholder: "A summary.",
  });
});
