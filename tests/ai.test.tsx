import { test } from "node:test";
import { assertEquals, assertRejects, assertStringIncludes, rejectionOf } from "./assert.ts";
import {
  AiGenerationError,
  buildDocument,
  buildProjectEngineDocument,
  createDeriverRegistry,
  defineDocumentProject,
  EchoAiClient,
  InMemoryDataProvider,
  resolveDocument,
  ScriptedAiClient,
} from "docxcelerate";
import type {
  DocumentModel,
  DocumentNode,
  GraphNode,
  ParagraphNode,
  PromptSpec,
  RuntimeState,
} from "docxcelerate";
import {
  aiProps,
  createPublishData,
  deriver,
  Document,
  Graph,
  Image,
  Paragraph,
  Section,
  template,
  useAi,
  useDeriver,
  useSetPrompts,
  useState,
} from "docxcelerate/template";

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

function firstOf(nodes: DocumentNode[]): DocumentNode {
  const section = nodes[0];

  if (section?.kind !== "section" || section.children[0] === undefined) {
    throw new Error("expected one node in the body");
  }

  return section.children[0];
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
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } });

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
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } });

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
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } });

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
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } });

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
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } });

  const facts = promptOf(paragraphOf(built.nodes).prompts, "info") ?? "";

  assertEquals(facts.includes("chaser"), false);
});

// ---------------------------------------------------------------------------
// Facts, published
// ---------------------------------------------------------------------------

/** A one-node project whose `from` is whatever the caller builds from the data. */
function publishedWith(facts: (data: InvoiceData) => Record<string, unknown>) {
  const Summary = () => {
    const [data] = useState((input: InvoiceData) => input);

    useAi({
      ask: "Summarise the invoice.",
      placeholder: "A summary.",
      from: facts(data) as never,
    });

    return <Paragraph id="summary" />;
  };

  return buildProjectEngineDocument(defineDocumentProject<InvoiceData>({
    id: "invoice",
    name: "Invoice",
    template: template<InvoiceData>(
      <Document id="invoice" title="Invoice">
        <Section id="body" title="Body">
          <Summary />
        </Section>
      </Document>,
    ),
    previewData: invoice,
  }));
}

const facts = (data: InvoiceData) => ({
  period: data.period,
  lines: data.lines,
  billing: { reference: data.reference },
  currency: "GBP",
});

test("`from` with request data publishes each value as the token the engine fills", async () => {
  const published = await publishedWith(facts);

  assertEquals(
    promptOf(paragraphOf(published.nodes).prompts, "info"),
    '{\n  "period": {{data.period}},\n  "lines": {{data.lines}},\n  "billing": {\n    "reference": {{data.reference}}\n  },\n  "currency": "GBP"\n}',
  );
});

test("a published `from` hands the model each recipient's own facts", async () => {
  const asked: Array<string | undefined> = [];
  const aiClient = {
    generateParagraph: (request: { prompts: PromptSpec[] }) => {
      asked.push(promptOf(request.prompts, "info"));
      return "Written.";
    },
  };
  const recipient = {
    reference: "INV-2026-0199",
    period: "September 2026",
    lines: [{ desc: "Audit" }],
  };

  await resolveDocument(await publishedWith(facts), {
    ctx: { ...recipient, availableTokens: 2_000 },
    derived: {},
    dataProvider: new InMemoryDataProvider(recipient),
    aiClient,
  } as never);

  assertEquals(asked, [
    '{\n  "period": September 2026,\n  "lines": [{"desc":"Audit"}],\n  "billing": {\n    "reference": INV-2026-0199\n  },\n  "currency": "GBP"\n}',
  ]);
});

test("a value read from request data can be the whole of `from`", async () => {
  const published = await publishedWith((data) => data.lines as never);

  assertEquals(promptOf(paragraphOf(published.nodes).prompts, "info"), "{{data.lines}}");
});

test("a local build writes `from` exactly as it did before publishing could", async () => {
  const built = await buildWith(() => {
    const [data] = useState((input: InvoiceData) => input);

    useAi({
      ask: "Summarise the invoice.",
      placeholder: "A summary.",
      from: { period: data.period, lines: data.lines },
    });

    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } });

  assertEquals(
    promptOf(paragraphOf(built.nodes).prompts, "info"),
    '{\n  "period": "August 2026",\n  "lines": [\n    {\n      "desc": "API build"\n    }\n  ]\n}',
  );
});

test("`from` given the whole request data while publishing says to name the fields", async () => {
  await assertRejects(
    () => publishedWith((data) => ({ all: data })),
    Error,
    "whole request data",
  );
});

test("`from` given a `.map()` while publishing says to pass the collection", async () => {
  await assertRejects(
    () => publishedWith((data) => ({ descs: data.lines.map((line) => line.desc) })),
    Error,
    "result of a `.map()`",
  );
});

// ---------------------------------------------------------------------------
// Examples
// ---------------------------------------------------------------------------

test("`example` travels as the shape a good answer takes", async () => {
  const built = await buildWith(() => {
    useAi({
      ask: "Summarise the invoice.",
      placeholder: "A summary.",
      example: "This invoice covers the June sprint on the payments integration.",
    });

    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } });

  assertEquals(
    promptOf(paragraphOf(built.nodes).prompts, "example"),
    "This invoice covers the June sprint on the payments integration.",
  );
});

test("a lone example travels as itself, with nothing wrapped around it", async () => {
  // Anything added around one example is text a model has to decide is not part
  // of the shape it was shown.
  const built = await buildWith(() => {
    useAi({ ask: "Summarise.", placeholder: "A summary.", example: "One short line." });

    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } });

  assertEquals(promptOf(paragraphOf(built.nodes).prompts, "example"), "One short line.");
});

test("several examples are numbered, so two do not read as one long answer", async () => {
  const built = await buildWith(() => {
    useAi({
      ask: "Summarise.",
      placeholder: "A summary.",
      example: ["The first shape.", "The second shape."],
    });

    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } });

  const shown = promptOf(paragraphOf(built.nodes).prompts, "example") ?? "";

  assertStringIncludes(shown, "Example 1:\nThe first shape.");
  assertStringIncludes(shown, "Example 2:\nThe second shape.");
});

test("a blank example is dropped rather than shown as the right answer", async () => {
  const built = await buildWith(() => {
    useAi({ ask: "Summarise.", placeholder: "A summary.", example: ["  ", ""] });

    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } });

  assertEquals(promptOf(paragraphOf(built.nodes).prompts, "example"), undefined);
});

test("the example reads last, after the request and the limits", async () => {
  // It is the thing the answer gets measured against, so it sits closest to
  // where the writing starts.
  const built = await buildWith(() => {
    useAi({
      ask: "Summarise.",
      placeholder: "A summary.",
      voice: "Plain English.",
      from: "The engagement ran from June to August.",
      avoid: "Do not restate the totals.",
      example: "This invoice covers the June sprint.",
    });

    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } });

  assertEquals(
    paragraphOf(built.nodes).prompts?.map((prompt) => prompt.kind),
    ["system", "general", "info", "negative", "example"],
  );
});

test("an example is enough on its own to make a node generated", async () => {
  const built = await buildWith(
    () => <Paragraph id="summary" examplePrompt="This invoice covers the June sprint." />,
    { dynamicMode: "placeholder" },
  );

  assertEquals(paragraphOf(built.nodes).mode, "dynamic");
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
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } });

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

test("an example set by a shared hook comes back with the rest of the prompts", async () => {
  let returned: unknown;

  await buildWith(() => {
    returned = useAi({
      ask: "Summarise.",
      placeholder: "A summary.",
      example: "This invoice covers the June sprint.",
    });

    return <Paragraph id="summary" />;
  }, { dynamicMode: "placeholder" });

  assertEquals(returned, {
    generalPrompt: "Summarise.",
    placeholder: "A summary.",
    examplePrompt: "This invoice covers the June sprint.",
  });
});

// ---------------------------------------------------------------------------
// Stand-in clients
// ---------------------------------------------------------------------------

test("ScriptedAiClient answers a paragraph by its id and records what was asked", async () => {
  const ai = new ScriptedAiClient({ summary: "Covers the API build." });
  const built = await buildWith(summary, { dynamicMode: "resolve", aiClient: ai });

  assertEquals(paragraphOf(built.nodes).text, "Covers the API build.");
  assertEquals(ai.requests.length, 1);
  assertEquals(ai.requests[0].id, "summary");
  assertEquals(ai.requests[0].kind, "paragraph");
  assertEquals(promptOf(ai.requests[0].prompts, "general"), "Three sentences summarising what this invoice covers.");
  assertStringIncludes(ai.requests[0].prompt, "Three sentences summarising what this invoice covers.");
});

test("ScriptedAiClient answers a chart and an image by their ids", async () => {
  const revenue = { data: { categories: ["Q1", "Q2"], series: [{ values: [120, 140] }] }, caption: "Revenue" };
  const ai = new ScriptedAiClient({ revenue, portrait: { path: "data:image/png;base64,AAAA", alt: "A portrait" } });
  const chart = firstOf(
    (await buildWith(() => <Graph id="revenue" generalPrompt="Plot revenue." placeholder="Revenue by quarter." />, {
      dynamicMode: "resolve",
      aiClient: ai,
    })).nodes,
  );
  const image = firstOf(
    (await buildWith(() => <Image id="portrait" generalPrompt="Draw it." placeholder="A portrait." />, {
      dynamicMode: "resolve",
      aiClient: ai,
    })).nodes,
  );

  assertEquals(chart.kind, "graph");
  if (chart.kind === "graph") {
    assertEquals(chart.data, revenue.data);
    assertEquals(chart.caption, "Revenue");
  }

  assertEquals(image.kind, "image");
  if (image.kind === "image") {
    assertEquals(image.path, "data:image/png;base64,AAAA");
  }

  assertEquals(ai.requests.map((r) => [r.id, r.kind]), [["revenue", "graph"], ["portrait", "image"]]);
});

test("a build whose client throws fails as an AiGenerationError naming the node", async () => {
  const down = new TypeError("fetch failed");

  const paragraphError = await rejectionOf(() =>
    buildWith(summary, { dynamicMode: "resolve", aiClient: { generateParagraph: () => { throw down; } } })
  );
  assertEquals(paragraphError instanceof AiGenerationError, true);
  const paragraphFailed = paragraphError as AiGenerationError;
  assertEquals(paragraphFailed.nodeId, "summary");
  assertEquals(paragraphFailed.kind, "paragraph");
  assertEquals(paragraphFailed.cause === down, true);
  assertEquals(paragraphFailed.message, 'Dynamic paragraph "summary" failed to generate: fetch failed');

  const graphError = await rejectionOf(() =>
    buildWith(() => <Graph id="revenue" generalPrompt="Plot revenue." placeholder="Revenue." />, {
      dynamicMode: "resolve",
      aiClient: { generateParagraph: () => "", generateGraph: () => Promise.reject(down) },
    })
  );
  assertEquals(graphError instanceof AiGenerationError, true);
  assertEquals((graphError as AiGenerationError).nodeId, "revenue");
  assertEquals((graphError as AiGenerationError).kind, "graph");
  assertEquals((graphError as AiGenerationError).cause === down, true);

  const imageError = await rejectionOf(() =>
    buildWith(() => <Image id="portrait" generalPrompt="Draw it." placeholder="A portrait." />, {
      dynamicMode: "resolve",
      aiClient: { generateParagraph: () => "", generateImage: () => { throw down; } },
    })
  );
  assertEquals(imageError instanceof AiGenerationError, true);
  assertEquals((imageError as AiGenerationError).nodeId, "portrait");
  assertEquals((imageError as AiGenerationError).kind, "image");
  assertEquals((imageError as AiGenerationError).cause === down, true);
});

test("a preview never calls the client, so a failing client cannot fail it", async () => {
  const built = await buildWith(summary, {
    dynamicMode: "placeholder",
    aiClient: { generateParagraph: () => { throw new Error("never"); } },
  });

  assertEquals(paragraphOf(built.nodes).text, "A short summary of the work this invoice covers.");
});

test("a ScriptedAiClient with no answer fails as an AiGenerationError, its own message on cause", async () => {
  const error = await rejectionOf(() =>
    buildWith(summary, { dynamicMode: "resolve", aiClient: new ScriptedAiClient({}) })
  );

  assertEquals(error instanceof AiGenerationError, true);
  const failed = error as AiGenerationError;
  assertEquals(failed.nodeId, "summary");
  assertEquals((failed.cause as Error).message, 'ScriptedAiClient has no answer for paragraph "summary". Known ids: (none).');
});

test("an answer that is not text is docxcelerate's refusal, not an AiGenerationError", async () => {
  const error = await rejectionOf(() =>
    buildWith(summary, { dynamicMode: "resolve", aiClient: { generateParagraph: () => ({ text: "hi" }) as never } })
  );

  assertEquals(error instanceof AiGenerationError, false);
  assertStringIncludes((error as Error).message, 'Dynamic paragraph "summary" got an answer that is not text');
});

test("ScriptedAiClient matches a repeat's suffixed ids by their base id, and an exact id wins", async () => {
  const model: DocumentModel = {
    schemaVersion: "docxcelerate.letter/v0",
    id: "notes",
    title: "Notes",
    nodes: [{
      id: "notes",
      kind: "repeat",
      source: { scope: "data", path: "notes" },
      as: "note",
      indexAs: "index",
      children: [{ id: "note", kind: "paragraph", mode: "dynamic", prompts: [{ kind: "general", text: "Write note {{ctx.index}}." }] }],
    }],
  };
  const ai = new ScriptedAiClient({ note: "Base.", "note-1": "Second." });
  const data = { notes: [{}, {}, {}] };
  const doc = await resolveDocument(model, {
    ctx: {},
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: ai,
  });

  assertEquals(doc.nodes.map((n) => (n.kind === "paragraph" ? n.text : undefined)), ["Base.", "Second.", "Base."]);
  assertEquals(ai.requests.map((r) => r.id), ["note-0", "note-1", "note-2"]);
  assertEquals(ai.requests[1].prompt, "GENERAL: Write note 1.");
});

test("ScriptedAiClient refuses an id it has no answer for, listing the ids it knows", async () => {
  const ai = new ScriptedAiClient({ other: "x", more: "y" });

  await assertRejects(
    () => buildWith(summary, { dynamicMode: "resolve", aiClient: ai }),
    Error,
    'ScriptedAiClient has no answer for paragraph "summary". Known ids: "other", "more".',
  );
  assertEquals(ai.requests.map((r) => r.id), ["summary"]);

  await assertRejects(
    () => buildWith(summary, { dynamicMode: "resolve", aiClient: new ScriptedAiClient({}) }),
    Error,
    "Known ids: (none).",
  );
});

test("ScriptedAiClient refuses an answer the node cannot use", async () => {
  await assertRejects(
    () => buildWith(summary, { dynamicMode: "resolve", aiClient: new ScriptedAiClient({ summary: { caption: "x" } }) }),
    Error,
    `ScriptedAiClient's answer for "summary" is not one a paragraph can use.`,
  );
  await assertRejects(
    () =>
      buildWith(() => <Graph id="revenue" generalPrompt="Plot revenue." placeholder="Revenue." />, {
        dynamicMode: "resolve",
        aiClient: new ScriptedAiClient({ revenue: "text" }),
      }),
    Error,
    `ScriptedAiClient's answer for "revenue" is not one a graph can use.`,
  );
  await assertRejects(
    () =>
      buildWith(() => <Image id="portrait" generalPrompt="Draw it." placeholder="A portrait." />, {
        dynamicMode: "resolve",
        aiClient: new ScriptedAiClient({ portrait: { caption: "no path" } }),
      }),
    Error,
    `ScriptedAiClient's answer for "portrait" is not one a image can use.`,
  );
});

test("ScriptedAiClient is asked nothing by a preview", async () => {
  const ai = new ScriptedAiClient({ summary: "Unused." });
  const built = await buildWith(summary, { dynamicMode: "placeholder", aiClient: ai });

  assertEquals(paragraphOf(built.nodes).text, "A short summary of the work this invoice covers.");
  assertEquals(ai.requests, []);
});

test("EchoAiClient fills a generated chart with labelled sample data, keeping its chart type", async () => {
  const built = await buildWith(
    () => <Graph id="revenue" graphType="line" generalPrompt="Plot revenue by quarter." placeholder="Revenue." />,
    { dynamicMode: "resolve", aiClient: new EchoAiClient() },
  );
  const chart = firstOf(built.nodes);

  assertEquals(chart.kind, "graph");
  if (chart.kind === "graph") {
    assertEquals(chart.graphType, "line");
    assertEquals(chart.data, {
      categories: ["A", "B", "C"],
      series: [{ label: "Sample", values: [3, 5, 4], sizes: [3, 5, 4] }],
    });
    assertStringIncludes(chart.caption ?? "", "[AI draft for revenue] ");
    assertStringIncludes(chart.caption ?? "", "Plot revenue by quarter.");
  }
});

test("EchoAiClient fills a generated image with a sample picture at the node's size", async () => {
  const built = await buildWith(
    () => <Image id="portrait" generalPrompt="Draw a portrait." placeholder="A portrait." width={200} height={100} />,
    { dynamicMode: "resolve", aiClient: new EchoAiClient() },
  );
  const image = firstOf(built.nodes);

  assertEquals(image.kind, "image");
  if (image.kind === "image") {
    assertStringIncludes(image.path ?? "", "data:image/png;base64,");
    assertStringIncludes(image.alt ?? "", "[AI draft for portrait] ");
    assertStringIncludes(image.alt ?? "", "Draw a portrait.");
    assertEquals([image.width, image.height], [200, 100]);
  }
});

function graphOf(nodes: DocumentNode[]): GraphNode {
  const section = nodes[0];
  const graph = section?.kind === "section" ? section.children[0] : undefined;
  if (graph?.kind !== "graph") {
    throw new Error("expected one graph in the body");
  }
  return graph;
}

test("a dynamic chart keeps its own caption when the model does not write one", async () => {
  const built = await buildWith(
    () => <Graph id="trend" generalPrompt="Plot the balance." caption="Own caption" />,
    {
      dynamicMode: "resolve",
      aiClient: {
        generateParagraph: () => "Written.",
        // The cast lets a string through GraphData, as a real client returns parsed JSON.
        generateGraph: () => ({ data: { series: [{ values: ["12"] }] } }) as never,
      },
    },
  );

  assertEquals(graphOf(built.nodes).caption, "Own caption");
  assertEquals(graphOf(built.nodes).data, { series: [{ values: [12] }] });
});

test("a preview of a dynamic chart carries its caption", async () => {
  const built = await buildWith(
    () => <Graph id="trend" generalPrompt="Plot the balance." caption="Own caption" />,
    { dynamicMode: "placeholder" },
  );

  const graph = graphOf(built.nodes);
  assertEquals(graph.caption, "Own caption");
  assertEquals(graph.mode, "dynamic");
  assertEquals("data" in graph, false);
});

test("a generated paragraph is trimmed and stripped of characters Word cannot open", async () => {
  const built = await buildWith(summary, {
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "\n  Hello\u0001 world\u000b end\uffff \n" },
  });
  assertEquals(paragraphOf(built.nodes).text, "Hello world end");
});

test("a generated paragraph with nothing left once trimmed names the node", async () => {
  await assertRejects(
    () => buildWith(summary, { dynamicMode: "resolve", aiClient: { generateParagraph: () => " \n " } }),
    Error,
    'Dynamic paragraph "summary" got an empty answer.',
  );
});

test("a generated chart with nothing to plot names the node", async () => {
  const chart = () => <Graph id="trend" generalPrompt="Plot the balance." />;

  for (const answer of [{}, { data: { series: [{ values: [null] }] } }]) {
    await assertRejects(
      () =>
        buildWith(chart, {
          dynamicMode: "resolve",
          aiClient: { generateParagraph: () => "Written.", generateGraph: () => answer },
        }),
      Error,
      'Dynamic graph "trend" got an answer it cannot plot:',
    );
  }
});

test("ScriptedAiClient's empty answer is refused like a model's", async () => {
  await assertRejects(
    () => buildWith(summary, { dynamicMode: "resolve", aiClient: new ScriptedAiClient({ summary: "" }) }),
    Error,
    'Dynamic paragraph "summary" got an empty answer.',
  );
});

test("a preview never asks, so it shows the placeholder, not an empty answer", async () => {
  const built = await buildWith(summary, {
    dynamicMode: "placeholder",
    aiClient: { generateParagraph: () => "" },
  });

  assertEquals(paragraphOf(built.nodes).text, "A short summary of the work this invoice covers.");
});

test("publishing a generated paragraph and chart writes neither text nor data", async () => {
  const Summary = summary as () => never;
  const published = await buildProjectEngineDocument(defineDocumentProject<InvoiceData>({
    id: "invoice",
    name: "Invoice",
    template: template<InvoiceData>(
      <Document id="invoice" title="Invoice">
        <Section id="body" title="Body">
          <Summary />
          <Graph id="trend" generalPrompt="Plot the balance." />
        </Section>
      </Document>,
    ),
    previewData: invoice,
  }));

  const section = published.nodes[0];
  if (section?.kind !== "section") throw new Error("expected the body section");
  const [paragraph, graph] = section.children;

  assertEquals(paragraph?.kind === "paragraph" && "text" in paragraph, false);
  assertEquals(graph?.kind === "graph" && "data" in graph, false);
  assertEquals(paragraph?.kind, "paragraph");
  assertEquals(graph?.kind, "graph");
});

test("a generated paragraph that is not text names the node", async () => {
  await assertRejects(
    () =>
      buildWith(summary, {
        dynamicMode: "resolve",
        aiClient: { generateParagraph: () => ({ text: "hi" }) as never },
      }),
    Error,
    'Dynamic paragraph "summary" got an answer that is not text: it returned an object.',
  );
});

// ---------------------------------------------------------------------------
// Options
// ---------------------------------------------------------------------------

const withOptions = (options: unknown) => () => {
  useAi({ ask: "Summarise.", placeholder: "A summary.", options: options as never });
  return <Paragraph id="summary" />;
};

/** A client that records the options of every node it is asked to write. */
function optionsClient(seen: unknown[]) {
  return {
    generateParagraph: (request: { node: { aiOptions?: unknown } }) => {
      seen.push(request.node.aiOptions);
      return "Written.";
    },
    generateImage: (request: { node: { aiOptions?: unknown } }) => {
      seen.push(request.node.aiOptions);
      return { path: "data:image/png;base64,AAAA" };
    },
    generateGraph: (request: { node: { aiOptions?: unknown } }) => {
      seen.push(request.node.aiOptions);
      return { data: { categories: ["Q1"], series: [{ values: [1] }] } };
    },
  };
}

test("a paragraph's options reach the client", async () => {
  const seen: unknown[] = [];
  const options = { model: "large", maxTokens: 400 };
  const built = await buildWith(withOptions(options), {
    dynamicMode: "resolve",
    aiClient: optionsClient(seen),
  });

  assertEquals(seen, [options]);
  assertEquals(paragraphOf(built.nodes).aiOptions, options);
});

test("an image's and a chart's options reach their clients", async () => {
  const seen: unknown[] = [];

  for (const node of [<Image id="portrait" />, <Graph id="revenue" />]) {
    await buildWith(() => {
      useAi({ ask: "Make it.", placeholder: "Stand-in.", options: { model: "img-1" } });
      return node;
    }, { dynamicMode: "resolve", aiClient: optionsClient(seen) });
  }

  assertEquals(seen, [{ model: "img-1" }, { model: "img-1" }]);
});

test("options survive publishing and a stored model", async () => {
  const Summary = () => {
    useAi({ ask: "Summarise.", placeholder: "A summary.", options: { model: "large" } });
    return <Paragraph id="summary" />;
  };
  const published = await buildProjectEngineDocument(defineDocumentProject<InvoiceData>({
    id: "invoice",
    name: "Invoice",
    template: template<InvoiceData>(
      <Document id="invoice" title="Invoice">
        <Section id="body" title="Body">
          <Summary />
        </Section>
      </Document>,
    ),
    previewData: invoice,
  }));

  assertEquals(paragraphOf(published.nodes).aiOptions, { model: "large" });

  const seen: unknown[] = [];

  await resolveDocument(JSON.parse(JSON.stringify(published)), {
    ctx: { ...invoice, availableTokens: 2_000 },
    derived: {},
    dataProvider: new InMemoryDataProvider({ ...invoice }),
    aiClient: optionsClient(seen),
  } as never);

  assertEquals(seen, [{ model: "large" }]);
});

test("omitted or empty options leave no aiOptions on the node", async () => {
  for (const component of [summary, withOptions({})]) {
    const built = await buildWith(component, {
      dynamicMode: "resolve",
      aiClient: { generateParagraph: () => "" },
    });

    assertEquals("aiOptions" in paragraphOf(built.nodes), false);
  }
});

test("a preview carries no options", async () => {
  const node = paragraphOf((await buildWith(withOptions({ model: "large" }), {
    dynamicMode: "placeholder",
  })).nodes);

  assertEquals("aiOptions" in node, false);
  assertEquals(node.text, "A summary.");
});

test("a static arm drops its options", async () => {
  const built = await buildWith(() => {
    useAi({ ask: "Summarise.", placeholder: "A summary.", options: { model: "large" } });
    return <Paragraph id="p">Written here.</Paragraph>;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "" } });
  const node = paragraphOf(built.nodes);

  assertEquals(node.mode, "static");
  assertEquals("aiOptions" in node, false);
});

test("a later layer replaces options whole", async () => {
  const built = await buildWith(() => {
    useSetPrompts({ aiOptions: { model: "small", temperature: 0 } });
    useAi({ ask: "Summarise.", placeholder: "A summary.", options: { model: "large" } });
    return <Paragraph id="summary" />;
  }, { dynamicMode: "resolve", aiClient: { generateParagraph: () => "" } });

  assertEquals(paragraphOf(built.nodes).aiOptions, { model: "large" });
});

test("options that are not an object are refused", async () => {
  await assertRejects(
    () => buildWith(withOptions("large"), { dynamicMode: "placeholder" }),
    Error,
    "must be an object",
  );
});

test("request data in options is refused at publish", async () => {
  await assertRejects(
    () => {
      const Summary = () => {
        const [data] = useState((input: InvoiceData) => input);
        useAi({
          ask: "Summarise.",
          placeholder: "A summary.",
          options: { model: data.reference as never },
        });
        return <Paragraph id="summary" />;
      };

      return buildProjectEngineDocument(defineDocumentProject<InvoiceData>({
        id: "invoice",
        name: "Invoice",
        template: template<InvoiceData>(
          <Document id="invoice" title="Invoice">
            <Section id="body" title="Body">
              <Summary />
            </Section>
          </Document>,
        ),
        previewData: invoice,
      }));
    },
    Error,
    "handed request data while publishing",
  );
});

// ---------------------------------------------------------------------------
// After a deriver: the same config, as props
// ---------------------------------------------------------------------------

/** Cheap, so it runs in a preview too. */
const lineFacts = deriver({
  name: "lineFacts",
  run: (lines: Array<{ desc: string }>) => ({ count: lines.length, first: lines[0]?.desc ?? "" }),
});

const Reminder = async () => {
  const [lines] = useState((data: InvoiceData) => data.lines);
  const facts = await useDeriver(lineFacts, [lines]);

  return (
    <Paragraph
      id="reminder"
      {...aiProps({
        ask: `Say that ${facts.count} line is billed.`,
        placeholder: `A reminder about ${facts.count} line.`,
        from: { first: facts.first },
      })}
    />
  );
};

const RESOLVE = { dynamicMode: "resolve", aiClient: { generateParagraph: () => "Written." } } as const;

function publishDoc(component: () => unknown): Promise<DocumentModel> {
  const Body = component as () => never;

  return buildDocument(
    template<InvoiceData>(
      <Document id="invoice" title="Invoice">
        <Section id="body" title="Body">
          <Body />
        </Section>
      </Document>,
    ),
    createPublishData() as InvoiceData,
    {
      branchMode: "publish",
      deriverMode: "preserve",
      dynamicMode: "resolve",
      aiClient: { generateParagraph: () => "Written." },
    },
  );
}

function stateWith(aiClient: RuntimeState["aiClient"]): RuntimeState {
  return {
    ctx: { ...invoice },
    derived: {},
    dataProvider: new InMemoryDataProvider(invoice as unknown as Record<string, unknown>),
    aiClient,
  };
}

test("aiProps returns the prompts useAi would set, and nothing it was not given", () => {
  assertEquals(aiProps({ ask: "Summarise.", placeholder: "A summary." }), {
    generalPrompt: "Summarise.",
    placeholder: "A summary.",
  });
});

test("aiProps maps every field the way useAi does", () => {
  assertEquals(
    aiProps({
      ask: "Summarise.",
      placeholder: "A summary.",
      voice: "Plain.",
      from: { period: "August 2026", chaser: undefined },
      avoid: "No totals.",
      example: ["One.", "Two."],
    }),
    {
      generalPrompt: "Summarise.",
      placeholder: "A summary.",
      systemPrompt: "Plain.",
      negativePrompt: "No totals.",
      infoPrompt: '{\n  "period": "August 2026"\n}',
      examplePrompt: "Example 1:\nOne.\n\nExample 2:\nTwo.",
    },
  );
});

test("aiProps spread on an element makes the node generated", async () => {
  const built = await buildWith(
    () => <Paragraph id="summary" {...aiProps({ ask: "Summarise.", placeholder: "A summary.", voice: "Plain." })} />,
    RESOLVE,
  );
  const node = paragraphOf(built.nodes);

  assertEquals(node.mode, "dynamic");
  assertEquals(promptOf(node.prompts, "general"), "Summarise.");
  assertEquals(promptOf(node.prompts, "system"), "Plain.");
});

test("aiProps written after await useDeriver reads the deriver's result", async () => {
  const built = await buildWith(Reminder, RESOLVE);
  const node = paragraphOf(built.nodes);

  assertEquals(promptOf(node.prompts, "general"), "Say that 1 line is billed.");
  assertEquals(promptOf(node.prompts, "info"), '{\n  "first": "API build"\n}');
});

test("a preview shows the placeholder, filled from the deriver's result", async () => {
  const built = await buildWith(Reminder, { dynamicMode: "placeholder" });

  assertEquals(paragraphOf(built.nodes).text, "A reminder about 1 line.");
});

test("published, the prompts carry the tokens the stand-in wrote", async () => {
  const published = await publishDoc(Reminder);
  const node = paragraphOf(published.nodes);

  assertEquals(promptOf(node.prompts, "general"), "Say that {{derived.lineFacts.count}} line is billed.");
  assertEquals(promptOf(node.prompts, "info"), '{\n  "first": {{derived.lineFacts.first}}\n}');
  assertEquals(node.derivers?.map((invocation) => invocation.name), ["lineFacts"]);
});

test("publishing then resolving asks the model what building directly asks", async () => {
  const asked: PromptSpec[][] = [];
  const client = {
    generateParagraph: (request: { prompts: PromptSpec[] }) => (asked.push(request.prompts), "Written."),
  };

  const published = await publishDoc(Reminder);
  const resolved = await resolveDocument(published, stateWith(client), {
    derivers: createDeriverRegistry([lineFacts]),
  });
  const direct = await buildWith(Reminder, { dynamicMode: "resolve", aiClient: client });

  assertEquals(asked.length, 2);
  // A string filled in from request data is written unquoted, as main has always
  // published it, so the only difference from a direct build is the quotes round it.
  assertEquals(asked[0][0], asked[1][0]);
  assertEquals(asked[0][1], { kind: "info", text: '{\n  "first": API build\n}' });
  assertEquals(asked[1][1], { kind: "info", text: '{\n  "first": "API build"\n}' });
  assertEquals(paragraphOf(resolved.nodes).text, "Written.");
  assertEquals(paragraphOf(direct.nodes).text, "Written.");
});

test("a deriver's whole result as `from` publishes as its one token", async () => {
  const published = await publishDoc(async () => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    const facts = await useDeriver(lineFacts, [lines]);

    return <Paragraph id="reminder" {...aiProps({ ask: "Remind.", placeholder: "A reminder.", from: facts })} />;
  });

  assertEquals(promptOf(paragraphOf(published.nodes).prompts, "info"), "{{derived.lineFacts}}");
});

test("useAi's facts from request data publish as tokens", async () => {
  const published = await publishDoc(() => {
    const [data] = useState((input: InvoiceData) => input);

    useAi({ ask: "Summarise.", placeholder: "A summary.", from: { period: data.period } });

    return <Paragraph id="summary" />;
  });

  assertEquals(promptOf(paragraphOf(published.nodes).prompts, "info"), '{\n  "period": {{data.period}}\n}');
});

test("aiProps refuses a missing placeholder, under its own name", () => {
  let message = "";

  try {
    (aiProps as (config: unknown) => unknown)({ ask: "Summarise." });
  } catch (error) {
    message = (error as Error).message;
  }

  assertStringIncludes(message, "aiProps needs a placeholder");
  assertStringIncludes(message, "skipped or failed");
});

test("aiProps refuses a blank ask", () => {
  let message = "";

  try {
    aiProps({ ask: "  ", placeholder: "A summary." });
  } catch (error) {
    message = (error as Error).message;
  }

  assertStringIncludes(message, "aiProps needs an `ask`");
});

test("aiProps on the element wins over what a hook set around it", async () => {
  const built = await buildWith(() => {
    useAi({ ask: "House ask.", placeholder: "House.", voice: "House voice." });

    return <Paragraph id="summary" {...aiProps({ ask: "Own ask.", placeholder: "Own.", voice: "Own voice." })} />;
  }, RESOLVE);
  const node = paragraphOf(built.nodes);

  assertEquals(promptOf(node.prompts, "system"), "Own voice.");
  assertEquals(promptOf(node.prompts, "general"), "Own ask.");
});

test("aiProps beside written text is the same contradiction as prompt props", async () => {
  await assertRejects(
    () =>
      buildWith(() => (
        <Paragraph id="summary" {...aiProps({ ask: "Summarise.", placeholder: "A summary." })}>
          Already written.
        </Paragraph>
      )),
    Error,
    "supplies both its content and a prompt",
  );
});

test("a client that throws while resolving an aiProps node fails the document", async () => {
  const published = await publishDoc(Reminder);

  const failing = {
    generateParagraph: () => {
      throw new Error("model down");
    },
  };

  await assertRejects(
    () => resolveDocument(published, stateWith(failing), { derivers: createDeriverRegistry([lineFacts]) }),
    Error,
    "model down",
  );
});
