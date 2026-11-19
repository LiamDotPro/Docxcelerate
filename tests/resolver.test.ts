import { test } from "node:test";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import {
  type AiGraphResult,
  createDeriverRegistry,
  createRuntimeState,
  type DocumentModel,
  type DocumentNode,
  type GraphNode,
  resolveDocument,
  type RuntimeState,
} from "docxcelerate";

/**
 * What an engine does with a published document.
 *
 * Everything here starts from a `DocumentModel` shaped the way publishing
 * leaves it — tokens unresolved, both arms of a branch present, loops still
 * loops — because that is the only thing an engine is ever handed.
 */
function engineState(data: Record<string, unknown>): RuntimeState {
  return createRuntimeState({
    data,
    aiClient: {
      generateParagraph: (request) => `[generated ${request.node.id}]`,
      generateImage: () => ({ path: "generated.png", alt: "generated" }),
      generateGraph: () => ({ data: { series: [{ values: [1] }] }, caption: "generated" }),
    },
  });
}

function published(nodes: DocumentNode[]): DocumentModel {
  return {
    schemaVersion: "docxcelerate.letter/v0",
    id: "published",
    title: "Published",
    nodes,
  };
}

function paragraph(id: string, text: string, extra: Partial<DocumentNode> = {}): DocumentNode {
  return { id, kind: "paragraph", mode: "static", text, ...extra } as DocumentNode;
}

function ids(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) =>
    node.kind === "section" || node.kind === "repeat" ? ids(node.children) : [node.id]
  );
}

function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    if (node.kind === "section" || node.kind === "repeat") return texts(node.children);
    return node.kind === "paragraph" ? [node.text ?? ""] : [];
  });
}

test("tokens resolve against the request data", async () => {
  const doc = await resolveDocument(
    published([paragraph("greeting", "Hello {{data.name}}, of {{data.city}}.")]),
    engineState({ name: "Avery", city: "Leeds" }),
  );

  assertEquals(texts(doc.nodes), ["Hello Avery, of Leeds."]);
});

test("a chart's alt resolves its tokens against the request data", async () => {
  const doc = await resolveDocument(
    published([{
      id: "revenue",
      kind: "graph",
      mode: "static",
      graphType: "bar",
      alt: "Revenue rose {{data.growth}}%",
      data: { series: [{ values: [1] }] },
    } as DocumentNode]),
    engineState({ growth: 40 }),
  );

  assertEquals((doc.nodes[0] as GraphNode).alt, "Revenue rose 40%");
});

test("only the arm whose condition holds survives", async () => {
  const nodes = [
    paragraph("arrears", "You owe money.", {
      when: { type: "compare", operator: "gt", left: { type: "ref", ref: { scope: "data", path: "balanceDue" } }, right: { type: "literal", value: 0 } },
    }),
    paragraph("settled", "Nothing outstanding.", {
      when: { type: "compare", operator: "lte", left: { type: "ref", ref: { scope: "data", path: "balanceDue" } }, right: { type: "literal", value: 0 } },
    }),
  ];

  const owing = await resolveDocument(published(nodes), engineState({ balanceDue: 240 }));
  const clear = await resolveDocument(published(nodes), engineState({ balanceDue: 0 }));

  assertEquals(ids(owing.nodes), ["arrears"]);
  assertEquals(ids(clear.nodes), ["settled"]);
});

test("a condition on a section takes its children with it", async () => {
  const doc = await resolveDocument(
    published([
      {
        id: "arrears",
        kind: "section",
        title: "Arrears",
        when: { type: "truthy", ref: { scope: "data", path: "hasArrears" } },
        children: [paragraph("chase", "Please pay.")],
      },
    ]),
    engineState({ hasArrears: false }),
  );

  assertEquals(doc.nodes, []);
});

test("a stored repeat binds its position when it names one, and restores it", async () => {
  const state = engineState({
    visits: [{ label: "A", kept: true }, { label: "B", kept: false }, { label: "C", kept: true }],
  });
  const doc = await resolveDocument(
    published([
      {
        id: "visits",
        kind: "repeat",
        source: { scope: "data", path: "visits" },
        as: "visit",
        indexAs: "index",
        positionAs: "position",
        where: { type: "truthy", ref: { scope: "ctx", path: "visit.kept" } },
        children: [paragraph("visit", "{{ctx.position}}/{{ctx.index}}: {{ctx.visit.label}}")],
      },
    ]),
    state,
  );

  assertEquals(texts(doc.nodes), ["1/0: A", "2/2: C"]);
  assertEquals(state.ctx.position, undefined);
});

test("a repeat is walked once per entry, with ids suffixed by position", async () => {
  const doc = await resolveDocument(
    published([
      {
        id: "visits",
        kind: "repeat",
        source: { scope: "data", path: "visits" },
        as: "visit",
        indexAs: "index",
        children: [paragraph("visit", "Visit {{ctx.index}}: {{ctx.visit.label}}")],
      },
    ]),
    engineState({ visits: [{ label: "First" }, { label: "Second" }, { label: "Third" }] }),
  );

  assertEquals(ids(doc.nodes), ["visit-0", "visit-1", "visit-2"]);
  assertEquals(texts(doc.nodes), [
    "Visit 0: First",
    "Visit 1: Second",
    "Visit 2: Third",
  ]);
});

test("a repeat reads its source by scope: derived and ctx, not only the provider", async () => {
  const loop = (scope: "derived" | "ctx") =>
    published([
      {
        id: "rows",
        kind: "repeat",
        source: { scope, path: "rows" },
        as: "row",
        indexAs: "index",
        children: [paragraph("row", "{{ctx.row}}")],
      },
    ]);
  const derivedState = { ...engineState({}), derived: { rows: ["a", "b"] } };
  assertEquals(texts((await resolveDocument(loop("derived"), derivedState)).nodes), ["a", "b"]);
  assertEquals(texts((await resolveDocument(loop("ctx"), engineState({ rows: ["c"] }))).nodes), ["c"]);
});

test("a repeat leaves no structure of its own behind", async () => {
  // Wrapping the passes in a section would print a heading nobody asked for.
  const doc = await resolveDocument(
    published([
      paragraph("before", "Before."),
      {
        id: "visits",
        kind: "repeat",
        source: { scope: "data", path: "visits" },
        as: "visit",
        indexAs: "index",
        children: [paragraph("visit", "{{ctx.visit.label}}")],
      },
      paragraph("after", "After."),
    ]),
    engineState({ visits: [{ label: "One" }, { label: "Two" }] }),
  );

  assertEquals(doc.nodes.map((node) => node.kind), [
    "paragraph",
    "paragraph",
    "paragraph",
    "paragraph",
  ]);
  assertEquals(ids(doc.nodes), ["before", "visit-0", "visit-1", "after"]);
});

test("an empty or missing collection contributes nothing", async () => {
  const loop = (source: string): DocumentNode => ({
    id: "visits",
    kind: "repeat",
    source: { scope: "data", path: source },
    as: "visit",
    indexAs: "index",
    children: [paragraph("visit", "{{ctx.visit.label}}")],
  });

  const empty = await resolveDocument(published([loop("visits")]), engineState({ visits: [] }));
  const missing = await resolveDocument(published([loop("nothing")]), engineState({ visits: [] }));

  assertEquals(empty.nodes, []);
  assertEquals(missing.nodes, []);
});

test("the loop binding does not leak past the loop", async () => {
  const doc = await resolveDocument(
    published([
      {
        id: "visits",
        kind: "repeat",
        source: { scope: "data", path: "visits" },
        as: "visit",
        indexAs: "index",
        children: [paragraph("inside", "{{ctx.visit.label}}")],
      },
      paragraph("outside", "After: [{{ctx.visit.label}}]"),
    ]),
    engineState({ visits: [{ label: "One" }] }),
  );

  assertEquals(texts(doc.nodes), ["One", "After: []"]);
});

test("a condition inside a loop is evaluated per entry", async () => {
  const doc = await resolveDocument(
    published([
      {
        id: "visits",
        kind: "repeat",
        source: { scope: "data", path: "visits" },
        as: "visit",
        indexAs: "index",
        children: [
          paragraph("flagged", "Flagged: {{ctx.visit.label}}", {
            when: { type: "truthy", ref: { scope: "ctx", path: "visit.flag" } },
          }),
        ],
      },
    ]),
    engineState({
      visits: [{ label: "One", flag: true }, { label: "Two" }, { label: "Three", flag: true }],
    }),
  );

  assertEquals(ids(doc.nodes), ["flagged-0", "flagged-2"]);
});

test("derivers carried on a node run before it resolves", async () => {
  const doc = await resolveDocument(
    published([
      paragraph("balance", "You owe {{derived.balanceLabel}}.", {
        derivers: [{
          name: "currencyLabel",
          output: "balanceLabel",
          inputs: [{ type: "ref", ref: { scope: "data", path: "balanceDue" } }],
        }],
      }),
    ]),
    engineState({ balanceDue: 240 }),
    {
      derivers: createDeriverRegistry({
        currencyLabel: ([amount]) => `£${Number(amount).toFixed(2)}`,
      }),
    },
  );

  assertEquals(texts(doc.nodes), ["You owe £240.00."]);
});

test("a deriver runs per entry inside a loop", async () => {
  const doc = await resolveDocument(
    published([
      {
        id: "visits",
        kind: "repeat",
        source: { scope: "data", path: "visits" },
        as: "visit",
        indexAs: "index",
        children: [
          paragraph("cost", "{{derived.costLabel}}", {
            derivers: [{
              name: "currencyLabel",
              output: "costLabel",
              inputs: [{ type: "ref", ref: { scope: "ctx", path: "visit.cost" } }],
            }],
          }),
        ],
      },
    ]),
    engineState({ visits: [{ cost: 10 }, { cost: 25 }] }),
    {
      derivers: createDeriverRegistry({
        currencyLabel: ([amount]) => `£${Number(amount).toFixed(2)}`,
      }),
    },
  );

  assertEquals(texts(doc.nodes), ["£10.00", "£25.00"]);
});

test("a dynamic node reaches the AI client with its prompts resolved", async () => {
  const doc = await resolveDocument(
    published([
      {
        id: "summary",
        kind: "paragraph",
        mode: "dynamic",
        prompts: [{ kind: "general", text: "Summarise for {{data.name}}." }],
      },
    ]),
    engineState({ name: "Avery" }),
  );

  const node = doc.nodes[0];
  assertEquals(node.kind, "paragraph");
  if (node.kind === "paragraph") {
    assertEquals(node.prompts?.[0].text, "Summarise for Avery.");
    assertEquals(node.text, "[generated summary]");
  }
});

test("a condition keeps a dynamic node from being generated at all", async () => {
  let calls = 0;
  const state = engineState({ include: false });
  state.aiClient = {
    generateParagraph: () => {
      calls += 1;
      return "";
    },
  };

  const doc = await resolveDocument(
    published([
      {
        id: "summary",
        kind: "paragraph",
        mode: "dynamic",
        when: { type: "truthy", ref: { scope: "data", path: "include" } },
        prompts: [{ kind: "general", text: "Summarise." }],
      },
    ]),
    state,
  );

  assertEquals(doc.nodes, []);
  // The point of deciding before generating: a node nobody will read costs
  // nothing to leave out.
  assertEquals(calls, 0);
});

test("an unknown deriver names itself rather than resolving to nothing", async () => {
  await assertRejectsWith(
    () =>
      resolveDocument(
        published([
          paragraph("balance", "{{derived.x}}", {
            derivers: [{ name: "notRegistered", output: "x", inputs: [] }],
          }),
        ]),
        engineState({}),
      ),
    "notRegistered",
  );
});

async function assertRejectsWith(fn: () => Promise<unknown>, includes: string): Promise<void> {
  try {
    await fn();
  } catch (error) {
    assertStringIncludes(error instanceof Error ? error.message : String(error), includes);
    return;
  }

  throw new Error(`Expected a rejection mentioning ${JSON.stringify(includes)}.`);
}

test("a shape's words resolve, because a banner is written against the data", async () => {
  const doc = await resolveDocument(
    published([
      {
        id: "banner",
        kind: "shape",
        height: 44,
        children: [paragraph("banner-line", "{{data.amount}} due by {{data.dueBy}}")],
      } as DocumentNode,
    ]),
    engineState({ amount: "£1,250.00", dueBy: "30 April" }),
  );

  const shape = doc.nodes[0];
  assertEquals(shape.kind, "shape");
  assertEquals(texts(shape.kind === "shape" ? shape.children : []), [
    "£1,250.00 due by 30 April",
  ]);
});

test("a shape produced by a loop gives each pass its own ids", async () => {
  const doc = await resolveDocument(
    published([
      {
        id: "banner",
        kind: "repeat",
        source: { scope: "data", path: "accounts" },
        as: "account",
        indexAs: "index",
        children: [
          {
            id: "banner",
            kind: "shape",
            children: [paragraph("banner-line", "{{ctx.account.label}}")],
          },
        ],
      } as DocumentNode,
    ]),
    engineState({ accounts: [{ label: "First" }, { label: "Second" }] }),
  );

  const lines = doc.nodes.flatMap((node) => node.kind === "shape" ? node.children : []);
  assertEquals(doc.nodes.map((node) => node.id), ["banner-0", "banner-1"]);
  assertEquals(lines.map((node) => node.id), ["banner-line-0", "banner-line-1"]);
  assertEquals(texts(lines), ["First", "Second"]);
});

function generatedChart(answer: unknown, own: Partial<GraphNode> = {}): Promise<DocumentModel> {
  const state = engineState({});
  state.aiClient = {
    generateParagraph: () => "",
    generateGraph: () => answer as AiGraphResult,
  };

  return resolveDocument(
    published([{
      id: "trend",
      kind: "graph",
      mode: "dynamic",
      graphType: "bar",
      prompts: [{ kind: "general", text: "Plot it." }],
      ...own,
    } as DocumentNode]),
    state,
  );
}

function graphOf(doc: DocumentModel): GraphNode {
  const node = doc.nodes[0];
  if (node?.kind !== "graph") throw new Error("expected one graph");
  return node;
}

test("a generated chart's numbers written as text are read as numbers", async () => {
  const doc = await generatedChart({
    data: {
      categories: [2025, "2026"],
      series: [{ label: "Paid", values: ["12", " 3.5 ", null], sizes: ["2"] }],
    },
  });

  assertEquals(graphOf(doc).data, {
    categories: ["2025", "2026"],
    series: [{ label: "Paid", values: [12, 3.5, null], sizes: [2] }],
  });
});

test("an absent caption, data or type keeps the node's own", async () => {
  const kept = graphOf(
    await generatedChart({}, {
      graphType: "line",
      caption: "Own caption",
      data: { series: [{ values: [9] }] },
    }),
  );

  assertEquals(kept.graphType, "line");
  assertEquals(kept.caption, "Own caption");
  assertEquals(kept.data, { series: [{ values: [9] }] });

  const partial = graphOf(
    await generatedChart(
      { caption: null, data: { series: [{ values: [1] }] } },
      { caption: "Own caption" },
    ),
  );

  assertEquals(partial.caption, "Own caption");
  assertEquals(partial.graphType, "bar");
  assertEquals(partial.data, { series: [{ values: [1] }] });
});

test("a well-formed chart answer comes back as it was given", async () => {
  const answer = {
    graphType: "pie",
    data: {
      categories: ["a", "b"],
      series: [{ label: "Share", values: [1, null], color: "FF0000" }],
    },
    caption: "Split",
  };
  const before = structuredClone(answer);

  const graph = graphOf(await generatedChart(answer));

  assertEquals(graph.graphType, "pie");
  assertEquals(graph.data, before.data as GraphNode["data"]);
  assertEquals(graph.caption, "Split");
  assertEquals(answer, before);
});

test("a generated chart that cannot be plotted names the node and the field", async () => {
  const cases: Array<[unknown, string]> = [
    [undefined, 'Dynamic graph "trend" got an answer it cannot plot: it is not an object.'],
    [{ graphType: "pyramid" }, 'graphType "pyramid" is not a chart type.'],
    [{ graphType: "toString" }, 'graphType "toString" is not a chart type.'],
    [{ data: [1, 2] }, "data is not an object."],
    [{ data: { series: "1,2,3" } }, "data.series is not an array."],
    [{ data: { series: [{ label: "x" }] } }, "data.series[0].values is not an array."],
    [
      { data: { series: [{ values: [1, "twelve"] }] } },
      'data.series[0].values[1] is "twelve", not a number.',
    ],
    [{ data: { series: [{ values: [""] }] } }, 'data.series[0].values[0] is "", not a number.'],
    [{ data: { categories: [{}], series: [] } }, "data.categories[0] is {}, not text."],
    [{ caption: 42 }, "caption is not text."],
  ];

  for (const [answer, expected] of cases) {
    await assertRejectsWith(() => generatedChart(answer), expected);
  }
});

test("a chart client that throws is not caught", async () => {
  const state = engineState({});
  state.aiClient = {
    generateParagraph: () => "",
    generateGraph: () => {
      throw new Error("model down");
    },
  };

  await assertRejectsWith(
    () =>
      resolveDocument(
        published([{
          id: "trend",
          kind: "graph",
          mode: "dynamic",
          graphType: "bar",
          prompts: [{ kind: "general", text: "Plot it." }],
        } as DocumentNode]),
        state,
      ),
    "model down",
  );
});

function generatedParagraph(answer: unknown): Promise<DocumentModel> {
  const state = engineState({});
  state.aiClient = { generateParagraph: () => answer as string };

  return resolveDocument(
    published([{
      id: "note",
      kind: "paragraph",
      mode: "dynamic",
      prompts: [{ kind: "general", text: "Write it." }],
    } as DocumentNode]),
    state,
  );
}

test("a generated paragraph's answer is trimmed and stripped of characters Word cannot open", async () => {
  const cases: Array<[string, string]> = [
    ["Clean answer", "Clean answer"],
    ["Hello\n", "Hello"],
    ["  Hello world  ", "Hello world"],
    ["Line one\nLine two\n", "Line one\nLine two"],
    ["Tab\there", "Tab\there"],
    ["Hello\u0001 world\u000b end", "Hello world end"],
    ["Hi\u0001\n", "Hi"],
    ["\u0000\uffff \n", ""],
    ["", ""],
  ];

  for (const [answer, expected] of cases) {
    assertEquals(texts((await generatedParagraph(answer)).nodes), [expected]);
  }
});

test("a generated paragraph that is not text names the node", async () => {
  const cases: Array<[unknown, string]> = [
    [{ text: "hi" }, 'Dynamic paragraph "note" got an answer that is not text: it returned an object.'],
    [undefined, "it returned undefined."],
    [null, "it returned null."],
    [42, "it returned a number."],
    [["hi"], "it returned an array."],
  ];

  for (const [answer, expected] of cases) {
    await assertRejectsWith(() => generatedParagraph(answer), expected);
  }
});

test("a paragraph client that throws is not caught", async () => {
  const state = engineState({});
  state.aiClient = {
    generateParagraph: () => {
      throw new Error("model down");
    },
  };

  await assertRejectsWith(
    () =>
      resolveDocument(
        published([{
          id: "note",
          kind: "paragraph",
          mode: "dynamic",
          prompts: [{ kind: "general", text: "Write it." }],
        } as DocumentNode]),
        state,
      ),
    "model down",
  );
});

test("a variant resolves per entry of a repeat", async () => {
  const doc = await resolveDocument(
    published([
      {
        id: "lines",
        kind: "repeat",
        source: { scope: "data", path: "lines" },
        as: "line",
        indexAs: "index",
        children: [paragraph("row", "{{ctx.line.desc}}", { variant: "{{ctx.line.status}}" })],
      },
    ]),
    engineState({ lines: [{ desc: "A", status: "overdue" }, { desc: "B", status: "paid" }] }),
  );

  assertEquals(doc.nodes.map((node) => node.variant), ["overdue", "paid"]);
  assertEquals(ids(doc.nodes), ["row-0", "row-1"]);
});

test("a variant with no token, and a node with none, come back as they were", async () => {
  const doc = await resolveDocument(
    published([paragraph("a", "x", { variant: "band" }), paragraph("b", "x")]),
    engineState({}),
  );

  assertEquals(doc.nodes[0].variant, "band");
  assertEquals("variant" in doc.nodes[1], false);
});

test("a variant in the furniture resolves too", async () => {
  const doc = await resolveDocument(
    { ...published([]), footer: [paragraph("f", "x", { variant: "{{data.tier}}" })] },
    engineState({ tier: "gold" }),
  );

  assertEquals(doc.footer?.[0].variant, "gold");
});

test("a container's own variant resolves", async () => {
  const doc = await resolveDocument(
    published([
      {
        id: "cell",
        kind: "tableCell",
        variant: "{{data.state}}",
        children: [paragraph("p", "x")],
      } as DocumentNode,
    ]),
    engineState({ state: "paid" }),
  );

  assertEquals(doc.nodes[0].variant, "paid");
});

const gt0 = {
  type: "compare",
  operator: "gt",
  left: { type: "ref", ref: { scope: "data", path: "balanceDue" } },
  right: { type: "literal", value: 0 },
} as const;

test("even-page furniture resolves like the rest of the furniture", async () => {
  const doc = await resolveDocument(
    {
      ...published([]),
      evenHeader: [paragraph("eh", "Ref {{data.reference}}")],
      evenFooter: [
        paragraph("ef", "Due {{derived.due}}", {
          derivers: [{
            name: "currencyLabel",
            output: "due",
            inputs: [{ type: "ref", ref: { scope: "data", path: "balanceDue" } }],
          }],
        }),
      ],
    },
    engineState({ reference: "INV-1", balanceDue: 240 }),
    {
      derivers: createDeriverRegistry({
        currencyLabel: ([amount]) => `£${Number(amount).toFixed(2)}`,
      }),
    },
  );

  assertEquals(texts(doc.evenHeader!), ["Ref INV-1"]);
  assertEquals(texts(doc.evenFooter!), ["Due £240.00"]);
});

test("absent even furniture stays absent, and an empty one stays empty", async () => {
  const absent = await resolveDocument(published([]), engineState({}));
  const empty = await resolveDocument({ ...published([]), evenHeader: [] }, engineState({}));

  assertEquals(absent.evenHeader, undefined);
  assertEquals(absent.evenFooter, undefined);
  assertEquals(empty.evenHeader, []);
});

test("a condition in even furniture is evaluated per document", async () => {
  const model = { ...published([]), evenFooter: [paragraph("ef", "Owing", { when: gt0 })] };
  const clear = await resolveDocument(model, engineState({ balanceDue: 0 }));
  const owing = await resolveDocument(model, engineState({ balanceDue: 240 }));

  assertEquals(clear.evenFooter, []);
  assertEquals(owing.evenFooter!.length, 1);
});

const summary: DocumentNode = {
  id: "summary",
  kind: "paragraph",
  mode: "dynamic",
  prompts: [{ kind: "general", text: "Summarise." }],
} as DocumentNode;

test("createRuntimeState seeds ctx.availableTokens the way a build does", async () => {
  const doc = published([paragraph("budget", "{{data.name}} has {{ctx.availableTokens}} tokens.")]);

  assertEquals(
    texts((await resolveDocument(doc, createRuntimeState({ data: { name: "Avery" } }))).nodes),
    ["Avery has 2000 tokens."],
  );
  assertEquals(
    texts((await resolveDocument(doc, createRuntimeState({ data: { name: "Avery" }, availableTokens: 500 }))).nodes),
    ["Avery has 500 tokens."],
  );
});

test("createRuntimeState hands the client every dynamic node, with the request data in reach", async () => {
  let seen: unknown;
  const state = createRuntimeState({
    data: { name: "Avery" },
    aiClient: {
      generateParagraph: (request) => {
        seen = request.state.ctx.name;
        return `[generated ${request.node.id}]`;
      },
    },
  });

  const doc = await resolveDocument(published([summary]), state);

  assertEquals(texts(doc.nodes), ["[generated summary]"]);
  assertEquals(seen, "Avery");
});

test("createRuntimeState without a client resolves a static document and refuses a dynamic one", async () => {
  const state = createRuntimeState({ data: { name: "Avery" } });

  assertEquals(
    texts((await resolveDocument(published([paragraph("greeting", "Hello {{data.name}}.")]), state)).nodes),
    ["Hello Avery."],
  );
  await assertRejects(() => resolveDocument(published([summary]), state), Error, "requires an aiClient");
});

test("createRuntimeState leaves the caller's data alone", async () => {
  const data = { name: "Avery" };
  const state = createRuntimeState({ data });

  assertEquals(data, { name: "Avery" });
  assertEquals(state.ctx, { name: "Avery", availableTokens: 2000 });
  assertEquals(state.derived, {});
  assertEquals(await state.dataProvider.get("name"), "Avery");
  assertEquals("locale" in state, false);
});

test("createRuntimeState's locale is canonical and wins over the document's", async () => {
  let locale: string | undefined;
  const client = {
    generateParagraph: (request: { state: RuntimeState }) => {
      locale = request.state.locale;
      return "x";
    },
  };
  const doc = { ...published([summary]), locale: "fr-FR" };

  await resolveDocument(doc, createRuntimeState({ data: {}, aiClient: client, locale: "de-de" }));
  assertEquals(locale, "de-DE");

  await resolveDocument(doc, createRuntimeState({ data: {}, aiClient: client }));
  assertEquals(locale, "fr-FR");

  await assertRejects(
    () => Promise.resolve().then(() => createRuntimeState({ data: {}, locale: "en_GB" })),
    Error,
    "BCP 47",
  );
});
