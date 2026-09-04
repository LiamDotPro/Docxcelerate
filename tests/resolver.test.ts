import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import {
  createDeriverRegistry,
  type DocumentModel,
  type DocumentNode,
  InMemoryDataProvider,
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
  return {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider({ ...data }),
    aiClient: {
      generateParagraph: (request) => `[generated ${request.node.id}]`,
      generateImage: () => ({ path: "generated.png", alt: "generated" }),
      generateGraph: () => ({ data: { series: [{ values: [1] }] }, caption: "generated" }),
    },
  };
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
