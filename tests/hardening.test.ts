import { test } from "node:test";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import {
  createDeriverBundle,
  createDeriverRegistry,
  createDeriverRegistryFromBundle,
  evaluateCondition,
  InMemoryDataProvider,
  invertCondition,
  resolveDocument,
} from "docxcelerate";
import type {
  ComparisonOperator,
  Condition,
  DocumentModel,
  RuntimeState,
} from "docxcelerate";
import { expr } from "docxcelerate/template";
import { transformDocumentSource } from "docxcelerate/transform";
import { documentXml } from "./docx.ts";

/**
 * The edges, and what happens at them.
 *
 * Everything here is reached from a published document rather than from code
 * somebody is reading: a path in a paragraph, an output key on a deriver, a
 * condition an engine evaluates. A published document is data — stored, copied
 * between systems, rendered by an engine that may serve more than one tenant —
 * so the interesting question about each of these is not whether it works on a
 * good day, but what it does when handed something nobody meant.
 */
function stateWith(data: Record<string, unknown> = {}): RuntimeState {
  return {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: { generateParagraph: () => "" },
  };
}

// ---------------------------------------------------------------------------
// Paths: reading
// ---------------------------------------------------------------------------

const reads: Array<[string, Record<string, unknown>, string, unknown]> = [
  ["a plain field", { name: "Avery" }, "name", "Avery"],
  ["a nested field", { a: { b: { c: 1 } } }, "a.b.c", 1],
  ["an array entry", { lines: [{ q: 2 }] }, "lines.0.q", 2],
  ["an array length", { lines: [1, 2, 3] }, "lines.length", 3],
  ["a string length", { name: "Avery" }, "name.length", 5],
  ["a field holding null", { a: null }, "a", null],
  ["a field holding false", { a: false }, "a", false],
  ["a field holding zero", { a: 0 }, "a", 0],
  ["a field holding an empty string", { a: "" }, "a", ""],
  ["a path off the end of the data", { a: 1 }, "a.b.c", undefined],
  ["a path through a null", { a: null }, "a.b", undefined],
  ["a path through a number", { a: 1 }, "a.b", undefined],
  ["a path through a boolean", { a: true }, "a.b", undefined],
  ["a field nobody supplied", {}, "missing", undefined],
  ["an index past the end", { lines: [1] }, "lines.5", undefined],
];

for (const [label, source, path, expected] of reads) {
  test(`reading ${label}`, () => {
    assertEquals(new InMemoryDataProvider(source).get(path), expected);
  });
}

const forbidden = ["__proto__", "constructor", "prototype"];

for (const segment of forbidden) {
  test(`reading \`${segment}\` gives nothing, because it is not data`, () => {
    assertEquals(new InMemoryDataProvider({ a: 1 }).get(segment), undefined);
  });

  test(`reading through \`${segment}\` gives nothing`, () => {
    assertEquals(new InMemoryDataProvider({ a: 1 }).get(`${segment}.anything`), undefined);
  });

  test(`a deriver writing to \`${segment}\` is refused`, async () => {
    // A deriver output key is the one route a published document has to a
    // write, and it travels inside the document itself.
    await assertRejects(
      () =>
        createDeriverRegistry({ pick: () => 1 }).run(
          { name: "pick", output: `${segment}.polluted`, inputs: [] },
          stateWith(),
        ),
      Error,
      "not a path a document may write to",
    );
  });
}

test("an inherited property is not readable as data", () => {
  const provider = new InMemoryDataProvider({});

  assertEquals(provider.get("toString"), undefined);
  assertEquals(provider.get("hasOwnProperty"), undefined);
});

test("a refused write leaves every other object alone", async () => {
  try {
    await createDeriverRegistry({ pick: () => 1 }).run(
      { name: "pick", output: "__proto__.polluted", inputs: [] },
      stateWith(),
    );
  } catch {
    // The refusal is the point; what matters is what did not happen.
  }

  assertEquals(({} as Record<string, unknown>).polluted, undefined);
});

test("a deriver cannot pollute the prototype through its output key", async () => {
  const registry = createDeriverRegistry({ pick: () => "yes" });

  await assertRejects(
    () => registry.run({ name: "pick", output: "__proto__.owned", inputs: [] }, stateWith()),
    Error,
    "not a path a document may write to",
  );

  assertEquals(({} as Record<string, unknown>).owned, undefined);
});

// ---------------------------------------------------------------------------
// Where a deriver result lands
// ---------------------------------------------------------------------------

/** Runs one deriver and hands back the `derived` scope it wrote into. */
async function derivedAfter(
  output: string,
  value: unknown,
  before: Record<string, unknown> = {},
): Promise<Record<string, unknown>> {
  const state = stateWith();
  Object.assign(state.derived, before);

  await createDeriverRegistry({ pick: () => value }).run(
    { name: "pick", output, inputs: [] },
    state,
  );

  return state.derived;
}

test("a plain output key holds the result", async () => {
  assertEquals(await derivedAfter("total", 12), { total: 12 });
});

test("a dotted output key builds the objects on the way", async () => {
  assertEquals(await derivedAfter("totals.due", 12), { totals: { due: 12 } });
});

test("a dotted output key over a plain value replaces it rather than failing", async () => {
  assertEquals(await derivedAfter("totals.due", 12, { totals: 1 }), { totals: { due: 12 } });
});

test("a second deriver keeps what the first left alongside it", async () => {
  assertEquals(
    await derivedAfter("totals.due", 12, { totals: { vat: 3 } }),
    { totals: { vat: 3, due: 12 } },
  );
});

test("a deriver that produced nothing still claims its key", async () => {
  // An absent key and a key holding nothing read differently downstream: the
  // first looks like a deriver that never ran.
  assertEquals(Object.hasOwn(await derivedAfter("a", undefined), "a"), true);
});

// ---------------------------------------------------------------------------
// Values carried into a document
// ---------------------------------------------------------------------------

const carried: Array<[string, unknown]> = [
  ["a string", "Avery"],
  ["an empty string", ""],
  ["zero", 0],
  ["a negative number", -12.5],
  ["false", false],
  ["true", true],
];

for (const [label, value] of carried) {
  test(`${label} is carried as a literal`, () => {
    assertEquals(expr(value), { type: "literal", value: value as string });
  });
}

const refused: Array<[string, unknown]> = [
  ["null", null],
  ["undefined", undefined],
  ["an object", { a: 1 }],
  ["an array", [1, 2]],
  ["a function", () => 1],
];

for (const [label, value] of refused) {
  test(`${label} cannot be carried into a document`, () => {
    let message = "";

    try {
      expr(value);
    } catch (error) {
      message = error instanceof Error ? error.message : String(error);
    }

    assertStringIncludes(message, "cannot be carried into a document");
  });
}

for (const [label, value] of [["NaN", NaN], ["Infinity", Infinity], ["-Infinity", -Infinity]] as const) {
  test(`${label} is refused, because JSON would turn it into null`, () => {
    let message = "";

    try {
      expr(value);
    } catch (error) {
      message = error instanceof Error ? error.message : String(error);
    }

    // Silently becoming null is the failure worth stopping: the number would
    // reach the engine as an absence and the document would be wrong.
    assertStringIncludes(message, "arrive at the engine as null");
  });
}

test("negative zero is carried as a number rather than refused", () => {
  assertEquals(expr(-0), { type: "literal", value: -0 });
});

// ---------------------------------------------------------------------------
// Conditions
// ---------------------------------------------------------------------------

const comparisons: Array<[unknown, ComparisonOperator, unknown, boolean]> = [
  [1, "eq", 1, true],
  [1, "eq", 2, false],
  [1, "ne", 2, true],
  [1, "gt", 0, true],
  [0, "gt", 1, false],
  [1, "gte", 1, true],
  [1, "lt", 2, true],
  [2, "lt", 1, false],
  [1, "lte", 1, true],
  ["a", "lt", "b", true],
  ["b", "lt", "a", false],
  ["a", "eq", "a", true],
  [true, "eq", true, true],
  [true, "ne", false, true],
  ["", "eq", "", true],
  [0, "eq", 0, true],
];

for (const [left, operator, right, expected] of comparisons) {
  test(`${JSON.stringify(left)} ${operator} ${JSON.stringify(right)} is ${expected}`, async () => {
    const condition: Condition = {
      type: "compare",
      operator,
      left: { type: "literal", value: left as string },
      right: { type: "literal", value: right as string },
    };

    assertEquals(await evaluateCondition(condition, stateWith()), expected);
  });
}

const unorderable: Array<[unknown, ComparisonOperator, unknown]> = [
  [1, "gt", "a"],
  ["a", "gt", 1],
  [true, "gt", false],
];

for (const [left, operator, right] of unorderable) {
  test(`ordering ${JSON.stringify(left)} against ${JSON.stringify(right)} is false, not a guess`, async () => {
    // Two values of different kinds have no order between them. Coercing one
    // would answer a question nobody asked, in a document nobody rereads.
    const condition: Condition = {
      type: "compare",
      operator,
      left: { type: "literal", value: left as string },
      right: { type: "literal", value: right as string },
    };

    assertEquals(await evaluateCondition(condition, stateWith()), false);
  });
}

const truthiness: Array<[string, unknown, boolean]> = [
  ["a non-empty string", "yes", true],
  ["an empty string", "", false],
  ["a positive number", 1, true],
  ["zero", 0, false],
  ["true", true, true],
  ["false", false, false],
  ["null", null, false],
  ["a missing field", undefined, false],
  ["an empty array", [], true],
  ["an object", {}, true],
];

for (const [label, value, expected] of truthiness) {
  test(`a truthy test on ${label} is ${expected}`, async () => {
    const condition: Condition = { type: "truthy", ref: { scope: "data", path: "v" } };

    assertEquals(await evaluateCondition(condition, stateWith({ v: value })), expected);
  });

  test(`a not test on ${label} is ${!expected}`, async () => {
    const condition: Condition = { type: "not", ref: { scope: "data", path: "v" } };

    assertEquals(await evaluateCondition(condition, stateWith({ v: value })), !expected);
  });
}

test("an absent condition holds, so a node with no test is always included", async () => {
  assertEquals(await evaluateCondition(undefined, stateWith()), true);
});

test("and with nothing in it holds", async () => {
  assertEquals(await evaluateCondition({ type: "and", conditions: [] }, stateWith()), true);
});

test("or with nothing in it does not hold", async () => {
  assertEquals(await evaluateCondition({ type: "or", conditions: [] }, stateWith()), false);
});

test("and needs every part", async () => {
  const state = stateWith({ a: true, b: false });
  const both: Condition = {
    type: "and",
    conditions: [
      { type: "truthy", ref: { scope: "data", path: "a" } },
      { type: "truthy", ref: { scope: "data", path: "b" } },
    ],
  };

  assertEquals(await evaluateCondition(both, state), false);
});

test("or needs one part", async () => {
  const state = stateWith({ a: true, b: false });
  const either: Condition = {
    type: "or",
    conditions: [
      { type: "truthy", ref: { scope: "data", path: "a" } },
      { type: "truthy", ref: { scope: "data", path: "b" } },
    ],
  };

  assertEquals(await evaluateCondition(either, state), true);
});

test("conditions nest to any depth", async () => {
  const state = stateWith({ a: true, b: false, c: true });
  const nested: Condition = {
    type: "and",
    conditions: [
      { type: "truthy", ref: { scope: "data", path: "a" } },
      {
        type: "or",
        conditions: [
          { type: "truthy", ref: { scope: "data", path: "b" } },
          { type: "negate", condition: { type: "not", ref: { scope: "data", path: "c" } } },
        ],
      },
    ],
  };

  assertEquals(await evaluateCondition(nested, state), true);
});

const invertible: Array<[string, Condition]> = [
  ["truthy", { type: "truthy", ref: { scope: "data", path: "v" } }],
  ["not", { type: "not", ref: { scope: "data", path: "v" } }],
  ["negate", { type: "negate", condition: { type: "truthy", ref: { scope: "data", path: "v" } } }],
  ["eq", {
    type: "compare",
    operator: "eq",
    left: { type: "ref", ref: { scope: "data", path: "v" } },
    right: { type: "literal", value: 1 },
  }],
  ["gt", {
    type: "compare",
    operator: "gt",
    left: { type: "ref", ref: { scope: "data", path: "v" } },
    right: { type: "literal", value: 1 },
  }],
  ["and", {
    type: "and",
    conditions: [
      { type: "truthy", ref: { scope: "data", path: "v" } },
      { type: "truthy", ref: { scope: "data", path: "w" } },
    ],
  }],
  ["or", {
    type: "or",
    conditions: [
      { type: "truthy", ref: { scope: "data", path: "v" } },
      { type: "truthy", ref: { scope: "data", path: "w" } },
    ],
  }],
];

for (const [label, condition] of invertible) {
  for (const [v, w] of [[1, 1], [1, 0], [0, 1], [0, 0]]) {
    test(`inverting a ${label} condition answers the opposite (v=${v}, w=${w})`, async () => {
      // Both arms of a compiled `if` are published, one under the condition and
      // one under its inverse. If those ever agreed, a document would print both
      // arms or neither.
      const state = stateWith({ v, w });
      const straight = await evaluateCondition(condition, state);
      const inverted = await evaluateCondition(invertCondition(condition), state);

      assertEquals(inverted, !straight);
    });
  }
}

// ---------------------------------------------------------------------------
// The transform, on input nobody meant
// ---------------------------------------------------------------------------

const untouched: Array<[string, string]> = [
  ["an empty file", ""],
  ["only comments", "// nothing here"],
  ["only whitespace", "   \n\n  "],
  ["a syntax error", "export const N = () => { if (a) { return <A />; }"],
  ["a switch", "export const N = () => { switch (s.x) { case 1: return <A />; } return <B />; };"],
  ["a try/catch", "export const N = () => { try { return <A />; } catch { return <B />; } };"],
  ["an if with no return at all", "export const N = () => { if (s.x) { go(); } return <A />; };"],
  ["a loop", "export const N = () => { for (const x of xs) { use(x); } return <A />; };"],
];

for (const [label, source] of untouched) {
  test(`the transform leaves ${label} alone`, () => {
    const result = transformDocumentSource(source, { fileName: "n.tsx" });

    assertEquals(result.changed, false);
    assertEquals(result.code, source);
  });
}

const compiled: Array<[string, string]> = [
  ["an async component", "export const N = async () => { if (s.x) { return <A />; } return <B />; };"],
  ["a generic arrow", "export const N = <T,>(v: T) => { if (s.x) { return <A />; } return <B />; };"],
  ["a component nested in another", "export const N = () => { const I = () => { if (s.y) { return <C />; } return <D />; }; return <I />; };"],
  ["a function declaration", "export function N() { if (s.x) { return <A />; } return <B />; }"],
  ["a decision inside a map", "export const N = () => xs.map((x) => { if (x.a) { return <A />; } return <B />; });"],
];

for (const [label, source] of compiled) {
  test(`the transform compiles the decision in ${label}`, () => {
    assertEquals(transformDocumentSource(source, { fileName: "n.tsx" }).branches, 1);
  });
}

test("compiling twice does not compile the branch twice", () => {
  const once = transformDocumentSource(
    "export const N = () => { if (s.x) { return <A />; } return <B />; };",
    { fileName: "n.tsx" },
  );
  const twice = transformDocumentSource(once.code, { fileName: "n.tsx" });

  // The first pass leaves a `branch(...)` call, which is not an `if`, so there
  // is nothing left to rewrite. A transform that compounded would double every
  // decision each time a file was touched.
  assertEquals(twice.branches, 0);
});

// ---------------------------------------------------------------------------
// Deriver bundles
// ---------------------------------------------------------------------------

test("an arrow deriver survives the round trip", async () => {
  const bundle = createDeriverBundle({ double: ([a]: unknown[]) => Number(a) * 2 });
  const registry = await createDeriverRegistryFromBundle(bundle);

  assertEquals(
    await registry.run(
      { name: "double", output: "o", inputs: [{ type: "literal", value: 21 }] },
      stateWith(),
    ),
    42,
  );
});

test("a method-shorthand deriver survives the round trip", async () => {
  const bundle = createDeriverBundle({
    count(inputs: unknown[]) {
      return inputs.length;
    },
  });
  const registry = await createDeriverRegistryFromBundle(bundle);

  assertEquals(
    await registry.run(
      { name: "count", output: "o", inputs: [{ type: "literal", value: 1 }] },
      stateWith(),
    ),
    1,
  );
});

test("an async deriver survives the round trip", async () => {
  const bundle = createDeriverBundle({
    // deno-lint-ignore require-await -- an async deriver is what the round trip is testing.
    later: async ([a]: unknown[]) => Number(a) + 1,
  });
  const registry = await createDeriverRegistryFromBundle(bundle);

  assertEquals(
    await registry.run(
      { name: "later", output: "o", inputs: [{ type: "literal", value: 1 }] },
      stateWith(),
    ),
    2,
  );
});

test("a deriver built from native code cannot be bundled, and says so", () => {
  let message = "";

  try {
    createDeriverBundle({ biggest: Math.max as unknown as (i: unknown[]) => unknown });
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }

  assertStringIncludes(message, "cannot be bundled from native code");
});

test("a deriver that closes over a variable fails where it runs, not where it was written", async () => {
  // The documented limit of bundling by source: a bundle carries the function
  // and nothing around it. This is pinned rather than fixed so that the failure
  // stays a loud one — the alternative is a deriver that quietly returns NaN.
  const outside = 5;
  const bundle = createDeriverBundle({ closes: ([a]: unknown[]) => Number(a) + outside });
  const registry = await createDeriverRegistryFromBundle(bundle);

  await assertRejects(
    () =>
      registry.run(
        { name: "closes", output: "o", inputs: [{ type: "literal", value: 1 }] },
        stateWith(),
      ),
    Error,
    "outside is not defined",
  );
});

test("bundling nothing produces nothing rather than an empty bundle", () => {
  assertEquals(createDeriverBundle(undefined), undefined);
  assertEquals(createDeriverBundle({}, { names: [] }), undefined);
});

test("running a deriver nobody registered names it", async () => {
  await assertRejects(
    () => createDeriverRegistry().run({ name: "nope", output: "o", inputs: [] }, stateWith()),
    Error,
    "Unknown deriver: nope",
  );
});

// ---------------------------------------------------------------------------
// Rendering things nobody meant
// ---------------------------------------------------------------------------

function documentOf(text: string, title = "Doc"): DocumentModel {
  return {
    schemaVersion: "docxcelerate.letter/v0",
    id: "d",
    title,
    nodes: [{ id: "p", kind: "paragraph", mode: "static", text }],
  };
}

const injections: Array<[string, string]> = [
  ["a script tag", "<script>alert(1)</script>"],
  ["an image handler", `<img src=x onerror="alert(1)">`],
  ["a closing title", "</title><script>alert(1)</script>"],
  ["an attribute break", `" onmouseover="alert(1)`],
  ["an ampersand", "Tom & Jerry"],
];

for (const [label, text] of injections) {
  test(`${label} in the text is packed as text, not as markup`, async () => {
    const xml = await documentXml(documentOf(text));

    // The characters that would make the text into XML are escaped, so what
    // was written reaches the page as the words somebody typed.
    assertEquals(xml.includes(text), false);
  });

  test(`${label} in the title is packed as text, not as markup`, async () => {
    const xml = await documentXml(documentOf("safe", text));

    assertEquals(xml.includes(text), false);
  });
}

test("a document with no nodes still packs a file", async () => {
  const xml = await documentXml({
    schemaVersion: "docxcelerate.letter/v0",
    id: "d",
    title: "Empty",
    nodes: [],
  });

  assertStringIncludes(xml, "<w:body>");
});

test("a paragraph with no text packs without printing undefined", async () => {
  const xml = await documentXml({
    schemaVersion: "docxcelerate.letter/v0",
    id: "d",
    title: "Doc",
    nodes: [{ id: "p", kind: "paragraph", mode: "static" }],
  });

  assertEquals(xml.includes("undefined"), false);
});

// ---------------------------------------------------------------------------
// Resolving a published document
// ---------------------------------------------------------------------------

test("a loop over a field that is not a collection resolves to nothing", async () => {
  const doc: DocumentModel = {
    schemaVersion: "docxcelerate.letter/v0",
    id: "d",
    title: "Doc",
    nodes: [{
      id: "loop",
      kind: "repeat",
      source: { scope: "data", path: "name" },
      as: "item",
      indexAs: "index",
      children: [{ id: "line", kind: "paragraph", mode: "static", text: "x" }],
    }],
  };

  const resolved = await resolveDocument(doc, stateWith({ name: "Avery" }));

  assertEquals(resolved.nodes, []);
});

test("a loop over a field nobody supplied resolves to nothing", async () => {
  const doc: DocumentModel = {
    schemaVersion: "docxcelerate.letter/v0",
    id: "d",
    title: "Doc",
    nodes: [{
      id: "loop",
      kind: "repeat",
      source: { scope: "data", path: "missing" },
      as: "item",
      indexAs: "index",
      children: [{ id: "line", kind: "paragraph", mode: "static", text: "x" }],
    }],
  };

  assertEquals((await resolveDocument(doc, stateWith({}))).nodes, []);
});

test("a node whose condition fails is left out", async () => {
  const doc: DocumentModel = {
    schemaVersion: "docxcelerate.letter/v0",
    id: "d",
    title: "Doc",
    nodes: [{
      id: "p",
      kind: "paragraph",
      mode: "static",
      text: "Only when owing.",
      when: { type: "truthy", ref: { scope: "data", path: "owing" } },
    }],
  };

  assertEquals((await resolveDocument(doc, stateWith({ owing: false }))).nodes, []);
  assertEquals((await resolveDocument(doc, stateWith({ owing: true }))).nodes.length, 1);
});
