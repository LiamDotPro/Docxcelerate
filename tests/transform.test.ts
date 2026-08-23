import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import {
  assertCompiledSources,
  compiledMarker,
  findUncompiledSources,
  isCompiledSource,
  transformDocumentSource,
} from "docxcelerate/transform";

/**
 * What the compiler is allowed to touch, and what it must leave alone.
 *
 * An `if` in a component is a decision a document carries; an `if` in a helper
 * that picks a string is control flow with nothing to publish. Rewriting the
 * second would turn a string into a branch element, so the line between them is
 * load-bearing and every case below is a place it could be drawn wrong.
 */
function compile(source: string, fileName = "node.tsx") {
  return transformDocumentSource(source, { fileName });
}

/** The compiled body, without the import and stamp the transform prepends. */
function bodyOf(source: string, fileName = "node.tsx"): string {
  return compile(source, fileName).code
    .split("\n")
    .filter((line) => !line.startsWith("import ") && !line.includes(compiledMarker))
    .join("\n")
    .trim();
}

const component = (body: string) => `export const N = () => {\n${body}\n};`;

// ---------------------------------------------------------------------------
// The shapes a component writes
// ---------------------------------------------------------------------------

test("an early return becomes a branch with both arms", () => {
  const out = bodyOf(component(
    `  if (state.paid) { return <Paid />; }\n  return <Awaiting />;`,
  ));

  assertStringIncludes(out, "branch(__test(state.paid), () => <Paid />, () => <Awaiting />)");
});

test("an explicit else becomes the same branch", () => {
  const out = bodyOf(component(
    `  if (state.paid) { return <Paid />; } else { return <Awaiting />; }`,
  ));

  assertStringIncludes(out, "branch(__test(state.paid), () => <Paid />, () => <Awaiting />)");
});

test("a return without braces is compiled too", () => {
  const out = bodyOf(component(`  if (state.paid) return <Paid />;\n  return <Awaiting />;`));

  assertStringIncludes(out, "branch(__test(state.paid)");
});

test("negation is read off the syntax rather than coerced away", () => {
  const out = bodyOf(component(
    `  if (!state.scholarship) { return <None />; }\n  return <Some />;`,
  ));

  assertStringIncludes(out, "__not(__test(state.scholarship))");
});

test("a double negation nests rather than cancelling", () => {
  const out = bodyOf(component(`  if (!!state.x) { return <A />; }\n  return <B />;`));

  assertStringIncludes(out, "__not(__not(__test(state.x)))");
});

test("parentheses around a condition are seen through", () => {
  const out = bodyOf(component(`  if ((state.x)) { return <A />; }\n  return <B />;`));

  assertStringIncludes(out, "branch(__test(state.x)");
});

test("statements before the branch are kept", () => {
  const out = bodyOf(component(
    `  const [state] = useState((d) => d);\n  if (state.x) { return <A />; }\n  return <B />;`,
  ));

  assertStringIncludes(out, "const [state] = useState((d) => d);");
  assertStringIncludes(out, "branch(__test(state.x)");
});

test("an arm that falls off the end yields nothing rather than a node", () => {
  const out = bodyOf(component(`  if (state.hide) { return undefined; }\n  return <A />;`));

  assertStringIncludes(out, "() => undefined");
});

test("the transform reports how many branches it compiled", () => {
  const out = compile(`
    export const A = () => {
      if (state.x) { return <A />; }
      return <B />;
    };
    export const B = () => {
      if (state.y) { return <C />; }
      return <D />;
    };
  `);

  assertEquals(out.branches, 2);
});

// ---------------------------------------------------------------------------
// The shapes a component writes inside its own JSX
//
// A component deciding between two nodes in the middle of its markup has
// nowhere to put a statement, so it writes the decision as an expression. Both
// of these coerce a stand-in to a boolean exactly as an uncompiled `if` would,
// and until they compiled, nothing downstream could tell — which made them the
// worst of the three, not the least important.
// ---------------------------------------------------------------------------

test("a ternary between two nodes becomes a branch with both arms", () => {
  const out = bodyOf(component(`  return <S>{state.paid ? <Paid /> : <Awaiting />}</S>;`));

  assertStringIncludes(out, "branch(__test(state.paid), () => <Paid />, () => <Awaiting />)");
});

test("a ternary returned on its own compiles the same way", () => {
  const out = bodyOf(component(`  return state.paid ? <Paid /> : <Awaiting />;`));

  assertStringIncludes(out, "branch(__test(state.paid), () => <Paid />, () => <Awaiting />)");
});

test("`&&` before a node becomes a branch whose other arm is nothing", () => {
  const out = bodyOf(component(`  return <S>{state.overdue && <Notice />}</S>;`));

  assertStringIncludes(out, "branch(__test(state.overdue), () => <Notice />, () => undefined)");
});

test("a guard reads its operators off the syntax, like an if does", () => {
  const out = bodyOf(component(`  return <S>{state.n > 2 && <Notice />}</S>;`));

  assertStringIncludes(out, `branch(__compare(state.n, "gt", 2)`);
});

test("a chain of tests before a node joins into one condition", () => {
  const out = bodyOf(component(`  return <S>{state.a && state.b && <Notice />}</S>;`));

  assertStringIncludes(out, "branch(__and(__test(state.a), __test(state.b))");
});

test("a chained ternary nests as deep as it is written", () => {
  const out = compile(component(
    `  return <S>{state.a ? <A /> : state.b ? <B /> : <C />}</S>;`,
  ));

  assertEquals(out.branches, 2);
  assertStringIncludes(out.code, "() => branch(__test(state.b), () => <B />, () => <C />)");
});

test("a decision inside a mapped body is compiled too", () => {
  // The body of a `.map()` is walked once while publishing, so a ternary in it
  // coerces its stand-in exactly as one anywhere else would.
  const out = bodyOf(component(
    `  return <S>{data.rows.map((r) => r.paid ? <Paid /> : <Awaiting />)}</S>;`,
  ));

  assertStringIncludes(out, "branch(__test(r.paid), () => <Paid />, () => <Awaiting />)");
});

// ---------------------------------------------------------------------------
// Operators
// ---------------------------------------------------------------------------

const comparisons: Array<[string, string]> = [
  ["state.n === 1", `__compare(state.n, "eq", 1)`],
  ["state.n == 1", `__compare(state.n, "eq", 1)`],
  ["state.n !== 1", `__compare(state.n, "ne", 1)`],
  ["state.n != 1", `__compare(state.n, "ne", 1)`],
  ["state.n > 1", `__compare(state.n, "gt", 1)`],
  ["state.n >= 1", `__compare(state.n, "gte", 1)`],
  ["state.n < 1", `__compare(state.n, "lt", 1)`],
  ["state.n <= 1", `__compare(state.n, "lte", 1)`],
];

for (const [written, expected] of comparisons) {
  test(`\`${written}\` compiles to ${expected.slice(0, 24)}…`, () => {
    const out = bodyOf(component(`  if (${written}) { return <A />; }\n  return <B />;`));

    assertStringIncludes(out, expected);
  });
}

test("a comparison keeps the order it was written in", () => {
  // Swapping the sides would invert every ordering test in every document.
  const out = bodyOf(component(`  if (1 < state.n) { return <A />; }\n  return <B />;`));

  assertStringIncludes(out, `__compare(1, "lt", state.n)`);
});

test("&& joins two tests", () => {
  const out = bodyOf(component(
    `  if (state.a && state.b) { return <A />; }\n  return <B />;`,
  ));

  assertStringIncludes(out, "__and(__test(state.a), __test(state.b))");
});

test("|| joins two tests", () => {
  const out = bodyOf(component(
    `  if (state.a || state.b) { return <A />; }\n  return <B />;`,
  ));

  assertStringIncludes(out, "__or(__test(state.a), __test(state.b))");
});

test("a mixed condition keeps its operators nested as written", () => {
  const out = bodyOf(component(
    `  if (state.a && state.n > 2) { return <A />; }\n  return <B />;`,
  ));

  assertStringIncludes(out, `__and(__test(state.a), __compare(state.n, "gt", 2))`);
});

test("a negated comparison compiles as a negated comparison", () => {
  const out = bodyOf(component(
    `  if (!(state.n > 2)) { return <A />; }\n  return <B />;`,
  ));

  assertStringIncludes(out, `__not(__compare(state.n, "gt", 2))`);
});

// ---------------------------------------------------------------------------
// What must be left alone
// ---------------------------------------------------------------------------

test("an if that assigns rather than returns is left exactly as written", () => {
  const source = component(`  let n = 0;\n  if (state.x) { n = 1; }\n  return <A>{n}</A>;`);

  assertEquals(compile(source).changed, false);
  assertEquals(compile(source).code, source);
});

test("a helper picking a string is not a document decision", () => {
  // This is the case that broke a real document: `formatList` returning a
  // string became a branch element, and a paragraph was handed an element where
  // it expected text.
  const source = `
    function formatList(items: string[]): string {
      if (items.length <= 1) { return items[0] ?? ""; }
      return items.join(", ");
    }
  `;

  assertEquals(compile(source, "helpers.ts").changed, false);
});

test("a guard clause returning a number is left alone", () => {
  const source = `
    function total(n: number): number {
      if (n < 0) { return 0; }
      return n * 2;
    }
  `;

  assertEquals(compile(source, "helpers.ts").changed, false);
});

test("one arm yielding a node is enough to compile the decision", () => {
  const out = compile(component(`  if (state.hide) { return null; }\n  return <A />;`));

  assertEquals(out.branches, 1);
});

test("a ternary picking a string for a prop is a value, not a decision", () => {
  // The same line an `if` is held to. Compiling this would hand a paragraph a
  // branch element where it expected text; a value that varies per document is
  // what the derivers are for.
  const source = component(`  return <P text={state.paid ? "Paid" : "Due"} />;`);

  assertEquals(compile(source).changed, false);
  assertEquals(compile(source).code, source);
});

test("`&&` picking a string is left alone for the same reason", () => {
  const source = component(
    `  const label = state.paid && "Paid";\n  return <P text={label} />;`,
  );

  assertEquals(compile(source).changed, false);
});

for (const written of ["state.a || <A />", "state.a ?? <A />"]) {
  test(`\`${written}\` is not a way of writing a decision, so it is left alone`, () => {
    // Neither yields the right-hand side when the test holds, so neither says
    // what `&&` says. Whatever it is doing in a component, it is not this.
    const source = component(`  return <S>{${written}}</S>;`);

    assertEquals(compile(source).changed, false);
    assertEquals(compile(source).code, source);
  });
}

test("a file with no conditional at all is handed back untouched", () => {
  const source = component(`  return <A />;`);

  assertEquals(compile(source).changed, false);
  assertEquals(compile(source).code, source);
});

// ---------------------------------------------------------------------------
// The stamp
// ---------------------------------------------------------------------------

test("a compiled module carries the mark", () => {
  const out = compile(component(`  if (state.x) { return <A />; }\n  return <B />;`));

  assertEquals(isCompiledSource(out.code), true);
});

test("an untouched module carries no mark, because nothing compiled it", () => {
  const out = compile(component(`  return <A />;`));

  assertEquals(isCompiledSource(out.code), false);
});

test("the mark is exported, so a bundle can be asked whether it was compiled", () => {
  const out = compile(component(`  if (state.x) { return <A />; }\n  return <B />;`));

  assertStringIncludes(out.code, `export { ${compiledMarker} }`);
});

test("only the helpers a file actually used are imported", () => {
  const out = compile(component(`  if (state.x) { return <A />; }\n  return <B />;`));

  assertStringIncludes(out.code, "import { __test, branch }");
});

test("the helpers come from the package by default", () => {
  const out = compile(component(`  if (state.x) { return <A />; }\n  return <B />;`));

  assertStringIncludes(out.code, `from "docxcelerate/template"`);
});

test("the helper module can be pointed elsewhere", () => {
  const out = transformDocumentSource(
    component(`  if (state.x) { return <A />; }\n  return <B />;`),
    { fileName: "n.tsx", runtimeModule: "@acme/docs/template" },
  );

  assertStringIncludes(out.code, `from "@acme/docs/template"`);
});

// ---------------------------------------------------------------------------
// Syntax survives
// ---------------------------------------------------------------------------

test("JSX survives the round trip", () => {
  const out = bodyOf(component(
    `  if (state.x) { return <Paragraph id="a">Hello {state.name}.</Paragraph>; }\n` +
      `  return <Paragraph id="b" />;`,
  ));

  assertStringIncludes(out, `<Paragraph id="a">Hello {state.name}.</Paragraph>`);
  // The printer closes a childless tag without the space, which is the same JSX.
  assertStringIncludes(out, `<Paragraph id="b"/>`);
});

test("type annotations survive the round trip", () => {
  const out = compile(
    `export const N: Paragraph = () => {\n` +
      `  if (state.x) { return <A />; }\n  return <B />;\n};`,
  );

  assertStringIncludes(out.code, "export const N: Paragraph");
});

test("a .ts file is parsed as TypeScript rather than TSX", () => {
  // `<string>value` is a cast in .ts and an unclosed tag in .tsx, so parsing the
  // wrong way round would mangle it.
  const out = transformDocumentSource(
    `function f(v: unknown) { return v as string; }`,
    { fileName: "helpers.ts" },
  );

  assertEquals(out.changed, false);
});

// ---------------------------------------------------------------------------
// else if chains
// ---------------------------------------------------------------------------

test("an else if becomes a branch inside the arm it belongs to", () => {
  const out = bodyOf(component(
    `  if (s.a) { return <A />; } else if (s.b) { return <B />; } else { return <C />; }`,
  ));

  assertStringIncludes(
    out,
    "branch(__test(s.a), () => <A />, () => branch(__test(s.b), () => <B />, () => <C />))",
  );
});

test("an else if chain counts every decision in it", () => {
  const out = compile(component(
    `  if (s.a) { return <A />; } else if (s.b) { return <B />; } else { return <C />; }`,
  ));

  assertEquals(out.branches, 2);
});

test("an else if chain can fall through to a trailing return", () => {
  const out = bodyOf(component(
    `  if (s.a) { return <A />; } else if (s.b) { return <B />; }
  return <C />;`,
  ));

  assertStringIncludes(out, "() => branch(__test(s.b), () => <B />, () => <C />)");
});

test("a chain nests as deep as it is written", () => {
  const out = compile(component(
    `  if (s.a) return <A />;
  else if (s.b) return <B />;
  else if (s.c) return <C />;
  else return <D />;`,
  ));

  assertEquals(out.branches, 3);
});

test("a chain of string-returning guards is still left alone", () => {
  const source = `
    function pick(n: number): string {
      if (n < 0) { return "low"; } else if (n > 10) { return "high"; } else { return "mid"; }
    }
  `;

  assertEquals(compile(source, "helpers.ts").changed, false);
});

// ---------------------------------------------------------------------------
// Catching what was never compiled
// ---------------------------------------------------------------------------

const decides = component(`  if (s.a) { return <A />; }
  return <B />;`);
const decidesNothing = component(`  return <A />;`);

test("a file holding an uncompiled decision is found", () => {
  const found = findUncompiledSources([{ fileName: "fees.node.tsx", source: decides }]);

  assertEquals(found, [{ fileName: "fees.node.tsx", branches: 1 }]);
});

test("a file with nothing to decide is not reported", () => {
  assertEquals(findUncompiledSources([{ fileName: "a.tsx", source: decidesNothing }]), []);
});

test("a file that was already compiled is not reported again", () => {
  const compiled = compile(decides).code;

  assertEquals(findUncompiledSources([{ fileName: "a.tsx", source: compiled }]), []);
});

test("a helper that returns strings is not mistaken for an uncompiled decision", () => {
  const source = `
    function formatList(items: string[]): string {
      if (items.length <= 1) { return items[0] ?? ""; }
      return items.join(", ");
    }
  `;

  assertEquals(findUncompiledSources([{ fileName: "helpers.ts", source }]), []);
});

test("every uncompiled file is reported, not just the first", () => {
  const found = findUncompiledSources([
    { fileName: "a.tsx", source: decides },
    { fileName: "b.tsx", source: decidesNothing },
    { fileName: "c.tsx", source: decides },
  ]);

  assertEquals(found.map((entry) => entry.fileName), ["a.tsx", "c.tsx"]);
});

test("the assertion passes when everything was compiled", () => {
  assertCompiledSources([{ fileName: "a.tsx", source: compile(decides).code }]);
});

test("the assertion names the files and says which build to fix", () => {
  let message = "";

  try {
    assertCompiledSources([{ fileName: "fees.node.tsx", source: decides }]);
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }

  assertStringIncludes(message, "fees.node.tsx");
  assertStringIncludes(message, "takes one arm");
  assertStringIncludes(message, "docxcelerateEsbuildTransform()");
});
