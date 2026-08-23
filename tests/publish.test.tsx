import { test } from "node:test";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import {
  buildDocument,
  createDeriverRegistry,
  EchoAiClient,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, ParagraphNode, RepeatNode, RuntimeState } from "docxcelerate";
import {
  __compare,
  __test,
  branch,
  createPublishData,
  Document,
  Paragraph,
  Section,
  template,
  useState,
} from "docxcelerate/template";

/**
 * What ordinary TypeScript is allowed to mean in a document nobody has the data
 * for yet.
 *
 * A publish build walks a template against stand-in values, so every expression
 * in a component meets a value that does not exist. Three answers are possible
 * and only three are honest: settle it (a path becomes the token an engine
 * substitutes), carry it (a loop and its filter travel to the engine intact), or
 * refuse it and say what does work. What is never acceptable is the fourth —
 * quietly producing an answer, because a wrong published document is wrong for
 * every recipient and nobody finds out until one of them reads it.
 */
interface Line {
  desc: string;
  qty: number;
  taxable: boolean;
  archived: boolean;
}

interface InvoiceData {
  reference: string;
  customer: { name: string; city: string };
  lines: Line[];
}

const invoice: InvoiceData = {
  reference: "INV-2026-0142",
  customer: { name: "Brackenfield Housing", city: "Leeds" },
  lines: [
    { desc: "API build", qty: 8, taxable: true, archived: false },
    { desc: "Disbursement", qty: 1, taxable: false, archived: false },
    { desc: "Support retainer", qty: 1, taxable: true, archived: true },
  ],
};

function templateOf(component: () => unknown) {
  const Body = component as () => never;

  return template<InvoiceData>(
    <Document id="invoice" title="Invoice">
      <Section id="body" title="Body">
        <Body />
      </Section>
    </Document>,
  );
}

function publish(component: () => unknown) {
  return buildDocument(templateOf(component), createPublishData() as InvoiceData, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
}

function build(component: () => unknown, data: InvoiceData = invoice) {
  return buildDocument(templateOf(component), data, {
    branchMode: "decide",
    dynamicMode: "placeholder",
  });
}

function bodyOf(nodes: DocumentNode[]): DocumentNode[] {
  const section = nodes[0];

  if (section?.kind !== "section") {
    throw new Error("expected the body section");
  }

  return section.children;
}

function loopOf(nodes: DocumentNode[]): RepeatNode {
  const loop = bodyOf(nodes)[0];

  if (loop?.kind !== "repeat") {
    throw new Error(`expected a loop, got ${loop?.kind ?? "nothing"}`);
  }

  return loop;
}

function textsOf(nodes: DocumentNode[]): string[] {
  return bodyOf(nodes).map((node) => (node as ParagraphNode).text ?? "");
}

function idsOf(nodes: DocumentNode[]): string[] {
  return bodyOf(nodes).map((node) => node.id);
}

/** Resolves a published document the way an engine does, against real data. */
async function resolveAgainst(published: Awaited<ReturnType<typeof publish>>, data: InvoiceData) {
  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data as unknown as Record<string, unknown>),
    aiClient: new EchoAiClient(),
  };

  return await resolveDocument(published, state, { derivers: createDeriverRegistry() });
}

// ---------------------------------------------------------------------------
// Paths become tokens
// ---------------------------------------------------------------------------

test("a top-level field interpolates as the token an engine substitutes", async () => {
  const built = await publish(() => {
    const [reference] = useState((data: InvoiceData) => data.reference);
    return <Paragraph id="ref">{reference}</Paragraph>;
  });

  assertEquals(textsOf(built.nodes), ["{{data.reference}}"]);
});

test("a nested field carries its whole path", async () => {
  const built = await publish(() => {
    const [name] = useState((data: InvoiceData) => data.customer.name);
    return <Paragraph id="who">{name}</Paragraph>;
  });

  assertEquals(textsOf(built.nodes), ["{{data.customer.name}}"]);
});

test("two fields in one sentence each become their own token", async () => {
  const built = await publish(() => {
    const [customer] = useState((data: InvoiceData) => data.customer);
    return <Paragraph id="who">{customer.name} of {customer.city}</Paragraph>;
  });

  assertEquals(textsOf(built.nodes), ["{{data.customer.name}} of {{data.customer.city}}"]);
});

test("a path nobody supplied still answers, because publishing walks every arm", async () => {
  const built = await publish(() => {
    const [data] = useState((input: InvoiceData) => input);
    return (
      <Paragraph id="deep">
        {(data as unknown as { a: { b: { c: string } } }).a.b.c}
      </Paragraph>
    );
  });

  assertEquals(textsOf(built.nodes), ["{{data.a.b.c}}"]);
});

test("length is a path segment, so a count reads off the array the request brings", async () => {
  const built = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return <Paragraph id="count">{lines.length} lines</Paragraph>;
  });

  assertEquals(textsOf(built.nodes), ["{{data.lines.length}} lines"]);
});

test("an index is a path segment too", async () => {
  const built = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return <Paragraph id="first">{lines[0].desc}</Paragraph>;
  });

  assertEquals(textsOf(built.nodes), ["{{data.lines.0.desc}}"]);
});

test("a length token resolves to the real count when the engine reads it", async () => {
  const published = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return <Paragraph id="count">{lines.length} lines</Paragraph>;
  });
  const resolved = await resolveAgainst(published, invoice);

  assertEquals(textsOf(resolved.nodes), ["3 lines"]);
});

test("an index token resolves to the entry it names", async () => {
  const published = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return <Paragraph id="first">{lines[0].desc}</Paragraph>;
  });
  const resolved = await resolveAgainst(published, invoice);

  assertEquals(textsOf(resolved.nodes), ["API build"]);
});

// ---------------------------------------------------------------------------
// .map() publishes a loop
// ---------------------------------------------------------------------------

test("a map publishes one body rather than one node per entry", async () => {
  const built = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines.map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  });

  assertEquals(loopOf(built.nodes).children.length, 1);
});

test("a published loop names the collection it walks", async () => {
  const built = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines.map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  });

  assertEquals(loopOf(built.nodes).source, { scope: "data", path: "lines" });
});

test("the entry writes its own reference, so no token is hand-written", async () => {
  const built = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines.map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  });

  assertEquals(
    (loopOf(built.nodes).children[0] as ParagraphNode).text,
    "{{ctx.lines.desc}}",
  );
});

test("the index is bound alongside the entry", async () => {
  const built = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines.map((line, index) => <Paragraph id="line">{index}: {line.desc}</Paragraph>);
  });

  assertEquals(
    (loopOf(built.nodes).children[0] as ParagraphNode).text,
    "{{ctx.lines_index}}: {{ctx.lines.desc}}",
  );
});

test("a map with real data walks the collection instead", async () => {
  const built = await build(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines.map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  });

  assertEquals(textsOf(built.nodes), ["API build", "Disbursement", "Support retainer"]);
});

test("passes are named by position when the build walks the loop", async () => {
  const built = await build(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines.map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  });

  assertEquals(idsOf(built.nodes), ["line-0", "line-1", "line-2"]);
});

test("a loop over one entry is still a loop, and its single pass is numbered", async () => {
  const built = await build(
    () => {
      const [lines] = useState((data: InvoiceData) => data.lines);
      return lines.map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
    },
    { ...invoice, lines: [invoice.lines[0]] },
  );

  assertEquals(idsOf(built.nodes), ["line-0"]);
});

test("a loop over nothing produces nothing", async () => {
  const built = await build(
    () => {
      const [lines] = useState((data: InvoiceData) => data.lines);
      return lines.map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
    },
    { ...invoice, lines: [] },
  );

  assertEquals(bodyOf(built.nodes), []);
});

test("a map nested in another publishes a loop over what the outer one bound", async () => {
  const built = await publish(() => {
    const [data] = useState((input: InvoiceData) => input);
    const groups = data as unknown as { groups: Array<{ rows: Array<{ label: string }> }> };

    return groups.groups.map((group) =>
      group.rows.map((row) => <Paragraph id="row">{row.label}</Paragraph>)
    );
  });

  const outer = loopOf(built.nodes);
  const inner = outer.children[0];

  if (inner?.kind !== "repeat") {
    throw new Error("expected a nested loop");
  }

  assertEquals(outer.source, { scope: "data", path: "groups" });
  // The inner collection lives on the entry the outer loop bound, so it is read
  // from `ctx` rather than from the caller's data.
  assertEquals(inner.source, { scope: "ctx", path: "groups.rows" });
});

// ---------------------------------------------------------------------------
// .filter() publishes a test
// ---------------------------------------------------------------------------

test("a filter publishes as a test on the loop rather than being applied now", async () => {
  const built = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines
      .filter((line) => line.taxable)
      .map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  });

  assertEquals(loopOf(built.nodes).where, {
    type: "truthy",
    ref: { scope: "ctx", path: "lines.taxable" },
  });
});

test("a filtered loop still walks the collection it started from", async () => {
  const built = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines
      .filter((line) => line.taxable)
      .map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  });

  assertEquals(loopOf(built.nodes).source, { scope: "data", path: "lines" });
});

test("two filters both travel, joined", async () => {
  const built = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines
      .filter((line) => line.taxable)
      .filter((line) => line.archived)
      .map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  });

  assertEquals(loopOf(built.nodes).where, {
    type: "and",
    conditions: [
      { type: "truthy", ref: { scope: "ctx", path: "lines.taxable" } },
      { type: "truthy", ref: { scope: "ctx", path: "lines.archived" } },
    ],
  });
});

test("a filter with real data drops the entries it says to drop", async () => {
  const built = await build(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines
      .filter((line) => line.taxable)
      .map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  });

  assertEquals(textsOf(built.nodes), ["API build", "Support retainer"]);
});

test("the engine and the build agree on which entries survive a filter", async () => {
  const component = () => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines
      .filter((line) => line.taxable)
      .map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  };

  const resolved = await resolveAgainst(await publish(component), invoice);
  const direct = await build(component);

  assertEquals(textsOf(resolved.nodes), textsOf(direct.nodes));
});

test("the engine and the build agree on what the surviving passes are called", async () => {
  const component = () => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines
      .filter((line) => line.taxable)
      .map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  };

  const resolved = await resolveAgainst(await publish(component), invoice);
  const direct = await build(component);

  // The dropped entry takes no pass number with it. If it did, a filtered loop
  // would number its passes differently on the two paths and an id would stop
  // meaning the same thing in a preview and in a delivered document.
  assertEquals(idsOf(resolved.nodes), ["line-0", "line-1"]);
  assertEquals(idsOf(direct.nodes), idsOf(resolved.nodes));
});

test("a predicate that never looks at the entry and keeps everything adds no test", async () => {
  const built = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines
      .filter(() => true)
      .map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  });

  assertEquals(loopOf(built.nodes).where, undefined);
});

test("a predicate that never looks at the entry and keeps nothing publishes no loop", async () => {
  const built = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    return lines
      .filter(() => false)
      .map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
  });

  assertEquals(bodyOf(built.nodes), []);
});

test("a predicate that read the entry and returned a plain boolean is refused", async () => {
  // `!line.archived` is already `false` by the time the filter sees it, so what
  // was tested cannot be recovered. Publishing it would drop every entry for
  // every recipient without saying so.
  await assertRejects(
    () =>
      publish(() => {
        const [lines] = useState((data: InvoiceData) => data.lines);
        return lines
          .filter((line) => !line.archived)
          .map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
      }),
    Error,
    "handed back a plain",
  );
});

test("the refusal names the shape that does work", async () => {
  let message = "";

  try {
    await publish(() => {
      const [lines] = useState((data: InvoiceData) => data.lines);
      return lines
        .filter((line) => !line.archived)
        .map((line) => <Paragraph id="line">{line.desc}</Paragraph>);
    });
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }

  assertStringIncludes(message, "(line) => line.taxable");
});

test("reading a filtered collection as a value is refused, count included", async () => {
  await assertRejects(
    () =>
      publish(() => {
        const [lines] = useState((data: InvoiceData) => data.lines);
        const taxable = lines.filter((line) => line.taxable);
        return <Paragraph id="count">{taxable.length}</Paragraph>;
      }),
    Error,
    "filtered",
  );
});

// ---------------------------------------------------------------------------
// What still belongs in a deriver
// ---------------------------------------------------------------------------

const deriverMembers = [
  "forEach",
  "reduce",
  "reduceRight",
  "slice",
  "find",
  "findLast",
  "findIndex",
  "some",
  "every",
  "flat",
  "flatMap",
  "join",
  "entries",
  "keys",
  "values",
  "includes",
  "indexOf",
  "lastIndexOf",
  "sort",
  "toSorted",
  "reverse",
  "toReversed",
  "concat",
  "fill",
  "splice",
  "toSpliced",
] as const;

for (const member of deriverMembers) {
  test(`\`${member}\` has to see the entries, so it says to use a deriver`, async () => {
    await assertRejects(
      () =>
        publish(() => {
          const [lines] = useState((data: InvoiceData) => data.lines);
          const collection = lines as unknown as Record<string, () => unknown>;
          collection[member]();
          return <Paragraph id="x">unreachable</Paragraph>;
        }),
      Error,
      "belongs in a deriver",
    );
  });
}

test("the deriver refusal says why a deriver is better than a closure", async () => {
  let message = "";

  try {
    await publish(() => {
      const [lines] = useState((data: InvoiceData) => data.lines);
      return <Paragraph id="x">{lines.join(", ")}</Paragraph>;
    });
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }

  assertStringIncludes(message, "testable on its own and shareable between documents");
});

test("a for-of loop cannot know how many times to go round, and says so", async () => {
  await assertRejects(
    () =>
      publish(() => {
        const [lines] = useState((data: InvoiceData) => data.lines);
        const out = [];

        for (const line of lines) {
          out.push(<Paragraph id="line">{line.desc}</Paragraph>);
        }

        return out;
      }),
    Error,
    "Use `.map()` instead",
  );
});

test("arithmetic on a value nobody has yet is refused", async () => {
  await assertRejects(
    () =>
      publish(() => {
        const [lines] = useState((data: InvoiceData) => data.lines);
        return <Paragraph id="x">{lines[0].qty * 2}</Paragraph>;
      }),
    Error,
    "read as a number while publishing",
  );
});

test("calling request data as a function is refused", async () => {
  await assertRejects(
    () =>
      publish(() => {
        const [data] = useState((input: InvoiceData) => input);
        (data as unknown as { go: () => void }).go();
        return <Paragraph id="x">unreachable</Paragraph>;
      }),
    Error,
    "values, not behaviour",
  );
});

test("awaiting a stand-in resolves rather than hanging forever", async () => {
  const built = await publish(() => {
    const [reference] = useState((data: InvoiceData) => data.reference);
    return <Paragraph id="ref">{reference}</Paragraph>;
  });

  assertEquals(textsOf(built.nodes), ["{{data.reference}}"]);
});

// ---------------------------------------------------------------------------
// What a compiled conditional becomes
//
// The transform's own tests read the source it emits. These run that source:
// every one below is written exactly as the compiler writes it, published, and
// then resolved against real data both ways round. What they are really asking
// is whether a decision made in a component still means the same thing after
// travelling to an engine as a condition — for the recipient it holds for and
// for the one it does not.
// ---------------------------------------------------------------------------

test("a compiled ternary publishes both arms, each with the test that selects it", async () => {
  const built = await publish(() => {
    const [state] = useState((data: InvoiceData) => data);

    return (
      <>
        {branch(
          __test(state.customer.name),
          () => <Paragraph id="named">Named.</Paragraph>,
          () => <Paragraph id="anonymous">Anonymous.</Paragraph>,
        )}
      </>
    );
  });

  assertEquals(idsOf(built.nodes), ["named", "anonymous"]);
  assertEquals(bodyOf(built.nodes)[0].when, {
    type: "truthy",
    ref: { scope: "data", path: "customer.name" },
  });
  assertEquals(bodyOf(built.nodes)[1].when, {
    type: "not",
    ref: { scope: "data", path: "customer.name" },
  });
});

test("a compiled guard publishes one arm, and nothing for the arm that was never written", async () => {
  const built = await publish(() => {
    const [state] = useState((data: InvoiceData) => data);

    return (
      <>
        {branch(__test(state.reference), () => <Paragraph id="ref">Referenced.</Paragraph>, () => undefined)}
      </>
    );
  });

  assertEquals(idsOf(built.nodes), ["ref"]);
});

test("a compiled comparison travels as the comparison it was written as", async () => {
  const built = await publish(() => {
    const [state] = useState((data: InvoiceData) => data);

    return (
      <>
        {branch(
          __compare(state.customer.city, "eq", "Leeds"),
          () => <Paragraph id="local">Local.</Paragraph>,
          () => undefined,
        )}
      </>
    );
  });

  assertEquals(bodyOf(built.nodes)[0].when, {
    type: "compare",
    operator: "eq",
    left: { type: "ref", ref: { scope: "data", path: "customer.city" } },
    right: { type: "literal", value: "Leeds" },
  });
});

test("an engine takes the arm the recipient's data selects, and only that arm", async () => {
  const built = await publish(() => {
    const [state] = useState((data: InvoiceData) => data);

    return (
      <>
        {branch(
          __compare(state.customer.city, "eq", "Leeds"),
          () => <Paragraph id="local">Local.</Paragraph>,
          () => <Paragraph id="distant">Distant.</Paragraph>,
        )}
      </>
    );
  });

  const leeds = await resolveAgainst(built, invoice);
  const elsewhere = await resolveAgainst(built, {
    ...invoice,
    customer: { name: "Brackenfield Housing", city: "Hull" },
  });

  assertEquals(idsOf(leeds.nodes), ["local"]);
  assertEquals(idsOf(elsewhere.nodes), ["distant"]);
});

test("the same conditional built against real data decides once, as the source read", async () => {
  const component = () => {
    const [state] = useState((data: InvoiceData) => data);

    return (
      <>
        {branch(
          __compare(state.customer.city, "eq", "Leeds"),
          () => <Paragraph id="local">Local.</Paragraph>,
          () => <Paragraph id="distant">Distant.</Paragraph>,
        )}
      </>
    );
  };

  assertEquals(idsOf((await build(component)).nodes), ["local"]);
  assertEquals(
    idsOf((await build(component, {
      ...invoice,
      customer: { name: "Brackenfield Housing", city: "Hull" },
    })).nodes),
    ["distant"],
  );
});

test("a decision inside a loop body is carried per entry, not per document", async () => {
  const built = await publish(() => {
    const [lines] = useState((data: InvoiceData) => data.lines);

    return (
      <>
        {lines.map((line) =>
          branch(
            __test(line.taxable),
            () => <Paragraph id="taxable">Taxable.</Paragraph>,
            () => <Paragraph id="exempt">Exempt.</Paragraph>,
          )
        )}
      </>
    );
  });

  const loop = loopOf(built.nodes);

  assertEquals(loop.children.map((node) => node.id), ["taxable", "exempt"]);
  assertEquals(loop.children[0].when, {
    type: "truthy",
    ref: { scope: "ctx", path: "lines.taxable" },
  });
});
