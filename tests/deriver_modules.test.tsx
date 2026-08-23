import { test } from "node:test";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import {
  buildDocument,
  createDeriverRegistry,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, ParagraphNode, RuntimeState } from "docxcelerate";
import {
  createPublishData,
  deriver,
  Document,
  Paragraph,
  Section,
  template,
  useDeriver,
  useState,
} from "docxcelerate/template";

/**
 * A deriver as a module, and what a component gets back from one.
 *
 * A deriver is the part of a document that runs where the data is, which makes
 * it the part most worth keeping out of a component: written as a module it is
 * an ordinary function with an ordinary test beside it, and two documents can
 * share it. What the hook has to hide is that the answer depends on which build
 * is asking — run it, stand in for it, or publish the invocation — without the
 * component ever learning which of the three happened.
 */
interface Line {
  desc: string;
  amount: number;
}

interface InvoiceData {
  reference: string;
  lines: Line[];
}

const invoice: InvoiceData = {
  reference: "INV-2026-0142",
  lines: [{ desc: "API build", amount: 100 }, { desc: "Support", amount: 50 }],
};

/** Cheap by construction, so it runs everywhere including a preview. */
const invoiceTotals = deriver({
  name: "invoiceTotals",
  run: (lines: Line[]) => {
    const subtotal = lines.reduce((total, line) => total + line.amount, 0);

    return { subtotal, vat: subtotal * 0.2, due: subtotal * 1.2 };
  },
});

/** Counts its own runs, so a preview can be held to not running it. */
let qrRuns = 0;

const paymentQr = deriver({
  name: "paymentQr",
  run: async (reference: string) => {
    qrRuns += 1;
    return `qr:${reference}`;
  },
  placeholder: "[scan-to-pay code]",
});

function templateOf(component: () => unknown) {
  const Body = component as () => never;

  return template<InvoiceData>(
    <Document title="Invoice">
      <Section title="Body">
        <Body />
      </Section>
    </Document>,
  );
}

function textOf(nodes: DocumentNode[]): string {
  const section = nodes[0];

  if (section?.kind !== "section") throw new Error("expected the body section");

  return (section.children[0] as ParagraphNode).text ?? "";
}

function nodeOf(nodes: DocumentNode[]): DocumentNode {
  const section = nodes[0];

  if (section?.kind !== "section") throw new Error("expected the body section");

  return section.children[0];
}

const Total = async () => {
  const [lines] = useState((data: InvoiceData) => data.lines);
  const totals = await useDeriver(invoiceTotals, [lines]);

  return <Paragraph id="total">Due {totals.due}.</Paragraph>;
};

function build(component: () => unknown, options = {}) {
  return buildDocument(templateOf(component), invoice, {
    dynamicMode: "placeholder",
    ...options,
  });
}

function publish(component: () => unknown) {
  return buildDocument(templateOf(component), createPublishData() as InvoiceData, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
}

// ---------------------------------------------------------------------------
// Declaring one
// ---------------------------------------------------------------------------

test("a deriver takes its arguments positionally, the way a test calls it", () => {
  // The engine calls every deriver with an array. Unpacking it here is what lets
  // the function be written and tested as the function it is.
  assertEquals(invoiceTotals.run([invoice.lines], {} as RuntimeState), {
    subtotal: 150,
    vat: 30,
    due: 180,
  });
});

test("a deriver carries the name a document calls it by", () => {
  assertEquals(invoiceTotals.name, "invoiceTotals");
});

test("a deriver with no name is refused, because the name is what travels", () => {
  let message = "";

  try {
    deriver({ name: "", run: () => 1 });
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }

  assertStringIncludes(message, "A deriver needs a name");
});

test("a costly deriver says so by declaring what stands in for it", () => {
  assertEquals(paymentQr.placeholder, "[scan-to-pay code]");
  assertEquals(invoiceTotals.placeholder, undefined);
});

// ---------------------------------------------------------------------------
// With data in hand
// ---------------------------------------------------------------------------

test("a deriver runs and hands back what it produced", async () => {
  assertEquals(textOf((await build(Total)).nodes), "Due 180.");
});

test("what it produced is also readable as a token", async () => {
  const Probe = async () => {
    const [lines] = useState((data: InvoiceData) => data.lines);
    await useDeriver(invoiceTotals, [lines]);

    return <Paragraph id="total">Due {"{{derived.invoiceTotals.due}}"}.</Paragraph>;
  };

  assertEquals(textOf((await build(Probe)).nodes), "Due 180.");
});

test("nothing is recorded on the node, because the answer is already in it", async () => {
  assertEquals((nodeOf((await build(Total)).nodes) as ParagraphNode).derivers, undefined);
});

test("two calls to one deriver get their own keys rather than overwriting", async () => {
  const Probe = async () => {
    await Promise.all([useDeriver(invoiceTotals, [[{ desc: "a", amount: 10 }]]), useDeriver(
      invoiceTotals,
      [[{ desc: "b", amount: 20 }]],
    )]);

    return (
      <Paragraph id="both">
        {"{{derived.invoiceTotals.subtotal}}"} and {"{{derived.invoiceTotals-2.subtotal}}"}
      </Paragraph>
    );
  };

  assertEquals(textOf((await build(Probe)).nodes), "10 and 20");
});

test("an async deriver is awaited rather than leaking a promise into the text", async () => {
  qrRuns = 0;

  const Probe = async () => {
    const [reference] = useState((data: InvoiceData) => data.reference);
    const code = await useDeriver(paymentQr, [reference]);

    return <Paragraph id="qr">{code}</Paragraph>;
  };

  assertEquals(textOf((await build(Probe)).nodes), "qr:INV-2026-0142");
});

// ---------------------------------------------------------------------------
// While previewing
// ---------------------------------------------------------------------------

test("a costly deriver stands in rather than running", async () => {
  qrRuns = 0;

  const Probe = async () => {
    const [reference] = useState((data: InvoiceData) => data.reference);
    const code = await useDeriver(paymentQr, [reference]);

    return <Paragraph id="qr">{code}</Paragraph>;
  };

  const built = await build(Probe, { deriverMode: "placeholder" });

  assertEquals(qrRuns, 0);
  assertEquals(textOf(built.nodes), "[scan-to-pay code]");
});

test("a cheap deriver still runs while previewing", async () => {
  const built = await build(Total, { deriverMode: "placeholder" });

  assertEquals(textOf(built.nodes), "Due 180.");
});

test("what stood in is readable as a token too, so the shape does not change", async () => {
  const Probe = async () => {
    const [reference] = useState((data: InvoiceData) => data.reference);
    await useDeriver(paymentQr, [reference]);

    return <Paragraph id="qr">{"{{derived.paymentQr}}"}</Paragraph>;
  };

  assertEquals(
    textOf((await build(Probe, { deriverMode: "placeholder" })).nodes),
    "[scan-to-pay code]",
  );
});

// ---------------------------------------------------------------------------
// While publishing
// ---------------------------------------------------------------------------

test("publishing does not run the deriver", async () => {
  qrRuns = 0;

  await publish(async () => {
    const [reference] = useState((data: InvoiceData) => data.reference);
    const code = await useDeriver(paymentQr, [reference]);

    return <Paragraph id="qr">{code}</Paragraph>;
  });

  assertEquals(qrRuns, 0);
});

test("what comes back knows where the real value will be", async () => {
  const built = await publish(Total);

  // Nobody wrote this token. The stand-in the hook handed back knows the key its
  // deriver writes to, so the reference writes itself.
  assertEquals(textOf(built.nodes), "Due {{derived.invoiceTotals.due}}.");
});

test("the invocation travels on the node for the engine to run", async () => {
  const built = await publish(Total);

  assertEquals((nodeOf(built.nodes) as ParagraphNode).derivers, [{
    name: "invoiceTotals",
    output: "invoiceTotals",
    inputs: [{ type: "ref", ref: { scope: "data", path: "lines" } }],
  }]);
});

test("an input that is a plain value travels as a literal", async () => {
  const rate = deriver({ name: "rate", run: (amount: number) => amount * 2 });
  const built = await publish(async () => {
    await useDeriver(rate, [21]);

    return <Paragraph id="t">{"{{derived.rate}}"}</Paragraph>;
  });

  assertEquals((nodeOf(built.nodes) as ParagraphNode).derivers, [{
    name: "rate",
    output: "rate",
    inputs: [{ type: "literal", value: 21 }],
  }]);
});

test("an input that is a whole collection cannot be published as one value", async () => {
  // A published input is a single value, because that is what a reference or a
  // literal can be. A collection reaches the engine by being pointed at — which
  // is what happens when the input is request data rather than something the
  // component built for itself.
  await assertRejects(
    () =>
      publish(async () => {
        await useDeriver(invoiceTotals, [[{ desc: "a", amount: 1 }]]);

        return <Paragraph id="t">x</Paragraph>;
      }),
    Error,
    "cannot be carried into a document",
  );
});

test("two calls publish two invocations with their own keys", async () => {
  const built = await publish(async () => {
    const [reference] = useState((data: InvoiceData) => data.reference);
    await Promise.all([useDeriver(paymentQr, [reference]), useDeriver(paymentQr, [reference])]);

    return <Paragraph id="both">{"{{derived.paymentQr}}"}</Paragraph>;
  });

  assertEquals(
    (nodeOf(built.nodes) as ParagraphNode).derivers?.map((invocation) => invocation.output),
    ["paymentQr", "paymentQr-2"],
  );
});

// ---------------------------------------------------------------------------
// The property that makes publishing worth anything
// ---------------------------------------------------------------------------

test("publishing then resolving matches building directly", async () => {
  const published = await publish(Total);
  const state: RuntimeState = {
    ctx: { ...invoice },
    derived: {},
    dataProvider: new InMemoryDataProvider(invoice as unknown as Record<string, unknown>),
    aiClient: { generateParagraph: () => "" },
  };
  const resolved = await resolveDocument(published, state, {
    derivers: createDeriverRegistry([invoiceTotals]),
  });
  const direct = await build(Total);

  assertEquals(textOf(resolved.nodes), textOf(direct.nodes));
});

test("an engine without the deriver fails on the document rather than guessing", async () => {
  const published = await publish(Total);
  const state: RuntimeState = {
    ctx: { ...invoice },
    derived: {},
    dataProvider: new InMemoryDataProvider(invoice as unknown as Record<string, unknown>),
    aiClient: { generateParagraph: () => "" },
  };

  await assertRejects(
    () => resolveDocument(published, state, { derivers: createDeriverRegistry() }),
    Error,
    "Unknown deriver: invoiceTotals",
  );
});

test("a deriver module registers straight into a registry", async () => {
  const registry = createDeriverRegistry([invoiceTotals, paymentQr]);

  assertEquals(registry.has("invoiceTotals"), true);
  assertEquals(registry.standsInForPreview("paymentQr"), true);
  assertEquals(registry.standsInForPreview("invoiceTotals"), false);
});

test("calling the hook outside a component says so", () => {
  let message = "";

  try {
    useDeriver(invoiceTotals, [[]]);
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }

  assertStringIncludes(message, "useDeriver");
});
