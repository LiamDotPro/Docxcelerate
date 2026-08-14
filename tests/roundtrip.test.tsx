/** @jsxImportSource docxcelerate/template */
import { test } from "node:test";
import { assertEquals } from "./assert.ts";
import {
  buildProjectEngineDocument,
  buildProjectFinalDocument,
  createDeriverRegistry,
  ctxRef,
  dataRef,
  defineDocumentProject,
  derive,
  type DocumentModel,
  InMemoryDataProvider,
  resolveDocument,
  type RuntimeState,
} from "docxcelerate";
import {
  branch,
  compare,
  dataPath,
  Document,
  literal,
  Paragraph,
  Repeat,
  Section,
  template,
  useSetPlaceholders,
  useSetPrompts,
  useState,
} from "docxcelerate/template";

/**
 * The property that makes publishing worth anything.
 *
 * A document built here with real data, and the published document resolved by
 * an engine against that same data, are the same document. If the two ever
 * disagree, what an author reviews in the preview is not what a recipient
 * receives — which is the failure the whole publish path exists to avoid.
 */
interface TenancyData {
  name: string;
  city: string;
  balanceDue: number;
  visits: Array<{ label: string; cost: number }>;
}

const Greeting: Paragraph = () => {
  const [state] = useState((data: TenancyData) => ({ name: data.name, city: data.city }));

  return (
    <Paragraph id="greeting">
      Dear {state.name} of {state.city},
    </Paragraph>
  );
};

const Balance: Paragraph = () => (
  <Paragraph
    id="balance"
    derivers={[derive("money", { output: "balanceLabel", inputs: [dataRef("balanceDue")] })]}
  >
    Your balance is {"{{derived.balanceLabel}}"}.
  </Paragraph>
);

const Summary: Paragraph = () => {
  const [state] = useState((data: TenancyData) => ({ city: data.city }));

  useSetPrompts({ generalPrompt: `Summarise the account for ${state.city}.` });
  useSetPlaceholders("A summary.");

  return <Paragraph id="summary" />;
};

const documentTemplate = template<TenancyData>(
  <Document id="tenancy" title="Tenancy">
    <Section id="opening" title="Opening">
      <Greeting />
      <Balance />
      <Summary />
    </Section>
    <Section id="visits" title="Visits">
      <Repeat over="visits" as="visit">
        <Paragraph
          id="visit"
          derivers={[derive("money", { output: "visitLabel", inputs: [ctxRef("visit.cost")] })]}
        >
          {"{{ctx.visit.label}}"} cost {"{{derived.visitLabel}}"}
        </Paragraph>
      </Repeat>
    </Section>
    {branch(
      compare({ type: "ref", ref: dataPath("balanceDue") }, "gt", literal(0)),
      () => <Paragraph id="chase">Please settle the balance.</Paragraph>,
      () => <Paragraph id="thanks">Thank you, nothing is outstanding.</Paragraph>,
    )}
  </Document>,
);

const derivers = {
  money: ([amount]: unknown[]) => `£${Number(amount ?? 0).toFixed(2)}`,
};

const project = defineDocumentProject<TenancyData>({
  id: "tenancy",
  name: "Tenancy",
  template: documentTemplate,
  derivers,
  previewData: {
    name: "Avery",
    city: "Leeds",
    balanceDue: 240,
    visits: [{ label: "First", cost: 10 }, { label: "Second", cost: 25 }],
  },
});

/** Answers deterministically, so the two paths can be compared at all. */
const aiClient = {
  generateParagraph: (request: { prompt: string }) => `[draft] ${request.prompt}`,
};

function engineState(data: TenancyData): RuntimeState {
  const ctx = { ...data } as Record<string, unknown>;
  ctx.availableTokens = 2_000;

  return {
    ctx,
    derived: {},
    dataProvider: new InMemoryDataProvider(ctx),
    aiClient,
  };
}

async function throughEngine(data: TenancyData): Promise<DocumentModel> {
  // Exactly what a workspace publishes, then exactly what an engine does to it.
  const publishedDocument = await buildProjectEngineDocument(project);

  return await resolveDocument(publishedDocument, engineState(data), {
    derivers: createDeriverRegistry(derivers),
  });
}

function directly(data: TenancyData): Promise<DocumentModel> {
  return buildProjectFinalDocument(project, { data, aiClient });
}

const owing: TenancyData = {
  name: "Avery",
  city: "Leeds",
  balanceDue: 240,
  visits: [{ label: "First", cost: 10 }, { label: "Second", cost: 25 }],
};

const settled: TenancyData = {
  name: "Rowan",
  city: "Cork",
  balanceDue: 0,
  visits: [{ label: "Only", cost: 5 }],
};

/** Node identity and content, flattened — what a reader of the document sees. */
function shape(doc: DocumentModel): string[] {
  const walk = (nodes: DocumentModel["nodes"]): string[] =>
    nodes.flatMap((node) => {
      if (node.kind === "section" || node.kind === "repeat") {
        return [`${node.kind} ${node.id}`, ...walk(node.children)];
      }

      if (node.kind === "paragraph") {
        return [`${node.id} :: ${node.text ?? ""}`];
      }

      return [`${node.kind} ${node.id}`];
    });

  return walk(doc.nodes);
}

test("publishing then resolving matches building directly, with a balance owing", async () => {
  assertEquals(shape(await throughEngine(owing)), shape(await directly(owing)));
});

test("publishing then resolving matches building directly, with nothing owing", async () => {
  assertEquals(shape(await throughEngine(settled)), shape(await directly(settled)));
});

test("one published document serves recipients whose branches differ", async () => {
  const chased = shape(await throughEngine(owing));
  const thanked = shape(await throughEngine(settled));

  assertEquals(chased.some((line) => line.startsWith("chase ::")), true);
  assertEquals(chased.some((line) => line.startsWith("thanks ::")), false);
  assertEquals(thanked.some((line) => line.startsWith("thanks ::")), true);
  assertEquals(thanked.some((line) => line.startsWith("chase ::")), false);
});

test("the published document really is unresolved, not a build in disguise", async () => {
  const publishedDocument = await buildProjectEngineDocument(project);
  const json = JSON.stringify(publishedDocument);

  // Tokens survive rather than being filled in with the preview data.
  assertEquals(json.includes("{{data.name}}"), true);
  assertEquals(json.includes("Avery"), false);
  // Both arms are present, each under a condition.
  assertEquals(json.includes('"id":"chase"'), true);
  assertEquals(json.includes('"id":"thanks"'), true);
  // The loop is still a loop.
  assertEquals(json.includes('"kind":"repeat"'), true);
  // And the deriver is recorded for the engine to run.
  assertEquals(json.includes('"name":"money"'), true);
});

test("a loop resolves to as many passes as the recipient has entries", async () => {
  const many = { ...owing, visits: [1, 2, 3, 4].map((n) => ({ label: `V${n}`, cost: n })) };

  assertEquals(
    shape(await throughEngine(many)).filter((line) => line.startsWith("visit-")),
    [
      "visit-0 :: V1 cost £1.00",
      "visit-1 :: V2 cost £2.00",
      "visit-2 :: V3 cost £3.00",
      "visit-3 :: V4 cost £4.00",
    ],
  );
  assertEquals(shape(await throughEngine(many)), shape(await directly(many)));
});

test("an empty collection drops its passes on both paths", async () => {
  const none = { ...settled, visits: [] };

  assertEquals(shape(await throughEngine(none)).filter((l) => l.startsWith("visit-")), []);
  assertEquals(shape(await throughEngine(none)), shape(await directly(none)));
});
