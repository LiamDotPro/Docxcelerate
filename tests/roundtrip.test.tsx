import { test } from "node:test";
import { assertEquals, assertRejects } from "./assert.ts";
import {
  type AiClient,
  buildProjectEngineDocument,
  buildProjectFinalDocument,
  createDeriverRegistry,
  createDocumentProjectArtifact,
  dataRef,
  defineDocumentProject,
  derive,
  type DocumentModel,
  type DocumentNode,
  InMemoryDataProvider,
  resolveDocument,
  type RuntimeState,
} from "docxcelerate";
import {
  branch,
  Cell,
  compare,
  dataPath,
  Document,
  Graph,
  Image,
  literal,
  Paragraph,
  Row,
  Section,
  Table,
  template,
  useAi,
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

/**
 * The loop, written as ordinary TypeScript.
 *
 * The engine walks a published loop and the build walks a real one, and this
 * test exists to prove the two agree — so the loop has to be the same source
 * either way, which is what a `.map()` in a component is.
 */
const Visits = () => {
  const [visits] = useState((data: TenancyData) => data.visits);

  return visits.map((visit) => (
    <Paragraph
      id="visit"
      derivers={[derive("money", { output: "visitLabel", inputs: [visit.cost] })]}
    >
      {visit.label} cost {"{{derived.visitLabel}}"}
    </Paragraph>
  ));
};

const documentTemplate = template<TenancyData>(
  <Document id="tenancy" title="Tenancy">
    <Section id="opening" title="Opening">
      <Greeting />
      <Balance />
      <Summary />
    </Section>
    <Section id="visits" title="Visits">
      <Visits />
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

// ---------------------------------------------------------------------------
// Generated charts and pictures, published
// ---------------------------------------------------------------------------

interface ReportData {
  city: string;
}

const Trend = () => {
  const [state] = useState((data: ReportData) => ({ city: data.city }));
  useAi({ ask: `Plot monthly visits for ${state.city}.`, placeholder: "A trend chart." });
  return <Graph id="trend" graphType="line" title={`Visits in ${state.city}`} width={300} />;
};

const Cover = () => {
  const [state] = useState((data: ReportData) => ({ city: data.city }));
  useAi({ ask: `Draw ${state.city} at dusk.`, placeholder: "A picture of the city." });
  return <Image id="cover" alt={`Picture of ${state.city}`} width={200} height={120} />;
};

const CellNote = () => {
  useAi({ ask: "One line on the figures.", placeholder: "A note." });
  return <Paragraph id="cell-note" />;
};

const reportProject = defineDocumentProject<ReportData>({
  id: "report",
  name: "Report",
  template: template<ReportData>(
    <Document id="report" title="Report">
      <Section id="figures" title="Figures">
        <Trend />
        <Cover />
      </Section>
      <Table id="grid" columns={[{ width: "auto" }]}>
        <Row>
          <Cell>
            <CellNote />
          </Cell>
        </Row>
      </Table>
    </Document>,
  ),
  previewData: { city: "Leeds" },
});

/** A model that draws, scripted so both paths can be compared. */
const drawingClient: AiClient = {
  generateParagraph: (request) => `[draft] ${request.prompt}`,
  generateGraph: (request) => ({
    graphType: "bar",
    data: { categories: ["Jan", "Feb"], series: [{ label: `${request.node.title} / ${request.prompt}`, values: [3, 5] }] },
    caption: "Visits by month.",
  }),
  generateImage: (request) => ({ path: "data:image/png;base64,iVBORw0KGgo=", alt: `${request.node.alt} / ${request.prompt}` }),
};

/** A model that answers every node with nothing. */
const silentClient: AiClient = {
  generateParagraph: () => "",
  generateGraph: () => ({}),
  generateImage: () => ({ path: "" }),
};

function reportState(data: ReportData, client: AiClient): RuntimeState {
  const ctx = { ...data } as Record<string, unknown>;

  return { ctx, derived: {}, dataProvider: new InMemoryDataProvider(ctx), aiClient: client };
}

function findNode(nodes: DocumentNode[], id: string): DocumentNode | undefined {
  for (const node of nodes) {
    if (node.id === id) {
      return node;
    }

    const found = "children" in node ? findNode(node.children, id) : undefined;
    if (found) {
      return found;
    }
  }

  return undefined;
}

/** What a stored model is: keys holding `undefined` are not part of it. */
const json = (value: unknown) => JSON.parse(JSON.stringify(value));

const leeds: ReportData = { city: "Leeds" };

test("a generated chart and image publish, then resolve to what a direct build writes", async () => {
  const published = await buildProjectEngineDocument(reportProject);
  const resolved = await resolveDocument(published, reportState(leeds, drawingClient));
  const direct = await buildProjectFinalDocument(reportProject, { data: leeds, aiClient: drawingClient });

  assertEquals(json(resolved.nodes), json(direct.nodes));

  const trend = findNode(resolved.nodes, "trend");
  assertEquals(
    trend?.kind === "graph" ? [trend.graphType, trend.caption, trend.width] : undefined,
    ["bar", "Visits by month.", 300],
  );

  const cover = findNode(resolved.nodes, "cover");
  assertEquals(
    cover?.kind === "image" ? [cover.path, cover.alt, cover.width, cover.height] : undefined,
    ["data:image/png;base64,iVBORw0KGgo=", "Picture of Leeds / GENERAL: Draw Leeds at dusk.", 200, 120],
  );
});

test("a published chart and image carry their prompts and nothing generated", async () => {
  const published = await buildProjectEngineDocument(reportProject);
  const trend = findNode(published.nodes, "trend");
  const cover = findNode(published.nodes, "cover");

  assertEquals(trend !== undefined && ("data" in trend || "caption" in trend), false);
  assertEquals(trend?.kind === "graph" && trend.prompts?.some((p) => p.text.includes("{{data.city}}")), true);
  assertEquals(cover !== undefined && "path" in cover, false);
  assertEquals(
    cover?.kind === "image" ? [cover.alt, cover.width, cover.height] : undefined,
    ["Picture of {{data.city}}", 200, 120],
  );
  assertEquals(JSON.stringify(published).includes("Leeds"), false);
});

test("a generated paragraph inside a table cell publishes without empty text", async () => {
  const note = findNode((await buildProjectEngineDocument(reportProject)).nodes, "cell-note");

  assertEquals(note?.kind, "paragraph");
  assertEquals(note !== undefined && "text" in note, false);
});

test("a project with a generated chart builds an artifact whose preview shows the placeholder", async () => {
  const artifact = await createDocumentProjectArtifact(reportProject, {
    builtAt: "2026-01-01T00:00:00.000Z",
  });

  assertEquals(json(artifact.engineDocument), json(await buildProjectEngineDocument(reportProject)));

  const previewTrend = findNode(artifact.previewDocument.nodes, "trend");
  assertEquals(previewTrend?.kind === "graph" ? previewTrend.placeholder : undefined, "A trend chart.");
  assertEquals(previewTrend !== undefined && "data" in previewTrend, false);
});

test("a model that answers a chart and image with nothing leaves them undrawn on both paths", async () => {
  const resolved = await resolveDocument(
    await buildProjectEngineDocument(reportProject),
    reportState(leeds, silentClient),
  );
  const direct = await buildProjectFinalDocument(reportProject, { data: leeds, aiClient: silentClient });

  assertEquals(json(resolved.nodes), json(direct.nodes));
  assertEquals("data" in json(findNode(resolved.nodes, "trend")), false);
});

test("an engine whose client cannot draw, or throws, fails the resolve naming why", async () => {
  const published = await buildProjectEngineDocument(reportProject);

  await assertRejects(
    () => resolveDocument(published, reportState(leeds, { generateParagraph: () => "x" })),
    Error,
    'Dynamic graph "trend" requires an aiClient.generateGraph method.',
  );
  await assertRejects(
    () =>
      resolveDocument(
        published,
        reportState(leeds, {
          ...drawingClient,
          generateGraph: () => {
            throw new Error("model down");
          },
        }),
      ),
    Error,
    "model down",
  );
});

test("a generated picture among a paragraph's words refuses to publish", async () => {
  const inlineProject = defineDocumentProject({
    id: "inline",
    name: "Inline",
    template: template(
      <Document id="inline" title="Inline">
        <Paragraph id="p">
          See <Image id="icon" generalPrompt="Draw an icon." placeholder="An icon." /> here.
        </Paragraph>
      </Document>,
    ),
    previewData: {},
  });

  await assertRejects(
    () => buildProjectEngineDocument(inlineProject),
    Error,
    'Paragraph "p" holds a generated picture, "icon"',
  );
});
