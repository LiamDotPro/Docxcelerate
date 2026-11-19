import { test } from "node:test";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, Section, template } from "docxcelerate/template";
import { COMPONENTS } from "docxcelerate/registry";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { PreferentialOriginStatement } = await loadRegistryModule<
  typeof import("../registry/components/preferential-origin-statement/preferential-origin-statement.node.tsx")
>("components/preferential-origin-statement/preferential-origin-statement.node.tsx");
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/preferential-origin-statement.deriver.ts")
>("components/derivers/preferential-origin-statement.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(
  createDeriverBundle({ preferentialOriginStatement: mod.run }),
);

type Obj = Record<string, unknown>;
interface Out {
  outcome: string;
  title: string;
  period: string | null;
  declaration: string;
  criteriaLine: string | null;
  exceptions: string | null;
  placeAndDate: string;
  exporterName: string;
  warning: string;
  originatingMinor: number;
  originatingEurCents: number;
}

const preview = COMPONENTS.find((c) => c.id === "preferential-origin-statement")!.previewData as {
  origin: Obj;
  consignment: Obj;
};

const baseOrigin: Obj = {
  agreement: "eu-uk-tca",
  originCountry: "UK",
  eurRate: 1.17,
  place: "Leeds",
  date: "2026-11-30",
  exporterName: "Acme Ltd",
};
const line = (ref: string, valueMinor: number, more: Obj = {}) => ({ ref, valueMinor, originating: true, criterion: "A", ...more });
const lines = (...l: Obj[]) => ({ currency: "GBP", lines: l });
const call = (o: Obj, c: Obj): Out => mod.run([{ ...baseOrigin, ...o }, c] as never, {} as RuntimeState) as unknown as Out;
const withheld = (o: Out) => assertEquals(o.outcome, "withheld");

test("the preview data makes a full TCA statement", async () => {
  const o = mod.run([preview.origin, preview.consignment] as never, {} as RuntimeState) as unknown as Out;
  assertEquals(o.outcome, "statement");
  assertEquals(o.originatingMinor, 486020125);
  assertEquals(o.originatingEurCents, Math.round(486020125 * 1.1712));
  assertEquals(o.originatingEurCents, 569226770);
  assertStringIncludes(o.declaration, "(Exporter Reference No GB123456789000)");
  assertStringIncludes(o.declaration, "of UK preferential origin.");
  assertEquals(o.criteriaLine, "Origin criteria: A (lines 10, 20); C, D (line 40); E (line 50)");
  assertEquals(o.exceptions, "Not of UK preferential origin: lines 30, 60.");
  assertEquals(o.period, "Period: from 1 December 2026 to 30 November 2027");
  assertEquals(o.placeAndDate, "Newcastle upon Tyne, Tyne and Wear, 30 November 2026");
  // The same through a shipped bundle.
  const via = await derivers.run(
    { name: "preferentialOriginStatement", output: "x", inputs: [{ type: "literal", value: preview.origin }, { type: "literal", value: preview.consignment }] } as never,
    { ctx: {}, derived: {}, dataProvider: new InMemoryDataProvider({}), aiClient: { generateParagraph: () => "" } } as RuntimeState,
  );
  assertEquals(via, o);
});

test("the threshold counts originating lines, not the invoice total", () => {
  const o = call({}, lines(line("1", 500000), line("2", 5000000, { originating: false })));
  assertEquals(o.outcome, "statement");
  assertEquals(o.originatingEurCents, 585000);
  assertEquals(o.declaration.includes("("), false);
});

test("exactly EUR 6,000.00 is not over; one cent more is", () => {
  const eur = (v: number) => call({ originCountry: "EU" }, { currency: "EUR", lines: [line("1", v)] });
  assertEquals(eur(600000).outcome, "statement");
  const w = eur(600001);
  withheld(w);
  assertStringIncludes(w.warning, "€6,000.01");
  assertStringIncludes(w.warning, "REX number");
});

test("the converted value is rounded once", () => {
  // 5128.20 GBP at 1.17 is 599,999.4 cents; one more minor unit is 600,000.57.
  assertEquals(call({}, lines(line("1", 512820))).originatingEurCents, 599999);
  assertEquals(call({}, lines(line("1", 512820))).outcome, "statement");
  const w = call({}, lines(line("1", 512821)));
  assertEquals([w.outcome, w.originatingEurCents], ["withheld", 600001]);
  assertEquals(call({ eurRate: 0.5 }, lines(line("1", 1))).originatingEurCents, 1);
});

test("a UK exporter over the threshold is told EORI, an EU one REX", () => {
  const w = call({}, lines(line("1", 99999999)));
  withheld(w);
  assertEquals(w.title, "Statement on origin: not made");
  assertEquals(w.declaration, "");
  assertEquals([w.period, w.criteriaLine, w.exceptions], [null, null, null]);
  assertEquals(w.placeAndDate, "Leeds, 30 November 2026");
  assertEquals(w.exporterName, "Acme Ltd");
  assertStringIncludes(w.warning, "EORI number");
  assertStringIncludes(call({ originCountry: "EU" }, lines(line("1", 99999999))).warning, "REX number");
});

test("a reference below the threshold is printed", () => {
  const o = call({ exporterReference: " GB1 " }, lines(line("1", 1000)));
  assertStringIncludes(o.declaration, "document (Exporter Reference No GB1) declares");
});

test("JPY has no minor units", () => {
  const o = call({ eurRate: 0.0062 }, { currency: "JPY", lines: [line("1", 800000)] });
  assertEquals([o.outcome, o.originatingEurCents], ["statement", 496000]);
});

test("a zero-value originating line still makes a statement", () => {
  const o = call({}, lines(line("1", 0)));
  assertEquals([o.outcome, o.originatingEurCents], ["statement", 0]);
});

test("nothing originating, no lines, or missing lines is none", async () => {
  for (const c of [lines(line("1", 100, { originating: false })), lines(), { currency: "GBP" }]) {
    const o = call({}, c);
    assertEquals([o.outcome, o.title, o.declaration, o.warning, o.period, o.exceptions], ["none", "", "", "", null, null]);
    assertEquals((await built(c)).length, 0);
  }
  // Validation still runs for a consignment that prints nothing.
  assertThrows(() => call({ place: "" }, lines()), "origin.place");
});

test("originating missing means not originating and is an exception", () => {
  const o = call({}, lines(line("1", 100), { ref: "2", valueMinor: 5 }));
  assertEquals(o.exceptions, "Not of UK preferential origin: line 2.");
  assertEquals(o.originatingMinor, 100);
});

test("TCA criteria", () => {
  assertEquals(call({}, lines(line("1", 100))).criteriaLine, "Origin criteria: A");
  assertEquals(call({}, lines(line("1", 100, { criterion: "c,d" }))).criteriaLine, "Origin criteria: C, D");
  assertEquals(call({}, lines(line("1", 100, { criterion: "dc" }))).criteriaLine, "Origin criteria: C, D");
  const w = call({}, lines(line("1", 100), line("7", 100, { criterion: "Z" })));
  withheld(w);
  assertStringIncludes(w.warning, "line 7");
  assertStringIncludes(w.warning, "A to E");
  assertStringIncludes(w.warning, '"Z"');
});

test("the first line without a criterion is the one named", () => {
  const w = call({}, lines(line("1", 1), line("2", 1, { criterion: undefined }), line("3", 1, { criterion: "" })));
  assertStringIncludes(w.warning, "line 2 is claimed as originating but has no origin criterion");
  assertEquals(w.warning.includes("line 3"), false);
});

test("GSP", () => {
  const g = (more: Obj, ls: Obj[]) =>
    call({ agreement: "eu-gsp", originCountry: "Bangladesh", exporterReference: "BDREX1", ...more }, lines(...ls));
  const o = g({}, [line("1", 100, { criterion: "p" }), line("2", 100, { criterion: "W", hsHeading: "6109" })]);
  assertEquals(o.criteriaLine, null);
  assertEquals(o.declaration.startsWith("The exporter Acme Ltd (Number of Registered Exporter BDREX1) of the products"), true);
  assertStringIncludes(o.declaration, "Generalised System of Preferences of the European Union");
  assertEquals(o.declaration.endsWith("origin criterion met is P (line 1); W 6109 (line 2)."), true);
  withheld(g({}, [line("1", 1, { criterion: "W" })]));
  const short = g({}, [line("1", 1, { criterion: "W", hsHeading: "61" })]);
  withheld(short);
  assertStringIncludes(short.warning, "HS heading");
  assertStringIncludes(g({}, [line("1", 1, { criterion: "A" })]).warning, "P or W");
});

test("CETA ignores criteria", () => {
  const o = call({ agreement: "ceta", originCountry: "EU", exporterReference: "X9" }, lines(line("1", 1, { criterion: "Z" }), line("2", 1, { criterion: undefined })));
  assertEquals([o.title, o.criteriaLine], ["Origin declaration", null]);
  assertStringIncludes(o.declaration, "(customs or competent governmental authorisation No X9)");
  assertStringIncludes(o.declaration, "of EU preferential origin.");
});

test("EU-Japan and UK DCTS", () => {
  const j = call({ agreement: "eu-japan", originCountry: "EU" }, lines(line("1", 1)));
  assertEquals(j.criteriaLine, "Origin criteria: A");
  assertStringIncludes(j.declaration, "The exporter of the products covered by this document declares");
  assertThrows(() => call({ agreement: "eu-japan", originCountry: "UK" }, lines(line("1", 1))), "origin.originCountry");
  const d = call({ agreement: "uk-dcts", originCountry: "Kenya" }, lines(line("1", 1, { criterion: "P" })));
  assertStringIncludes(d.declaration, "Developing Countries Trading Scheme of the United Kingdom");
  assertThrows(() => call({ agreement: "uk-dcts", originCountry: "UK" }, lines(line("1", 1, { criterion: "P" }))), "origin.originCountry");
});

test("a period is at most 12 months, both ends inclusive", () => {
  const p = (from: string, to: string) => call({ period: { from, to } }, lines(line("1", 100)));
  assertEquals(p("2026-03-01", "2027-02-28").outcome, "statement");
  const w = p("2026-03-01", "2027-03-01");
  withheld(w);
  assertStringIncludes(w.warning, "longer than the 12 months");
  assertEquals(p("2024-02-29", "2025-02-28").outcome, "statement");
  withheld(p("2024-02-29", "2025-03-01"));
  assertThrows(() => p("2026-03-01", "2026-02-01"), "origin.period.to");
  assertThrows(() => call({ period: { from: "2026-03-01" } }, lines(line("1", 1))), "origin.period.to");
  assertThrows(() => call({ period: { to: "2026-03-01" } }, lines(line("1", 1))), "origin.period.from");
});

test("withheld checks run period, then criterion, then reference", () => {
  const all = lines(line("1", 99999999, { criterion: "Z" }));
  const long = { period: { from: "2026-01-01", to: "2027-06-01" } };
  assertStringIncludes(call(long, all).warning, "longer than the 12 months");
  assertStringIncludes(call({}, all).warning, "origin criterion");
  assertStringIncludes(call({}, lines(line("1", 99999999))).warning, "exporter's");
});

test("bad data throws, naming the field", () => {
  const ok = lines(line("1", 100));
  const bad: Array<[Obj, Obj, string]> = [
    [{ agreement: "pem" }, ok, "origin.agreement"],
    [{ originCountry: "Canada" }, ok, "origin.originCountry"],
    [{ agreement: "eu-gsp", originCountry: "EU" }, ok, "origin.originCountry"],
    [{ place: " " }, ok, "origin.place"],
    [{ exporterName: "" }, ok, "origin.exporterName"],
    [{ date: "30/11/2026" }, ok, "origin.date"],
    [{}, { ...ok, currency: "gbp" }, "consignment.currency"],
    [{ eurRate: undefined }, ok, "origin.eurRate"],
    [{ eurRate: 0 }, ok, "origin.eurRate"],
    [{ eurRate: -1 }, ok, "origin.eurRate"],
    [{}, lines(line("1", 1), line(" ", 1)), "consignment.lines[1].ref"],
    [{}, lines(line("1", 1), line("1", 2)), "consignment.lines[1].ref"],
    [{}, lines(line("1", -1)), "consignment.lines[0].valueMinor"],
    [{}, lines(line("1", 1.5)), "consignment.lines[0].valueMinor"],
  ];
  for (const [o, c, field] of bad) assertThrows(() => call(o, c), field);
  assertThrows(() => mod.run([undefined, ok] as never, {} as RuntimeState), "origin");
});

test("EUR ignores a missing rate", () => {
  const o = call({ eurRate: undefined }, { currency: "EUR", lines: [line("1", 100)] });
  assertEquals(o.outcome, "statement");
});

test("forty lines keep their grouping and ref order", () => {
  const ls = Array.from({ length: 40 }, (_, i) => line(String(i + 1), 100, { criterion: i % 2 ? "C" : "a" }));
  const o = call({}, lines(...ls));
  const odd = ls.filter((_, i) => i % 2 === 0).map((l) => l.ref).join(", ");
  const even = ls.filter((_, i) => i % 2 === 1).map((l) => l.ref).join(", ");
  assertEquals(o.criteriaLine, `Origin criteria: A (lines ${odd}); C (lines ${even})`);
});

function assertThrows(fn: () => unknown, text: string) {
  try {
    fn();
  } catch (e) {
    assertStringIncludes((e as Error).message, text);
    return;
  }
  throw new Error("expected a throw naming " + text);
}

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Invoice">
      <Section id="body" title="Body">
        <PreferentialOriginStatement />
      </Section>
    </Document>,
  );

function state(data: Obj): RuntimeState {
  return { ctx: { ...data }, derived: {}, dataProvider: new InMemoryDataProvider(data), aiClient: { generateParagraph: () => "" } };
}

function collect(nodes: DocumentNode[], into: DocumentNode[] = []): DocumentNode[] {
  for (const n of nodes) {
    into.push(n);
    const children = (n as { children?: DocumentNode[] }).children;
    if (children) collect(children, into);
  }
  return into;
}

async function built(consignment: Obj, origin: Obj = baseOrigin) {
  const doc = await buildDocument(templateOf(), { origin, consignment }, {
    branchMode: "decide",
    dynamicMode: "placeholder",
    derivers,
  } as never);
  return collect((doc.nodes[0] as unknown as { children: DocumentNode[] }).children);
}

const ids = (nodes: DocumentNode[]) => nodes.map((n) => n.id);

test("a statement prints every id in the prescribed order", async () => {
  assertEquals(ids(await built(preview.consignment, preview.origin)), [
    "origin-statement",
    "origin-statement-period",
    "origin-statement-declaration",
    "origin-statement-exceptions",
    "origin-statement-criteria",
    "origin-statement-place-date",
    "origin-statement-exporter",
  ]);
});

test("a withheld case prints only the withheld ids", async () => {
  assertEquals(ids(await built(lines(line("1", 99999999)))), [
    "origin-statement-withheld",
    "origin-statement-withheld-warning",
    "origin-statement-withheld-place-date",
    "origin-statement-withheld-exporter",
  ]);
});

test("a publish build carries one deriver call and resolves like a direct build", async () => {
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  const arms = all.filter((n) => (n as { when?: unknown }).when !== undefined);
  assertEquals(arms.length, 5);
  const invocations = all.flatMap((n) => (n as { derivers?: unknown[] }).derivers ?? []);
  // The hook is called once; both arms' sections carry that same invocation.
  const plain = invocations.map((i) => JSON.parse(JSON.stringify(i)));
  assertEquals(new Set(plain.map((i) => JSON.stringify(i))).size, 1);
  assertEquals(plain[0].name, "preferentialOriginStatement");

  const seen = (nodes: DocumentNode[]) => collect(nodes).map((n) => [n.id, (n as { text?: string }).text]);
  for (const [origin, consignment] of [
    [preview.origin, preview.consignment],
    [baseOrigin, lines(line("1", 99999999))],
    [baseOrigin, lines()],
  ]) {
    const resolved = await resolveDocument(published, state({ origin, consignment }), { derivers });
    const direct = await buildDocument(templateOf(), { origin, consignment }, {
      branchMode: "decide",
      dynamicMode: "placeholder",
      derivers,
    } as never);
    assertEquals(seen(resolved.nodes), seen(direct.nodes));
  }
  await assertRejects(() => resolveDocument(published, state({ origin: { ...baseOrigin, place: "" }, consignment: lines() }), { derivers }));
});
