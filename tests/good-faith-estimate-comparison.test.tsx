import { test } from "node:test";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, Section, template } from "docxcelerate/template";
import { transformDocumentSource } from "docxcelerate/transform";
import { registryRoot } from "docxcelerate/registry/install";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";
import { COMPONENTS } from "docxcelerate/registry";

const nodeSource = "components/good-faith-estimate-comparison/good-faith-estimate-comparison.node.tsx";
const { GoodFaithEstimateComparison } = await loadRegistryModule<
  typeof import("../registry/components/good-faith-estimate-comparison/good-faith-estimate-comparison.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/good-faith-estimate-comparison.deriver.ts")
>("components/derivers/good-faith-estimate-comparison.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ goodFaithEstimateComparison: mod.run }));
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
interface Out {
  outcome: string;
  estimateDate: string;
  rows: Array<{ name: string; estimate: string; billed: string; difference: string }>;
  disputeSubject: string;
  initialBillDate: string;
  deadline: string;
  stateName: string;
  stateContact: string;
}

const run = (bill: unknown): Out => mod.run([bill], {} as RuntimeState) as Out;
const prov = (name: string, estimateCents: number | undefined, ...billed: number[]) => ({
  name,
  ...(estimateCents === undefined ? {} : { estimateCents }),
  billedLines: billed.map((amountCents) => ({ amountCents })),
});
const bill = (providers: Obj[], more: Obj = {}) => ({
  coverage: "uninsured",
  initialBillDate: "2026-11-01",
  estimate: { date: "2026-10-02" },
  providers,
  ...more,
});
const one = (est: number | undefined, ...billed: number[]) => run(bill([prov("X", est, ...billed)]));
const throws = (b: unknown, field: string) =>
  assertRejects(() => Promise.resolve().then(() => run(b)), Error, field);

test("insured and federal-program print nothing and read nothing else", () => {
  for (const o of [run({ coverage: "insured" }), run(bill([prov("X", 0, 1000000)], { coverage: "federal-program" }))]) {
    assertEquals(o.outcome, "none");
    assertEquals(o.rows.length, 6);
    assertEquals(
      [o.estimateDate, o.disputeSubject, o.initialBillDate, o.deadline, o.stateName, o.stateContact, ...o.rows.flatMap((r) => Object.values(r))].every((s) => s === ""),
      true,
    );
  }
});

test("no estimate, or no providers, prints nothing", () => {
  assertEquals(run(bill([prov("X", 1, 1)], { estimate: undefined })).outcome, "none");
  assertEquals(run(bill([prov("X", 1, 1)], { estimate: null })).outcome, "none");
  assertEquals(run(bill([], { coverage: "self-pay" })).outcome, "none");
  assertEquals(run({ coverage: "self-pay", estimate: { date: "2026-10-02" } }).outcome, "none");
});

test("coverage is required and exact", async () => {
  await throws({ providers: [] }, "patientBill.coverage");
  await throws(bill([], { coverage: "Uninsured" }), "patientBill.coverage");
  await throws(undefined, "patientBill");
});

test("the $400 line is inclusive and counted in cents", () => {
  const o = one(100000, 139999);
  assertEquals([o.outcome, o.rows[0].difference, o.deadline, o.disputeSubject], ["over", "$399.99 more", "", ""]);
  const d = one(100000, 140000);
  assertEquals([d.outcome, d.rows[0].difference, d.disputeSubject], ["dispute", "$400.00 more", "the bill from X"]);
  const w = one(100000, 100000);
  assertEquals([w.outcome, w.rows[0].difference], ["within", "No difference"]);
  const z = one(100000);
  assertEquals([z.outcome, z.rows[0].billed, z.rows[0].difference], ["within", "$0.00", "$1,000.00 less"]);
  assertEquals(w.rows[0].estimate, "$1,000.00");
});

test("a missing estimate is $0, and a present 0 is not 'not on your estimate'", () => {
  const zero = one(0, 40000);
  assertEquals([zero.rows[0].estimate, zero.outcome], ["$0.00", "dispute"]);
  const missing = one(undefined, 40000);
  assertEquals([missing.rows[0].estimate, missing.outcome], ["Not on your estimate", "dispute"]);
  const under = one(undefined, 39999);
  assertEquals([under.rows[0].estimate, under.outcome], ["Not on your estimate", "over"]);
});

test("each provider is compared with its own estimate, never in aggregate", () => {
  const o = run(bill([prov("A", 500000, 450000), prov("B", 100000, 145000)]));
  assertEquals([o.outcome, o.disputeSubject, o.rows[0].difference], ["dispute", "the bill from B", "$500.00 less"]);
  const agg = run(bill([prov("A", 100000, 130000), prov("B", 100000, 130000)]));
  assertEquals(agg.outcome, "over");
});

test("disputed names are listed in input order with an Oxford comma", () => {
  const o = run(bill([prov("A", 0, 50000), prov("Q", 0, 1), prov("B", 0, 50000), prov("C", 0, 50000)]));
  assertEquals(o.disputeSubject, "the bills from A, B, and C");
  assertEquals(run(bill([prov("A", 0, 50000), prov("B", 0, 50000)])).disputeSubject, "the bills from A and B");
});

test("six providers fit in order, seven throw, duplicates throw", async () => {
  const six = run(bill(["a", "b", "c", "d", "e", "f"].map((n) => prov(n, 1, 1))));
  assertEquals(six.rows.map((r) => r.name), ["a", "b", "c", "d", "e", "f"]);
  const three = run(bill([prov("a", 1, 1), prov("b", 1, 1), prov("c", 1, 1)]));
  assertEquals(three.rows.length, 6);
  assertEquals(three.rows[3], { name: "", estimate: "", billed: "", difference: "" });
  await throws(bill(["a", "b", "c", "d", "e", "f", "g"].map((n) => prov(n, 1, 1))), "providers");
  await throws(bill([prov("Dr Lee", 1, 1), prov(" Dr Lee ", 1, 1)]), "providers[1].name");
});

test("invalid amounts and shapes throw naming the path", async () => {
  await throws(bill([prov("A", 1, 1), prov("B", 1, -1)]), "providers[1].billedLines[0].amountCents");
  await throws(bill([prov("A", 1, 1, 12.5)]), "providers[0].billedLines[1].amountCents");
  await throws(bill([prov("A", -1, 1)]), "providers[0].estimateCents");
  await throws(bill([{ name: "A", estimateCents: 1 }]), "providers[0].billedLines");
  await throws(bill([prov("  ", 1, 1)]), "providers[0].name");
});

test("the deadline is 120 days on, in date-only arithmetic", () => {
  const dl = (iso: string) => run(bill([prov("X", 0, 50000)], { initialBillDate: iso })).deadline;
  assertEquals(dl("2026-11-01"), "March 1, 2027");
  assertEquals(dl("2027-11-01"), "February 29, 2028");
  assertEquals(dl("2026-12-31"), "April 30, 2027");
  assertEquals(dl("2026-11-01T23:30:00-08:00"), "March 1, 2027");
  assertEquals(run(bill([prov("X", 0, 50000)], { initialBillDate: "2026-11-01T23:30:00-08:00" })).initialBillDate, "November 1, 2026");
});

test("dates and state process are validated", async () => {
  const ok = [prov("X", 0, 50000)];
  await throws(bill(ok, { initialBillDate: undefined }), "patientBill.initialBillDate");
  await throws(bill(ok, { initialBillDate: "11/01/2026" }), "patientBill.initialBillDate");
  await throws(bill(ok, { estimate: {} }), "patientBill.estimate.date");
  const s = run(bill(ok, { stateProcess: { name: " X Dept ", contact: "1-555-0100" } }));
  assertEquals([s.stateName, s.stateContact], ["X Dept", "1-555-0100"]);
  assertEquals([one(1, 1).stateName, one(1, 1).stateContact], ["", ""]);
  await throws(bill(ok, { stateProcess: { name: "X" } }), "stateProcess.contact");
  await throws(bill(ok, { stateProcess: { name: " ", contact: "c" } }), "stateProcess.name");
});

test("the catalog's preview data", () => {
  const entry = COMPONENTS.find((c) => c.id === "good-faith-estimate-comparison")!;
  const o = run((entry.previewData as { patientBill: Obj }).patientBill);
  assertEquals(o.rows.slice(0, 3).map((r) => [r.estimate, r.billed, r.difference]), [
    ["$48,500.00", "$48,392.49", "$107.51 less"],
    ["$6,400.00", "$8,207.89", "$1,807.89 more"],
    ["Not on your estimate", "$482.50", "$482.50 more"],
  ]);
  assertEquals([o.outcome, o.deadline, o.estimateDate], ["dispute", "March 2, 2027", "September 14, 2026"]);
  assertStringIncludes(o.disputeSubject, "the bills from Associated Anesthesiologists");
  assertStringIncludes(o.disputeSubject, " and Lakeshore Pathology");
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Statement">
      <Section id="body" title="Body">
        <GoodFaithEstimateComparison />
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

const entry = COMPONENTS.find((c) => c.id === "good-faith-estimate-comparison")!;
const preview = (entry.previewData as { patientBill: Obj }).patientBill;

async function published() {
  return await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
}

async function resolved(patientBill: Obj) {
  const doc = await resolveDocument(await published(), state({ patientBill }), { derivers });
  return { ids: collect(doc.nodes).map((n) => n.id), text: JSON.stringify(doc.nodes) };
}

test("a publish build carries 12 arms and one deriver call", async () => {
  assertEquals(branches, 12);
  const doc = await published();
  const all = collect(doc.nodes);
  assertEquals(all.filter((n) => (n as { when?: unknown }).when !== undefined).length, 12);
  const section = all.find((n) => n.id === "gfe-comparison") as unknown as { derivers: unknown[] };
  assertEquals(section.derivers.length, 1);
  const d = JSON.parse(JSON.stringify(section.derivers[0]));
  assertEquals(d.name, "goodFaithEstimateComparison");
  assertEquals(d.inputs, [{ type: "ref", ref: { scope: "data", path: "patientBill" } }]);
});

test("the preview resolves to three rows and the federal contact", async () => {
  const { ids, text } = await resolved(preview);
  for (const id of ["gfe-comparison", "gfe-comparison-provider-1", "gfe-comparison-provider-2", "gfe-comparison-provider-3", "gfe-dispute", "gfe-dispute-how-federal"]) {
    assertEquals(ids.includes(id), true);
  }
  for (const id of ["gfe-comparison-provider-4", "gfe-comparison-provider-5", "gfe-comparison-provider-6", "gfe-comparison-within", "gfe-comparison-over", "gfe-dispute-how-state"]) {
    assertEquals(ids.includes(id), false);
  }
  assertStringIncludes(text, "March 2, 2027");
});

test("building locally with the catalog preview data yields the same nodes", async () => {
  const built = await buildDocument(templateOf(), entry.previewData as Obj, { dynamicMode: "placeholder" });
  const ids = collect(built.nodes).map((n) => n.id);
  for (const id of ["gfe-comparison", "gfe-comparison-provider-1", "gfe-comparison-provider-2", "gfe-comparison-provider-3", "gfe-dispute", "gfe-dispute-how-federal"]) {
    assertEquals(ids.includes(id), true);
  }
  assertEquals(ids.includes("gfe-comparison-provider-4"), false);
  assertStringIncludes(JSON.stringify(built.nodes), "March 2, 2027");
});

test("insured resolves to nothing; within, over and state variants pick their arms", async () => {
  assertEquals((await resolved({ ...preview, coverage: "insured" })).ids.some((id) => id.startsWith("gfe-")), false);
  const within = await resolved(bill([prov("X", 100000, 100000)]));
  assertEquals([within.ids.includes("gfe-comparison-within"), within.ids.includes("gfe-dispute")], [true, false]);
  const over = await resolved(bill([prov("X", 100000, 139999)]));
  assertEquals([over.ids.includes("gfe-comparison-over"), over.ids.includes("gfe-dispute")], [true, false]);
  const st = await resolved({ ...preview, stateProcess: { name: "X Dept", contact: "1-555-0100" } });
  assertEquals([st.ids.includes("gfe-dispute-how-state"), st.ids.includes("gfe-dispute-how-federal")], [true, false]);
  assertStringIncludes(st.text, "X Dept");
});
