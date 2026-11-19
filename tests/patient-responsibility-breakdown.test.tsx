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
import { COMPONENTS } from "docxcelerate/registry";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const nodeSource = "components/patient-responsibility-breakdown/patient-responsibility-breakdown.node.tsx";
const { PatientResponsibilityBreakdown } = await loadRegistryModule<
  typeof import("../registry/components/patient-responsibility-breakdown/patient-responsibility-breakdown.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/patient-responsibility-breakdown.deriver.ts")
>("components/derivers/patient-responsibility-breakdown.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(
  createDeriverBundle({ patientResponsibilityBreakdown: mod.run }),
);
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
interface Row {
  service: string;
  detail: string;
  billed: string;
  allowed: string;
  deductible: string;
  copay: string;
  coinsurance: string;
  youOwe: string;
  note: string | null;
}
interface Out {
  outcome: string;
  rows: Row[];
  total: Row;
  youOweMinor: number;
  adjustment: string;
  hasAdjustment: boolean;
  excluded: string;
  hasExcluded: boolean;
  deductibleRemaining: string;
  oopRemaining: string;
  oopMet: boolean;
}

const run = (pr: Obj): Out => mod.run([pr], {} as RuntimeState) as Out;
const plan = (more: Obj = {}) => ({ deductibleRemainingMinor: 50000, oopRemainingMinor: 200000, coinsuranceRate: 0.2, ...more });
const line = (more: Obj = {}) => ({
  serviceDate: "2026-11-03",
  description: "Service",
  billedMinor: 120000,
  allowedMinor: 80000,
  network: "in",
  covered: true,
  ...more,
});
const one = (lineMore: Obj = {}, planMore: Obj = {}) => run({ plan: plan(planMore), lines: [line(lineMore)] });
// Dollars as the formatter prints them, so assertions read in cents.
const usd = (cents: number) => new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(cents / 100);
const cells = (o: Out, i = 0) => [o.rows[i].deductible, o.rows[i].copay, o.rows[i].coinsurance, o.rows[i].youOwe];
const cents = (...v: number[]) => v.map(usd);

test("one ordinary line", () => {
  const o = one();
  assertEquals(cells(o), cents(50000, 0, 6000, 56000));
  assertEquals([o.adjustment, o.deductibleRemaining, o.oopRemaining, o.outcome, o.rows[0].note], [usd(40000), usd(0), usd(144000), "owes", null]);
});

test("plan order: copay first, counting or not, and deductible first", () => {
  const l = { allowedMinor: 60000, copayMinor: 3000 };
  assertEquals(cells(one(l)), cents(50000, 3000, 1400, 54400));
  assertEquals(cells(one(l, { copayCountsTowardDeductible: true })), cents(47000, 3000, 2000, 52000));
  for (const counts of [false, true]) {
    assertEquals(cells(one(l, { copayAppliesBeforeDeductible: false, copayCountsTowardDeductible: counts })), cents(50000, 3000, 1400, 54400));
  }
});

test("copay only, and a copay above the allowed amount", () => {
  const l = { allowedMinor: 25000, copayMinor: 4000, copayOnly: true };
  const o = one(l);
  assertEquals([o.rows[0].youOwe, o.deductibleRemaining], [usd(4000), usd(50000)]);
  assertEquals(one(l, { copayCountsTowardDeductible: true }).deductibleRemaining, usd(46000));
  const c = one({ billedMinor: 3000, allowedMinor: 1800, copayMinor: 4000, copayOnly: true }, { deductibleRemainingMinor: 0 });
  assertEquals(c.rows[0].copay, usd(1800));
});

test("the cap carries across lines in date order", () => {
  const o = run({
    plan: plan({ deductibleRemainingMinor: 0, oopRemainingMinor: 10000 }),
    lines: [30000, 40000, 50000].map((a, i) => line({ serviceDate: "2026-11-0" + (i + 1), billedMinor: a, allowedMinor: a })),
  });
  assertEquals(o.rows.map((r) => r.coinsurance), cents(6000, 4000, 0));
  assertEquals(o.rows.map((r) => r.note), [
    null,
    "You reached your out-of-pocket maximum with this service. Your plan pays the rest.",
    "Your plan pays 100%: you have already reached your out-of-pocket maximum.",
  ]);
  assertEquals([o.total.youOwe, o.youOweMinor, o.oopMet], [usd(10000), 10000, true]);
});

test("the cap cuts a line mid-way and satisfies the deductible", () => {
  const o = one({ allowedMinor: 20000, copayMinor: 3000 }, { oopRemainingMinor: 5000 });
  assertEquals(cells(o), cents(2000, 3000, 0, 5000));
  assertEquals(o.deductibleRemaining, usd(0));
});

test("an already met cap bills only what sits outside it", () => {
  const o = run({
    plan: plan({ oopRemainingMinor: 0 }),
    lines: [line(), line({ serviceDate: "2026-11-04", billedMinor: 15000, covered: false })],
  });
  assertEquals(o.rows[0].youOwe, usd(0));
  assertStringIncludes(o.rows[0].note!, "Your plan pays 100%");
  assertEquals(o.rows[1].youOwe, usd(15000));
  assertStringIncludes(o.rows[1].note!, "Not covered by your plan. The full charge of $150.00");
  assertEquals([o.total.youOwe, o.excluded, o.hasExcluded], [usd(15000), usd(15000), true]);
});

test("a met deductible bills coinsurance alone", () => {
  const o = one({}, { deductibleRemainingMinor: 0 });
  assertEquals(cells(o), cents(0, 0, 16000, 16000));
  assertEquals(o.rows[0].note, null);
});

test("out of network: balance bill, or protected", () => {
  const l = { billedMinor: 100000, allowedMinor: 60000, network: "out" };
  const bill = one({ ...l, nsaProtected: false }, { deductibleRemainingMinor: 0 });
  assertEquals(cells(bill), cents(0, 0, 12000, 52000));
  assertEquals([bill.adjustment, bill.excluded, bill.hasAdjustment], [usd(0), usd(40000), false]);
  assertStringIncludes(bill.rows[0].note!, "Out of network. The provider bills you the $400.00");
  const safe = one(l, { deductibleRemainingMinor: 0 });
  assertEquals([safe.rows[0].youOwe, safe.adjustment], [usd(12000), usd(40000)]);
  assertStringIncludes(safe.rows[0].note!, "No Surprises Act");
});

test("a denial is not billed, and a zero charge has no note", () => {
  const d = one({ billedMinor: 50000, allowedMinor: 0 });
  assertEquals([d.rows[0].youOwe, d.adjustment, d.outcome, d.deductibleRemaining, d.oopRemaining], [usd(0), usd(0), "nothing-owed", usd(50000), usd(200000)]);
  assertStringIncludes(d.rows[0].note!, "allowed nothing");
  assertEquals(one({ billedMinor: 0, allowedMinor: 0 }).rows[0].note, null);
});

test("coinsurance rounds half up once per line, and totals do not drift", () => {
  assertEquals(one({ allowedMinor: 12345, billedMinor: 12345 }, { deductibleRemainingMinor: 0, coinsuranceRate: 0.15 }).rows[0].coinsurance, usd(1852));
  assertEquals(one({ allowedMinor: 3, billedMinor: 3 }, { deductibleRemainingMinor: 0, coinsuranceRate: 0.5 }).rows[0].coinsurance, usd(2));
  const o = run({
    plan: plan({ deductibleRemainingMinor: 0, coinsuranceRate: 0.2 }),
    lines: [1, 2, 3].map((i) => line({ serviceDate: "2026-11-0" + i, billedMinor: 1, allowedMinor: 1 })),
  });
  assertEquals([o.rows.map((r) => r.coinsurance), o.total.coinsurance], [cents(0, 0, 0), usd(0)]);
});

test("lines print in date order, ties in input order", () => {
  const o = run({
    plan: plan({ deductibleRemainingMinor: 30000 }),
    lines: [
      line({ serviceDate: "2026-11-20", allowedMinor: 20000, billedMinor: 20000, description: "late" }),
      line({ serviceDate: "2026-11-05", allowedMinor: 20000, billedMinor: 20000, description: "early" }),
    ],
  });
  assertEquals(o.rows.map((r) => r.service), ["early", "late"]);
  assertEquals([o.rows[0].youOwe, o.rows[1].deductible, o.rows[1].coinsurance, o.rows[1].youOwe], cents(20000, 10000, 2000, 12000));
  const t = run({ plan: plan(), lines: ["b", "a", "c"].map((d) => line({ description: d })) });
  assertEquals(t.rows.map((r) => r.service), ["b", "a", "c"]);
});

test("no lines prints nothing", () => {
  assertEquals(run({ plan: plan() }).outcome, "none");
  const e = run({ plan: plan(), lines: [] });
  assertEquals([e.outcome, e.oopMet, e.rows.length], ["none", false, 0]);
});

test("forty lines, and rates of 0 and 1", () => {
  const o = run({
    plan: plan({ deductibleRemainingMinor: 0, oopRemainingMinor: 10_000_000 }),
    lines: Array.from({ length: 40 }, (_, i) => line({ serviceDate: "2026-11-01", description: "L" + i, billedMinor: 1001, allowedMinor: 1001 })),
  });
  assertEquals(o.rows.length, 40);
  assertEquals(o.youOweMinor, 40 * 200);
  assertEquals(o.total.billed, usd(40 * 1001));
  assertEquals(one({}, { deductibleRemainingMinor: 0, coinsuranceRate: 0 }).rows[0].coinsurance, usd(0));
  assertEquals(one({}, { deductibleRemainingMinor: 0, coinsuranceRate: 1 }).rows[0].coinsurance, usd(80000));
});

test("invalid input throws, naming the field", async () => {
  const bad: Array<[Obj, string]> = [
    [{ plan: plan(), lines: [line({ billedMinor: -1 })] }, "lines[0].billedMinor"],
    [{ plan: plan(), lines: [line(), line({ allowedMinor: 1.5 })] }, "lines[1].allowedMinor"],
    [{ plan: plan(), lines: [line({ allowedMinor: 130000 })] }, "lines[0].allowedMinor"],
    [{ plan: plan(), lines: [line({ allowedMinor: undefined })] }, "lines[0].allowedMinor"],
    [{ plan: plan(), lines: [line(), line(), line({ serviceDate: undefined })] }, "lines[2].serviceDate"],
    [{ plan: plan(), lines: [line({ description: "  " })] }, "lines[0].description"],
    [{ plan: plan(), lines: [line({ copayMinor: -5 })] }, "lines[0].copayMinor"],
    [{ plan: plan({ coinsuranceRate: 1.2 }), lines: [] }, "plan.coinsuranceRate"],
    [{ plan: plan(), lines: [line({ network: "x" })] }, "lines[0].network"],
    [{ currency: "XXX!", plan: plan(), lines: [line()] }, "currency"],
    [{ plan: { oopRemainingMinor: 0, coinsuranceRate: 0 }, lines: [] }, "plan.deductibleRemainingMinor"],
  ];
  for (const [pr, field] of bad) {
    await assertRejects(() => Promise.resolve().then(() => run(pr)), Error, field);
  }
});

test("formatting", () => {
  const o = run({
    plan: plan(),
    lines: [
      line({ serviceDate: "2026-09-14T10:00:00Z", code: "27447", billedMinor: 24875000, allowedMinor: 100 }),
      line({ serviceDate: "2026-09-15" }),
    ],
  });
  assertEquals([o.rows[0].billed, o.rows[0].detail, o.rows[1].detail], ["$248,750.00", "Sep 14, 2026 · 27447", "Sep 15, 2026"]);
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = (arrangement?: string) =>
  template<Obj>(
    <Document id="d" title="Statement">
      <Section id="body" title="Body">
        <PatientResponsibilityBreakdown arrangement={arrangement as never} />
      </Section>
    </Document>,
  );

const state = (data: Obj): RuntimeState => ({
  ctx: { ...data },
  derived: {},
  dataProvider: new InMemoryDataProvider(data),
  aiClient: { generateParagraph: () => "" },
});

function collect(nodes: DocumentNode[], into: DocumentNode[] = []): DocumentNode[] {
  for (const n of nodes) {
    into.push(n);
    const children = (n as { children?: DocumentNode[] }).children;
    if (children) collect(children, into);
  }
  return into;
}

const preview = (COMPONENTS.find((c) => c.id === "patient-responsibility-breakdown")!.previewData) as Obj;
const publish = (arrangement?: string) =>
  buildDocument(templateOf(arrangement), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
const resolveWith = async (data: Obj, arrangement?: string) =>
  collect((await resolveDocument(await publish(arrangement), state(data), { derivers })).nodes);

test("the publish build has one deriver and a loop over the derived rows", async () => {
  const all = collect((await publish()).nodes);
  const section = all.find((n) => n.id === "patient-responsibility") as unknown as { derivers: unknown[] };
  assertEquals(section.derivers.length, 1);
  assertEquals(JSON.parse(JSON.stringify(section.derivers[0])).name, "patientResponsibilityBreakdown");

  const loop = all.find((n) => n.kind === "repeat") as unknown as { source: unknown };
  assertEquals(JSON.parse(JSON.stringify(loop.source)), {
    scope: "derived",
    path: "patientResponsibilityBreakdown.rows",
  });
  const note = all.find((n) => n.id === "patient-responsibility-line-note") as unknown as { when: { ref: { scope: string } } };
  assertEquals(note.when.ref.scope, "ctx");
  assertEquals(branches <= 32, true);
});

test("the preview resolves to six rows and the expected paragraphs", async () => {
  const o = run((preview as { patientResponsibility: Obj }).patientResponsibility);
  assertEquals(o.rows.map((r) => r.youOwe), cents(6500, 906000, 0, 0, 108000, 36000));
  assertEquals([o.total.youOwe, o.adjustment, o.excluded, o.oopMet], [usd(1056500), usd(20761548), usd(144000), true]);
  assertEquals(o.rows[1].coinsurance, usd(621000));

  for (const arrangement of ["table", "stacked"]) {
    const nodes = await resolveWith(preview, arrangement);
    const ids = nodes.map((n) => n.id);
    const text = JSON.stringify(nodes);
    for (const id of ["owes", "adjustment", "excluded", "oop-met"]) assertEquals(ids.includes("patient-responsibility-" + id), true);
    for (const id of ["remaining", "nothing-owed"]) assertEquals(ids.includes("patient-responsibility-" + id), false);
    for (const r of o.rows.slice(1)) assertStringIncludes(text, r.note!.slice(0, 40));
    assertEquals(ids.filter((id) => id.startsWith("patient-responsibility-line-service")).length, 6);
    assertEquals(ids.filter((id) => id.startsWith("patient-responsibility-line-note")).length, 5);
  }
  const table = await resolveWith(preview, "table");
  assertEquals(table.filter((n) => /^patient-responsibility-line-\d+$/.test(n.id)).length, 6);
});

test("no lines resolves to nothing", async () => {
  const data = { patientResponsibility: { plan: plan(), lines: [] } };
  assertEquals((await resolveWith(data)).filter((n) => n.id.startsWith("patient-responsibility")), []);
});

test("both arrangements build locally and differ", async () => {
  const seen = [];
  for (const arrangement of ["table", "stacked"]) {
    const built = await buildDocument(templateOf(arrangement), preview, {
      branchMode: "decide",
      dynamicMode: "placeholder",
      derivers,
    } as never);
    seen.push(JSON.stringify(built.nodes));
  }
  assertEquals(seen[0] === seen[1], false);
});
