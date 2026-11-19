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

const nodeSource = "components/commitment-drawdown/commitment-drawdown.node.tsx";
const { CommitmentDrawdown } = await loadRegistryModule<
  typeof import("../registry/components/commitment-drawdown/commitment-drawdown.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/commitment-drawdown.deriver.ts")
>("components/derivers/commitment-drawdown.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ commitmentDrawdown: mod.run }));

// deno-lint-ignore no-explicit-any
type Obj = Record<string, any>;
type Out = ReturnType<NonNullable<(typeof mod)["__types"]>>;
const run = (dd: Obj): Out => mod.run([dd], {} as RuntimeState) as Out;

const base = {
  currency: "GBP",
  invoiceDate: "2026-07-02",
  period: { start: "2026-06-01", end: "2026-06-30" },
  creditTaxBasis: "taxed-at-purchase",
};
const line = (amount: number, more: Obj = {}) => ({ description: "Usage", quantity: 100, amount, ...more });
const grant = (id: string, balance: number, more: Obj = {}) => ({
  id,
  name: id,
  kind: "commit",
  denomination: "money",
  balance,
  effective: "2026-01-01",
  expires: "2026-12-31",
  ...more,
});
const withBoth = (lines: Obj[], grants: Obj[], more: Obj = {}) => ({ ...base, lines, grants, ...more });
const drawOrder = (o: Out) => o.allocation.flatMap((a: Obj) => a.draws.map((d: Obj) => d.grantId));

test("a balance running out splits the line exactly", () => {
  const o = run(withBoth(
    [line(10000, { description: "A", quantity: 1000 }), line(5000, { description: "B", quantity: 500 })],
    [grant("g", 12000)],
  ));
  assertEquals(o.allocation.map((a: Obj) => [a.coveredMinor, a.chargedMinor]), [[10000, 0], [2000, 3000]]);
  assertEquals(o.allocation[1].chargedQuantity, 300);
  assertEquals(o.lines[1].charged, "£30.00 (300)");
  assertEquals(o.lines[0].charged, "—");
  assertEquals([o.minor.usage, o.minor.covered, o.minor.charged], [15000, 12000, 3000]);
  assertEquals(o.balances[0].remaining, 0);
});

test("a grant is tested on the usage date, not the invoice date", () => {
  const covered = (g: Obj, l: Obj = {}) => run(withBoth([line(1000, { usageDate: "2026-06-30", ...l })], [g])).minor.covered;
  assertEquals(covered(grant("g", 5000, { expires: "2026-07-01" })), 1000);
  assertEquals(covered(grant("g", 5000, { expires: "2026-06-30" })), 0);
  assertEquals(covered(grant("g", 5000, { expires: "2026-06-30" }), { usageDate: undefined }), 0);
  assertEquals(covered(grant("g", 5000, { expires: "2026-07-01" }), { usageDate: undefined }), 1000);
  assertEquals(covered(grant("g", 5000, { effective: "2026-07-01" })), 0);
});

test("burn order", () => {
  const first = (grants: Obj[]) => drawOrder(run(withBoth([line(100)], grants)))[0];
  assertEquals(first([
    grant("promo", 5000, { kind: "promotional", priority: 0, expires: "2026-07-05" }),
    grant("commit", 5000, { priority: 5 }),
    grant("roll", 5000, { kind: "rollover", priority: 9 }),
  ]), "roll");
  assertEquals(first([
    grant("dec", 5000, { priority: 1, expires: "2026-12-31" }),
    grant("aug", 5000, { priority: 1, expires: "2026-08-31" }),
  ]), "aug");
  assertEquals(first([grant("late", 5000, { effective: "2026-03-01" }), grant("early", 5000, { effective: "2026-02-01" })]), "early");
  assertEquals(first([grant("b", 5000, { created: "2026-02-02" }), grant("a", 5000, { created: "2026-02-01" })]), "a");
  assertEquals(first([grant("x", 5000), grant("y", 5000)]), "x");
  assertEquals(first([grant("none", 5000), grant("p99", 5000, { priority: 99 })]), "p99");
});

test("a units grant covers part of a line and prices the overage", async () => {
  const l = line(2000, { quantity: 1000000, unitPrice: 0.002, priceId: "p" });
  const o = run(withBoth([l], [grant("m", 750000, { denomination: "units", unit: "minutes", priceIds: ["p"] })]));
  assertEquals([o.allocation[0].coveredMinor, o.allocation[0].chargedMinor, o.allocation[0].chargedQuantity], [1500, 500, 250000]);
  assertEquals(o.allocation[0].draws, [{ grantId: "m", amountMinor: 1500, units: 750000 }]);
  assertEquals(o.grants[0].applied, "750,000 minutes (£15.00)");
  const r = run(withBoth([line(1500, { quantity: 1000000, unitPrice: 0.0015, priceId: "p" })], [
    grant("m", 666667, { denomination: "units", priceIds: ["p"] }),
  ]));
  assertEquals([r.allocation[0].chargedMinor, r.allocation[0].coveredMinor], [500, 1000]);
  assertEquals(run(withBoth([l], [grant("m", 5, { denomination: "units", priceIds: ["other"] })])).minor.covered, 0);
  await assertRejects(() => Promise.resolve().then(() => run(withBoth([l], [grant("m", 5, { denomination: "units" })]))), Error, "priceIds");
  await assertRejects(
    () => Promise.resolve().then(() => run(withBoth([line(2000, { priceId: "p" })], [grant("m", 5, { denomination: "units", priceIds: ["p"] })]))),
    Error,
    "unitPrice",
  );
});

test("several grants on one line", () => {
  const o = run(withBoth([line(10000)], [grant("a", 3000, { priority: 1 }), grant("b", 4000, { priority: 2 })]));
  assertEquals([o.allocation[0].coveredMinor, o.allocation[0].chargedMinor, o.allocation[0].chargedQuantity], [7000, 3000, 30]);
  assertEquals(o.allocation[0].draws.map((d: Obj) => [d.grantId, d.amountMinor]), [["a", 3000], ["b", 4000]]);
});

const trueUp = (extra: Obj = {}, term: Obj = {}) =>
  run(withBoth([line(20000)], [grant("c", 8000), grant("p", 4000, { kind: "promotional" })], {
    term: { start: "2025-07-01", end: "2026-06-30", commitment: 100000, eligibleSpendToDate: 60000, trueUpsBilled: 5000, ...term },
    ...extra,
  }));

test("true-up", () => {
  const o = trueUp();
  assertEquals([o.minor.coveredPaid, o.minor.coveredPromotional, o.minor.charged], [8000, 4000, 8000]);
  assertEquals([o.minor.eligibleSpend, o.minor.trueUp, o.trueUpArm], [76000, 19000, "shortfall"]);
  assertStringIncludes(o.trueUpNote, "£50.00 of true-ups were billed earlier");
  assertEquals(trueUp({}, { end: "2026-12-31" }).trueUpArm, "none");
  const met = trueUp({}, { commitment: 70000 });
  assertEquals([met.trueUpArm, met.minor.trueUp], ["met", 0]);
  assertStringIncludes(met.trueUpNote, "has been met");
  assertEquals(trueUp({}, { commitment: 0 }).trueUpArm, "none");
});

test("expiry warnings", () => {
  const note = (expires: string | undefined, more: Obj = {}, warnDays?: number) => {
    const o = run(withBoth([line(100)], [grant("g", 50100, { expires, ...more })], { warnDays }));
    return [o.grants[0].note, o.balances[0].warns, o.expiring] as const;
  };
  assertEquals(note("2026-07-31", { dailyBurn: 1000 })[0], "Expires in 29 days, on 31 July 2026. At recent usage about £210.00 will go unused.");
  assertEquals(note("2026-07-31", { dailyBurn: 2000 })[0], "Expires in 29 days, on 31 July 2026.");
  assertEquals(note("2026-07-31")[0], "Expires in 29 days, on 31 July 2026.");
  assertEquals(note("2026-08-01")[1], true);
  assertEquals(note("2026-08-02")[1], false);
  assertEquals(note("2026-07-31", {}, 0)[1], false);
  assertEquals(note(undefined), ["No expiry.", false, ""]);
  assertEquals(note("2026-07-31")[2], "1 credit expires within 30 days of this invoice: g on 31 July 2026.");
  assertEquals(note("2026-07-01")[0], "Expired on 1 July 2026. The rest has lapsed.");
});

test("tax basis", () => {
  const t = (basis: string, grants: Obj[]) =>
    run(withBoth([line(15000 + grants.reduce((n, g) => n + (g.kind === "promotional" ? g.balance : 0), 0))], grants, { creditTaxBasis: basis }));
  const paid = [grant("c", 12000)];
  const a = t("taxed-at-purchase", paid);
  assertEquals([a.minor.vatBase, a.minor.deductionAfterTax, a.taxArm], [3000, 0, "purchase"]);
  const b = t("multi-purpose", paid);
  assertEquals([b.minor.vatBase, b.minor.deductionAfterTax, b.taxArm], [15000, 12000, "redemption"]);
  const c = t("multi-purpose", [...paid, grant("p", 4000, { kind: "promotional" })]);
  assertEquals([c.minor.vatBase, c.minor.deductionAfterTax], [15000, 12000]);
  assertStringIncludes(c.taxNote, "Promotional credit of £40.00 is a discount");
  const only = t("multi-purpose", [grant("p", 4000, { kind: "promotional" })]);
  assertEquals(only.taxNote.startsWith("VAT applies to the £150.00 of usage charged."), true);
  assertEquals(run(withBoth([line(100)], [])).taxArm, "none");
});

test("zero and empty", async () => {
  const none = run({ ...base });
  assertEquals(none.show, false);
  assertEquals(await ids(withBoth([], [])), []);
  const z = run(withBoth([line(0), line(100)], [grant("g", 1000), grant("zero", 0)]));
  assertEquals(z.lines[0].usage, "£1.00");
  assertEquals(z.lines[1].description, "");
  assertEquals(z.allocation[0].draws, []);
  assertEquals(z.grants.filter((g: Obj) => g.name).length, 1);
  const g = run(withBoth([], [grant("g", 1000)]));
  assertEquals([g.show, g.showLines, g.showGrants], [true, false, true]);
});

test("invalid input throws, naming the field", async () => {
  const ok = withBoth([line(100)], [grant("g", 100)]);
  const bad: Array<[Obj, string]> = [
    [{ ...ok, lines: [line(-1)] }, "lines[0].amount"],
    [{ ...ok, lines: [line(1.5)] }, "lines[0].amount"],
    [{ ...ok, grants: [grant("g", -1)] }, "grants[0].balance"],
    [{ ...ok, grants: [grant("g", 1.5)] }, "grants[0].balance"],
    [{ ...ok, grants: [grant("g", 1, { kind: "bonus" })] }, "grants[0].kind"],
    [{ ...ok, creditTaxBasis: undefined }, "creditTaxBasis"],
    [{ ...ok, currency: "XX" }, "currency"],
    [{ ...ok, period: { start: "2026-06-30", end: "2026-06-01" } }, "period.end"],
    [{ ...ok, warnDays: -1 }, "warnDays"],
    [{ ...ok, lines: [{ ...line(1), description: " " }] }, "description"],
    [{ ...ok, term: { start: "2025-07-01", end: "2026-06-30", commitment: 1.5 } }, "commitment"],
  ];
  for (const [dd, field] of bad) await assertRejects(() => Promise.resolve().then(() => run(dd)), Error, field);
});

test("slots", () => {
  const many = (n: number, covered: number[]) => run(withBoth(
    Array.from({ length: n }, (_, i) => line(100, { description: `L${i}`, priceId: covered.includes(i) ? "c" : "x" })),
    [grant("g", 100000, { priceIds: ["c"] })],
  ));
  const names = (o: Out) => o.lines.map((l) => l.description);
  assertEquals(names(many(1, [0])), ["L0", "", "", "", "", ""]);
  assertEquals(names(many(6, [0])), ["L0", "L1", "L2", "L3", "L4", "L5"]);
  const three = many(9, [1, 3, 5]);
  assertEquals(names(three), ["L1", "L3", "L5", "", "", "6 lines not covered by credit"]);
  assertEquals([three.lines[5].usage, three.lines[5].covered, three.lines[5].charged], ["£6.00", "—", "£6.00"]);
  assertEquals(names(many(9, [0, 1, 2, 3, 4, 5, 6, 7, 8])).slice(4), ["L4", "4 further lines"]);
  const gs = run(withBoth([line(100)], ["a", "b", "c", "d", "e", "f"].map((id) => grant(id, 1000))));
  assertEquals(gs.grants.map((g: Obj) => g.name), ["a (prepaid commitment)", "b (prepaid commitment)", "c (prepaid commitment)", "3 other credits"]);
  assertEquals(gs.grants[3].remaining, "£30.00");
});

test("missing optionals", () => {
  const o = run(withBoth([line(100, { quantity: 12.25 })], [grant("g", 40)]));
  assertEquals(o.allocation[0].key, "0");
  assertEquals(o.allocation[0].chargedQuantity, 7.35);
  assertEquals(o.trueUpArm, "none");
  const w = run(withBoth([line(100)], [grant("g", 1000, { expires: "2026-08-01" })]));
  assertEquals(w.balances[0].warns, true); // 30 days: warnDays defaults to 30
});

test("a zero-decimal currency has no minor units", () => {
  const o = run(withBoth([line(1000)], [grant("g", 400)], { currency: "JPY" }));
  assertEquals(o.totals.usage, new Intl.NumberFormat("en-GB", { style: "currency", currency: "JPY" }).format(1000));
  assertEquals([o.allocation[0].coveredMinor, o.allocation[0].chargedMinor], [400, 600]);
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Invoice">
      <Section id="body" title="Body">
        <CommitmentDrawdown />
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

const built = (dd: Obj) =>
  buildDocument(templateOf(), { drawdown: dd }, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);

async function ids(dd: Obj): Promise<string[]> {
  const body = (await built(dd)).nodes[0] as { children: DocumentNode[] };
  return collect(body.children).map((n) => n.id).filter((id) => !id.startsWith("table-cell"));
}

const entry = COMPONENTS.find((c) => c.id === "commitment-drawdown")!;
const preview = (entry.previewData as { drawdown: Obj }).drawdown;

test("the catalog's preview data", () => {
  const o = run(preview);
  assertEquals(drawOrder(o), ["g-roll", "g-roll", "g-roll", "g-minutes", "g-commit", "g-commit"]);
  assertEquals(o.allocation.map((a: Obj) => a.draws.map((d: Obj) => d.amountMinor)), [
    [5520000], [4425000], [2555000, 528401], [1109767], [46890233], [], [],
  ]);
  assertEquals(o.allocation[4].chargedMinor, 23669767);
  assertEquals(o.lines[4].charged, "US$236,697.67 (242 hours)");
  assertEquals(o.lines[5], { description: "2 lines not covered by credit", usage: "US$58,800.00", covered: "—", charged: "US$58,800.00" });
  assertEquals([o.minor.usage, o.minor.covered, o.minor.charged], [90578168, 61028401, 29549767]);
  assertEquals([o.taxArm, o.minor.vatBase, o.minor.deductionAfterTax], ["redemption", 90578168, 61028401]);
  assertEquals([o.minor.eligibleSpend, o.minor.trueUp, o.trueUpArm], [289028168, 8471832, "shortfall"]);
  assertStringIncludes(o.trueUpNote, "US$84,718.32");
  const bal = Object.fromEntries(o.balances.map((b: Obj) => [b.id, b]));
  assertEquals(bal["g-roll"].remaining, 0);
  assertEquals(bal["g-commit"].remaining, 0);
  assertEquals(Math.abs(bal["g-minutes"].remaining - 779833.5) < 1, true);
  assertEquals(bal["g-minutes"].warns, true);
  assertStringIncludes(o.grants.find((g: Obj) => g.name.startsWith("Rollover"))!.note, "Expired on 1 July 2026.");
  assertStringIncludes(o.grants.find((g: Obj) => g.name.startsWith("Prepaid transcription"))!.note, "Expires in 13 days");
  assertEquals(o.expiring.startsWith("1 credit expires"), true);
});

test("a publish build carries 18 arms and one deriver call", async () => {
  const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), { fileName: nodeSource }).branches;
  assertEquals(branches, 18);
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  assertEquals(all.filter((n) => (n as { when?: unknown }).when !== undefined).length, 18);
  const section = all.find((n) => n.id === "commitment-drawdown") as unknown as { derivers: unknown[] };
  assertEquals(section.derivers.length, 1);
  assertEquals(JSON.parse(JSON.stringify(section.derivers[0])).name, "commitmentDrawdown");

  const cases: Obj[] = [preview, withBoth([line(100)], [grant("g", 50)]), withBoth([], [])];
  for (const dd of cases) {
    const direct = await built(dd);
    const resolved = await resolveDocument(published, state({ drawdown: dd }), { derivers });
    const seen = (nodes: DocumentNode[]) => collect(nodes).map((n) => [n.kind, (n as { text?: string }).text]);
    assertEquals(seen(resolved.nodes), seen(direct.nodes));
  }

  const resolvedIds = collect((await resolveDocument(published, state({ drawdown: preview }), { derivers })).nodes).map((n) => n.id);
  for (const id of ["line-5", "line-6", "tax-redemption", "true-up", "expiring"]) {
    assertEquals(resolvedIds.includes(`commitment-drawdown-${id}`), true);
  }
  for (const id of ["commitment-met", "tax-purchase"]) assertEquals(resolvedIds.includes(`commitment-drawdown-${id}`), false);
  assertStringIncludes(JSON.stringify((await built(preview)).nodes), "2 lines not covered by credit");
});
