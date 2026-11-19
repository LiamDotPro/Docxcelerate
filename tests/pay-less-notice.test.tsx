import { test } from "node:test";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, template } from "docxcelerate/template";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";
import { COMPONENTS } from "../src/registry/catalog.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/pay-less-notice.deriver.ts")
>("components/derivers/pay-less-notice.deriver.ts");
const { PayLessNotice } = await loadRegistryModule<
  typeof import("../registry/components/pay-less-notice/pay-less-notice.node.tsx")
>("components/pay-less-notice/pay-less-notice.node.tsx");

// deno-lint-ignore no-explicit-any
type Obj = Record<string, any>;
const preview = (COMPONENTS.find((c) => c.id === "pay-less-notice")!.previewData as Obj).payLess as Obj;
const base: Obj = {
  contractRef: "C1",
  payer: "Payer Ltd",
  payee: "Payee Ltd",
  notifiedSum: { amountMinor: 1_000_000, source: "payment-notice" },
  finalDateForPayment: "2026-11-30",
  prescribedPeriodDays: 7,
  servedDate: "2026-11-20",
};
const run = (extra: Obj = {}, drop: string[] = []) => {
  const input = { ...base, ...extra };
  for (const k of drop) delete input[k];
  return mod.run([input], {} as RuntimeState) as Obj;
};
const ded = (kind: string, amountMinor: number, description = "Work") => ({ kind, description, amountMinor });
const rejects = (extra: Obj, text: string, drop: string[] = []) =>
  assertRejects(() => Promise.resolve().then(() => run(extra, drop)), Error, text);

test("preview data gives the worked figures", () => {
  const r = run(preview as Obj, []);
  assertEquals(r.timing, "in-time");
  assertEquals(r.deadline, "23 November 2026");
  assertEquals(r.lines.map((l: Obj) => l.kind).slice(0, 6), [
    "Liquidated damages",
    "Contra-charge",
    "Contra-charge",
    "Abatement",
    "Set-off",
    "",
  ]);
  assertEquals(r.lines[0].amount, "£200,000.00");
  assertStringIncludes(r.lines[0].description, "£235,000.00 capped at £200,000.00");
  assertEquals(r.lines[2].description.includes("CC-0201"), true);
  assertEquals(r.totalDeductions, "£842,772.68");
  assertEquals(r.sumDue, "£4,029,543.77");
  assertEquals(r.excess, "");
});

test("no deductions leaves twelve empty slots", () => {
  const r = run();
  assertEquals(r.lines.length, 12);
  assertEquals(r.lines.every((l: Obj) => l.kind === "" && l.description === "" && l.amount === ""), true);
  assertEquals([r.totalDeductions, r.sumDue, r.excess], ["£0.00", "£10,000.00", ""]);
  assertStringIncludes(r.basisFacts, "No deductions.");
});

test("one contra deduction fills slot one only", () => {
  const r = run({ deductions: [ded("contra", 100)] });
  assertEquals(r.lines[0].kind, "Contra-charge");
  assertEquals(r.lines.slice(1).every((l: Obj) => l.description === ""), true);
});

test("twelve lines pass and thirteen throw", async () => {
  const ld = { rateMinor: 100, per: "day", delayDays: 1 };
  run({ ld, deductions: Array.from({ length: 11 }, () => ded("contra", 1)) });
  await rejects({ ld, deductions: Array.from({ length: 12 }, () => ded("contra", 1)) }, "at most 12");
});

test("kinds are grouped, input order kept within a kind", () => {
  const r = run({
    ld: { rateMinor: 100, per: "day", delayDays: 2 },
    deductions: [ded("setoff", 1, "s1"), ded("contra", 2, "c1"), ded("abatement", 3, "a1"), ded("contra", 4, "c2")],
  });
  assertEquals(r.lines.slice(0, 6).map((l: Obj) => l.description.slice(0, 2)), ["2 ", "c1", "c2", "a1", "s1", ""]);
});

test("a zero deduction is dropped", () => {
  assertEquals(run({ deductions: [ded("contra", 0)] }).lines[0].description, "");
});

test("bad amounts, kinds and descriptions throw", async () => {
  await rejects({ deductions: [ded("contra", -1)] }, "payLess.deductions[0].amountMinor");
  await rejects({ deductions: [ded("contra", 1.5)] }, "payLess.deductions[0].amountMinor");
  await rejects({ deductions: [ded("contra", 1), ded("cis", 1), ded("contra", 1)] }, "CIS");
  await rejects({ deductions: [ded("contra", 1), ded("contra", 1), ded("cis", 1)] }, "not a ground");
  await rejects({ deductions: [ded("contra", 1), ded("contra", 1), ded("fine", 1)] }, "payLess.deductions[2].kind");
  await rejects({ deductions: [ded("contra", 1, "  ")] }, "payLess.deductions[0].description");
  await rejects({ notifiedSum: { amountMinor: -1, source: "payment-notice" } }, "notifiedSum.amountMinor");
});

test("deductions equal to the notified sum give nil with no excess", () => {
  const r = run({ deductions: [ded("contra", 1_000_000)] });
  assertEquals([r.sumDue, r.excess], ["nil (£0.00)", ""]);
});

test("deductions over the notified sum give nil and an excess, never a negative", () => {
  const r = run({ deductions: [ded("contra", 2_250_000)] });
  assertEquals([r.sumDue, r.excess], ["nil (£0.00)", "£12,500.00"]);
  assertEquals(JSON.stringify(r).includes("-£"), false);
  assertStringIncludes(r.basisFacts, "exceed the notified sum by £12,500.00");
  const z = run({ notifiedSum: { amountMinor: 0, source: "application" }, deductions: [ded("setoff", 500)] });
  assertEquals([z.sumDue, z.excess, z.totalDeductions], ["nil (£0.00)", "£5.00", "£5.00"]);
});

test("liquidated damages by day and week", () => {
  const d = run({ ld: { rateMinor: 50000, per: "day", delayDays: 10 } });
  assertEquals(d.lines[0].amount, "£5,000.00");
  assertEquals(d.lines[0].description.includes("cap"), false);
  assertEquals(run({ ld: { rateMinor: 1000, per: "week", delayDays: 3 } }).lines[0].amount, "£4.29");
  assertEquals(run({ ld: { rateMinor: 100, per: "day", delayDays: 0 } }).lines[0].description, "");
  assertStringIncludes(run({ ld: { rateMinor: 100, per: "day", delayDays: 1 } }).lines[0].description, "1 day of");
  assertEquals(run({ ld: { rateMinor: 100, per: "day", delayDays: 1 } }).lines[0].description.includes("1 days"), false);
  assertStringIncludes(
    run({ ld: { rateMinor: 100, per: "week", delayDays: 7 } }).lines[0].description,
    "per week, pro rata by the day",
  );
});

test("the cap is cumulative", () => {
  const ld = { rateMinor: 10000, per: "day", delayDays: 10, capMinor: 150000 };
  const part = run({ ld: { ...ld, previouslyDeductedMinor: 100000 } }).lines[0];
  assertEquals(part.amount, "£500.00");
  assertStringIncludes(part.description, "capped at");
  for (const previouslyDeductedMinor of [150000, 200000]) {
    const none = run({ ld: { ...ld, previouslyDeductedMinor } }).lines[0];
    assertEquals(none.amount, "£0.00");
    assertStringIncludes(none.description, "nothing deducted");
  }
  const uncapped = run({ ld: { rateMinor: 10000, per: "day", delayDays: 10, previouslyDeductedMinor: 5 } }).lines[0];
  assertEquals(uncapped.amount, "£1,000.00");
  assertEquals(uncapped.description.includes("cap"), false);
  assertEquals(run({ ld }).lines[0].amount, "£1,000.00");
  assertEquals(run({ ld: { ...ld, capMinor: 400 } }).lines[0].amount, "£4.00");
});

test("timing around the deadline", () => {
  assertEquals(run({ servedDate: "2026-11-23" }).timing, "in-time");
  assertEquals(run({ servedDate: "2026-11-24" }).timing, "late");
  assertEquals(run({ servedDate: "2026-12-01" }).timing, "late");
});

test("deadline: explicit wins, one is required, and it cannot follow the final date", async () => {
  assertEquals(run({ payLessDeadline: "2026-11-10" }).deadline, "10 November 2026");
  await rejects({}, "prescribedPeriodDays", ["prescribedPeriodDays"]);
  await rejects({ payLessDeadline: "2026-12-01" }, "payLessDeadline");
  await rejects({ prescribedPeriodDays: 0 }, "prescribedPeriodDays");
});

test("excluded days push the deadline earlier", () => {
  const change = { finalDateForPayment: "2026-12-29", servedDate: "2026-12-01" };
  assertEquals(run(change).deadline, "22 December 2026");
  assertEquals(run({ ...change, excludedDays: ["2026-12-25", "2026-12-28"] }).deadline, "20 December 2026");
});

test("malformed dates throw", async () => {
  await rejects({ servedDate: "yesterday" }, "payLess.servedDate");
  await rejects({ excludedDays: ["2026-12-25", "xmas"] }, "payLess.excludedDays[1]");
});

test("notified source text", async () => {
  assertEquals(run().notifiedSource, "the payment notice");
  assertEquals(
    run({ notifiedSum: { amountMinor: 1, source: "payment-notice", reference: "  " } }).notifiedSource,
    "the payment notice",
  );
  assertEquals(
    run({ notifiedSum: { amountMinor: 1, source: "application", reference: " IA 14 " } }).notifiedSource,
    "the payee's payment application, which stands as the notice of the sum due (IA 14)",
  );
  assertEquals(
    run({ notifiedSum: { amountMinor: 1, source: "default-notice" } }).notifiedSource,
    "the payee's default payment notice",
  );
  await rejects({ notifiedSum: { amountMinor: 1, source: "other" } }, "notifiedSum.source");
});

// ---------------------------------------------------------------------------
// The node
// ---------------------------------------------------------------------------

const wrap = () => template<Obj>(<Document title="Notice"><PayLessNotice /></Document>);
const ids = (value: unknown): string[] => {
  const out: string[] = [];
  const walk = (n: unknown) => {
    if (n && typeof n === "object") {
      const id = (n as Obj).id;
      if (typeof id === "string") out.push(id);
      Object.values(n).forEach(walk);
    }
  };
  walk(value);
  return out;
};
const textOf = (value: unknown): string => JSON.stringify(value, (k, v) => (k === "key" ? undefined : v));
const buildWith = (payLess: Obj) =>
  buildDocument(wrap(), { payLess }, { dynamicMode: "placeholder" });

test("built with real data: late banner, or in-time banner", async () => {
  const late = ids((await buildWith({ ...base, servedDate: "2026-11-30" })).nodes);
  assertEquals([late.includes("pay-less-notice-late"), late.includes("pay-less-notice-in-time")], [true, false]);
  const ok = ids((await buildWith(base)).nodes);
  assertEquals([ok.includes("pay-less-notice-late"), ok.includes("pay-less-notice-in-time")], [false, true]);
  assertEquals(ok.includes("pay-less-notice-excess"), false);
});

test("built with real data: nil case shows the excess", async () => {
  const built = ids((await buildWith({ ...base, deductions: [ded("contra", 2_000_000)] })).nodes);
  assertEquals(built.includes("pay-less-notice-excess"), true);
  assertEquals(built.includes("pay-less-notice-line-1"), true);
  assertEquals(built.includes("pay-less-notice-line-2"), false);
});

const publish = () =>
  buildDocument(wrap(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
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

test("publishing keeps 15 arms, one deriver call and a dynamic basis prompt", async () => {
  const all = collect((await publish()).nodes);
  assertEquals(all.filter((n) => (n as { when?: unknown }).when !== undefined).length, 15);
  const basis = all.find((n) => n.id === "pay-less-notice-basis") as Obj;
  assertEquals(basis.mode, "dynamic");
  assertStringIncludes(textOf(basis), "{{derived.payLessNotice.basisFacts}}");
  const section = all.find((n) => n.id === "pay-less-notice") as unknown as { derivers: unknown[] };
  assertEquals(section.derivers.length, 1);
  assertEquals(JSON.parse(JSON.stringify(section.derivers[0])).inputs, [
    { type: "ref", ref: { scope: "data", path: "payLess" } },
  ]);
});

test("publish, bundle and resolve match a direct build", async () => {
  const published = await publish();
  const derivers = { derivers: await createDeriverRegistryFromBundle(createDeriverBundle({ payLessNotice: mod.run })) };
  const cases: Obj[] = [
    preview,
    { ...base, servedDate: "2026-11-25" },
    { ...base, deductions: [ded("contra", 2_000_000)] },
  ];
  // The generated basis paragraph differs by design (stand-in vs. the AI client), so leave it out.
  const seen = (nodes: DocumentNode[]) =>
    collect(nodes).filter((n) => n.id !== "pay-less-notice-basis").map((n) => [n.kind, (n as { text?: string }).text]);
  for (const payLess of cases) {
    const data = { payLess };
    const state: RuntimeState = {
      ctx: { ...data },
      derived: {},
      dataProvider: new InMemoryDataProvider(data),
      aiClient: { generateParagraph: () => "generated" },
    };
    const resolved = await resolveDocument(published, state, derivers);
    const direct = await buildDocument(wrap(), data, { dynamicMode: "placeholder" });
    assertEquals(seen(resolved.nodes), seen(direct.nodes));
  }
});
