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
import type { RuntimeState } from "docxcelerate";
import { createPublishData, Document, template } from "docxcelerate/template";
import { transformDocumentSource } from "docxcelerate/transform";
import { registryRoot } from "docxcelerate/registry/install";
import { COMPONENTS } from "docxcelerate/registry";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const nodeSource = "components/timesheet-fee-summary/timesheet-fee-summary.node.tsx";
const { TimesheetFeeSummary } = await loadRegistryModule<
  typeof import("../registry/components/timesheet-fee-summary/timesheet-fee-summary.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/timesheet-fee-summary.deriver.ts")
>("components/derivers/timesheet-fee-summary.deriver.ts");

type Obj = Record<string, unknown>;
interface Out {
  currency: string;
  hasRows: boolean;
  basis: string;
  rows: Array<{ label: string; role: string; hours: string; rate: string; amount: string }>;
  adjustments: Array<{ label: string; amount: string }>;
  subtotal: string;
  subtotalMinor: number;
}
const state = {} as RuntimeState;
const run = (time: Obj) => mod.run([time], state) as Out;
const keeper = (more: Obj = {}) => ({ id: "A", name: "Ann", role: "Partner", rates: [{ from: "2026-01-01", rate: 250 }], ...more });
const at = (minutes: number, more: Obj = {}) => ({ date: "2026-03-01", timekeeperId: "A", minutes, ...more });
/** One timekeeper at 250, GBP, default billing, unless overridden. */
const mk = (entries: Obj[], over: Obj = {}, k: Obj = {}) => run({ timekeepers: [keeper(k)], entries, ...over });
const hours = (minutes: number, over: Obj = {}) => mk([at(minutes)], over).rows[0].hours;
const throwsOn = (time: Obj, text: string) => {
  let message = "";
  try {
    run(time);
  } catch (e) {
    message = (e as Error).message;
  }
  assertStringIncludes(message, text);
};

const preview = COMPONENTS.find((c) => c.id === "timesheet-fee-summary")!.previewData as { time: Obj };

test("the catalog preview", () => {
  const o = run(preview.time);
  assertEquals(
    o.rows.map((r) => [r.label, r.hours, r.rate, r.amount]),
    [
      ["Alexandra Montgomery-Fitzwilliam KC (rate from 1 April 2025)", "16.10", "1,150.00", "18,515.00"],
      ["Alexandra Montgomery-Fitzwilliam KC (rate from 1 October 2026)", "9.80", "1,295.00", "12,691.00"],
      ["Alexandra Montgomery-Fitzwilliam KC (rate from 1 October 2026, travel at 50%)", "3.10", "647.50", "2,007.25"],
      ["Chukwuemeka Okonkwo-Hargreaves", "20.10", "685.00", "13,768.50"],
      ["Chukwuemeka Okonkwo-Hargreaves (travel at 50%)", "3.10", "342.50", "1,061.75"],
      ["Chukwuemeka Okonkwo-Hargreaves (no charge)", "0.80", "0.00", "0.00"],
      ["Priyanka Subramaniam-Lindqvist", "10.10", "245.00", "2,474.50"],
    ],
  );
  assertEquals(o.rows[0].role, "Partner, Head of Commercial Litigation");
  assertEquals(o.adjustments, [
    { label: "Reduction per billing guidelines: Alexandra Montgomery-Fitzwilliam KC (rate from 1 October 2026), 9.80 hours at 1,200.00", amount: "-931.00" },
    { label: "Reduction per billing guidelines: Alexandra Montgomery-Fitzwilliam KC (rate from 1 October 2026, travel at 50%), 3.10 hours at 600.00", amount: "-147.25" },
    { label: "Write-down: Agreed reduction for duplicated research on the limitation point", amount: "-4,750.00" },
  ]);
  assertEquals([o.subtotal, o.subtotalMinor, o.currency], ["44,689.75", 4468975, "GBP"]);
  assertEquals(
    o.basis,
    "Each time entry is rounded up to the next 6 minutes (0.10 hours), with a minimum of 6 minutes per entry. Travel is billed at 50% of the hourly rate.",
  );
});

test("rounding is per entry, not per total", () => {
  const o = mk([at(7), at(7), at(7)]);
  assertEquals([o.rows.length, o.rows[0].hours, o.rows[0].amount], [1, "0.60", "150.00"]);
});

test("rounding up, nearest, minimum and fractions", () => {
  assertEquals([1, 6, 7].map((m) => hours(m)), ["0.10", "0.10", "0.20"]);
  const nearest = { billing: { rounding: "nearest" } };
  assertEquals([8, 9, 2].map((m) => hours(m, nearest)), ["0.10", "0.20", "0.10"]);
  assertEquals(hours(0.5), "0.10");
  assertEquals(hours(16, { billing: { incrementMinutes: 15 } }), "0.50");
});

test("float exactness", () => {
  const o = mk([at(6), at(12)]);
  assertEquals([o.rows[0].hours, o.rows[0].amount, o.subtotalMinor], ["0.30", "75.00", 7500]);
});

test("a bad increment throws", () => {
  for (const incrementMinutes of [10, 0, -6, 4.5]) {
    throwsOn({ timekeepers: [keeper()], entries: [], billing: { incrementMinutes } }, "time.billing.incrementMinutes");
  }
});

test("a rate change splits the timekeeper, in any rate order", () => {
  const rates = [{ from: "2026-01-01", rate: 400 }, { from: "2026-10-01", rate: 450 }];
  const entries = [at(60, { date: "2026-10-01" }), at(60, { date: "2026-09-30" })];
  for (const r of [rates, [...rates].reverse()]) {
    const o = mk(entries, {}, { rates: r });
    assertEquals(
      o.rows.map((x) => [x.label, x.hours, x.rate, x.amount]),
      [["Ann (rate from 1 January 2026)", "1.00", "400.00", "400.00"], ["Ann (rate from 1 October 2026)", "1.00", "450.00", "450.00"]],
    );
    assertEquals(o.subtotal, "850.00");
  }
  throwsOn({ timekeepers: [keeper()], entries: [at(60, { date: "2025-12-31" })] }, "time.entries[0]");
});

test("an approved rate below standard is a reduction line, above is nothing", () => {
  const o = mk([at(120)], {}, { rates: [{ from: "2026-01-01", rate: 500 }], approvedRate: 450 });
  assertEquals([o.rows[0].hours, o.rows[0].rate, o.rows[0].amount], ["2.00", "500.00", "1,000.00"]);
  assertEquals(o.adjustments, [{ label: "Reduction per billing guidelines: Ann, 2.00 hours at 450.00", amount: "-100.00" }]);
  assertEquals(o.subtotal, "900.00");
  const above = mk([at(120)], {}, { rates: [{ from: "2026-01-01", rate: 500 }], approvedRate: 550 });
  assertEquals([above.rows[0].amount, above.adjustments.length], ["1,000.00", 0]);
});

test("travel is its own row at the factor", () => {
  const o = mk([at(60, { kind: "travel" })], { billing: { travelFactor: 0.5 } }, { rates: [{ from: "2026-01-01", rate: 401.01 }] });
  assertEquals([o.rows[0].label, o.rows[0].rate, o.rows[0].amount], ["Ann (travel at 50%)", "200.51", "200.51"]);
  const full = mk([at(60, { kind: "travel" })]);
  assertEquals([full.rows[0].label, full.rows[0].rate], ["Ann (travel at 100%)", "250.00"]);
  assertEquals(full.basis.includes("Travel"), false);
  for (const travelFactor of [0, 1.5, -0.5]) throwsOn({ timekeepers: [keeper()], entries: [], billing: { travelFactor } }, "travelFactor");
});

test("no-charge time is listed at 0.00", () => {
  const o = mk([at(60), at(30, { kind: "no-charge" })]);
  assertEquals(o.rows.map((r) => [r.label, r.hours, r.rate, r.amount]), [
    ["Ann", "1.00", "250.00", "250.00"],
    ["Ann (no charge)", "0.50", "0.00", "0.00"],
  ]);
  assertEquals(o.subtotal, "250.00");
});

test("a write-down", () => {
  const o = mk([at(60)], { billing: { writeDown: { amount: 50, reason: " Duplicated review " } } });
  assertEquals(o.adjustments, [{ label: "Write-down: Duplicated review", amount: "-50.00" }]);
  assertEquals(o.subtotal, "200.00");
  assertEquals(mk([at(60)], { billing: { writeDown: { amount: 0, reason: "x" } } }).adjustments, []);
  for (const writeDown of [{ amount: -1, reason: "x" }, { amount: 5, reason: "" }, { amount: 5, reason: "  " }, { amount: 300, reason: "x" }]) {
    throwsOn({ timekeepers: [keeper()], entries: [at(60)], billing: { writeDown } }, "time.billing.writeDown");
  }
});

test("a write-down is limited by the fees after reductions", () => {
  const rates = { rates: [{ from: "2026-01-01", rate: 500 }], approvedRate: 450 };
  // Rows total 1,000.00 and the reduction takes 100.00, so 950 is under the rows but over the fees.
  throwsOn({ timekeepers: [keeper(rates)], entries: [at(120)], billing: { writeDown: { amount: 950, reason: "x" } } }, "time.billing.writeDown");
  assertEquals(mk([at(120)], { billing: { writeDown: { amount: 900, reason: "x" } } }, rates).subtotal, "0.00");
});

test("rows in a rate period go work, travel, no-charge whatever the entry order", () => {
  const o = mk([at(60, { kind: "no-charge" }), at(60, { kind: "travel" }), at(60)]);
  assertEquals(o.rows.map((r) => r.label.includes("no charge") ? 2 : r.label.includes("travel") ? 1 : 0), [0, 1, 2]);
});

test("no entries", () => {
  const o = mk([]);
  assertEquals([o.hasRows, o.rows, o.adjustments, o.basis, o.subtotal], [false, [], [], "", "0.00"]);
  throwsOn({ timekeepers: [keeper()], entries: [], billing: { writeDown: { amount: 5, reason: "x" } } }, "writeDown");
  throwsOn({ timekeepers: [keeper()], entries: [], billing: { incrementMinutes: 4 } }, "incrementMinutes");
});

test("bad entries and timekeepers throw naming the field", () => {
  for (const m of [0, -30, NaN]) throwsOn({ timekeepers: [keeper()], entries: [at(60), at(m)] }, "time.entries[1].minutes");
  throwsOn({ timekeepers: [keeper()], entries: [at(60, { timekeeperId: "XYZ" })] }, 'time.entries[0].timekeeperId "XYZ" is not in time.timekeepers');
  throwsOn({ timekeepers: [keeper(), keeper()], entries: [] }, "time.timekeepers[1].id");
  throwsOn({ timekeepers: [keeper({ rates: [{ from: "2026-01-01", rate: 1 }, { from: "2026-01-01T10:00", rate: 2 }] })], entries: [] }, "duplicate");
  throwsOn({ timekeepers: [keeper()], entries: [at(60, { kind: "expense" })] }, "time.entries[0].kind");
  throwsOn({ timekeepers: [keeper()], entries: [at(60, { date: "2026-02-30" })] }, "time.entries[0].date");
  throwsOn({ timekeepers: [keeper()] }, "time.entries");
  throwsOn({ timekeepers: [keeper()], entries: [], currency: "pounds" }, "time.currency");
});

test("missing optional fields mean the defaults", () => {
  const o = mk([at(60)]);
  assertEquals([o.currency, o.adjustments.length, o.basis.startsWith("Each time entry is rounded up to the next 6 minutes")], ["GBP", 0, true]);
  assertEquals([o.rows[0].label, o.rows[0].hours, o.rows[0].rate, o.rows[0].amount, o.subtotal], ["Ann", "1.00", "250.00", "250.00", "250.00"]);
});

test("a timekeeper without entries prints nothing", () => {
  const o = run({ timekeepers: [keeper({ id: "A" }), keeper({ id: "B", name: "Bo" }), keeper({ id: "C", name: "Cy" })], entries: [at(60, { timekeeperId: "C" }), at(60)] });
  assertEquals(o.rows.map((r) => r.label), ["Ann", "Cy"]);
});

test("zero-decimal and sub-minor currencies", () => {
  const jpy = mk([at(60)], { currency: "jpy" }, { rates: [{ from: "2026-01-01", rate: 45000 }] });
  assertEquals([jpy.currency, jpy.rows[0].rate, jpy.rows[0].amount, jpy.rows[0].hours], ["JPY", "45,000", "45,000", "1.00"]);
  throwsOn({ timekeepers: [keeper({ rates: [{ from: "2026-01-01", rate: 450.5 }] })], entries: [], currency: "JPY" }, "450.5");
  throwsOn({ timekeepers: [keeper({ rates: [{ from: "2026-01-01", rate: 450.005 }] })], entries: [] }, "450.005");
});

test("large figures", () => {
  const o = mk([at(72270)], {}, { rates: [{ from: "2026-01-01", rate: 1295 }] });
  assertEquals([o.rows[0].hours, o.rows[0].amount], ["1,204.50", "1,559,827.50"]);
});

test("many entries across shuffled timekeepers", () => {
  const ids = ["A", "B", "C", "D", "E"];
  const timekeepers = ids.map((id, i) => ({
    id, name: id, role: "r",
    rates: [{ from: "2026-01-01", rate: 100 + i }, { from: "2026-06-01", rate: 200 + i }],
    approvedRate: 150 + i,
  }));
  const kinds = ["work", "travel", "no-charge"];
  const entries = Array.from({ length: 200 }, (_, i) => ({
    date: i % 2 ? "2026-03-01" : "2026-07-01",
    timekeeperId: ids[(i * 7) % 5],
    minutes: 1 + ((i * 13) % 97),
    kind: kinds[(i * 5) % 3],
  })).sort((a, b) => Math.sin(a.minutes * 7 + a.date.length) - Math.sin(b.minutes * 3));
  const o = run({ timekeepers, entries, billing: { travelFactor: 0.5 } });
  const order = o.rows.map((r) => ids.indexOf(r.label.split(" ")[0]));
  assertEquals(order, [...order].sort((a, b) => a - b));
  const n = (s: string) => Math.round(Number(s.replace(/,/g, "")) * 100);
  const sum = (a: Array<{ amount: string }>) => a.reduce((t, r) => t + n(r.amount), 0);
  assertEquals(o.subtotalMinor, sum(o.rows) + sum(o.adjustments));
  for (const r of o.rows) {
    const exact = (Number(r.hours.replace(/,/g, "")) * 100 * n(r.rate)) / 100;
    assertEquals(Math.abs(exact - n(r.amount)) <= 0.5, true);
  }
});

// ---------------------------------------------------------------------------
// The publish path
// ---------------------------------------------------------------------------

const tpl = () => {
  const Body = TimesheetFeeSummary as unknown as () => never;
  return template(<Document title="Fees"><Body /></Document>);
};
const published = await buildDocument(tpl(), createPublishData(), {
  branchMode: "publish",
  deriverMode: "preserve",
  dynamicMode: "resolve",
  aiClient: { generateParagraph: () => "" },
});
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ timesheetFeeSummary: mod.run }));

async function resolveJson(time: Obj): Promise<string> {
  const data = { time };
  const resolved = await resolveDocument(
    published,
    { ctx: { ...data }, derived: {}, dataProvider: new InMemoryDataProvider(data), aiClient: { generateParagraph: () => "" } },
    { derivers },
  );
  return JSON.stringify(resolved.nodes);
}
const count = (json: string, id: string) => json.split(`"id":"${id}`).length - 1;

test("the node publishes three branches", async () => {
  const source = await readFile(join(registryRoot(), nodeSource), "utf8");
  assertEquals(transformDocumentSource(source, { fileName: nodeSource }).branches, 3);
});

test("a published document resolves through the shipped deriver", async () => {
  const json = await resolveJson(preview.time);
  assertEquals(count(json, "timesheet-fee-line"), 7);
  assertEquals(count(json, "timesheet-fee-adjustment"), 3);
  assertStringIncludes(json, "44,689.75");
  assertEquals(json.includes("timesheet-fee-none"), false);

  const none = await resolveJson({ ...preview.time, entries: [], billing: {} });
  assertStringIncludes(none, "No time was recorded for this period.");
  assertEquals(none.includes("timesheet-fee-table"), false);
  assertEquals(none.includes("timesheet-fee-basis"), false);
});

test("basis wording: nearest, hours of the increment, travel only when billed", () => {
  assertEquals(
    mk([at(8)], { billing: { rounding: "nearest" } }).basis,
    "Each time entry is rounded to the nearest 6 minutes (0.10 hours), with a minimum of 6 minutes per entry.",
  );
  assertStringIncludes(mk([at(8)], { billing: { incrementMinutes: 15 } }).basis, "(0.25 hours)");
  assertEquals(mk([at(60)], { billing: { travelFactor: 0.5 } }).basis.includes("Travel"), false);
});

test("a rate difference below one minor unit prints no reduction line", () => {
  const o = mk([at(6)], {}, { rates: [{ from: "2026-01-01", rate: 10.01 }], approvedRate: 10 });
  assertEquals(o.adjustments, []);
});

test("more bad billing and timekeeper input throws naming the field", () => {
  const ks = (more: Obj) => [keeper(more)];
  throwsOn({ timekeepers: ks({}), entries: [], billing: { rounding: "down" } }, "time.billing.rounding");
  throwsOn({ timekeepers: ks({ rates: [] }), entries: [] }, "time.timekeepers[0].rates");
  for (const rate of [-1, Infinity]) {
    throwsOn({ timekeepers: ks({ rates: [{ from: "2026-01-01", rate }] }), entries: [] }, "time.timekeepers[0].rates[0].rate");
  }
  for (const approvedRate of [-1, 450.005]) throwsOn({ timekeepers: ks({ approvedRate }), entries: [] }, "approvedRate");
  throwsOn({ timekeepers: ks({ id: "" }), entries: [] }, "time.timekeepers[0].id");
  throwsOn({ timekeepers: ks({ name: "" }), entries: [] }, "time.timekeepers[0].name");
  throwsOn({ timekeepers: "x", entries: [] }, "time.timekeepers");
});
