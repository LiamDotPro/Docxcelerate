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

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/dunning-reminder.deriver.ts")
>("components/derivers/dunning-reminder.deriver.ts");
const { DunningReminder } = await loadRegistryModule<
  typeof import("../registry/components/dunning-reminder/dunning-reminder.node.tsx")
>("components/dunning-reminder/dunning-reminder.node.tsx");
const { COMPONENTS } = await import("../src/registry/catalog.ts");

const registry = await createDeriverRegistryFromBundle(createDeriverBundle({ dunningReminder: mod.run }));
const preview = COMPONENTS.find((c) => c.id === "dunning-reminder")!.previewData as { reminder: Obj };

// deno-lint-ignore no-explicit-any -- request data is unchecked JSON.
type Obj = Record<string, any>;
const state = {} as RuntimeState;
const run = (reminder: Obj) => mod.run([reminder], state) as Obj;
// 19 November 2026 is a Thursday; sending on the 20th (Friday) unless said.
const base: Obj = { sendDate: "2026-11-20", paymentsCheckedTo: "2026-11-19", currency: "GBP" };
const old = (extra: Obj = {}): Obj => ({ number: "A", dueDate: "2026-10-01", openAmount: 100, ...extra });
const withInvoices = (invoices: Obj[], extra: Obj = {}) => run({ ...base, ...extra, invoices });
const final = (extra: Obj = {}) => old({ lastLevel: 2, lastReminderDate: "2026-10-01", ...extra });

// A sync throw, as the promise assertRejects wants.
const throws = (fn: () => unknown, message: string) => assertRejects(() => Promise.resolve().then(fn), Error, message);

const emptyResult = {
  outcome: "none",
  level: 0,
  title: "",
  opening: "",
  invoices: Array.from({ length: 8 }, () => ({ number: "", dueDate: "", overdue: "", amount: "", line: "" })),
  overflow: "",
  total: "",
  request: "",
  contact: "",
  consequence: "",
  paidSince: "",
  promptFacts: "",
  deadline: "",
  sent: [],
};

test("nothing to chase prints nothing", () => {
  assertEquals(run(base), emptyResult);
  assertEquals(withInvoices([]), emptyResult);
  assertEquals(
    withInvoices([
      old({ disputed: true }),
      old({ onHold: true }),
      old({ openAmount: 0 }),
      old({ openAmount: -50 }),
      old({ dueDate: "2026-11-15" }),
      old({ lastLevel: 1, lastReminderDate: "2026-11-15" }),
      old({ lastLevel: 3, lastReminderDate: "2026-01-01" }),
    ]),
    emptyResult,
  );
});

test("a credit is not netted against another invoice", () => {
  const r = withInvoices([old({ openAmount: -50 }), old({ number: "B", openAmount: 100 })]);
  assertEquals(r.total, "£100.00");
});

test("rounding is half-up in minor units", () => {
  assertEquals(withInvoices([old({ openAmount: 0.004 })]).outcome, "none");
  assertEquals(withInvoices([old({ openAmount: 0.005 })]).total, "£0.01");
  assertEquals(withInvoices([old({ openAmount: 0.1 }), old({ number: "B", openAmount: 0.2 })]).total, "£0.30");
  assertEquals(withInvoices([old({ openAmount: 1500.4 })], { currency: "JPY" }).total, "JP¥1,500");
});

test("grace days and the due date", () => {
  assertEquals(withInvoices([old({ dueDate: "2026-11-13" })], { graceDays: 7 }).outcome, "remind");
  assertEquals(withInvoices([old({ dueDate: "2026-11-14" })], { graceDays: 7 }).outcome, "none");
  assertEquals(withInvoices([old({ dueDate: "2026-11-20" })], { graceDays: 0 }).outcome, "none");
  assertEquals(withInvoices([old({ dueDate: "2026-11-19" })], { graceDays: 0 }).outcome, "remind");
});

test("the interval between reminders", () => {
  const at = (date: string, extra: Obj = {}) => withInvoices([old({ lastLevel: 1, lastReminderDate: date })], extra);
  assertEquals(at("2026-11-07").outcome, "none");
  assertEquals(at("2026-11-06").level, 2);
  assertEquals(at("2026-11-19").outcome, "none");
  assertEquals(at("2026-11-13", { minIntervalDays: 7 }).outcome, "remind");
  assertEquals(withInvoices([old({ lastLevel: 3, lastReminderDate: "2026-01-01" })]).outcome, "none");
});

test("bad levels and missing dates throw, naming the field", async () => {
  const msg = (inv: Obj) => throws(() => withInvoices([old(), old(), inv]), "reminder.invoices[2]");
  await throws(
    () => withInvoices([old({ lastLevel: 2 })]),
    "reminder.invoices[0].lastReminderDate is required when lastLevel is 2",
  );
  await msg({ lastLevel: 4, lastReminderDate: "2026-01-01" });
  await msg({ lastLevel: 1.5, lastReminderDate: "2026-01-01" });
  await msg({ number: " " });
  await msg({ dueDate: "soon" });
  await msg({ openAmount: NaN });
  await msg({ lastLevel: 1, lastReminderDate: "2026-12-01" });
});

test("bad data throws even when nothing would be sent", async () => {
  await throws(() => run({ ...base, sendDate: undefined }), "reminder.sendDate");
  await throws(() => run({ ...base, paymentsCheckedTo: "2026-11-21" }), "paymentsCheckedTo");
  await throws(() => run({ ...base, currency: "GB" }), "reminder.currency");
  await throws(() => run({ ...base, currency: undefined }), "reminder.currency");
  await throws(() => run({ ...base, graceDays: -1 }), "graceDays");
  await throws(() => run({ ...base, minIntervalDays: 0 }), "minIntervalDays");
  await throws(() => run({ ...base, deadlineDays: { final: 0 } }), "deadlineDays.final");
  await throws(() => run({ ...base, legalFloor: "FR" }), "legalFloor");
  await throws(() => run({ ...base, channel: "fax" }), "channel");
  await throws(() => run({ ...base, intendedConsequences: ["bailiff"] }), "intendedConsequences");
  await throws(() => run({ ...base, customerSince: "2027-01-01" }), "customerSince");
  await throws(() => run({ ...base, customerSince: "x" }), "customerSince");
});

test("one first-time invoice", () => {
  const r = withInvoices([old({ dueDate: "2026-11-01" })], { intendedConsequences: ["court"] });
  assertEquals([r.level, r.title, r.contact, r.consequence], [1, "Payment reminder", "", ""]);
  assertEquals(r.opening, "Our records show that the invoice below is past its due date. This may simply have been overlooked, so this is a friendly reminder.");
  assertEquals(r.invoices[0].line, "Invoice A, due 1 November 2026, 19 days overdue: £100.00");
  assertEquals(r.invoices[1].number, "");
  assertEquals(r.sent, [{ number: "A", level: 1 }]);
  assertEquals(withInvoices([old({ dueDate: "2026-11-19" })], { graceDays: 1 }).invoices[0].overdue, "1 day");
});

test("a second reminder", () => {
  const r = withInvoices([old({ lastLevel: 1, lastReminderDate: "2026-10-01" })]);
  assertEquals(r.title, "Second reminder");
  assertEquals(r.opening, "This is our second reminder about the invoice below. We have not yet received payment.");
  assertStringIncludes(r.contact, "If something is wrong with the invoice,");
  assertEquals(r.request, "Please pay £100.00 by 30 November 2026.");
});

test("mixed levels", () => {
  const r = withInvoices([old({ number: "B" }), final({ number: "C" })], { intendedConsequences: ["court"] });
  assertEquals(r.level, 3);
  assertStringIncludes(r.opening, "about some of the invoices below. Despite earlier reminders, they remain unpaid.");
  assertEquals(r.consequence, "If invoice C is not paid by 27 November 2026, we intend to start a court claim to recover it.");
  assertStringIncludes(r.contact, "any of these invoices");
  assertEquals(r.request, "Please make sure £200.00 reaches us by 27 November 2026.");
});

test("consequences are deduplicated, ordered and only for a final letter", () => {
  const r = withInvoices([final(), final({ number: "B" })], { intendedConsequences: ["court", "agency", "agency"] });
  assertEquals(
    r.consequence,
    "If we have not received payment by 27 November 2026, we intend to pass the debt to a collection agency and start a court claim to recover it.",
  );
  assertEquals(withInvoices([final()], { intendedConsequences: ["court", "suspension", "agency"] }).consequence,
    "If we have not received payment by 27 November 2026, we intend to suspend your service, pass the debt to a collection agency and start a court claim to recover it.");
  assertEquals(withInvoices([final()]).consequence, "");
  assertEquals(withInvoices([final()], { intendedConsequences: [] }).consequence, "");
  assertEquals(withInvoices([old()], { intendedConsequences: ["court"] }).consequence, "");
  assertEquals(withInvoices([old({ lastLevel: 1, lastReminderDate: "2026-10-01" })], { intendedConsequences: ["court"] }).consequence, "");
});

test("deadlines", () => {
  const dl = (invoice: Obj, extra: Obj = {}) => withInvoices([invoice], extra).deadline;
  const l1 = old();
  const l2 = old({ lastLevel: 1, lastReminderDate: "2026-10-01" });
  const l3 = final();
  assertEquals([dl(l1), dl(l2), dl(l3)], ["4 December 2026", "30 November 2026", "27 November 2026"]);
  assertEquals(dl(l2, { sendDate: "2026-11-18", paymentsCheckedTo: "2026-11-18" }), "30 November 2026");
  assertEquals(dl(l2, { sendDate: "2026-11-19" }), "30 November 2026");
  assertEquals([dl(l3, { legalFloor: "NL-consumer", channel: "email" }), dl(l3, { legalFloor: "NL-consumer" })], [
    "4 December 2026",
    "7 December 2026",
  ]);
  assertEquals([dl(l3, { legalFloor: "BE-consumer", channel: "email" }), dl(l3, { legalFloor: "BE-consumer" })], [
    "7 December 2026",
    "9 December 2026",
  ]);
  assertEquals(dl(l3, { legalFloor: "UK-pre-action" }), "21 December 2026");
  assertEquals(dl(l3, { legalFloor: "UK-pre-action", deadlineDays: { final: 40 } }), "30 December 2026");
});

test("nine invoices fill eight slots, oldest first, and overflow", () => {
  const invoices = Array.from({ length: 9 }, (_, i) => old({ number: "N" + (9 - i), dueDate: i < 2 ? "2026-09-01" : "2026-10-01" }));
  const r = withInvoices(invoices);
  assertEquals(r.invoices.map((s: Obj) => s.number), ["N8", "N9", "N1", "N2", "N3", "N4", "N5", "N6"]);
  assertEquals(r.overflow, "1 further invoice, totalling £100.00, is also overdue and included in the total.");
  assertEquals(r.total, "£900.00");
  assertEquals(r.sent.length, 9);
  assertStringIncludes(withInvoices([...invoices, old({ number: "Z" })]).overflow, "2 further invoices, totalling £200.00, are also");
});

test("customerSince and paidSince", () => {
  const facts = (customerSince?: string) => withInvoices([old()], { customerSince }).promptFacts;
  assertEquals(facts(), "This is the first reminder. The oldest invoice in it is 50 days overdue.");
  assertStringIncludes(facts("2025-12-01"), "with us for 11 months.");
  assertStringIncludes(facts("2026-11-01"), "with us for less than a month.");
  assertStringIncludes(facts("2019-03-01"), "with us for 7 years.");
  assertStringIncludes(facts("2025-11-20"), "with us for 1 year.");
  assertStringIncludes(withInvoices([final()]).promptFacts, "final reminder");
  for (const invoice of [old(), final()]) {
    assertStringIncludes(withInvoices([invoice]).paidSince, "up to and including 19 November 2026.");
  }
});

test("the preview data gives the expected letter", () => {
  const r = run(preview.reminder);
  assertEquals(r.invoices.map((s: Obj) => s.number.slice(8)).filter(Boolean), ["07-004817/REV-2", "08-005102-FACILITIES-NORTH", "09-005390", "10-005702", "11-005940"]);
  assertEquals(r.invoices.map((s: Obj) => s.overdue).slice(0, 5), ["112 days", "81 days", "51 days", "20 days", "10 days"]);
  assertEquals([r.level, r.title, r.total, r.deadline], [3, "Final reminder", "£321,205.74", "21 December 2026"]);
  assertEquals(
    r.consequence,
    "If invoices SI-2026-07-004817/REV-2 and SI-2026-08-005102-FACILITIES-NORTH are not paid by 21 December 2026, we intend to pass the debt to a collection agency and start a court claim to recover it.",
  );
  assertEquals(r.sent.map((s: Obj) => s.level), [3, 3, 2, 1, 1]);
  assertStringIncludes(r.promptFacts, "112 days overdue. The customer has been with us for 7 years.");
});

// The node.
type Arrangement = "table" | "stacked";
const documentOf = (arrangement?: Arrangement, generateOpening?: boolean) =>
  template(
    <Document title="Reminder">
      <DunningReminder arrangement={arrangement} generateOpening={generateOpening} />
    </Document>,
  );

function ids(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { id?: string; children?: DocumentNode[]; rows?: DocumentNode[] };
    return [...(own.id ? [own.id] : []), ...ids(own.children ?? []), ...ids(own.rows ?? [])];
  });
}

const build = (data: Obj, arrangement?: Arrangement, generateOpening?: boolean) =>
  buildDocument(documentOf(arrangement, generateOpening), data as never, { dynamicMode: "placeholder" });

test("both arrangements build, with their own ids", async () => {
  const table = ids((await build(preview, "table")).nodes);
  const stacked = ids((await build(preview, "stacked")).nodes);
  assertEquals(table.includes("dunning-reminder-invoices-header"), true);
  assertEquals(stacked.includes("dunning-reminder-invoices-header"), false);
  for (const list of [table, stacked]) {
    for (const id of ["dunning-reminder", "dunning-reminder-opening", "dunning-reminder-invoice-5", "dunning-reminder-total",
      "dunning-reminder-request", "dunning-reminder-contact", "dunning-reminder-consequence", "dunning-reminder-paid-since"]) {
      assertEquals([id, list.includes(id)], [id, true]);
    }
    assertEquals(list.includes("dunning-reminder-invoice-6"), false);
    assertEquals(list.includes("dunning-reminder-overflow"), false);
  }
  assertEquals(JSON.stringify((await build(preview, "table")).nodes) === JSON.stringify((await build(preview, "stacked")).nodes), false);
  assertEquals(JSON.stringify((await build(preview)).nodes), JSON.stringify((await build(preview, "table")).nodes));
});

test("the fixed opening has no dynamic node, the generated one shows its placeholder", async () => {
  const fixed = JSON.stringify((await build(preview)).nodes);
  assertStringIncludes(fixed, "This is our final reminder");
  assertEquals(fixed.includes("dunning-reminder-opening-"), false);
  const generated = JSON.stringify((await build(preview, undefined, true)).nodes);
  assertStringIncludes(generated, "dunning-reminder-opening-final");
  assertStringIncludes(generated, "A serious opening, written per document, stating this is the final reminder.");
  assertEquals(generated.includes('"dunning-reminder-opening"'), false);
});

test("publishing keeps every generated arm and points the facts at the deriver", async () => {
  const published = await buildDocument(documentOf("table", true), createPublishData() as never, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const json = JSON.stringify(published.nodes);
  for (const level of ["first", "second", "final"]) assertStringIncludes(json, `dunning-reminder-opening-${level}`);
  assertStringIncludes(json, "{{derived.dunningReminder.promptFacts}}");
});

test("a published node resolves against a bundled deriver", async () => {
  const published = await buildDocument(documentOf("table"), createPublishData() as never, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const resolve = async (data: Obj) => {
    const runtime: RuntimeState = {
      ctx: { ...data },
      derived: {},
      dataProvider: new InMemoryDataProvider(data),
      aiClient: { generateParagraph: () => "" },
    };
    return ids((await resolveDocument(published, runtime, { derivers: registry })).nodes);
  };

  const level1 = await resolve({ reminder: { ...base, invoices: [old({ dueDate: "2026-11-01" })] } });
  assertEquals(level1.includes("dunning-reminder-opening"), true);
  assertEquals(level1.includes("dunning-reminder-contact"), false);
  assertEquals(level1.includes("dunning-reminder-consequence"), false);

  const full = await resolve(preview);
  for (const id of ["dunning-reminder-contact", "dunning-reminder-consequence", "dunning-reminder-total"]) {
    assertEquals([id, full.includes(id)], [id, true]);
  }

  assertEquals((await resolve({ reminder: base })).includes("dunning-reminder"), false);
});
