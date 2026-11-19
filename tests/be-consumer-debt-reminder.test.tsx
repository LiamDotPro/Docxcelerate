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
import { assertEquals } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const nodeSource = "components/be-consumer-debt-reminder/be-consumer-debt-reminder.node.tsx";
const { BeConsumerDebtReminder } = await loadRegistryModule<
  typeof import("../registry/components/be-consumer-debt-reminder/be-consumer-debt-reminder.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/be-consumer-debt-reminder.deriver.ts")
>("components/derivers/be-consumer-debt-reminder.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ beConsumerDebtReminder: mod.run }));
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
interface Out {
  outcome: string;
  amountDue: string;
  periodStart: string;
  lastDay: string;
  lumpSum: string;
  rate: string;
  interestFrom: string;
  interest: string;
  interestDays: number;
  interestIncomplete: boolean;
  charged: boolean;
  addsNothing: boolean;
  missing: string;
  compensation: string;
  invoicesLabel: string;
  total: string;
}

const run = (debt: Obj): Out => mod.run([debt], {} as RuntimeState) as Out;
const throwsOn = (debt: Obj, field: string) => {
  try {
    run(debt);
  } catch (e) {
    assertEquals((e as Error).message.includes(field), true);
    return;
  }
  throw new Error("expected a throw naming " + field);
};

const base: Obj = {
  debtorType: "consumer",
  stage: "first",
  asOf: "2026-11-19",
  amountDue: 1000,
  dueDate: "2026-11-01",
  creditorName: "Meubles Janssens SRL",
  creditorEnterpriseNumber: "0712.345.678",
  description: "a sofa (invoice F-1)",
};
const electronic = { sentDate: "2026-11-19", channel: "electronic" };
const both = { contractProvidesDamages: true, contractProvidesInterest: true, referenceRate: 4.25 };
const paper = (sentDate: string, more: Obj = {}) => run({ ...base, firstReminder: { sentDate, channel: "paper" }, ...more });

test("the period starts on the third working day after a paper reminder", () => {
  const cases: Array<[string, string, string, Obj?]> = [
    ["2026-11-30", "3 December 2026", "16 December 2026"],
    ["2026-12-24", "30 December 2026", "12 January 2027"],
    ["2027-03-25", "31 March 2027", "13 April 2027"],
    ["2026-11-27", "2 December 2026", "15 December 2026"],
  ];
  for (const [sent, start, last, more] of cases) {
    const o = paper(sent, more);
    assertEquals([o.periodStart, o.lastDay], [start, last]);
  }
  assertEquals(paper("2026-11-30", { holidays: ["2026-12-02"] }).periodStart, "4 December 2026");
});

test("an electronic reminder starts the day after it is sent", () => {
  const o = run({ ...base, firstReminder: electronic });
  assertEquals([o.periodStart, o.lastDay], ["20 November 2026", "3 December 2026"]);
});

test("lump-sum bands", () => {
  const sum = (amountDue: number, more: Obj = {}) =>
    run({ ...base, amountDue, dueDate: "2020-01-01", firstReminder: electronic, contractProvidesDamages: true, ...more })
      .lumpSum;
  const bands: Array<[number, string]> = [
    [0.01, "€20.00"],
    [150, "€20.00"],
    [150.01, "€30.00"],
    [300, "€45.00"],
    [500, "€65.00"],
    [500.01, "€65.00"],
    [1000, "€90.00"],
    [39200, "€2,000.00"],
    [1284750.37, "€2,000.00"],
  ];
  for (const [amount, expected] of bands) assertEquals(sum(amount), expected);
  assertEquals(sum(1000, { agreedLumpSum: 15 }), "€15.00");
  assertEquals(sum(1000, { agreedLumpSum: 500 }), "€90.00");
  assertEquals(sum(1000, { agreedLumpSum: 0 }), "");
  assertEquals(sum(1000, { contractProvidesDamages: false, agreedLumpSum: 15 }), "");
  assertEquals(sum(1000, { contractProvidesDamages: undefined, agreedLumpSum: 15 }), "");
  throwsOn({ ...base, firstReminder: electronic, contractProvidesDamages: true, agreedLumpSum: -1 }, "agreedLumpSum");
});

test("the first reminder announces what a lapse would cost", () => {
  const o = run({ ...base, firstReminder: electronic, ...both });
  assertEquals(
    [o.outcome, o.lumpSum, o.rate, o.interestFrom, o.addsNothing, o.interest],
    ["first", "€90.00", "12.25%", "4 December 2026", false, ""],
  );
  assertEquals(run({ ...base, firstReminder: electronic, ...both, creditorIsSme: true }).interestFrom, "20 November 2026");
  const none = run({ ...base, firstReminder: electronic });
  assertEquals([none.outcome, none.addsNothing, none.lumpSum, none.rate], ["first", true, "", ""]);
  const rateOf = (more: Obj) => run({ ...base, firstReminder: electronic, ...both, ...more }).rate;
  assertEquals(rateOf({ agreedRate: 10 }), "10%");
  assertEquals(rateOf({ agreedRate: 20 }), "12.25%");
  assertEquals(rateOf({ referenceRate: -9 }), "");
});

test("a first reminder missing a mention is incomplete", () => {
  const o = run({ ...base, creditorEnterpriseNumber: undefined, description: "", firstReminder: electronic });
  assertEquals([o.outcome, o.missing], [
    "first-incomplete",
    "the creditor's enterprise number and a description of what the debt is for",
  ]);
  assertEquals(run({ ...base, creditorName: "", firstReminder: electronic }).outcome, "first-incomplete");
  assertEquals(run({ ...base }).missing, "the date and channel the reminder is sent by");
  const rate = run({ ...base, firstReminder: electronic, contractProvidesInterest: true });
  assertEquals(rate.missing, "the reference rate that caps the late interest");
});

const later = (asOf: string, more: Obj = {}) =>
  run({ ...base, stage: "subsequent", asOf, firstReminder: electronic, ...more });

test("a later reminder waits for the period to lapse", () => {
  const o = later("2026-12-03", both);
  assertEquals([o.outcome, o.charged, o.lumpSum, o.interest, o.lastDay], ["waiting", false, "", "", "3 December 2026"]);
  assertEquals(run({ ...base, stage: "subsequent" }).outcome, "no-first-reminder");
});

test("interest is simple, on the principal only, actual/365", () => {
  const one = later("2026-12-04", { contractProvidesInterest: true, referenceRate: 4.25 });
  assertEquals([one.outcome, one.interestDays, one.interest], ["lapsed", 1, "€0.33"]);
  const o = later("2027-01-02", both);
  assertEquals([o.lumpSum, o.interestDays, o.interest, o.total, o.charged], ["€90.00", 30, "€10.06", "€1,100.06", true]);
  const sme = later("2027-01-02", { ...both, creditorIsSme: true });
  assertEquals([sme.interestFrom, sme.interestDays, sme.interest], ["20 November 2026", 44, "€14.76"]);
});

test("a lapsed reminder with no clause, or no rate, adds only what it can", () => {
  const o = later("2027-01-02");
  assertEquals([o.outcome, o.charged, o.total], ["lapsed", false, "€1,000.00"]);
  const i = later("2027-01-02", { contractProvidesDamages: true, contractProvidesInterest: true });
  assertEquals([i.lumpSum, i.interest, i.interestIncomplete, i.charged], ["€90.00", "", true, true]);
});

test("a business debtor owes EUR 40 per invoice", () => {
  const biz = { ...base, debtorType: "business", asOf: "2026-12-17", dueDate: "2026-12-16", stage: "bogus", ...both };
  const o = run(biz);
  assertEquals([o.outcome, o.compensation, o.invoicesLabel, o.total], ["business", "€40.00", "1 invoice", "€1,040.00"]);
  const three = run({ ...biz, invoiceCount: 3 });
  assertEquals([three.compensation, three.invoicesLabel, three.total], ["€120.00", "3 invoices", "€1,120.00"]);
  throwsOn({ ...biz, invoiceCount: 0 }, "invoiceCount");
  throwsOn({ ...biz, invoiceCount: 1.5 }, "invoiceCount");
  assertEquals(run({ ...biz, asOf: "2026-12-16" }).outcome, "not-yet-due");
});

test("nothing due and not yet due", () => {
  assertEquals(run({ ...base, amountDue: 0 }).outcome, "nothing-due");
  assertEquals(run({ ...base, amountDue: -25.5 }).outcome, "nothing-due");
  assertEquals(run({ ...base, asOf: "2026-11-01" }).outcome, "not-yet-due");
});

test("malformed input throws, naming the field", () => {
  const { asOf: _, ...noAsOf } = base;
  throwsOn(noAsOf, "asOf");
  throwsOn({ ...base, amountDue: NaN }, "amountDue");
  throwsOn({ ...base, amountDue: "1000" }, "amountDue");
  throwsOn({ ...base, stage: "second" }, "stage");
  throwsOn({ ...base, debtorType: "other" }, "debtorType");
  throwsOn({ ...base, firstReminder: { sentDate: "2026-11-19", channel: "fax" } }, "channel");
  throwsOn({ ...base, firstReminder: { sentDate: "14/10/2026", channel: "paper" } }, "sentDate");
  throwsOn({ ...base, holidays: ["soon"] }, "holidays");
  throwsOn({ ...base, contractProvidesInterest: true, agreedRate: -1 }, "agreedRate");
});

test("the catalog's preview data", () => {
  const entry = COMPONENTS.find((c) => c.id === "be-consumer-debt-reminder")!;
  const debt = (entry.previewData as { debt: Obj }).debt;
  const o = run(debt);
  assertEquals(
    [o.outcome, o.periodStart, o.lastDay, o.lumpSum, o.interestFrom, o.interestDays, o.interest, o.total],
    ["lapsed", "19 October 2026", "1 November 2026", "€2,000.00", "2 November 2026", 29, "€474.28", "€51,204.88"],
  );
  assertEquals(o.rate, "12.25%");
});

// ---------------------------------------------------------------------------
// Building and publishing
// ---------------------------------------------------------------------------

const templateOf = () =>
  template<Obj>(
    <Document id="d" title="Reminder">
      <Section id="body" title="Body">
        <BeConsumerDebtReminder />
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

const built = (debt: Obj) =>
  buildDocument(templateOf(), { debt }, { branchMode: "decide", dynamicMode: "placeholder", derivers } as never);

async function ids_of(debt: Obj): Promise<string[]> {
  const body = (await built(debt)).nodes[0] as { children: DocumentNode[] };
  return collect(body.children).map((n) => n.id).filter((id) => !id.startsWith("table-cell"));
}

test("each outcome builds only its own section", async () => {
  const idsFor = async (debt: Obj) => (await ids_of(debt)).filter((id) => /^be-debt-[a-z-]+$/.test(id));
  const first = await idsFor({ ...base, firstReminder: electronic });
  assertEquals(first.includes("be-debt-first-adds-nothing"), true);
  assertEquals(first.some((id) => id === "be-debt-first-lump-sum" || id === "be-debt-first-interest"), false);
  const withClauses = await idsFor({ ...base, firstReminder: electronic, ...both });
  assertEquals(withClauses.includes("be-debt-first-adds-nothing"), false);
  assertEquals(withClauses.includes("be-debt-first-lump-sum") && withClauses.includes("be-debt-first-interest"), true);
  assertEquals((await idsFor({ ...base })).includes("be-debt-first-incomplete"), true);
  const plain = await idsFor({ ...base, stage: "subsequent", asOf: "2027-01-02", firstReminder: electronic });
  assertEquals(plain.includes("be-debt-lapsed-no-charges"), true);
  assertEquals(plain.includes("be-debt-lapsed-charges"), false);
  const incomplete = await idsFor({
    ...base, stage: "subsequent", asOf: "2027-01-02", firstReminder: electronic, contractProvidesInterest: true,
  });
  assertEquals(incomplete.includes("be-debt-lapsed-interest-incomplete"), true);
  assertEquals(incomplete.includes("be-debt-lapsed-no-charges"), false);
  assertEquals(incomplete.includes("be-debt-lapsed-charges"), false);
  const warned = await idsFor({
    ...base, stage: "subsequent", asOf: "2027-01-02", firstReminder: electronic,
    contractProvidesDamages: true, contractProvidesInterest: true,
  });
  assertEquals(
    ["be-debt-lapsed-charges", "be-debt-lapsed-interest-incomplete"].every((id) => warned.includes(id)) &&
      !warned.includes("be-debt-lapsed-interest"),
    true,
  );
});

test("a publish build carries 16 arms, and resolving matches a direct build", async () => {
  assertEquals(branches, 16);
  const published = await buildDocument(templateOf(), createPublishData() as Obj, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const all = collect(published.nodes);
  assertEquals(all.filter((n) => (n as { when?: unknown }).when !== undefined).length, 16);
  const section = all.find((n) => n.id === "be-debt-first") as unknown as { derivers: unknown[] };
  assertEquals(JSON.parse(JSON.stringify(section.derivers[0])).name, "beConsumerDebtReminder");

  const preview = (COMPONENTS.find((c) => c.id === "be-consumer-debt-reminder")!.previewData as { debt: Obj }).debt;
  const cases: Array<[Obj, string]> = [
    [{ ...base, firstReminder: electronic, ...both }, "be-debt-first"],
    [preview, "be-debt-lapsed"],
    [{ ...base, debtorType: "business", asOf: "2026-12-17", dueDate: "2026-12-16" }, "be-debt-business"],
  ];
  for (const [debt, sectionId] of cases) {
    const direct = await built(debt);
    const resolved = await resolveDocument(published, state({ debt }), { derivers });
    const seen = (nodes: DocumentNode[]) => collect(nodes).map((n) => [n.kind, (n as { text?: string }).text]);
    assertEquals(seen(resolved.nodes), seen(direct.nodes));
    const ids = collect(resolved.nodes).map((n) => n.id);
    assertEquals(ids.includes(sectionId), true);
    assertEquals(ids.filter((id) => /^be-debt-[a-z-]+$/.test(id) && !id.startsWith(sectionId)).length, 0);
    assertEquals(JSON.stringify(resolved.nodes).includes("{{derived"), false);
  }
});
