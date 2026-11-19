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
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/direct-debit-collection-notice.deriver.ts")
>("components/derivers/direct-debit-collection-notice.deriver.ts");
const { DirectDebitCollectionNotice } = await loadRegistryModule<
  typeof import("../registry/components/direct-debit-collection-notice/direct-debit-collection-notice.node.tsx")
>("components/direct-debit-collection-notice/direct-debit-collection-notice.node.tsx");
const registry = await createDeriverRegistryFromBundle(
  createDeriverBundle({ directDebitCollectionNotice: mod.run }),
);

type Obj = Record<string, unknown>;

const sepa: Obj = {
  scheme: "sepa-core",
  creditorName: "Northwind Facilities & Estates Management Services GmbH",
  creditorId: "IT58ZZZ0000012345678901",
  mandateReference: "MNDT-2026-NORTHFACILITIES-000048172",
  payerIban: "IT60 X054 2811 1010 0000 0123 456",
  requestedDate: "2027-03-26",
  amount: 1284750.0,
  first: true,
};
const guarantee = ["The Direct Debit Guarantee", "First paragraph.", "Second paragraph."];
const bacs: Obj = {
  scheme: "bacs",
  creditorName: "Northwind Facilities Ltd",
  sun: "123 456",
  mandateReference: "NW-0001",
  requestedDate: "2026-12-24",
  amount: 99.5,
  bankHolidays: ["2026-12-25", "2026-12-28"],
  guaranteeText: guarantee,
};
const ach: Obj = {
  scheme: "ach",
  creditorName: "Northwind Facilities Inc",
  mandateReference: "AUTH-77",
  requestedDate: "2026-11-25",
  amount: 182.4,
  previousAmount: 175.1,
  payerAccountLast4: "6789",
};

interface Result {
  heading: string;
  outcome: string;
  schemeName: string;
  mandateNoun: string;
  amount: string;
  previousAmount: string | null;
  amountChanged: boolean;
  late: boolean;
  noticeLabel: string;
  date: string;
  creditorName: string;
  creditorIdLabel: string | null;
  creditorId: string | null;
  mandateLabel: string;
  mandateReference: string;
  account: string | null;
  guaranteeMissing: boolean;
}

const run = (base: Obj, over: Obj = {}, issuedAt: unknown = "2027-03-08"): Result =>
  mod.run([{ ...base, ...over }, issuedAt], {} as RuntimeState) as Result;
function throwsWith(fn: () => unknown, text: string): void {
  let message: string | undefined;
  try {
    fn();
  } catch (e) {
    message = (e as Error).message;
  }
  assertStringIncludes(message ?? "(did not throw)", text);
}
const dateOf = (base: Obj, over: Obj, issuedAt: string) => run(base, over, issuedAt).date;

test("the preview: on time, Good Friday moves to Tuesday after Easter Monday", () => {
  const r = run(sepa);
  assertEquals([r.outcome, r.late, r.date, r.amount], ["collect", false, "30 March 2027", "€1,284,750.00"]);
  assertEquals([r.account, r.creditorId, r.heading], [
    "IBAN ending 3456",
    "IT58ZZZ0000012345678901",
    "Payment by SEPA Direct Debit",
  ]);
});

test("SEPA late notice rolls past New Year and the weekend", () => {
  const r = run(sepa, { requestedDate: "2026-12-24" }, "2026-12-18");
  assertEquals([r.late, r.date, r.noticeLabel], [true, "4 January 2027", "14 calendar days"]);
});

test("SEPA exactly 14 days is in time; agreed days replace 14", () => {
  const r = run(sepa, { requestedDate: "2026-12-03" }, "2026-11-19");
  assertEquals([r.late, r.date], [false, "3 December 2026"]);
  const a = run(sepa, { agreedNoticeDays: 2, requestedDate: "2026-11-23" }, "2026-11-19");
  assertEquals([a.late, a.date], [false, "23 November 2026"]);
  assertEquals(run(sepa, { agreedNoticeDays: 1 }).noticeLabel, "1 calendar day");
});

test("SEPA creditor identifiers are checked without the business code", () => {
  assertEquals(run(sepa, { creditorId: "DE98ZZZ09999999999" }).creditorId, "DE98ZZZ09999999999");
  assertEquals(run(sepa, { creditorId: "de98 abc 0999 9999 999" }).creditorId, "DE98ABC09999999999");
  throwsWith(() => run(sepa, { creditorId: "DE12ZZZ09999999999" }), "creditorId");
  throwsWith(() => run(sepa, { creditorId: "DE12ZZZ09999999999" }), "98");
  throwsWith(() => run(sepa, { creditorId: undefined }), "creditorId");
  throwsWith(() => run(sepa, { creditorId: "XX" }), "creditorId");
});

test("mandate reference lengths", () => {
  throwsWith(() => run(sepa, { mandateReference: "x".repeat(36) }), "mandateReference");
  throwsWith(() => run(bacs, { mandateReference: "x".repeat(19) }), "mandateReference");
  run(sepa, { mandateReference: "x".repeat(35) });
  run(bacs, { mandateReference: "x".repeat(18) });
});

test("Bacs counts working days and skips supplied holidays", () => {
  const r = run(bacs, {}, "2026-12-14");
  assertEquals([r.late, r.date, r.creditorId, r.creditorIdLabel], [true, "30 December 2026", "123456", "Service User Number"]);
  assertEquals(r.amount, "£99.50");
  assertEquals(dateOf(bacs, { bankHolidays: [] }, "2026-12-14"), "28 December 2026");
  const f = run(bacs, { agreedNoticeDays: 1, requestedDate: "2026-11-21" }, "2026-11-19");
  assertEquals([f.late, f.date, f.noticeLabel], [true, "24 November 2026", "3 working days"]);
  throwsWith(() => run(bacs, { bankHolidays: undefined }), "bankHolidays");
  throwsWith(() => run(bacs, { sun: "12345" }), "sun");
});

test("Bacs guarantee missing", () => {
  for (const guaranteeText of [undefined, [], ["", "x"]]) {
    assertEquals(run(bacs, { guaranteeText }).guaranteeMissing, true);
  }
  assertEquals(run(bacs).guaranteeMissing, false);
  assertEquals(run(sepa).guaranteeMissing, false);
});

test("ACH: a changed amount needs 10 days, an unchanged one none", () => {
  const r = run(ach, {}, "2026-11-19");
  assertEquals(
    [r.amountChanged, r.late, r.date, r.previousAmount, r.account, r.creditorId, r.noticeLabel],
    [true, true, "November 30, 2026", "$175.10", "Account ending 6789", null, "10 days"],
  );
  for (const previousAmount of [182.4, 182.4000000001]) {
    const s = run(ach, { previousAmount }, "2026-11-19");
    assertEquals([s.amountChanged, s.late, s.date, s.previousAmount], [false, false, "November 25, 2026", null]);
  }
  const h = run(ach, { previousAmount: undefined, requestedDate: "2026-11-26", bankHolidays: ["2026-11-26"] }, "2026-11-19");
  assertEquals([h.late, h.date], [false, "November 27, 2026"]);
});

test("nothing to collect", () => {
  for (const base of [sepa, bacs, ach]) {
    for (const amount of [0, 0.004]) {
      const r = run(base, { amount });
      assertEquals([r.outcome, r.late, r.date], ["clear", false, ""]);
    }
  }
  const c = run(sepa, { amount: -45.2 });
  assertEquals([c.outcome, c.amount, c.late, c.date], ["credit", "€45.20", false, ""]);
  throwsWith(() => run(sepa, { amount: 0, creditorId: "DE12ZZZ09999999999" }), "creditorId");
});

test("nouns and accounts", () => {
  assertEquals([run(sepa).mandateNoun, run(bacs).mandateNoun, run(ach).mandateNoun], [
    "mandate",
    "Direct Debit Instruction",
    "authorization",
  ]);
  assertEquals(run(sepa, { payerIban: undefined }).account, null);
  assertEquals(run(sepa, { payerIban: "DE89" }).account, "IBAN ending DE89");
  throwsWith(() => run(sepa, { payerIban: "DE8" }), "payerIban");
  throwsWith(() => run(ach, { payerAccountLast4: "12345" }), "payerAccountLast4");
  assertEquals(run(ach, { payerAccountLast4: undefined }).account, null);
});

test("bad input names its field", () => {
  throwsWith(() => run(sepa, { requestedDate: "2027-02-30" }), "requestedDate");
  throwsWith(() => run(sepa, {}, "next week"), "issuedAt");
  throwsWith(() => run(sepa, { amount: NaN }), "amount");
  throwsWith(() => run(sepa, { scheme: "becs" }), "scheme");
  throwsWith(() => run(sepa, { agreedNoticeDays: -1 }), "agreedNoticeDays");
  throwsWith(() => run(sepa, { agreedNoticeDays: 2.5 }), "agreedNoticeDays");
});

test("a requested date before issue is late", () => {
  const r = run(sepa, { requestedDate: "2026-11-01" }, "2026-11-19");
  assertEquals([r.late, r.date], [true, "3 December 2026"]);
});

test("Easter is worked out per year", () => {
  assertEquals(dateOf(sepa, { requestedDate: "2026-04-03" }, "2026-01-01"), "7 April 2026");
  assertEquals(dateOf(sepa, { requestedDate: "2028-04-14" }, "2028-01-01"), "18 April 2028");
});

// The node.
const document = (arrangement?: "stacked" | "table") =>
  template(
    <Document title="Notice">
      <DirectDebitCollectionNotice arrangement={arrangement} />
    </Document>,
  );

function ids(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { id?: string; children?: DocumentNode[]; rows?: DocumentNode[]; cells?: DocumentNode[] };
    return [
      ...(own.id ? [own.id] : []),
      ...ids(own.children ?? []),
      ...ids(own.rows ?? []),
    ];
  });
}

function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { title?: string; text?: string; children?: DocumentNode[]; rows?: DocumentNode[] };
    return [
      ...(node.kind === "section" && own.title ? [own.title] : []),
      ...(own.text ? [own.text] : []),
      ...texts(own.children ?? []),
      ...texts(own.rows ?? []),
    ];
  });
}

const dataOf = (debit: Obj, issuedAt = "2027-03-08") => ({ debit, document: { issuedAt } });
const built = async (debit: Obj, issuedAt?: string, arrangement?: "stacked" | "table") =>
  (await buildDocument(document(arrangement), dataOf(debit, issuedAt), { dynamicMode: "placeholder" })).nodes;
const idsOf = async (debit: Obj, issuedAt?: string, arrangement?: "stacked" | "table") =>
  ids(await built(debit, issuedAt, arrangement));

test("the preview prints its arms", async () => {
  const got = await idsOf(sepa);
  for (const id of ["dd-notice-collect-on", "dd-notice-first", "dd-notice-sepa-refund", "dd-notice-account"]) {
    assertEquals(got.includes(id), true);
  }
  assertEquals(got.some((id) => id.startsWith("dd-notice-guarantee")), false);
});

test("Bacs late prints its guarantee verbatim and no refund right", async () => {
  const nodes = await built(bacs, "2026-12-14");
  const got = ids(nodes);
  assertEquals(got.filter((id) => id.startsWith("dd-notice-guarantee")).length, 3);
  assertEquals(got.includes("dd-notice-collect-after"), true);
  assertEquals(got.includes("dd-notice-sepa-refund"), false);
  assertEquals(got.includes("dd-notice-recurring"), true);
  const lines = texts(nodes);
  for (const g of guarantee) assertEquals(lines.includes(g), true);
  assertStringIncludes(lines.join("\n"), "on or after 30 December 2026");
});

test("Bacs without approved text warns", async () => {
  for (const guaranteeText of [undefined, [], ["", "x"]]) {
    const got = await idsOf({ ...bacs, guaranteeText }, "2026-12-14");
    assertEquals(got.includes("dd-notice-guarantee-missing"), true);
    assertEquals(got.includes("dd-notice-guarantee"), false);
  }
});

test("ACH shows the changed amount and has no identifier line", async () => {
  const got = await idsOf(ach, "2026-11-19");
  assertEquals(got.includes("dd-notice-collect-after"), true);
  assertEquals(got.includes("dd-notice-amount-changed"), true);
  assertEquals(got.includes("dd-notice-creditor-id"), false);
  assertEquals((await idsOf({ ...ach, previousAmount: 182.4 }, "2026-11-19")).includes("dd-notice-amount-changed"), false);
});

test("a zero or negative amount prints only its one paragraph", async () => {
  const clear = await idsOf({ ...sepa, amount: 0 });
  assertEquals(clear.filter((id) => id !== "direct-debit-notice"), ["dd-notice-nothing-clear"]);
  const credit = await idsOf({ ...sepa, amount: -45.2 });
  assertEquals(credit.filter((id) => id !== "direct-debit-notice"), ["dd-notice-nothing-credit"]);
});

test("no account, no account line, in either arrangement", async () => {
  for (const arrangement of ["stacked", "table"] as const) {
    assertEquals((await idsOf({ ...sepa, payerIban: undefined }, undefined, arrangement)).includes("dd-notice-account"), false);
  }
});

test("the arrangements differ but print the same details", async () => {
  const detail = (id: string) => id.startsWith("dd-notice-") && ["creditor", "creditor-id", "mandate", "account"].includes(id.slice(10));
  const s = await built(sepa, undefined, "stacked");
  const t = await built(sepa, undefined, "table");
  assertEquals(ids(s).filter(detail), ids(t).filter(detail));
  assertEquals(ids(t).includes("dd-notice-details"), true);
  assertEquals(ids(s).includes("dd-notice-details"), false);
  assertEquals(JSON.stringify(s) === JSON.stringify(t), false);
  assertEquals(JSON.stringify(await built(sepa)), JSON.stringify(s));
});

// Publishing, then resolving by an engine that only has the bundle.
const published = () =>
  buildDocument(document(), createPublishData() as ReturnType<typeof dataOf>, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });

test("publishing carries the deriver invocation", async () => {
  const doc = await published();
  assertStringIncludes(JSON.stringify(doc.nodes), "directDebitCollectionNotice");
});

for (const [name, debit, issuedAt, absent] of [
  ["SEPA", sepa, "2027-03-08", "dd-notice-guarantee"],
  ["Bacs", bacs, "2026-12-14", "dd-notice-sepa-refund"],
] as const) {
  test(`${name}: resolving the published document matches the local build`, async () => {
    const data = dataOf(debit, issuedAt);
    const state: RuntimeState = {
      ctx: { ...data },
      derived: {},
      dataProvider: new InMemoryDataProvider(data),
      aiClient: { generateParagraph: () => "" },
    };
    const resolved = await resolveDocument(await published(), state, { derivers: registry });
    const direct = await buildDocument(document(), data, { dynamicMode: "placeholder" });
    assertEquals(texts(resolved.nodes), texts(direct.nodes));
    assertEquals(ids(resolved.nodes).includes(absent), false);
  });
}

test("publishing stays within the 13 branches the doc comment states, whichever arrangement", async () => {
  for (const arrangement of ["stacked", "table"] as const) {
    await buildDocument(
      template(
        <Document title="Notice">
          <DirectDebitCollectionNotice arrangement={arrangement} />
        </Document>,
      ),
      createPublishData() as ReturnType<typeof dataOf>,
      {
        branchMode: "publish",
        deriverMode: "preserve",
        dynamicMode: "resolve",
        branchLimit: 13,
        aiClient: { generateParagraph: () => "" },
      },
    );
  }
});
