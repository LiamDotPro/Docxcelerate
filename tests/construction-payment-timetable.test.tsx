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
  typeof import("../registry/components/derivers/construction-payment-timetable.deriver.ts")
>("components/derivers/construction-payment-timetable.deriver.ts");
const { ConstructionPaymentTimetable } = await loadRegistryModule<
  typeof import("../registry/components/construction-payment-timetable/construction-payment-timetable.node.tsx")
>("components/construction-payment-timetable/construction-payment-timetable.node.tsx");

type Obj = Record<string, unknown>;

const EW = [
  "2026-01-01", "2026-04-03", "2026-04-06", "2026-05-04", "2026-05-25", "2026-08-31", "2026-12-25", "2026-12-28",
  "2027-01-01", "2027-03-26", "2027-03-29", "2027-05-03", "2027-05-31", "2027-08-30", "2027-12-27", "2027-12-28",
];
const SC = [
  "2026-01-01", "2026-01-02", "2026-04-03", "2026-05-04", "2026-05-25", "2026-08-03", "2026-11-30", "2026-12-25", "2026-12-28",
  "2027-01-01", "2027-01-04", "2027-03-26", "2027-05-03", "2027-05-31", "2027-08-02", "2027-11-30", "2027-12-27", "2027-12-28",
];
// 2026-07-13 is the Battle of the Boyne substitute day.
const NI = [
  "2026-01-01", "2026-03-17", "2026-04-03", "2026-04-06", "2026-05-04", "2026-05-25", "2026-07-13", "2026-08-31",
  "2026-12-25", "2026-12-28",
];
const holidays = { EW, SC, NI };

interface Result {
  applies: boolean;
  notApplicable: string | null;
  basis: string | null;
  dueDate: string | null;
  paymentNoticeBy: string | null;
  finalDate: string | null;
  payLessBy: string | null;
  noticeLate: string | null;
  noticeGiven: string | null;
  applicationStands: string | null;
  defaultNotice: string | null;
  defaultOpen: string | null;
  reckoning: string | null;
}

// async, so a throw rejects, which is how a published document sees it.
const run = async (contract: Obj, payment: Obj, bankHolidays: Obj = holidays) =>
  await mod.run([contract, payment, bankHolidays], {} as RuntimeState) as Result;

const scheme = (region: string, valuationPeriodEnd: string, claimReceived: string, more: Obj = {}) => ({
  contract: { region },
  payment: { valuationPeriodEnd, claimReceived, ...more },
});
const go = (c: { contract: Obj; payment: Obj }, bankHolidays?: Obj) => run(c.contract, c.payment, bankHolidays);

/** The four dates a case is held to, in the order they print. */
const dates = (r: Result) => [r.dueDate, r.paymentNoticeBy, r.finalDate, r.payLessBy];

const A = scheme("EW", "2026-11-06", "2026-11-09");
const aDates = [
  "Friday, 13 November 2026",
  "Wednesday, 18 November 2026",
  "Monday, 30 November 2026",
  "Monday, 23 November 2026",
];

test("A: the Scheme, with the claim standing as the notified sum", async () => {
  const r = await go(A);
  assertEquals(dates(r), aDates);
  assertStringIncludes(r.applicationStands ?? "", "received on Monday, 9 November 2026 stands as the notified sum");
  assertEquals([r.noticeGiven, r.noticeLate, r.defaultNotice, r.defaultOpen], [null, null, null, null]);
  assertStringIncludes(
    r.basis ?? "",
    "Under the Scheme for Construction Contracts (England and Wales), payment falls due on the later of 7 days after",
  );
  assertStringIncludes(r.reckoning ?? "", "section 116 of the Housing Grants, Construction and Regeneration Act 1996");
  assertStringIncludes(r.reckoning ?? "", "bank holidays in England and Wales");
});

test("A1 and A2: a notice on the deadline counts, a day later it does not", async () => {
  const on = await go(scheme("EW", "2026-11-06", "2026-11-09", { paymentNoticeIssued: "2026-11-18" }));
  assertEquals(dates(on), aDates);
  assertStringIncludes(on.noticeGiven ?? "", "gave its payment notice on Wednesday, 18 November 2026");
  assertEquals([on.noticeLate, on.applicationStands], [null, null]);

  const late = await go(scheme("EW", "2026-11-06", "2026-11-09", { paymentNoticeIssued: "2026-11-19" }));
  assertEquals(dates(late), aDates);
  assertEquals(late.noticeGiven, null);
  assertStringIncludes(late.noticeLate ?? "", "Thursday, 19 November 2026 came after the deadline");
  assertStringIncludes(late.applicationStands ?? "", "stands as the notified sum");
});

test("B: under the Scheme a later claim sets the due date", async () => {
  assertEquals(dates(await go(scheme("EW", "2026-11-06", "2026-11-20"))), [
    "Friday, 20 November 2026",
    "Wednesday, 25 November 2026",
    "Monday, 7 December 2026",
    "Monday, 30 November 2026",
  ]);
});

const C = scheme("EW", "2026-12-11", "2026-12-14");

test("C and C-SC: Christmas is skipped, and the region picks the list", async () => {
  assertEquals(dates(await go(C)), [
    "Friday, 18 December 2026",
    "Wednesday, 23 December 2026",
    "Thursday, 7 January 2027",
    "Wednesday, 30 December 2026",
  ]);
  const sc = await go(scheme("SC", "2026-12-11", "2026-12-14"));
  assertEquals(dates(sc), [
    "Friday, 18 December 2026",
    "Wednesday, 23 December 2026",
    "Friday, 8 January 2027",
    "Wednesday, 30 December 2026",
  ]);
});

test("D: a Scottish holiday is skipped there and counted in England", async () => {
  assertEquals(dates(await go(scheme("SC", "2026-11-20", "2026-11-20"))), [
    "Friday, 27 November 2026",
    "Thursday, 3 December 2026",
    "Tuesday, 15 December 2026",
    "Tuesday, 8 December 2026",
  ]);
  assertEquals(dates(await go(scheme("EW", "2026-11-20", "2026-11-20"))), [
    "Friday, 27 November 2026",
    "Wednesday, 2 December 2026",
    "Monday, 14 December 2026",
    "Monday, 7 December 2026",
  ]);
});

test("G: Good Friday and Easter Monday are not counted", async () => {
  assertEquals(dates(await go(scheme("EW", "2026-03-20", "2026-03-23"))), [
    "Friday, 27 March 2026",
    "Wednesday, 1 April 2026",
    "Wednesday, 15 April 2026",
    "Wednesday, 8 April 2026",
  ]);
});

test("Good Friday is excluded without being listed, in 2026 and 2027", async () => {
  // G with Good Friday (and Easter Monday) taken out of the list: only Good Friday still goes.
  const bare = { EW: ["2026-01-01"] };
  const r = await go(scheme("EW", "2026-03-20", "2026-03-23"), bare);
  assertEquals(r.finalDate, "Tuesday, 14 April 2026");
  // 2027: Good Friday is 26 March, so due 7 days after 2027-03-19 skips it.
  const r27 = await go(scheme("EW", "2027-03-19", "2027-03-19"), { EW: ["2027-01-01"] });
  assertEquals(r27.dueDate, "Saturday, 27 March 2027");
});

test("Christmas Day is excluded without being listed", async () => {
  // C with 25 and 28 December out of the list: 25 December still goes, 28 now counts, so 6 January not 7 (5 if Christmas Day counted).
  const r = await go(C, { EW: ["2026-01-01", "2027-01-01"] });
  assertEquals(r.finalDate, "Wednesday, 6 January 2027");
});

test("H: a due date on a weekend stays there, across a year boundary", async () => {
  assertEquals(dates(await go(scheme("SC", "2026-12-31", "2026-12-31"))), [
    "Saturday, 9 January 2027",
    "Thursday, 14 January 2027",
    "Tuesday, 26 January 2027",
    "Tuesday, 19 January 2027",
  ]);
});

test("NI: the Order is named, and the region's own list is used", async () => {
  const r = await go(scheme("NI", "2026-07-03", "2026-07-03"));
  assertEquals(dates(r), [
    "Friday, 10 July 2026",
    "Thursday, 16 July 2026",
    "Tuesday, 28 July 2026",
    "Tuesday, 21 July 2026",
  ]);
  assertStringIncludes(r.reckoning ?? "", "Construction Contracts (Northern Ireland) Order 1997");
  assertStringIncludes(r.basis ?? "", "the Scheme for Construction Contracts (Northern Ireland)");
});

const E = {
  contract: { region: "EW", dueDateDays: 7, finalDateDays: 14, payLessDays: 5 },
  payment: { valuationPeriodEnd: "2026-11-06", defaultNoticeIssued: "2026-11-25" },
};

test("E: the payee's default notice moves the final date, and the pay less deadline with it", async () => {
  const r = await go(E);
  assertEquals(dates(r), [
    "Friday, 13 November 2026",
    "Wednesday, 18 November 2026",
    "Friday, 4 December 2026",
    "Sunday, 29 November 2026",
  ]);
  assertEquals(
    r.defaultNotice,
    "No payment notice was given by Wednesday, 18 November 2026. The payee's notice of Wednesday, 25 November 2026 " +
      "stands as the notified sum, and moves the final date for payment back 7 days, from " +
      "Friday, 27 November 2026 to Friday, 4 December 2026.",
  );
  assertEquals(r.basis, "Under the contract, payment falls due 7 days after the valuation period ended on Friday, 6 November 2026.");
  assertEquals([r.noticeGiven, r.applicationStands, r.defaultOpen, r.noticeLate], [null, null, null, null]);
});

test("E-open: with no default notice yet the dates are the earliest", async () => {
  const r = await go({ ...E, payment: { valuationPeriodEnd: "2026-11-06" } });
  assertEquals(r.finalDate, "Friday, 27 November 2026 at the earliest");
  assertEquals(r.payLessBy, "Sunday, 22 November 2026 at the earliest");
  assertStringIncludes(r.defaultOpen ?? "", "has not yet given a notice in its place");
  assertEquals(r.defaultNotice, null);
});

test("E-early: a default notice on or before the deadline throws", async () => {
  await assertRejects(
    () => go({ ...E, payment: { ...E.payment, defaultNoticeIssued: "2026-11-18" } }),
    Error,
    "payment.defaultNoticeIssued must be after the payment notice deadline, 2026-11-18",
  );
});

test("E-claim: a claim before the deadline stands, and the default notice is ignored", async () => {
  const r = await go({ ...E, payment: { ...E.payment, claimReceived: "2026-11-17" } });
  assertEquals(r.finalDate, "Friday, 27 November 2026");
  assertStringIncludes(r.applicationStands ?? "", "received on Tuesday, 17 November 2026");
  assertEquals([r.defaultNotice, r.defaultOpen], [null, null]);
  // Ignored means not even validated.
  const ignored = await go({ ...E, payment: { ...E.payment, claimReceived: "2026-11-17", defaultNoticeIssued: "nonsense" } });
  assertEquals(ignored.finalDate, "Friday, 27 November 2026");
});

test("a claim on the deadline day does not stand (strictly before)", async () => {
  const r = await go({ ...E, payment: { ...E.payment, claimReceived: "2026-11-18" } });
  assertEquals(r.applicationStands, null);
  assertStringIncludes(r.defaultNotice ?? "", "moves the final date for payment back 7 days");
});

test("a default notice whose every day in between is excluded does not move the final date", async () => {
  // The payment notice deadline is Thursday 24 December; Christmas Day is not counted.
  const at = (defaultNoticeIssued: string) =>
    run(
      { region: "EW", dueDateDays: 0, finalDateDays: 14, payLessDays: 0 },
      { valuationPeriodEnd: "2026-12-19", defaultNoticeIssued },
    );
  const same = await at("2026-12-25");
  assertStringIncludes(same.defaultNotice ?? "", "stands as the notified sum. The final date for payment does not move.");
  assertStringIncludes((await at("2026-12-26")).defaultNotice ?? "", "back 1 day, from");
});

test("Zero: a due date of the period end itself, and a pay less deadline equal to the final date", async () => {
  const r = await run(
    { region: "EW", dueDateDays: 0, payLessDays: 0, finalDateDays: 14 },
    { valuationPeriodEnd: "2026-11-06", claimReceived: "2026-11-09" },
  );
  assertEquals(r.dueDate, "Friday, 6 November 2026");
  assertEquals(r.basis, "Under the contract, payment falls due on the day the valuation period ended, Friday, 6 November 2026.");
  assertEquals(r.payLessBy, r.finalDate?.replace(" at the earliest", ""));
});

test("one day is singular", async () => {
  const r = await run(
    { region: "EW", dueDateDays: 1, payLessDays: 0, finalDateDays: 14 },
    { valuationPeriodEnd: "2026-11-06" },
  );
  assertStringIncludes(r.basis ?? "", "falls due 1 day after");
});

test("PL<PN: a pay less deadline before the payment notice deadline throws", async () => {
  await assertRejects(
    () => run({ region: "EW", dueDateDays: 0, finalDateDays: 1, payLessDays: 0 }, { valuationPeriodEnd: "2026-11-06" }),
    Error,
    "pay less notice deadline falls before the payment notice deadline",
  );
});

test("Neg/frac: each bad period names its field", async () => {
  const pay = { valuationPeriodEnd: "2026-11-06", claimReceived: "2026-11-06" };
  const bad = (contract: Obj, message: string) =>
    assertRejects(() => run({ region: "EW", ...contract }, pay), Error, message);
  await bad({ dueDateDays: -1 }, "contract.dueDateDays must be a whole number of days, at least 0");
  await bad({ dueDateDays: "7" }, "contract.dueDateDays must be a whole number of days, at least 0");
  await bad({ dueDateDays: 1.5 }, "contract.dueDateDays");
  await bad({ finalDateDays: 0 }, "contract.finalDateDays must be a whole number of days, at least 1");
  await bad({ payLessDays: 2.5 }, "contract.payLessDays must be a whole number of days, at least 0");
  // null counts as missing.
  assertEquals((await run({ region: "EW", dueDateDays: null, finalDateDays: null, payLessDays: null }, pay)).applies, true);
});

test("Residential: one sentence, nothing else, and nothing validated", async () => {
  const r = await run({ region: "EW", residentialOccupier: true }, {}, {});
  assertEquals(r, {
    applies: false,
    notApplicable:
      "This contract is with a residential occupier, so the Housing Grants, Construction and Regeneration Act 1996 " +
      "does not apply to it, and neither does the payment timetable it sets.",
    basis: null,
    dueDate: null,
    paymentNoticeBy: null,
    finalDate: null,
    payLessBy: null,
    noticeLate: null,
    noticeGiven: null,
    applicationStands: null,
    defaultNotice: null,
    defaultOpen: null,
    reckoning: null,
  });
  assertStringIncludes(
    (await run({ region: "NI", residentialOccupier: true }, {}, {})).notApplicable ?? "",
    "Construction Contracts (Northern Ireland) Order 1997",
  );
  // The region still has to be one, because it names the Act.
  await assertRejects(() => run({ region: "GB", residentialOccupier: true }, {}, {}), Error, "contract.region");
  // Only true is residential.
  await assertRejects(() => run({ region: "EW", residentialOccupier: "true" }, {}, {}), Error, "bankHolidays.EW is required");
});

test("Missing: each bad input throws with its own message", async () => {
  const ok = { valuationPeriodEnd: "2026-11-06", claimReceived: "2026-11-09" };
  await assertRejects(() => run({ region: "GB" }, ok), Error, "contract.region must be one of EW, SC, NI");
  await assertRejects(
    () => run({ region: "EW" }, { valuationPeriodEnd: "2026-11-06" }),
    Error,
    "payment.claimReceived is required when contract.dueDateDays is not given",
  );
  await assertRejects(
    () => run({ region: "EW" }, { valuationPeriodEnd: "2026-11-06", claimReceived: "" }),
    Error,
    "payment.claimReceived is required",
  );
  await assertRejects(
    () => run({ region: "EW" }, { ...ok, valuationPeriodEnd: "2026-02-30" }),
    Error,
    "payment.valuationPeriodEnd must be an ISO 8601 date",
  );
  await assertRejects(() => run({ region: "EW" }, { ...ok, valuationPeriodEnd: undefined }), Error, "valuationPeriodEnd");
  await assertRejects(() => run({ region: "EW" }, { ...ok, claimReceived: "09/11/2026" }), Error, "payment.claimReceived must be");
  await assertRejects(
    () => run({ region: "EW" }, { ...ok, paymentNoticeIssued: "soon" }),
    Error,
    "payment.paymentNoticeIssued must be an ISO 8601 date",
  );
  await assertRejects(
    () => run({ region: "SC" }, ok, { EW }),
    Error,
    "bankHolidays.SC is required for a contract in Scotland",
  );
  await assertRejects(() => run({ region: "SC" }, ok, null as never), Error, "bankHolidays.SC is required");
  await assertRejects(() => run({ region: "EW" }, ok, { EW: ["2026-13-01"] }), Error, "bankHolidays.EW[0] must be an ISO 8601 date");
});

test("a date's time is ignored, and its first ten characters are what count", async () => {
  const r = await go(scheme("EW", "2026-11-06T23:30:00Z", "2026-11-09T01:00:00+01:00"));
  assertEquals(dates(r), aDates);
});

test("Stale: a list with no dates in a year the timetable runs through throws", async () => {
  const only2026 = { SC: SC.filter((d) => d.startsWith("2026")) };
  await assertRejects(
    () => go(scheme("SC", "2026-12-31", "2026-12-31"), only2026),
    Error,
    "bankHolidays.SC has no dates in 2027",
  );
  await assertRejects(() => go(A, { EW: [] }), Error, "bankHolidays.EW has no dates in 2026");
});

test("Stale: the year check starts at the valuation period end's year", async () => {
  await assertRejects(
    () => go(scheme("SC", "2026-12-31", "2026-12-31"), { SC: SC.filter((d) => d.startsWith("2027")) }),
    Error,
    "bankHolidays.SC has no dates in 2026",
  );
});

test("Stale: a default notice in a later year than the final date is checked too", async () => {
  // The final date lands on 31 December 2026; only the notice (4 January) is in 2027.
  await assertRejects(
    () =>
      run(
        { region: "EW", dueDateDays: 0, finalDateDays: 1, payLessDays: 0 },
        { valuationPeriodEnd: "2026-12-20", defaultNoticeIssued: "2027-01-04" },
        { EW: EW.filter((d) => d.startsWith("2026")) },
      ),
    Error,
    "bankHolidays.EW has no dates in 2027",
  );
});

test("Stale: a default notice pushes the span into the next year", async () => {
  await assertRejects(
    () =>
      run(
        { region: "EW", dueDateDays: 7, finalDateDays: 14, payLessDays: 5 },
        { valuationPeriodEnd: "2026-12-04", defaultNoticeIssued: "2027-01-05" },
        { EW: EW.filter((d) => d.startsWith("2026")) },
      ),
    Error,
    "bankHolidays.EW has no dates in 2027",
  );
});

test("Many/one: duplicates and order do not matter, and one year is enough for one year", async () => {
  const messy = { EW: [...EW].reverse().concat(EW) };
  assertEquals(await go(C, messy), await go(C));
  const single = { EW: EW.filter((d) => d.startsWith("2026")) };
  assertEquals(dates(await go(A, single)), aDates);
});

test("PL<PN is checked before the stale-list year check", async () => {
  await assertRejects(
    () => run({ region: "EW", dueDateDays: 0, finalDateDays: 1, payLessDays: 7 }, { valuationPeriodEnd: "2026-11-06" }),
    Error,
    "pay less notice deadline",
  );
});

// The preview exercises the longest text: the late notice and the default notice, over Christmas.
const previewData = {
  contract: { region: "SC", dueDateDays: 14, finalDateDays: 30, payLessDays: 7 },
  payment: {
    valuationPeriodEnd: "2026-12-04",
    claimReceived: "2027-01-12",
    paymentNoticeIssued: "2027-01-06",
    defaultNoticeIssued: "2027-01-14",
  },
  bankHolidays: { SC },
};

test("the preview data prints the long case", async () => {
  const r = await run(previewData.contract, previewData.payment, previewData.bankHolidays);
  assertEquals(dates(r), [
    "Friday, 18 December 2026",
    "Wednesday, 23 December 2026",
    "Monday, 8 February 2027",
    "Monday, 1 February 2027",
  ]);
  assertStringIncludes(r.defaultNotice ?? "", "back 18 days, from Thursday, 21 January 2027 to Monday, 8 February 2027");
  assertEquals(r.noticeLate !== null, true);
});

// The node, built and then published and resolved by an engine that only has the bundle.
const ARRANGEMENTS = ["table", "stacked"] as const;
const NOTICE_ARMS = ["notice-given", "application-stands", "default-notice", "default-open"];

function documentOf(arrangement?: string) {
  const Node = ConstructionPaymentTimetable as (props: { arrangement?: string }) => never;
  return template(
    <Document title="Timetable">
      <Node arrangement={arrangement} />
    </Document>,
  );
}

const build = (arrangement: string | undefined, data: Obj = previewData) =>
  buildDocument(documentOf(arrangement), data as typeof previewData, { dynamicMode: "placeholder" });

function ids(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { id?: string; children?: DocumentNode[] };
    return [...(own.id ? [own.id] : []), ...ids(own.children ?? [])];
  });
}

function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { title?: string; text?: string; children?: DocumentNode[] };
    return [
      ...(node.kind === "section" && own.title ? [own.title] : []),
      ...(own.text ? [own.text] : []),
      ...texts(own.children ?? []),
    ];
  });
}

test("every arrangement carries the same ids", async () => {
  const [table, stacked] = await Promise.all(ARRANGEMENTS.map(async (a) => ids((await build(a)).nodes)));
  const expected = [
    "payment-timetable",
    "payment-timetable-basis",
    "payment-timetable-due",
    "payment-timetable-notice-by",
    "payment-timetable-final",
    "payment-timetable-pay-less-by",
    "payment-timetable-notice-late",
    "payment-timetable-default-notice",
    "payment-timetable-reckoning",
  ];
  // A table gives its own cells ids of their own; only the component's are held to a list.
  const ours = (all: string[]) => all.filter((id) => id.startsWith("payment-timetable"));
  assertEquals(ours(table).filter((id) => id !== "payment-timetable-dates"), expected);
  assertEquals(ours(table).includes("payment-timetable-dates"), true);
  assertEquals(ours(stacked), expected);
});

test("the arrangements print the same dates, placed differently", async () => {
  const table = texts((await build("table")).nodes);
  const stacked = texts((await build("stacked")).nodes);
  assertEquals(table.includes("Final date for payment"), true);
  assertEquals(table.includes("Monday, 8 February 2027"), true);
  assertEquals(stacked.includes("Final date for payment: Monday, 8 February 2027"), true);
  assertEquals(table.length > stacked.length, true);
  assertEquals((await build(undefined)).nodes, (await build("table")).nodes);
});

test("a residential contract builds only the one sentence", async () => {
  const data = { contract: { region: "EW", residentialOccupier: true }, payment: {}, bankHolidays: {} };
  for (const arrangement of ARRANGEMENTS) {
    assertEquals(ids((await build(arrangement, data)).nodes), ["payment-timetable", "payment-timetable-not-applicable"]);
  }
});

test("a bogus arrangement rejects", async () => {
  await assertRejects(() => build("split"), Error, "Unknown arrangement");
});

const registry = await createDeriverRegistryFromBundle(
  createDeriverBundle({ constructionPaymentTimetable: mod.run }),
);

async function resolved(arrangement: string, data: Obj) {
  const published = await buildDocument(documentOf(arrangement), createPublishData() as typeof previewData, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: { generateParagraph: () => "" },
  };
  return resolveDocument(published, state, { derivers: registry });
}

for (const arrangement of ARRANGEMENTS) {
  test(`${arrangement}: publishing then resolving matches building directly, case E`, async () => {
    const data = {
      contract: E.contract,
      payment: E.payment,
      bankHolidays: { EW },
    };
    const out = await resolved(arrangement, data);
    const direct = await build(arrangement, data);

    assertEquals(texts(out.nodes), texts(direct.nodes));
    const text = texts(out.nodes).join("\n");
    assertStringIncludes(text, "Friday, 4 December 2026");
    assertStringIncludes(text, "Sunday, 29 November 2026");
    const present = ids(out.nodes);
    assertEquals(NOTICE_ARMS.filter((arm) => present.includes("payment-timetable-" + arm)), ["default-notice"]);
  });

  test(`${arrangement}: publishing then resolving the residential case`, async () => {
    const data = { contract: { region: "EW", residentialOccupier: true }, payment: {}, bankHolidays: {} };
    const out = await resolved(arrangement, data);

    assertEquals(ids(out.nodes), ["payment-timetable", "payment-timetable-not-applicable"]);
  });
}
