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
import { createPublishData, Document, template } from "docxcelerate/template";
import { transformDocumentSource } from "docxcelerate/transform";
import { registryRoot } from "docxcelerate/registry/install";
import { COMPONENTS } from "docxcelerate/registry";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { loadRegistryModule } from "./registry_support.ts";

const nodeSource = "components/nl-collection-costs/nl-collection-costs.node.tsx";
const { NlCollectionCosts } = await loadRegistryModule<
  typeof import("../registry/components/nl-collection-costs/nl-collection-costs.node.tsx")
>(nodeSource);
const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/nl-collection-costs.deriver.ts")
>("components/derivers/nl-collection-costs.deriver.ts");
const derivers = await createDeriverRegistryFromBundle(createDeriverBundle({ nlCollectionCosts: mod.run }));
const branches = transformDocumentSource(await readFile(join(registryRoot(), nodeSource), "utf8"), {
  fileName: nodeSource,
}).branches;

type Obj = Record<string, unknown>;
interface Out {
  outcome: string;
  invoices: string;
  interest: string;
  vat: string;
  costsPhrase: string;
  deadline: string;
  previousDeadline: string;
  baseCents: number;
  costsCents: number;
  vatCents: number;
  totalCents: number;
}

const invoice = (principal: number, number = "F1"): Obj => ({ number, principal, dueDate: "2026-10-01" });
const base = (over: Obj = {}): Obj => ({
  debtorType: "consumer",
  sendDate: "2026-11-19",
  channel: "email",
  invoices: [invoice(1000)],
  ...over,
});

function run(over: Obj = {}): Out {
  return mod.run([base(over)], {} as RuntimeState) as Out;
}
const costs = (principal: number) => run({ invoices: [invoice(principal)] }).costsCents;
const fails = (over: Obj, field: string) => assertRejects(() => Promise.resolve().then(() => run(over)), Error, field);

test("staffel arithmetic", () => {
  const cases: Array<[number, number]> = [
    [100, 4000],
    [266.66, 4000],
    [300.1, 4502],
    [1234.57, 18519],
    [2500, 37500],
    [3000, 42500],
    [10000, 87500],
    [12345.67, 89846],
    [200000, 277500],
    [1000000, 677500],
    [2000000, 677500],
  ];
  for (const [principal, expected] of cases) assertEquals(costs(principal), expected);
});

test("several invoices are added before the staffel, and joined", () => {
  const three = run({ invoices: [invoice(200, "A"), invoice(200, "B"), invoice(200, "C")] });
  assertEquals(three.costsCents, 9000);
  assertEquals(three.invoices, "A, B en C");
  const eight = run({ invoices: "ABCDEFGH".split("").map((n) => invoice(10, n)) });
  assertEquals(eight.invoices, "A, B, C, D, E, F, G en H");
});

test("invoices and their fields are validated", async () => {
  await fails({ invoices: [] }, "invoices");
  await fails({ invoices: undefined }, "invoices");
  for (const bad of [0, -50, NaN]) await fails({ invoices: [invoice(bad)] }, "principal");
  await fails({ invoices: [{ ...invoice(10), number: " " }] }, "number");
  await fails({ invoices: [{ ...invoice(10), dueDate: "2026-11-19" }] }, "dueDate");
  assertEquals(run({ invoices: [{ ...invoice(10), dueDate: "2026-11-18" }] }).outcome, "consumer-notice");
});

test("btw is added only when the creditor cannot reclaim it", () => {
  const no = run({ creditorCanReclaimVat: false, invoices: [invoice(2500)] });
  assertEquals([no.costsCents, no.vatCents], [37500, 7875]);
  assertStringIncludes(no.costsPhrase, "inclusief €\u00a078,75 btw");
  assertEquals(run({ creditorCanReclaimVat: false, invoices: [invoice(12345.67)] }).vatCents, 18868);
  for (const reclaim of [undefined, true]) {
    const out = run({ creditorCanReclaimVat: reclaim });
    assertEquals([out.vatCents, out.vat], [0, ""]);
  }
});

test("interest joins the base only after a full year", () => {
  const inv = { invoices: [invoice(2400)], interestAccrued: 150 };
  const year = run({ ...inv, interestRunningSince: "2025-07-01" });
  assertEquals([year.baseCents, year.costsCents], [255000, 38000]);
  const young = run({ ...inv, interestRunningSince: "2026-01-01" });
  assertEquals([young.baseCents, young.costsCents, young.interest], [240000, 36000, ""]);
  assertEquals(run({ ...inv }).baseCents, 240000);
  // 29 February rolls to 1 March: 2025-02-28 is too early, 2025-03-01 is enough.
  const leap = (sendDate: string) => run({ ...inv, interestRunningSince: "2024-02-29", sendDate, invoices: [{ ...invoice(2400), dueDate: "2025-01-01" }] })
      .baseCents;
  assertEquals([leap("2025-02-28"), leap("2025-03-01")], [240000, 255000]);
});

test("negative interest throws", () => fails({ interestAccrued: -1 }, "interestAccrued"));

test("the payment date is rolled past weekends and holidays", () => {
  const deadline = (over: Obj) => run(over).deadline;
  assertEquals(deadline({}), "maandag 7 december 2026");
  assertEquals(deadline({ channel: "post" }), "dinsdag 8 december 2026");
  assertEquals(deadline({ channel: "post", postDeliveryDays: 0 }), "vrijdag 4 december 2026");
  assertEquals(deadline({ sendDate: "2026-12-09" }), "maandag 28 december 2026");
  assertEquals(deadline({ sendDate: "2027-03-13" }), "dinsdag 30 maart 2027");
  assertEquals(deadline({ sendDate: "2027-04-20" }), "vrijdag 7 mei 2027");
});

test("channel and post days are validated", async () => {
  await fails({ channel: undefined }, "channel");
  await fails({ channel: "fax" }, "channel");
  await fails({ postDeliveryDays: -1 }, "postDeliveryDays");
  await fails({ postDeliveryDays: 1.5 }, "postDeliveryDays");
});

test("a consumer instalment is capped", async () => {
  const inst = (priorCharged: number, windowStart?: string, principal = 150) =>
    run({ invoices: [invoice(principal)], instalment: { priorCharged, windowStart } }).costsCents;
  assertEquals(inst(0), 4000);
  assertEquals(inst(2, "2026-08-01"), 2000);
  assertEquals(inst(2, "2026-05-01"), 4000);
  assertEquals(inst(0, undefined, 300), 4500);
  await fails({ invoices: [invoice(150)], instalment: { priorCharged: 2 } }, "windowStart");
  await fails({ invoices: [invoice(150), invoice(150, "F2")], instalment: { priorCharged: 0 } }, "instalment");
  assertEquals(
    run({ debtorType: "business", invoices: [invoice(150), invoice(150, "F2")], instalment: { priorCharged: 0 } })
      .costsCents,
    8000, // two invoices at EUR 40 each: the instalment cap never applies
  );
});

test("a later consumer letter claims the stated figure only when due", async () => {
  const later = (previousLetter: Obj, over: Obj = {}) => run({ sendDate: "2026-12-09", previousLetter, ...over });
  const due = later({ deadline: "2026-12-08", statedCosts: 48.4 }, { creditorCanReclaimVat: false });
  assertEquals([due.outcome, due.costsCents, due.vatCents, due.previousDeadline], [
    "consumer-due",
    4840,
    0,
    "dinsdag 8 december 2026",
  ]);
  assertEquals(run({ sendDate: "2026-12-08", previousLetter: { deadline: "2026-12-08", statedCosts: 48.4 } }).outcome, "none");
  assertEquals(later({ deadline: "2026-12-08" }).outcome, "none");
  assertEquals(later({ deadline: "2026-12-08", statedCosts: 0 }).outcome, "none");
  assertEquals(
    later({ deadline: "2026-12-08", statedCosts: 48.4, receivedDate: "2026-11-30" }).outcome,
    "none",
  );
  await assertRejects(() => Promise.resolve().then(() => later({ deadline: "2026-12-08", statedCosts: -1 })), Error, "statedCosts");
  await assertRejects(() => Promise.resolve().then(() => later({ statedCosts: 1 })), Error, "previousLetter.deadline");
});

test("a business owes the larger of staffel and EUR 40 per invoice, or its contract", () => {
  const biz = (over: Obj) => run({ debtorType: "business", channel: undefined, ...over });
  const fixed = biz({ invoices: [invoice(120, "A"), invoice(80, "B"), invoice(60, "C")] });
  assertEquals([fixed.outcome, fixed.costsCents], ["business-fixed", 12000]);
  const staffel = biz({ invoices: [invoice(10000)] });
  assertEquals([staffel.outcome, staffel.costsCents], ["business-staffel", 87500]);
  assertEquals(biz({ invoices: [invoice(100)] }).outcome, "business-staffel");
  const contract = biz({ contractualCosts: 1500, creditorCanReclaimVat: false });
  assertEquals([contract.outcome, contract.costsCents, contract.vatCents], ["business-contract", 150000, 31500]);
  assertEquals(biz({ contractualCosts: 0 }).outcome, "business-staffel");
  assertEquals(biz({ previousLetter: { deadline: "2026-12-08" } }).outcome, "business-staffel");
});

test("an agency needs both its name and its number", async () => {
  await fails({ agency: { name: "Bureau", registrationNumber: " " } }, "agency.registrationNumber");
  await fails({ agency: { name: "", registrationNumber: "1" } }, "agency.name");
});

test("bad inputs throw", async () => {
  await fails({ debtorType: "person" }, "debtorType");
  await fails({ sendDate: "19-11-2026" }, "sendDate");
});

// ---------------------------------------------------------------------------
// The node
// ---------------------------------------------------------------------------

const documentOf = (arrangement?: "stacked" | "table") =>
  template(
    <Document title="Incasso">
      <NlCollectionCosts arrangement={arrangement} />
    </Document>,
  );

/** Every id in the tree, in reading order. */
function idsOf(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { id?: string; children?: DocumentNode[] };
    return [...(own.id ? [own.id] : []), ...idsOf(own.children ?? [])];
  });
}
function textsOf(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    const own = node as { text?: string; children?: DocumentNode[] };
    return [...(own.text ? [own.text] : []), ...textsOf(own.children ?? [])];
  });
}

const build = (collection: Obj, arrangement?: "stacked" | "table") =>
  buildDocument(documentOf(arrangement), { collection }, { dynamicMode: "placeholder" });

const arms = ["notice", "due", "business-staffel", "business-fixed", "business-contract"].map(
  (a) => "nl-collection-costs-" + a,
);
const armsIn = (ids: string[]) => ids.filter((id) => arms.includes(id));
const agency = { name: "Bureau B.V.", registrationNumber: "2024-1" };

test("each outcome prints exactly its own arm", async () => {
  const cases: Array<[Obj, string]> = [
    [base(), "notice"],
    [base({ sendDate: "2026-12-09", previousLetter: { deadline: "2026-12-08", statedCosts: 48.4 } }), "due"],
    [base({ debtorType: "business" }), "business-staffel"],
    [base({ debtorType: "business", invoices: [invoice(50, "A"), invoice(50, "B")] }), "business-fixed"],
    [base({ debtorType: "business", contractualCosts: 100 }), "business-contract"],
  ];
  for (const [collection, arm] of cases) {
    const ids = idsOf((await build(collection)).nodes);
    assertEquals(armsIn(ids), ["nl-collection-costs-" + arm]);
    assertEquals(ids.includes("nl-collection-costs"), true);
  }
});

test("none prints no section, but an agency still prints its number", async () => {
  const none = base({ sendDate: "2026-12-08", previousLetter: { deadline: "2026-12-08", statedCosts: 48.4 } });
  assertEquals(idsOf((await build(none)).nodes), []);
  const withAgency = await build({ ...none, agency });
  assertEquals(idsOf(withAgency.nodes), ["nl-collection-costs-agency"]);
  assertStringIncludes(
    textsOf(withAgency.nodes).join("\n"),
    "Bureau B.V. is ingeschreven in het register incassodienstverlening onder nummer 2024-1.",
  );
});

test("the interest and btw lines appear only when they have a value", async () => {
  for (const arrangement of ["stacked", "table"] as const) {
    const plain = idsOf((await build(base(), arrangement)).nodes);
    assertEquals([plain.includes("nl-collection-costs-interest"), plain.includes("nl-collection-costs-vat")], [false, false]);
    const full = idsOf(
      (await build(
        base({ creditorCanReclaimVat: false, interestAccrued: 10, interestRunningSince: "2025-01-01" }),
        arrangement,
      )).nodes,
    );
    assertEquals([full.includes("nl-collection-costs-interest"), full.includes("nl-collection-costs-vat")], [true, true]);
  }
});

// ---------------------------------------------------------------------------
// Publishing, then resolving where only the bundle exists
// ---------------------------------------------------------------------------

async function resolved(collection: Obj, arrangement?: "stacked" | "table") {
  const data = { collection };
  const published = await buildDocument(documentOf(arrangement), createPublishData() as typeof data, {
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
  return await resolveDocument(published, state, { derivers });
}

const preview = COMPONENTS.find((c) => c.id === "nl-collection-costs")!.previewData as { collection: Obj };

test("the preview data resolves to the stated figures, in either arrangement", async () => {
  for (const arrangement of ["stacked", "table"] as const) {
    const text = textsOf((await resolved(preview.collection, arrangement)).nodes).join("\n");
    for (const expected of ["dinsdag 8 december 2026", "€\u00a01.172,03", "€\u00a020.534,51", "€\u00a019.362,48", "2024-0048176"]) {
      assertStringIncludes(text, expected);
    }
  }
});

test("a none input resolves to no section", async () => {
  const none = { ...preview.collection, previousLetter: { deadline: "2026-11-19" }, sendDate: "2026-11-19" };
  const ids = idsOf((await resolved(none)).nodes);
  assertEquals(ids.includes("nl-collection-costs"), false);
});

test("a publish build carries the branches the doc comment quotes", async () => {
  const arrangementBranches = async (name: string) => {
    const file = `components/arrangements/nl-collection-costs/${name}.arrangement.tsx`;
    return transformDocumentSource(await readFile(join(registryRoot(), file), "utf8"), { fileName: file }).branches;
  };
  assertEquals([branches, await arrangementBranches("stacked"), await arrangementBranches("table")], [7, 2, 2]);
});
