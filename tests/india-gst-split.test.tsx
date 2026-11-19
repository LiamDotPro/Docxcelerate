import { test } from "node:test";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import {
  buildDocument,
  createDeriverBundle,
  createDeriverRegistryFromBundle,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentNode, RuntimeState } from "docxcelerate";
import { createPublishData, Document, template } from "docxcelerate/template";
import { COMPONENTS } from "../src/registry/catalog.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { default: mod } = await loadRegistryModule<
  typeof import("../registry/components/derivers/india-gst-split.deriver.ts")
>("components/derivers/india-gst-split.deriver.ts");
const { IndiaGstSplit } = await loadRegistryModule<
  typeof import("../registry/components/india-gst-split/india-gst-split.node.tsx")
>("components/india-gst-split/india-gst-split.node.tsx");

type Obj = Record<string, unknown>;
type Line = { hsnSac: string; taxableMinor: number; rate: number };
const MH = "27AAACR5055K1Z5";

interface Opts {
  gstin?: unknown;
  stateCode?: unknown;
  pos?: string;
  stateName?: unknown;
  customer?: unknown;
  lut?: boolean;
  rc?: unknown;
  lines?: unknown;
}

const call = (inputs: unknown[]) => (mod.run as (i: unknown[], s: RuntimeState) => unknown)(inputs, {} as RuntimeState);

const args = (o: Opts) =>
  [
    { gstin: "gstin" in o ? o.gstin : MH, stateCode: o.stateCode },
    { stateCode: "pos" in o ? o.pos : "27", stateName: "stateName" in o ? o.stateName : "Maharashtra" },
    o.customer,
    { exportUnderLut: o.lut, reverseCharge: "rc" in o ? o.rc : false },
    "lines" in o ? o.lines : [{ hsnSac: "998314", taxableMinor: 10003, rate: 18 }],
  ] as const;

// deno-lint-ignore no-explicit-any
const run = (o: Opts = {}): any => call([...args(o)]);
const rejects = (o: Opts, field: string) => assertRejects(async () => await run(o), Error, field);
const item = (taxableMinor: number, rate = 18, hsnSac = "998314"): Line[] => [{ hsnSac, taxableMinor, rate }];

const preview = COMPONENTS.find((c) => c.id === "india-gst-split")!.previewData as Obj;

test("the catalog preview lands on UTGST with the expected rows and totals", () => {
  // deno-lint-ignore no-explicit-any
  const r = call([preview.seller, preview.placeOfSupply, preview.customer, preview.supply, preview.lines]) as any;

  assertEquals(r.intraState, true);
  assertEquals(r.stateTaxLabel, "UTGST");
  assertEquals(r.placeOfSupply, "Dadra and Nagar Haveli and Daman and Diu (26)");
  assertEquals(r.rows.length, 4);
  assertEquals(r.rows[0], {
    hsnSac: "84714190",
    taxable: "₹12,35,56,789.49",
    rate: "18%",
    halfRate: "9%",
    centralTax: "₹1,11,20,111.05",
    stateTax: "₹1,11,20,111.05",
    integratedTax: "",
    totalTax: "₹2,22,40,222.10",
  });
  assertEquals(r.rows[3].hsnSac, "71023910");
  assertEquals(r.rows[3].halfRate, "0.125%");
  assertEquals(r.rows[3].centralTax, "₹123.46");
  assertEquals(r.rows[3].stateTax, "₹123.46");
  assertEquals(r.taxableTotal, "₹17,16,80,135.08");
  assertEquals(r.taxTotal, "₹3,05,79,955.18");
  assertEquals(r.roundOff, "-₹0.26");
  assertEquals(r.grandTotal, "₹20,22,60,090.00");
  assertEquals(r.grandTotalInWords, "Rupees Twenty Crore Twenty Two Lakh Sixty Thousand Ninety Only");
});

test("an odd paisa splits into two halves that differ from IGST on the same base", () => {
  const intra = run();

  assertEquals(intra.rows[0].centralTax, "₹9.00");
  assertEquals(intra.rows[0].stateTax, "₹9.00");
  assertEquals(intra.taxTotal, "₹18.00");
  assertEquals(intra.stateTaxLabel, "SGST");
  assertEquals(intra.roundOff, "-₹0.03");
  assertEquals(intra.grandTotal, "₹118.00");
  assertEquals(intra.grandTotalInWords, "Rupees One Hundred Eighteen Only");

  const inter = run({ pos: "29" });

  assertEquals(inter.intraState, false);
  assertEquals(inter.stateTaxLabel, "");
  assertEquals(inter.rows[0].integratedTax, "₹18.01");
  assertEquals(inter.rows[0].rate, "18%");
  assertEquals(inter.rows[0].halfRate, "");
  assertEquals(inter.rows[0].centralTax, "");
  assertEquals(inter.taxTotal, "₹18.01");
  assertEquals(inter.roundOff, "-₹0.04");
  assertEquals(inter.grandTotal, "₹118.00");
});

test("the state name is printed and never read", () => {
  const r = run({ stateName: "  Karnataka " });

  assertEquals(r.intraState, true);
  assertEquals(r.placeOfSupply, "Karnataka (27)");
});

test("Delhi and Jammu and Kashmir are SGST, the other territories UTGST", () => {
  assertEquals(run({ gstin: "07AAACR5055K1Z5", pos: "07" }).stateTaxLabel, "SGST");
  assertEquals(run({ gstin: "01AAACR5055K1Z5", pos: "01" }).stateTaxLabel, "SGST");
  assertEquals(run({ gstin: "34AAACR5055K1Z5", pos: "34" }).stateTaxLabel, "SGST");
  for (const code of ["04", "25", "26", "31", "35", "38", "97"]) {
    assertEquals(run({ gstin: `${code}AAACR5055K1Z5`, pos: code }).stateTaxLabel, "UTGST");
  }
});

test("an export under LUT is inter-state at nil", () => {
  const r = run({ pos: "96", lut: true });

  assertEquals(r.intraState, false);
  assertEquals(r.rows[0].rate, "Nil");
  assertEquals(r.rows[0].integratedTax, "₹0.00");
  assertStringIncludes(r.zeroRatedMention, "SUPPLY MEANT FOR EXPORT UNDER BOND");
  assertEquals(r.grandTotal, "₹100.00");
});

test("an export without LUT, and an SEZ either way, carry the Rule 46 mention", () => {
  const exp = run({ pos: "96" });

  assertEquals(exp.zeroRatedMention, "SUPPLY MEANT FOR EXPORT ON PAYMENT OF INTEGRATED TAX");
  assertEquals(exp.rows[0].integratedTax, "₹18.01");

  const sez = run({ customer: { isSez: true } });

  assertEquals(sez.intraState, false);
  assertEquals(sez.rows[0].integratedTax, "₹18.01");
  assertStringIncludes(sez.zeroRatedMention, "ON PAYMENT OF INTEGRATED TAX");
  assertStringIncludes(sez.zeroRatedMention, "SUPPLY TO SEZ UNIT");

  const lut = run({ customer: { isSez: true }, lut: true });

  assertStringIncludes(lut.zeroRatedMention, "UNDER BOND OR LETTER OF UNDERTAKING WITHOUT PAYMENT");
  assertEquals(lut.rows[0].rate, "Nil");
  assertEquals(run().zeroRatedMention, "");
});

test("reverse charge shows the tax and leaves it out of the amount payable", () => {
  const r = run({ rc: true });

  assertEquals(r.reverseCharge, "Yes");
  assertEquals(r.taxTotal, "₹18.00");
  assertStringIncludes(r.reverseChargeNote, "₹18.00");
  assertEquals(r.grandTotal, "₹100.00");
  assertEquals(r.roundOff, "-₹0.03");
  assertEquals(run().reverseCharge, "No");
  assertEquals(run().reverseChargeNote, "");
});

test("zero value, a nil rate and a missing customer", () => {
  const z = run({ lines: item(0) });

  assertEquals(z.taxableTotal, "₹0.00");
  assertEquals(z.taxTotal, "₹0.00");
  assertEquals(z.roundOff, "₹0.00");
  assertEquals(z.grandTotal, "₹0.00");
  assertEquals(z.grandTotalInWords, "Rupees Zero Only");

  const nil = run({ lines: item(12345, 0, "4901") });

  assertEquals(nil.rows[0].rate, "0%");
  assertEquals(nil.rows[0].halfRate, "0%");
  assertEquals(nil.rows[0].totalTax, "₹0.00");
  assertEquals(run({ customer: undefined }).intraState, true);
});

test("a credit mirrors its invoice", () => {
  const r = run({ lines: item(-10003) });

  assertEquals(r.rows[0].centralTax, "-₹9.00");
  assertEquals(r.taxTotal, "-₹18.00");
  assertEquals(r.roundOff, "+₹0.03");
  assertEquals(r.grandTotal, "-₹118.00");
  assertEquals(r.grandTotalInWords, "Minus Rupees One Hundred Eighteen Only");
});

test("a half rupee rounds away from zero", () => {
  const up = run({ lines: item(10050, 0, "9983") });

  assertEquals(up.roundOff, "+₹0.50");
  assertEquals(up.grandTotal, "₹101.00");

  const down = run({ lines: item(-10050, 0, "9983") });

  assertEquals(down.roundOff, "-₹0.50");
  assertEquals(down.grandTotal, "-₹101.00");
});

const many: Line[] = [
  { hsnSac: "9983", taxableMinor: 100000, rate: 18 },
  { hsnSac: "9983", taxableMinor: 50000, rate: 5 },
  { hsnSac: "9983", taxableMinor: 25000, rate: 18 },
  { hsnSac: "4901", taxableMinor: 12345, rate: 0 },
];

test("lines group by HSN/SAC and rate, in first-seen order", () => {
  const r = run({ lines: many });

  assertEquals(r.rows.length, 3);
  assertEquals(r.rows.map((x: Obj) => [x.hsnSac, x.halfRate]), [
    ["9983", "9%"],
    ["9983", "2.5%"],
    ["4901", "0%"],
  ]);
  assertEquals(r.rows[0].taxable, "₹1,250.00");
  assertEquals(r.rows[0].centralTax, "₹112.50");
  assertEquals(r.rows[1].taxable, "₹500.00");
  assertEquals(r.taxTotal, "₹250.00");
  assertEquals(r.roundOff, "-₹0.45");
  assertEquals(r.grandTotal, "₹2,123.00");
  assertEquals(r.grandTotalInWords, "Rupees Two Thousand One Hundred Twenty Three Only");
});

test("the total is written in lakh and crore", () => {
  const words = (rupees: number) => run({ lines: item(rupees * 100, 0, "9983") }).grandTotalInWords;

  assertEquals(
    words(1234567890),
    "Rupees One Hundred Twenty Three Crore Forty Five Lakh Sixty Seven Thousand Eight Hundred Ninety Only",
  );
  assertEquals(words(100000), "Rupees One Lakh Only");
  assertEquals(words(1e7), "Rupees One Crore Only");
  assertEquals(words(1e12), "Rupees One Lakh Crore Only");
  assertEquals(words(19), "Rupees Nineteen Only");
});

test("bad input names the field it rejects", async () => {
  await rejects({ gstin: undefined }, "seller.gstin");
  await rejects({ gstin: "27AAACR5055K1Z" }, "seller.gstin");
  await rejects({ gstin: "00AAACR5055K1Z5" }, "seller.gstin");
  await rejects({ stateCode: "29" }, "seller.stateCode");
  await rejects({ stateName: undefined }, "placeOfSupply.stateName");
  await rejects({ stateName: "  " }, "placeOfSupply.stateName");
  await rejects({ pos: 7 as unknown as string }, "placeOfSupply.stateCode");
  await rejects({ pos: "7" }, "placeOfSupply.stateCode");
  await rejects({ rc: undefined }, "supply.reverseCharge");
  await rejects({ lines: [] }, "lines");
  await rejects({ lines: undefined }, "lines");
  await rejects({ lines: item(1.5) }, "taxableMinor");
  await rejects({ lines: item(100, 18.005) }, "rate");
  await rejects({ lines: item(100, -1) }, "rate");
  await rejects({ lines: item(100, 18, "847") }, "hsnSac");
  await rejects({ lines: item(100, 18, "84714A") }, "hsnSac");
  await rejects({ lut: true }, "supply.exportUnderLut");
  await rejects({ pos: "96", customer: { isSez: true } }, "customer.isSez");
  await rejects({ lines: item(Number.MAX_SAFE_INTEGER) }, "overflows");
});

// ---------------------------------------------------------------------------
// Publishing
// ---------------------------------------------------------------------------

type Data = Record<string, unknown>;

const template_ = template<Data>(
  <Document id="invoice" title="Invoice">
    <IndiaGstSplit />
  </Document>,
);

const published = await buildDocument(template_, createPublishData() as Data, {
  branchMode: "publish",
  deriverMode: "preserve",
  dynamicMode: "resolve",
  aiClient: { generateParagraph: () => "" },
});

const walk = (nodes: DocumentNode[], into = new Map<string, DocumentNode>()) => {
  for (const node of nodes) {
    into.set(node.id, node);
    const children = (node as { children?: DocumentNode[] }).children;
    if (children) walk(children, into);
  }

  return into;
};

async function resolve(o: Opts) {
  const [seller, placeOfSupply, customer, supply, lines] = args(o);
  const data = { seller, placeOfSupply, customer, supply, lines };
  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: { generateParagraph: () => "" },
  };
  const resolved = await resolveDocument(published, state, {
    derivers: await createDeriverRegistryFromBundle(createDeriverBundle({ indiaGstSplit: mod.run })),
  });

  return walk(resolved.nodes);
}

test("a publish build carries both table arms, a loop over derived output and one invocation", () => {
  const found = walk(published.nodes);

  for (const id of ["gst-split-intra", "gst-split-inter"]) {
    if ((found.get(id) as { when?: unknown } | undefined)?.when === undefined) throw new Error(`${id} has no condition`);
  }
  for (const id of ["gst-split-zero-rated", "gst-split-reverse-charge-note"]) {
    if ((found.get(id) as { when?: unknown } | undefined)?.when === undefined) throw new Error(`${id} has no condition`);
  }

  for (const id of ["gst-split-intra", "gst-split-inter"]) {
    const repeats = (found.get(id) as unknown as { children: Array<{ kind: string; source?: { scope: string } }> })
      .children.filter((c) => c.kind === "repeat");

    assertEquals(repeats.length, 1);
    assertEquals(repeats[0].source?.scope, "derived");
  }

  const section = JSON.stringify(found.get("india-gst-split"));

  assertEquals(section.split('"name":"indiaGstSplit"').length - 1, 1);
});

test("a resolved document keeps one table with one row per summary group", async () => {
  const intra = await resolve({ lines: many });
  const local = run({ lines: many });

  assertEquals([...intra.keys()].filter((id) => id.startsWith("gst-split-intra-row")), [
    "gst-split-intra-row-0",
    "gst-split-intra-row-1",
    "gst-split-intra-row-2",
  ]);
  assertEquals(intra.has("gst-split-inter"), false);
  assertEquals(JSON.stringify(intra.get("gst-split-intra-total")).includes(local.taxTotal), true);
  assertEquals(JSON.stringify(intra.get("gst-split-grand-total")).includes(local.grandTotal), true);

  const inter = await resolve({ pos: "29", lines: many });

  assertEquals(inter.has("gst-split-intra"), false);
  assertEquals(inter.has("gst-split-inter"), true);
  assertEquals([...inter.keys()].filter((id) => id.startsWith("gst-split-inter-row")).length, 3);
});
