import type { JsonObject } from "../domain/types.ts";

/**
 * The component registry: prebuilt nodes, described once.
 *
 * The entries here are metadata. The code itself lives under `registry/` at the
 * root of the package, as ordinary `.tsx` files that import `docxcelerate` the
 * way a project would — which means they are typechecked against the real
 * framework rather than kept as strings, and `dxcl add` installs them by
 * copying rather than by generating.
 *
 * Copying is the point. An installed component is your file: it has no version,
 * it is not upgraded behind you, and editing it is the expected next step
 * rather than a fork. What the registry gives you is the first draft and the
 * decisions already made in it.
 *
 * @module
 */

/** Where in a document a component usually sits. */
export type ComponentCategory =
  | "Opening"
  | "Body"
  | "Closing"
  | "Legal"
  | "Generated";

/** One field an installed component expects to find on the project's data. */
export interface RegistryDataField {
  /** The path, as it is read — `sender.name`, not `name`. */
  path: string;
  /** The TypeScript type the component reads it as. */
  type: string;
  /** What it is for, and what happens when it is missing. */
  summary: string;
}

/**
 * One way a component can place what it prints, chosen with its `arrangement`
 * prop.
 *
 * Arrangements change where the lines sit, never which lines print: every one
 * of a component's arrangements prints the same lines under the same
 * conditions. The component picks between them with `arrange()` from
 * `docxcelerate/template`, which treats the first as the default, the same
 * order these are listed in. Each one is a file of its own, installed under
 * `arrangements/<component id>/`.
 */
export interface RegistryArrangement {
  /** The value the `arrangement` prop takes. */
  id: string;
  /** Its name, as it is printed. */
  title: string;
  /** One line: where things sit, and when you would want that. */
  summary: string;
}

/** One file an install copies. */
export interface RegistryFile {
  /** Path under `registry/`, inside the package. */
  source: string;
  /** Where it lands, relative to the document project. */
  target: string;
}

/** A prebuilt component, as the CLI and the site read it. */
export interface RegistryComponent {
  /** Discriminator, so themes and components can share one lookup. */
  kind: "component";
  /** The id: the URL slug, and what `dxcl add` is given. */
  id: string;
  /** The component's name, as it is printed. */
  title: string;
  /** One line, for cards and lists. */
  summary: string;
  /** A paragraph: what it does, and the decision it has already made for you. */
  detail: string;
  /** Where in a document it usually sits. */
  category: ComponentCategory;
  /** Free-form labels a reader might filter on. */
  tags: string[];
  /** The component names the files export, in the order they should be used. */
  exports: string[];
  /** What it reads, so a project knows what to add to its data type. */
  dataFields: RegistryDataField[];
  /**
   * The placements its `arrangement` prop chooses between, the default first.
   * Missing for a component that has only the one.
   */
  arrangements?: RegistryArrangement[];
  /** The files copied in. */
  files: RegistryFile[];
  /** Data the previews on the site are built against, and a project can crib. */
  previewData: JsonObject;
  /** Other registry ids installed alongside this one. */
  requires: string[];
  /** The theme it was drawn against, where it was drawn against one. */
  themeHint?: string;
}

/**
 * One file per component, named after the component, under `nodes/`.
 *
 * Flat rather than a directory each: a document project keeps every node in one
 * place, and an installed component is meant to stop being distinguishable from
 * one you wrote as soon as you have edited it.
 *
 * A component that computes on request data ships one deriver beside it, the
 * default export of `derivers/<id>.deriver.ts`, named the id in camel case. It
 * sits in a `derivers/` folder in the registry too, so the node's
 * `../derivers/<id>.deriver.ts` import resolves the same before and after it is
 * copied.
 *
 * A component with arrangements keeps each one in its own file, laid out the
 * same way and for the same reason: `arrangements/<id>/<name>.arrangement.tsx`,
 * beside `arrangements/<id>/lines.ts`, the type of what the node hands every
 * one of them. The node imports them from `../arrangements/<id>/`, so it reads
 * as what it decides, and each file reads as where things sit.
 */
function node(id: string, options: { deriver?: boolean; arrangements?: string[] } = {}): RegistryFile[] {
  const arranged = (file: string) => ({
    source: `components/arrangements/${id}/${file}`,
    target: `arrangements/${id}/${file}`,
  });

  return [
    { source: `components/${id}/${id}.node.tsx`, target: `nodes/${id}.node.tsx` },
    ...(options.deriver ? [{ source: `components/derivers/${id}.deriver.ts`, target: `derivers/${id}.deriver.ts` }] : []),
    ...(options.arrangements
      ? [arranged("lines.ts"), ...options.arrangements.map((name) => arranged(`${name}.arrangement.tsx`))]
      : []),
  ];
}

/**
 * Every component `dxcl add` can install, in the order the catalog lists them.
 *
 * Each entry describes what the component does, what it reads from a project's
 * data, and which files installing writes. The source itself is not here — it
 * lives under `registry/` and is read from disk, so the catalog and the file
 * that gets copied cannot disagree about what a component is.
 */
export const COMPONENTS: RegistryComponent[] = [
  {
    kind: "component",
    id: "status-banner",
    title: "Status banner",
    summary: "A drawn block at the top of the page saying where things stand.",
    detail:
      "The component that shows what a shape is for. A banner has to be the " +
      "same depth on every document it goes out on — the one saying 'Approved' " +
      "and the one saying 'Awaiting signature, 14 days' have to be the same " +
      "block, or a reader flicking through a stack sees the shorter one as a " +
      "different kind of notice. A paragraph with a background grows with its " +
      "words; a shape is the size you gave it, so the three tones are three " +
      "wordings of one block rather than three blocks. What the banner is " +
      "about is yours: it prints the line it is given under one of three " +
      "tones, so the same component carries an approval, a deadline, a draft " +
      "mark or a refusal. The colours are the theme's, named for the state " +
      "they report rather than the colour they happen to be.",
    category: "Opening",
    tags: ["shape", "banner", "status", "branching"],
    exports: ["StatusBanner"],
    dataFields: [
      {
        path: "status.label",
        type: "string",
        summary: "The line printed across the banner.",
      },
      {
        path: "status.tone",
        type: '"positive" | "attention" | "critical" | undefined',
        summary: "Which of the three it is drawn as. Anything else draws the attention tone.",
      },
      {
        path: "status.note",
        type: "string | undefined",
        summary: "A second thing on the line — a date, a reference, whatever follows.",
      },
    ],
    files: node("status-banner"),
    previewData: {
      status: {
        label: "Awaiting signature",
        tone: "attention",
        note: "please return by 30 April 2026",
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "letterhead",
    title: "Letterhead",
    summary: "Who sent this, from where, and when.",
    detail:
      "The block at the top of the page. The address is one paragraph with " +
      "the lines joined rather than one node per line, so a sender with two " +
      "address lines and one with five both come out as a block instead of a " +
      "ragged run of nodes.",
    category: "Opening",
    tags: ["letter", "header", "static"],
    exports: ["Letterhead"],
    dataFields: [
      {
        path: "sender.name",
        type: "string",
        summary: "The organisation, as it should be printed.",
      },
      {
        path: "sender.addressLines",
        type: "string[]",
        summary: "Street, town, postcode — one entry per line, joined for print.",
      },
      {
        path: "sentOn",
        type: "string | number | Date",
        summary: "When the document was sent. Formatted for the build locale.",
      },
    ],
    files: node("letterhead"),
    previewData: {
      sender: {
        name: "Riverside Leisure Centre",
        addressLines: ["14 Mill Lane", "Bristol", "BS1 4TG"],
      },
      sentOn: "2026-09-01",
    },
    requires: [],
  },
  {
    kind: "component",
    id: "recipient-block",
    title: "Recipient block",
    summary: "The address, and a greeting that survives a missing name.",
    detail:
      "The salutation is the half worth having. A document generated in bulk " +
      "meets recipients whose names it does not have — a joint tenancy, a " +
      "company, a record where the field was never filled in — so the " +
      "greeting branches, and the fallback is formal rather than clever.",
    category: "Opening",
    tags: ["letter", "address", "branching"],
    exports: ["RecipientBlock"],
    dataFields: [
      {
        path: "recipient.name",
        type: "string | undefined",
        summary: "The addressee. Absent is handled: the greeting falls back.",
      },
      {
        path: "recipient.addressLines",
        type: "string[]",
        summary: "Street, town, postcode — one entry per line.",
      },
      {
        path: "recipient.formalName",
        type: "string | undefined",
        summary: "Used in the greeting where a first name would be too familiar.",
      },
    ],
    files: node("recipient-block"),
    previewData: {
      recipient: {
        name: "Adaeze Nkemelu",
        formalName: "Ms Nkemelu",
        addressLines: ["Flat 6", "22 Colston Street", "Bristol", "BS1 5AE"],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "payment-summary",
    title: "Payment summary",
    summary: "What is owed, by when — and what to say when nothing is.",
    detail:
      "Three outcomes, three ids: in credit, clear, or owing. Branching on " +
      "the amount rather than papering over it with one sentence that reads " +
      "oddly at zero is the point — somebody who owes nothing should not be " +
      "given a payment deadline.",
    category: "Body",
    tags: ["money", "branching", "section", "localised"],
    exports: ["PaymentSummary"],
    dataFields: [
      {
        path: "account.reference",
        type: "string",
        summary: "Your reference, printed so a caller can quote it.",
      },
      {
        path: "account.balanceDue",
        type: "number",
        summary: "What is owed. Negative means in credit; zero means clear.",
      },
      {
        path: "account.dueBy",
        type: "string | number | Date | undefined",
        summary: "When it is due. Printed only where something is owed.",
      },
      {
        path: "account.currency",
        type: "string | undefined",
        summary: "ISO 4217 code. Defaults to GBP.",
      },
    ],
    files: node("payment-summary"),
    previewData: {
      account: {
        reference: "RIV-88214",
        balanceDue: 128.42,
        dueBy: "2026-09-30",
        currency: "GBP",
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "invoice-totals",
    title: "Invoice totals",
    summary: "Line items and VAT totals, worked in minor units and rounded once per rate.",
    detail:
      "The arithmetic every invoice gets subtly wrong, done once in a deriver. " +
      "Money is worked in integer minor units at the currency's own exponent, " +
      "so yen and dinars come out right and no float reaches a total. Each " +
      "line is rounded once, lines are grouped by rate, and VAT is worked once " +
      "per band on the band's sum. That is rate-level rounding, which HMRC and " +
      "the EU accept, and it is the figure an accounting system recomputes; " +
      "per-line VAT summed drifts a penny from it. Inclusive prices have their " +
      "VAT backed out of the band's gross, so the total is exactly what was " +
      "charged. Discounts reduce their own band. Exempt lines print under " +
      "their own label rather than as 0%, and more than one band draws the " +
      "per-rate VAT analysis an invoice is required to carry.",
    category: "Body",
    tags: ["money", "table", "tax", "invoice", "branching", "section"],
    exports: ["InvoiceTotals"],
    dataFields: [
      {
        path: "invoice.currency",
        type: "string",
        summary: "ISO 4217 code. Required; sets the decimal places.",
      },
      {
        path: "invoice.pricing",
        type: '"exclusive" | "inclusive" | undefined',
        summary: "Whether unit prices exclude or include VAT. Defaults to exclusive.",
      },
      {
        path: "invoice.lines",
        type:
          "Array<{ description: string; quantity: number; unit?: string; unitPrice: number; taxRate: number | null; taxLabel?: string }>",
        summary:
          "One per line, in print order. `taxRate` is a percentage; `null` is exempt or outside scope, labelled by `taxLabel`. Empty prints \"no charges\".",
      },
    ],
    files: node("invoice-totals", { deriver: true }),
    previewData: {
      invoice: {
        currency: "GBP",
        pricing: "exclusive",
        lines: [
          { description: "Quarterly planned preventive maintenance of rooftop air handling units AHU-1 to AHU-6, including filter replacement and belt tensioning", quantity: 4, unit: "visits", unitPrice: 12850, taxRate: 20 },
          { description: "Out-of-hours emergency call-out, chilled water pump failure at Riverside House, 14–15 March", quantity: 7.25, unit: "h", unitPrice: 142.5, taxRate: 20 },
          { description: "Replacement variable-speed drive and mounting kit for pump P-03", quantity: 1, unitPrice: 3487.6, taxRate: 20 },
          { description: "Communal heating fuel and power, March (reduced rate)", quantity: 18432.5, unit: "kWh", unitPrice: 0.2847, taxRate: 5 },
          { description: "Printed site logbooks and operating manuals", quantity: 6, unit: "copies", unitPrice: 48.5, taxRate: 0 },
          { description: "Plant breakdown insurance premium recharge, 2026–27", quantity: 1, unitPrice: 2310, taxRate: null, taxLabel: "Exempt" },
          { description: "Loyalty discount on planned maintenance", quantity: 1, unitPrice: -2570, taxRate: 20 },
        ],
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "balance-trend",
    title: "Balance trend",
    summary: "A line of how a balance has moved, and a sentence saying the same thing.",
    detail:
      "The chart and the prose are worked out from one array, which is the " +
      "reason to install this rather than write a <Graph> yourself: a line " +
      "that falls beside a sentence that says it rose is the failure this " +
      "shape rules out. Three decisions come with it. A reading nobody took " +
      "is null rather than zero, so a month that has not been billed draws as " +
      "a gap instead of a cliff. Fewer than two readings is not a trend, so " +
      "one reading prints as a balance and draws nothing — a line through a " +
      "single point is a chart that looks like it says something. And the " +
      "currency goes in the sentence rather than down the axis, where it " +
      "would be six copies of a fact already given.",
    category: "Body",
    tags: ["chart", "line", "money", "branching", "section"],
    exports: ["BalanceTrend"],
    dataFields: [
      {
        path: "balance.history",
        type: "Array<{ period: string; amount: number | null }>",
        summary:
          "One entry per period, oldest first — the order it is drawn in. " +
          "`null` is a period nothing was read for, and draws as a gap.",
      },
      {
        path: "balance.currency",
        type: "string | undefined",
        summary: "ISO 4217 code, for the sentence. Defaults to GBP.",
      },
      {
        path: "balance.numberFormat",
        type: "string | undefined",
        summary:
          "How the value axis prints its figures, as an OOXML number format. " +
          'Defaults to "#,##0".',
      },
    ],
    files: node("balance-trend"),
    previewData: {
      balance: {
        currency: "GBP",
        history: [
          { period: "Apr", amount: 412.5 },
          { period: "May", amount: 388.2 },
          { period: "Jun", amount: null },
          { period: "Jul", amount: 296.75 },
          { period: "Aug", amount: 241.1 },
          { period: "Sep", amount: 128.42 },
        ],
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "usage-breakdown",
    title: "Usage breakdown",
    summary: "A pie of where a total went, with the tail folded into Other.",
    detail:
      "The decision this makes for you is what happens past the sixth slice. " +
      "A breakdown out of real data has as many categories as the system had " +
      "rows, and a chart drawn straight from one runs off the end of the " +
      "palette — where the ninth slice is painted the same as the first and " +
      "two unrelated things look like one. Everything past the largest few is " +
      "added up and drawn as Other, which is as much as a reader holds in " +
      "their head anyway. The slices are sorted largest first, because a pie " +
      "read clockwise is read in order; each one prints its share, because " +
      "three of the shipped palette's hues sit under 3:1 on white and colour " +
      "alone is not an answer.",
    category: "Body",
    tags: ["chart", "pie", "share", "section"],
    exports: ["UsageBreakdown"],
    dataFields: [
      {
        path: "usage.items",
        type: "Array<{ label: string; amount: number }>",
        summary:
          "One entry per category, in any order — this component sorts them. " +
          "Zero and negative amounts are dropped rather than drawn.",
      },
      {
        path: "usage.unit",
        type: "string | undefined",
        summary: 'What is being counted: "kWh", "visits", "hours". Printed in the prose.',
      },
      {
        path: "usage.period",
        type: "string | undefined",
        summary: 'What the breakdown covers: "this quarter", "since April".',
      },
    ],
    files: node("usage-breakdown"),
    previewData: {
      usage: {
        unit: "visits",
        period: "this year",
        items: [
          { label: "Swimming", amount: 96 },
          { label: "Gym", amount: 64 },
          { label: "Classes", amount: 41 },
          { label: "Squash", amount: 18 },
          { label: "Sauna", amount: 12 },
          { label: "Badminton", amount: 9 },
          { label: "Climbing", amount: 4 },
          { label: "Table tennis", amount: 2 },
        ],
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "period-comparison",
    title: "Period comparison",
    summary: "This period against the last, category by category, as clustered bars.",
    detail:
      "The chart every report reaches for, and the one that goes wrong the " +
      "same two ways every time. It turns its bars on their side when the " +
      "labels are long: category names out of a real system are 'Ground " +
      "floor maintenance', not 'Q1', and under a vertical bar those overlap, " +
      "shrink or tilt. Laid down, each label sits beside its bar at full size " +
      "and the chart grows downwards, which is the direction a page has room " +
      "in. And both series are the same measure at two different times — the " +
      "one comparison a shared axis is honest about. Spend against headcount " +
      "is two charts, because a second axis makes the crossing point look " +
      "like a finding when it is an artefact of the scales.",
    category: "Body",
    tags: ["chart", "bar", "comparison", "section"],
    exports: ["PeriodComparison"],
    dataFields: [
      {
        path: "comparison.rows",
        type: "Array<{ label: string; previous: number | null; current: number | null }>",
        summary:
          "One row per category, in the order they are drawn. `null` is a " +
          "category not measured in that period, and draws as a gap.",
      },
      {
        path: "comparison.previousLabel",
        type: "string",
        summary: 'What the earlier run is called: "2024", "Last year".',
      },
      {
        path: "comparison.currentLabel",
        type: "string",
        summary: 'What the later run is called: "2025", "This year".',
      },
      {
        path: "comparison.measure",
        type: "string | undefined",
        summary: "What the figures measure. Printed beside the value axis.",
      },
      {
        path: "comparison.dimension",
        type: "string | undefined",
        summary: "What the categories are. Printed beside the category axis.",
      },
      {
        path: "comparison.numberFormat",
        type: "string | undefined",
        summary:
          "How the value axis prints its figures, as an OOXML number format. " +
          'Defaults to "#,##0".',
      },
    ],
    files: node("period-comparison"),
    previewData: {
      comparison: {
        previousLabel: "2024",
        currentLabel: "2025",
        measure: "Visits",
        dimension: "Activity",
        rows: [
          { label: "Swimming", previous: 84, current: 96 },
          { label: "Gym", previous: 71, current: 64 },
          { label: "Classes", previous: 33, current: 41 },
          { label: "Squash", previous: 18, current: 18 },
          { label: "Climbing", previous: null, current: 4 },
        ],
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "mix-over-time",
    title: "Mix over time",
    summary: "How a make-up shifted across periods, as a hundred per cent stack.",
    detail:
      "A stacked chart answers one of two questions and cannot answer both. " +
      "Stacked to a total, the height of a column is the size of the period. " +
      "Stacked to a hundred per cent, every column is the same height and " +
      "what is read is the mix — which is what anybody asking 'has this " +
      "changed?' is really asking. This takes the second and gives the " +
      "totals back in the sentence, because that is the fact the chart " +
      "deliberately drops. The raw figures go into the file and Word works " +
      "the shares out itself, so nothing here divides. Which categories are " +
      "large is decided once across the whole span rather than per column: " +
      "folding the tail into Other per period would put a category in Other " +
      "in one column and beside it in the next, drawing a band that appears " +
      "and disappears for no reason a reader can see. A period with nothing " +
      "recorded is dropped rather than drawn, because stacked to a hundred " +
      "per cent an empty column is whatever the rounding produced.",
    category: "Body",
    tags: ["chart", "bar", "stacked", "share", "section"],
    exports: ["MixOverTime"],
    dataFields: [
      {
        path: "mix.periods",
        type: "Array<{ label: string; amounts: Record<string, number> }>",
        summary:
          "One entry per period, oldest first — the order it is drawn in. " +
          "`amounts` is keyed by category, and the keys need not agree " +
          "between periods: a category absent from one counts as nothing in it.",
      },
      {
        path: "mix.unit",
        type: "string | undefined",
        summary: 'What is being counted: "visits", "tickets". Printed in the prose.',
      },
    ],
    files: node("mix-over-time"),
    previewData: {
      mix: {
        unit: "visits",
        periods: [
          { label: "Q1", amounts: { Swimming: 41, Gym: 38, Classes: 12, Squash: 9, Sauna: 4 } },
          { label: "Q2", amounts: { Swimming: 52, Gym: 34, Classes: 19, Squash: 7, Sauna: 6 } },
          {
            label: "Q3",
            amounts: { Swimming: 68, Gym: 29, Classes: 26, Squash: 5, Sauna: 7, Climbing: 3 },
          },
          {
            label: "Q4",
            amounts: { Swimming: 74, Gym: 24, Classes: 31, Squash: 4, Sauna: 5, Climbing: 6 },
          },
        ],
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "profile-compare",
    title: "Profile compare",
    summary: "One subject against a benchmark across several measures, as a radar.",
    detail:
      "A radar answers 'strong where', not 'how much' — and only if every " +
      "spoke shares a scale. That is the decision this makes for you. Handed " +
      "a hundred pounds, four incidents and nine per cent, a radar drawn " +
      "straight from the numbers puts the hundred at the rim and the four at " +
      "the centre, and draws a shape saying nothing except which measure uses " +
      "the largest units. So each measure is scored against its own best and " +
      "plotted out of 100, and the caption says so. The shape is the reading " +
      "and the sentence carries the figures, because a reader cannot compare " +
      "the areas of two rings and should not be asked to. Fewer than three " +
      "measures is not a radar — two spokes draw a line through the middle — " +
      "so below three it prints the comparison as prose and draws nothing.",
    category: "Body",
    tags: ["chart", "radar", "comparison", "benchmark", "section"],
    exports: ["ProfileCompare"],
    dataFields: [
      {
        path: "profile.subject",
        type: "string",
        summary: "What the first ring is called: a name, a site, a period.",
      },
      {
        path: "profile.benchmark",
        type: "string",
        summary: 'What the second is called: "Sector average", "Last year", "Target".',
      },
      {
        path: "profile.measures",
        type:
          "Array<{ label: string; value: number; against?: number | null; " +
          "best?: number; unit?: string }>",
        summary:
          "One entry per spoke, in the order they are read round the ring. " +
          "Three at least. `best` is full marks for that measure, and is what " +
          "lets measures in different units share one scale; absent, the " +
          "larger of the two readings is used. `against` absent draws a gap " +
          "rather than a benchmark of zero.",
      },
    ],
    files: node("profile-compare"),
    previewData: {
      profile: {
        subject: "Riverside",
        benchmark: "Sector average",
        measures: [
          { label: "Pool use", value: 82, against: 54, best: 100, unit: "%" },
          { label: "Gym use", value: 61, against: 70, best: 100, unit: "%" },
          { label: "Class fill", value: 45, against: 63, best: 100, unit: "%" },
          { label: "Retention", value: 88, against: 74, best: 100, unit: "%" },
          { label: "Peak capacity", value: 71, against: 66, best: 100, unit: "%" },
        ],
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "weighted-scatter",
    title: "Weighted scatter",
    summary: "Two measures against each other, with a third drawn as the point's size.",
    detail:
      "What a scatter becomes when 'which of these matters' is the real " +
      "question: plotting cost against usage says where things sit, and " +
      "sizing the points by how many people they affect says which is worth " +
      "the page. The size is an area and this keeps it one — a bubble drawn " +
      "with its radius proportional to the figure shows twice the reading as " +
      "four times the ink, which is the commonest way this chart lies, so " +
      "the weight is passed to Word untouched. A point with no weight is not " +
      "drawn, because a bubble of no size is a reading nobody took rather " +
      "than a reading of nothing. Labels stay in the prose: a bubble chart " +
      "labelled point by point is unreadable at a text column's width the " +
      "moment two bubbles touch, so the chart shows the field and the " +
      "sentence names the corner worth looking at.",
    category: "Body",
    tags: ["chart", "bubble", "scatter", "section"],
    exports: ["WeightedScatter"],
    dataFields: [
      {
        path: "scatter.xLabel",
        type: "string",
        summary: 'What the horizontal axis measures: "Cost per visit".',
      },
      {
        path: "scatter.yLabel",
        type: "string",
        summary: 'What the vertical axis measures: "Visits per week".',
      },
      {
        path: "scatter.weightLabel",
        type: "string",
        summary: 'What the size of a point means: "members affected". Used in the caption.',
      },
      {
        path: "scatter.points",
        type: "Array<{ label: string; x: number; y: number; weight: number }>",
        summary:
          "One entry per point, in any order. `label` is for the sentence and " +
          "is never drawn on the plot. `weight` is the point's area — do not " +
          "square it first.",
      },
      {
        path: "scatter.numberFormat",
        type: "string | undefined",
        summary: "How the axes print their figures, as an OOXML number format.",
      },
    ],
    files: node("weighted-scatter"),
    previewData: {
      scatter: {
        xLabel: "Cost per visit",
        yLabel: "Visits per week",
        weightLabel: "members affected",
        numberFormat: "#,##0",
        points: [
          { label: "Swimming", x: 3, y: 41, weight: 420 },
          { label: "Gym", x: 2, y: 38, weight: 610 },
          { label: "Classes", x: 6, y: 19, weight: 240 },
          { label: "Squash", x: 9, y: 7, weight: 85 },
          { label: "Climbing", x: 12, y: 5, weight: 40 },
        ],
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "next-steps",
    title: "Next steps",
    summary: "A generated paragraph with all four prompts already fenced off.",
    detail:
      "The node most likely to invent something, so all four prompts are set: " +
      "the info prompt hands the engine the facts it may use, and the " +
      "negative prompt closes off the two failures that matter in a document " +
      "somebody acts on — an invented deadline and an invented way to " +
      "contact you.",
    category: "Generated",
    tags: ["dynamic", "prompts", "ai"],
    exports: ["NextSteps"],
    dataFields: [
      {
        path: "actions",
        type: "string[]",
        summary: "What the reader has to do, in your words. Given to the engine as fact.",
      },
      {
        path: "contact",
        type: "string",
        summary: "How to reach you. Given to the engine rather than printed.",
      },
    ],
    files: node("next-steps"),
    previewData: {
      actions: [
        "check the balance shown above",
        "pay online or by bank transfer",
        "tell us if anything looks wrong",
      ],
      contact: "hello@riversideleisure.example on 0117 496 0000",
    },
    requires: [],
  },
  {
    kind: "component",
    id: "signature-block",
    title: "Signature block",
    summary: "A closing, a signature image, and who signed it.",
    detail:
      "The image is optional and the name is not, which is the right way " +
      "round: a letter signed by nobody is a letter nobody owns, whereas a " +
      "missing image is a rendering detail. Without one the block closes with " +
      "the typed name.",
    category: "Closing",
    tags: ["letter", "image", "closing"],
    exports: ["SignatureBlock"],
    dataFields: [
      {
        path: "signatory.name",
        type: "string",
        summary: "The person signing, as they should be printed.",
      },
      {
        path: "signatory.role",
        type: "string",
        summary: "Their role, printed beneath the name.",
      },
      {
        path: "signatory.signatureImage",
        type: "string | undefined",
        summary: "Path to a signature image, relative to the document project.",
      },
      {
        path: "signatory.closing",
        type: "string | undefined",
        summary: 'How the letter closes. Defaults to "Yours sincerely".',
      },
    ],
    files: node("signature-block"),
    previewData: {
      signatory: {
        name: "Tomas Lindqvist",
        role: "Centre Manager",
        signatureImage: "assets/signature-lindqvist.png",
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "terms-notice",
    title: "Terms notice",
    summary: "Numbered small print, one paragraph and one stable id per clause.",
    detail:
      "Boilerplate is the part of a document nobody rereads and everybody " +
      "copies, so it is worth one node that owns it. Clauses arrive as data, " +
      "each becomes its own paragraph, and each id is stable — which is what " +
      "lets next year's terms be diffed against this year's clause by clause.",
    category: "Legal",
    tags: ["legal", "boilerplate", "section"],
    exports: ["TermsNotice"],
    dataFields: [
      {
        path: "terms.heading",
        type: "string | undefined",
        summary: 'The heading above the clauses. Defaults to "Terms".',
      },
      {
        path: "terms.clauses",
        type: "string[]",
        summary: "One entry per clause, printed in order and numbered.",
      },
      {
        path: "terms.version",
        type: "string | undefined",
        summary: "Version or date, printed last so a reader can cite it.",
      },
    ],
    files: node("terms-notice"),
    previewData: {
      terms: {
        heading: "Terms of membership",
        clauses: [
          "Membership runs for twelve months from the renewal date shown above.",
          "Prices are held for the term and reviewed once a year.",
          "You may cancel within fourteen days of renewal for a full refund.",
        ],
        version: "2026.1",
      },
    },
    requires: [],
    themeHint: "legal-serif",
  },

  {
    kind: "component",
    id: "invoice-document-type",
    title: "Invoice document type",
    summary: "What an invoice may call itself, and the dates it has to carry.",
    detail:
      "The heading is a legal decision, not a label. A seller who is not " +
      "VAT-registered gets 'Invoice' and no VAT wording, because calling it a " +
      "VAT invoice is an offence in the UK. A registered seller gets 'VAT " +
      "invoice', or 'Simplified VAT invoice' under the local ceiling where they " +
      "allow it. A pro forma says it is not a VAT invoice, so nobody reclaims on " +
      "it, and a self-billed invoice carries its jurisdiction's statement. The " +
      "supply date prints only when it differs from the issue date, except in " +
      "Germany, which always wants a Leistungsdatum. All of it is one deriver, " +
      "not a dozen published arms.",
    category: "Opening",
    tags: ["invoice", "vat", "tax", "deriver", "branching", "section"],
    exports: ["InvoiceDocumentType"],
    dataFields: [
      { path: "invoice.number", type: "string", summary: "The invoice number. Required." },
      { path: "invoice.issueDate", type: "string", summary: "ISO 8601 date of issue. Required." },
      { path: "invoice.supplyDate", type: "string | undefined", summary: "ISO 8601 tax point. Missing means the issue date; printed only when it differs, always in Germany." },
      { path: "invoice.supplyPeriod", type: "{ from: string; to: string } | undefined", summary: "A supply over a period. Printed in place of the supply date." },
      { path: "invoice.kind", type: '"invoice" | "proforma" | "self-billed"', summary: "What the document is. Anything else fails the document." },
      { path: "invoice.jurisdiction", type: '"UK" | "EU" | "DE" | "FR" | "US"', summary: "Whose rules apply. US skips every VAT rule." },
      { path: "invoice.grossTotal", type: "number", summary: "VAT-inclusive total, local currency. Decides only whether it can be simplified." },
      { path: "invoice.purchaseOrder", type: "string | undefined", summary: "The customer's PO. Printed when given." },
      { path: "seller.vatRegistered", type: "boolean", summary: "Missing means not registered: no VAT number, no VAT wording." },
      { path: "seller.vatNumber", type: "string | undefined", summary: "Required when registered; the document fails without it." },
      { path: "seller.simplifiedAllowed", type: "boolean | undefined", summary: "Whether simplified invoices are issued at all. Missing means no." },
    ],
    arrangements: [
      {
        id: "stacked",
        title: "Stacked",
        summary: "One labelled line under another, down the text column.",
      },
      {
        id: "table",
        title: "Table",
        summary: "Labels in a narrow column with the values beside them, so the values line up.",
      },
      {
        id: "split",
        title: "Split",
        summary: "Number, purchase order and VAT number on the left; the dates ranged right.",
      },
    ],
    files: node("invoice-document-type", { deriver: true, arrangements: ["stacked", "table", "split"] }),
    previewData: {
      invoice: {
        number: "SI-2026-11-004817/REV-2",
        issueDate: "2026-11-30",
        supplyDate: "2026-09-28",
        kind: "invoice",
        jurisdiction: "UK",
        grossTotal: 1284750.0,
        purchaseOrder: "PO-4500098817-FACILITIES-NORTH",
      },
      seller: { vatRegistered: true, vatNumber: "GB 987 6543 21", simplifiedAllowed: true },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "tax-treatment-notice",
    title: "Tax treatment notice",
    summary: "The one mention an invoice's tax treatment requires by law, or a warning when it can't be given.",
    detail:
      "Prints the statutory mention for the invoice's tax treatment in that " +
      "treatment's own words, never '0% VAT', which reads as zero-rating, a " +
      "different legal treatment. Reverse charge, intra-Community supply and " +
      "sale for resale print the VAT numbers or certificate the exemption " +
      "depends on, and when one is missing the notice prints an 'incomplete' " +
      "warning instead of the mention: an invoice without the ID loses the " +
      "exemption. The UK construction reverse charge states the VAT that would " +
      "have been due without charging it. The French and German small-business " +
      "mentions are printed verbatim. A standard invoice gets nothing, so the " +
      "node can sit under the totals on every invoice. The UK amount is " +
      "formatted in sterling by a deriver that ships with it.",
    category: "Legal",
    tags: ["tax", "vat", "invoice", "legal", "branching", "deriver"],
    exports: ["TaxTreatmentNotice"],
    dataFields: [
      {
        path: "invoice.taxTreatment",
        type:
          '"standard" | "zero-rated-export" | "intra-community-supply" | "reverse-charge-services" | "uk-domestic-reverse-charge" | "small-business-exempt" | "us-resale-exempt" | "margin-scheme"',
        summary: "Which treatment the invoice is issued under. Standard prints nothing; an unrecognised value prints a warning.",
      },
      {
        path: "invoice.jurisdiction",
        type: "string | undefined",
        summary: "Upper-case country code, FR or DE, read for small-business-exempt only. Anything else prints a warning.",
      },
      {
        path: "invoice.exportEvidence",
        type: "string | undefined",
        summary: "The export evidence reference, such as the MRN, printed after a zero-rated export mention when given.",
      },
      {
        path: "invoice.notionalVat",
        type: "number | undefined",
        summary: "UK domestic reverse charge only: the VAT that would have been due, in sterling. Missing or zero prints a warning.",
      },
      {
        path: "seller.vatNumber",
        type: "string | undefined",
        summary: "Printed on a reverse-charge supply. Missing there prints a warning.",
      },
      {
        path: "customer.vatNumber",
        type: "string | undefined",
        summary: "Printed on an intra-Community or reverse-charge supply. Missing there prints a warning.",
      },
      {
        path: "customer.resaleCertificate",
        type: "{ number: string; state: string } | undefined",
        summary: "The resale certificate and issuing state for a US sale for resale. Missing prints a warning.",
      },
    ],
    files: [
      ...node("tax-treatment-notice", { deriver: true }),
    ],
    previewData: {
      invoice: {
        taxTreatment: "reverse-charge-services",
        jurisdiction: "DE",
        exportEvidence: "26DE470123456789A3",
        notionalVat: 184250.75,
      },
      seller: { vatNumber: "FR12345678901" },
      customer: {
        vatNumber: "NL123456789B01",
        resaleCertificate: { number: "SR EAA 102-345678", state: "California" },
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "invoice-billing-periods",
    title: "Invoice billing periods",
    summary: "The periods a subscription invoice bills for, in advance and in arrears.",
    detail:
      "Works out the periods an invoice states from the billing anchor and " +
      "interval. Month-end anchors are clamped without being forgotten, so an " +
      "anchor of the 31st bills on 31 Jan, 28 Feb and 31 Mar, and 29 Feb on an " +
      "annual plan falls on 28 Feb in a common year. End dates are inclusive, " +
      "the EN 16931 convention. Licensed items are billed in advance for the " +
      "coming period and metered items in arrears for the previous one, so " +
      "one invoice can state two periods; the header period then runs from " +
      "the earliest start to the latest end, so it contains every line " +
      "period, as Peppol requires. A first period that starts off the anchor " +
      "is labelled a part period. Dates are calendar dates in the seller's " +
      "billing time zone, never UTC. All of it is one deriver.",
    category: "Opening",
    tags: ["invoice", "subscription", "billing", "dates", "deriver", "section"],
    exports: ["InvoiceBillingPeriods"],
    dataFields: [
      { path: "billing.anchor", type: "string", summary: "The original anchor date, YYYY-MM-DD or an instant with an offset. Required; its day is kept for every period." },
      { path: "billing.interval", type: '"week" | "month" | "year"', summary: "The unit the cycle repeats in. Anything else fails the document." },
      { path: "billing.intervalCount", type: "number | undefined", summary: "Units per period; 3 months is quarterly. Missing means 1." },
      { path: "billing.invoiceDate", type: "string", summary: "When the invoice is issued, read in the billing time zone. Required." },
      { path: "billing.timeZone", type: "string", summary: "The seller's IANA billing time zone. Required." },
      { path: "billing.signupDate", type: "string | undefined", summary: "When the subscription started. Inside a billed period, it starts a part period." },
      { path: "billing.hasAdvanceItems", type: "boolean", summary: "Licensed items, billed in advance for the coming period. Missing means false." },
      { path: "billing.hasArrearsItems", type: "boolean", summary: "Metered items, billed in arrears for the previous period. Missing means false." },
    ],
    files: node("invoice-billing-periods", { deriver: true }),
    previewData: {
      billing: {
        anchor: "2025-12-31",
        interval: "month",
        intervalCount: 1,
        invoiceDate: "2027-01-01T02:30:00Z",
        timeZone: "America/Argentina/Buenos_Aires",
        signupDate: "2025-12-14",
        hasAdvanceItems: true,
        hasArrearsItems: true,
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "plan-change-proration",
    title: "Plan change proration",
    summary: "The credit and charge lines a mid-period plan change causes, and what the net means.",
    detail:
      "Every change made partway through a billing period becomes a credit for " +
      "the unused time on the old plan and a charge for the remaining time on " +
      "the new one, counted on the days the period actually has, never 30, with " +
      "an inclusive end date. Each line is rounded once in minor units, not built " +
      "from a rounded daily rate, so the pennies agree. The credit is what the " +
      "customer actually paid after discounts, not the list price, and a second " +
      "change in the same period credits the first change's plan rather than the " +
      "original one again. A change on the renewal date, a metered item or a line " +
      "worth nothing prints nothing. The net picks the sentence: a charge on this " +
      "invoice, a credit carried to the balance and stated as not refunded unless " +
      "the refund policy says otherwise, or a lateral change that only says so. " +
      "All of it is one deriver.",
    category: "Body",
    tags: ["money", "invoice", "subscription", "proration", "table", "deriver", "branching", "section"],
    exports: ["PlanChangeProration"],
    dataFields: [
      { path: "subscription.currency", type: "string", summary: "ISO 4217 code. Sets the minor units. Required." },
      { path: "subscription.period", type: "{ start: string; end: string }", summary: "ISO 8601 billing period; end is inclusive. Required." },
      {
        path: "subscription.changes",
        type:
          "Array<{ at: string; item?: string; metered?: boolean; from: { label: string; quantity?: number; paidMinor: number }; to: { label: string; quantity?: number; periodAmountMinor: number } }>",
        summary: "Every change in the period, amounts in integer minor units after discount. Missing or empty prints nothing.",
      },
      { path: "subscription.refundPolicy", type: '"balance" | "refund" | undefined', summary: "What a net credit becomes. Missing means balance: not refunded." },
    ],
    files: node("plan-change-proration", { deriver: true }),
    previewData: {
      subscription: {
        currency: "EUR",
        period: { start: "2026-12-01", end: "2026-12-31" },
        refundPolicy: "balance",
        changes: [
          {
            at: "2026-12-09", item: "platform",
            from: { label: "Business Analytics Suite (Annual, billed monthly)", quantity: 1, paidMinor: 1249900 },
            to: { label: "Enterprise Analytics Suite with Extended Data Retention", quantity: 1, periodAmountMinor: 2499900 },
          },
          {
            at: "2026-12-17", item: "seats",
            from: { label: "Named analyst seats", quantity: 1180, paidMinor: 5428000 },
            to: { label: "Named analyst seats", quantity: 1450, periodAmountMinor: 6670000 },
          },
          {
            at: "2026-12-23", item: "platform",
            from: { label: "Enterprise Analytics Suite with Extended Data Retention", quantity: 1, paidMinor: 2499900 },
            to: { label: "Enterprise Analytics Suite with Extended Data Retention and Dedicated Support", quantity: 1, periodAmountMinor: 3149900 },
          },
        ],
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "commodity-code-summary",
    title: "Commodity code summary",
    summary: "Tariff code, origin and mass on every line, and the summary a customs declaration is split on.",
    detail:
      "Prints each goods line's tariff code, origin, preference claim, quantity, " +
      "supplementary unit, net mass and value, then one summary row per tariff " +
      "code, origin and preference, which is how customs splits a declaration. " +
      "Codes are text, so 0901 stays 0901, and each is checked against the " +
      "destination's length: 6 for HS, 8 for EU CN, 10 for UK and US. A short " +
      "or missing code warns on its line. Masses in lb are converted, summed " +
      "exactly and rounded once per group, and the total is the sum of the " +
      "groups, so the summary always adds up. Net mass above the declared gross " +
      "mass prints a warning. All of it is one deriver.",
    category: "Body",
    tags: ["customs", "trade", "tariff", "invoice", "deriver", "table", "section"],
    exports: ["CommodityCodeSummary"],
    dataFields: [
      { path: "customs.currency", type: "string", summary: "ISO 4217 code the values are in. Required." },
      { path: "customs.destinationCodeLength", type: "6 | 8 | 10", summary: "Digits the destination's tariff needs: HS 6, EU CN 8, UK and US 10. Required." },
      { path: "customs.grossMassKg", type: "number | undefined", summary: "Declared gross mass in kg. Missing means no gross line and no net-over-gross check." },
      { path: "customs.lines", type: 'Array<{ ref?; description?; tariffCode: string; origin: string; preference?: boolean; quantity: number; supplementaryUnit?: { code: string; quantity: number }; netMass: { value: number; unit: "kg" | "g" | "lb" }; valueMinor: number }>', summary: "The goods lines, in invoice order. Empty prints one sentence and no tables." },
    ],
    arrangements: [
      { id: "lines-then-summary", title: "Lines, then summary", summary: "Every goods line, with the summary by tariff code and origin under it." },
      { id: "summary-first", title: "Summary first", summary: "The summary and its warnings first, the lines after, for a long invoice." },
    ],
    files: node("commodity-code-summary", { deriver: true, arrangements: ["lines-then-summary", "summary-first"] }),
    // The tariff codes are illustrative, not checked against a live tariff.
    previewData: {
      customs: {
        currency: "GBP",
        destinationCodeLength: 10,
        grossMassKg: 27850.5,
        lines: [
          { ref: "INV-2026-004817/0001", description: "Roasted arabica coffee beans, not decaffeinated, in 1 kg one-way valve bags", tariffCode: "0901210000", origin: "CO", preference: true, quantity: 12000, netMass: { value: 12000, unit: "kg" }, valueMinor: 15600000 },
          { ref: "INV-2026-004817/0002", description: "Roasted arabica coffee beans, not decaffeinated, in 1 kg one-way valve bags", tariffCode: "0901210000", origin: "BR", quantity: 6000, netMass: { value: 6000, unit: "kg" }, valueMinor: 7245000 },
          { ref: "INV-2026-004817/0003", description: "Men's leather lace-up boots with vulcanised rubber soles, assorted sizes UK 6 to 13", tariffCode: "6403919800", origin: "PT", preference: true, quantity: 2400, supplementaryUnit: { code: "pa", quantity: 2400 }, netMass: { value: 1420.5, unit: "kg" }, valueMinor: 10080000 },
          { ref: "INV-2026-004817/0004", description: "Men's leather lace-up boots with vulcanised rubber soles, assorted sizes UK 6 to 13", tariffCode: "6403919800", origin: "PT", preference: true, quantity: 1500, supplementaryUnit: { code: "pa", quantity: 1500 }, netMass: { value: 1850.75, unit: "lb" }, valueMinor: 6300000 },
          { ref: "INV-2026-004817/0005", description: "Red wine of fresh grapes with a protected designation of origin, in 75 cl glass bottles", tariffCode: "2204214200", origin: "FR", preference: true, quantity: 4800, supplementaryUnit: { code: "l", quantity: 3600 }, netMass: { value: 5280, unit: "kg" }, valueMinor: 2052500 },
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "preferential-origin-statement",
    title: "Preferential origin statement",
    summary: "The statement on origin a trade agreement prescribes, or a warning when its exporter reference is missing.",
    detail:
      "Prints the statement on origin each agreement prescribes, in its own words: " +
      "the EU-UK TCA statement on origin, the EU GSP REX statement, the CETA origin " +
      "declaration, the EU-Japan statement and the UK DCTS statement, with the " +
      "exporter reference, the origin, the origin criteria that agreement uses and " +
      "the multiple-shipment period. The EUR 6,000 threshold is tested against the " +
      "originating lines only, converted at the rate you supply, never against the " +
      "invoice total or a fixed sterling figure. Above it the exporter's REX or EORI " +
      "number is required, and when it is missing a warning prints in place of the " +
      "statement, because an importer who claims preference on an incomplete " +
      "statement faces a duty demand later. Non-originating lines are listed by " +
      "reference, so \"except where otherwise clearly indicated\" is true. A period " +
      "longer than 12 months is refused the same way. Nothing prints when no line is " +
      "originating. Worked out by one deriver.",
    category: "Legal",
    tags: ["origin", "customs", "export", "invoice", "legal", "branching", "deriver"],
    exports: ["PreferentialOriginStatement"],
    dataFields: [
      { path: "origin.agreement", type: '"eu-uk-tca" | "eu-gsp" | "ceta" | "eu-japan" | "uk-dcts"', summary: "Which agreement's prescribed wording, reference and criteria apply. Required." },
      { path: "origin.originCountry", type: "string", summary: "UK or EU under the TCA, EU under CETA and EU-Japan, the beneficiary country under GSP and DCTS. Required." },
      { path: "origin.exporterReference", type: "string | undefined", summary: "EORI for UK exporters, REX otherwise. Required when originating goods exceed EUR 6,000, or the statement is withheld." },
      { path: "origin.eurRate", type: "number | undefined", summary: "Euros per unit of the consignment currency, for the EUR 6,000 test. Required unless the currency is EUR." },
      { path: "origin.period", type: "{ from: string; to: string } | undefined", summary: "ISO dates for a multiple-shipment statement. Over 12 months withholds it." },
      { path: "origin.place", type: "string", summary: "Where the statement is made. Required." },
      { path: "origin.date", type: "string", summary: "ISO 8601 date of the statement. Required." },
      { path: "origin.exporterName", type: "string", summary: "The exporter, closing the statement. Required." },
      { path: "consignment.currency", type: "string", summary: "ISO 4217 code of the line values. Required." },
      {
        path: "consignment.lines",
        type: "Array<{ ref: string; valueMinor: number; originating?: boolean; criterion?: string; hsHeading?: string }>",
        summary: "Every line in integer minor units. Only lines marked originating count toward the threshold; the rest are listed as exceptions.",
      },
    ],
    files: node("preferential-origin-statement", { deriver: true }),
    previewData: {
      origin: {
        agreement: "eu-uk-tca",
        originCountry: "UK",
        exporterReference: "GB123456789000",
        eurRate: 1.1712,
        period: { from: "2026-12-01", to: "2027-11-30" },
        place: "Newcastle upon Tyne, Tyne and Wear",
        date: "2026-11-30",
        exporterName: "Northumbrian Precision Engineering and Fabrication Holdings Limited",
      },
      consignment: {
        currency: "GBP",
        lines: [
          { ref: "10", valueMinor: 128450000, originating: true, criterion: "A" },
          { ref: "20", valueMinor: 96320075, originating: true, criterion: "A" },
          { ref: "30", valueMinor: 18999900, originating: false },
          { ref: "40", valueMinor: 254000000, originating: true, criterion: "C, D" },
          { ref: "50", valueMinor: 7250050, originating: true, criterion: "E" },
          { ref: "60", valueMinor: 4410000, originating: false },
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "dunning-reminder",
    title: "Dunning reminder",
    summary: "A payment reminder that escalates one level at a time and never threatens what you won't do.",
    detail:
      "Chases only invoices past their grace period that are not disputed, on " +
      "hold or paid down to zero, and prints nothing at all when none qualify. " +
      "Each invoice moves up one level per reminder and only after the minimum " +
      "interval, so a daily run never jumps from first reminder to final notice " +
      "or re-sends the same one. The level sets the title and tone. The deadline " +
      "is never shorter than the Dutch, Belgian or UK pre-action minimum you name " +
      "and never falls on a weekend. A final reminder states only the " +
      "consequences you mark as intended, and none if you mark none, because an " +
      "empty threat is unlawful in the UK and the US. Every level says which " +
      "date payments were checked up to. The opening can be generated per " +
      "document. One deriver decides it all.",
    category: "Body",
    tags: ["reminder", "dunning", "invoice", "deriver", "branching", "section", "dynamic", "prompts", "ai"],
    exports: ["DunningReminder"],
    dataFields: [
      { path: "reminder.sendDate", type: "string", summary: "ISO 8601 date the reminder is sent. Required." },
      { path: "reminder.paymentsCheckedTo", type: "string", summary: "ISO 8601 date payments are reflected up to. Printed at every level. Required." },
      { path: "reminder.currency", type: "string", summary: "ISO 4217 code for every amount. Required." },
      { path: "reminder.graceDays", type: "number | undefined", summary: "Days past due before an invoice is chased. Missing means 7." },
      { path: "reminder.minIntervalDays", type: "number | undefined", summary: "Days between reminders for one invoice. Missing means 14." },
      { path: "reminder.deadlineDays", type: "{ first?: number; second?: number; final?: number } | undefined", summary: "Days to the deadline per level. Missing means 14, 10 and 7." },
      { path: "reminder.legalFloor", type: '"NL-consumer" | "BE-consumer" | "UK-pre-action" | undefined', summary: "A statutory minimum the deadline is raised to. Missing means none." },
      { path: "reminder.channel", type: '"post" | "email" | undefined', summary: "Delivery, read for the NL and BE floors only. Missing means post." },
      { path: "reminder.intendedConsequences", type: '("agency" | "court" | "suspension")[] | undefined', summary: "What you will actually do after a final reminder. Missing prints no threat." },
      { path: "reminder.customerSince", type: "string | undefined", summary: "ISO 8601 start of the relationship, told to the generated opening only." },
      { path: "reminder.invoices", type: "Array<{ number; dueDate; openAmount; disputed?; onHold?; lastLevel?; lastReminderDate? }> | undefined", summary: "Every open invoice. The component picks which this reminder covers." },
    ],
    arrangements: [
      { id: "table", title: "Table", summary: "Number, due date, days overdue and amount in columns, with a total row." },
      { id: "stacked", title: "Stacked", summary: "One sentence per invoice down the text column, then the total." },
    ],
    files: node("dunning-reminder", { deriver: true, arrangements: ["table", "stacked"] }),
    previewData: {
      reminder: {
        sendDate: "2026-11-20",
        paymentsCheckedTo: "2026-11-19",
        currency: "GBP",
        graceDays: 7,
        minIntervalDays: 14,
        legalFloor: "UK-pre-action",
        channel: "post",
        intendedConsequences: ["agency", "court"],
        customerSince: "2019-03-01",
        invoices: [
          { number: "SI-2026-07-004817/REV-2", dueDate: "2026-07-31", openAmount: 184250.75, lastLevel: 2, lastReminderDate: "2026-11-02" },
          { number: "SI-2026-08-005102-FACILITIES-NORTH", dueDate: "2026-08-31", openAmount: 96480.0, lastLevel: 2, lastReminderDate: "2026-11-02" },
          { number: "SI-2026-09-005390", dueDate: "2026-09-30", openAmount: 12999.99, lastLevel: 1, lastReminderDate: "2026-10-30" },
          { number: "SI-2026-09-005391", dueDate: "2026-09-30", openAmount: 7450.0, lastLevel: 1, lastReminderDate: "2026-11-12" },
          { number: "SI-2026-10-005688", dueDate: "2026-10-31", openAmount: 48120.4, disputed: true },
          { number: "SI-2026-10-005702", dueDate: "2026-10-31", openAmount: 23875.0 },
          { number: "SI-2026-10-005703", dueDate: "2026-10-31", openAmount: 1250.0, onHold: true },
          { number: "SI-2026-10-005750", dueDate: "2026-10-20", openAmount: 0 },
          { number: "SI-2026-11-005940", dueDate: "2026-11-10", openAmount: 3600.0 },
          { number: "SI-2026-11-006011", dueDate: "2026-11-15", openAmount: 9999.0 },
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "open-items-ageing",
    title: "Open items and ageing",
    summary: "Every invoice still open, aged by days past due, with unallocated credit kept out of the buckets.",
    detail:
      "Lists every invoice not yet fully settled with the amount still open on it, then ages the " +
      "open balance into not yet due, 1-30, 31-60, 61-90 and over 90 days overdue. Age is counted " +
      "from the due date, not the invoice date, and each bucket includes its top edge, so 30 days is " +
      "1-30 and an invoice due today is not yet due; one with no due date ages from its date plus the " +
      "default terms, and the strip says so. A payment or credit note not set against an invoice is " +
      "unallocated credit on a line of its own, never netted into a bucket, so no bucket goes negative " +
      "and an old debt never disappears. A disputed invoice is listed, marked and aged in its own " +
      "column, so it cannot drive the over-90 figure. The shares are largest-remainder rounded to sum " +
      "to exactly 100, an over-allocated invoice fails the document, and when nothing is overdue the " +
      "strip is one line. All of it is one deriver.",
    category: "Body",
    tags: ["money", "statement", "ageing", "table", "deriver", "branching", "section"],
    exports: ["OpenItemsAgeing"],
    dataFields: [
      { path: "account.currency", type: "string", summary: "ISO 4217 code. Sets the minor unit. Required." },
      { path: "account.statementDate", type: "string", summary: "ISO 8601 date every age is counted to. Required." },
      { path: "account.defaultTermsDays", type: "number | undefined", summary: "Days to due for an invoice with no due date. Required only when one is open without it." },
      { path: "account.invoices", type: "Array<{ number: string; date: string; dueDate?: string; amount: number; disputed?: boolean }>", summary: "Every invoice, settled ones an allocation names included. Amounts in major units." },
      { path: "account.credits", type: "Array<{ type: \"payment\" | \"credit-note\"; date: string; number: string; amount: number; allocations?: Array<{ invoice: string; amount: number }> }>", summary: "Payments and credit notes. What their allocations do not cover is unallocated credit." },
    ],
    arrangements: [
      { id: "across", title: "Across", summary: "The buckets side by side: name, amount and share in three rows." },
      { id: "down", title: "Down", summary: "One row per bucket with its amount and share, for narrow columns and large figures." },
    ],
    files: node("open-items-ageing", { deriver: true, arrangements: ["across", "down"] }),
    previewData: {
      account: {
        currency: "GBP", statementDate: "2026-11-30", defaultTermsDays: 30,
        invoices: [
          { number: "INV-2026-0004211", date: "2026-06-12", dueDate: "2026-07-12", amount: 248750, disputed: true },
          { number: "INV-2026-0004290", date: "2026-07-01", dueDate: "2026-07-31", amount: 54000 },
          { number: "INV-2026-0004388", date: "2026-07-20", dueDate: "2026-08-19", amount: 1186420.5 },
          { number: "INV-2026-0004502", date: "2026-08-28", dueDate: "2026-09-27", amount: 312980.75 },
          { number: "INV-2026-0004617", date: "2026-09-30", amount: 97640 },
          { number: "INV-2026-0004733", date: "2026-10-15", dueDate: "2026-11-14", amount: 1048215.4 },
          { number: "INV-2026-0004851", date: "2026-11-02", dueDate: "2026-12-02", amount: 655300 },
        ],
        credits: [
          { type: "credit-note", date: "2026-10-08", number: "CN-2026-000412", amount: 54000, allocations: [{ invoice: "INV-2026-0004290", amount: 54000 }] },
          { type: "payment", date: "2026-11-03", number: "BACS-REMIT-20261103-000918", amount: 400000, allocations: [{ invoice: "INV-2026-0004388", amount: 400000 }] },
          { type: "credit-note", date: "2026-11-18", number: "CN-2026-000431", amount: 18500, allocations: [{ invoice: "INV-2026-0004502", amount: 6250 }] },
          { type: "payment", date: "2026-11-26", number: "BACS-REMIT-20261126-001147", amount: 125000 },
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "statement-of-account",
    title: "Statement of account",
    summary: "Every transaction in a period with a running balance, and what the closing balance means.",
    detail:
      "A balance-forward statement: the opening balance as you state it, never " +
      "rebuilt from older rows, then every invoice, credit note, payment, refund " +
      "and adjustment dated inside the period with a running balance. Invoices and " +
      "refunds are debits, payments and credit notes are credits, so a credit note " +
      "is never a negative invoice and both columns add up. On the same day an " +
      "invoice prints before the payment that settles it, so the balance never " +
      "dips into credit first. The running balance is kept unrounded and each " +
      "figure rounded half away from zero to the currency's minor unit, so the " +
      "last line equals the closing balance; a closing balance you supply is " +
      "checked against it and a mismatch fails the build. The closing sentence " +
      "says the amount due, that the account is clear, or that a credit balance " +
      "is in the customer's favour. Mixed currencies are refused, and an empty " +
      "period prints the opening balance and says there were no transactions. " +
      "All of it is one deriver.",
    category: "Body",
    tags: ["money", "statement", "account", "ledger", "table", "deriver", "branching", "section"],
    exports: ["StatementOfAccount"],
    dataFields: [
      { path: "statement.currency", type: "string", summary: "ISO 4217 code. Sets the decimals. Required." },
      { path: "statement.periodStart", type: "string", summary: "ISO 8601 first day covered, inclusive. Required." },
      { path: "statement.statementDate", type: "string", summary: "ISO 8601 last day covered, inclusive. Required." },
      { path: "statement.openingBalance", type: "number", summary: "Balance brought forward in major units; negative is in credit. Required." },
      { path: "statement.closingBalance", type: "number | undefined", summary: "Your ledger's closing balance, checked to half a minor unit. Missing means no check." },
      { path: "statement.dueBy", type: "string | undefined", summary: "ISO 8601 due date for an amount owed. Ignored unless owing." },
      {
        path: "statement.entries",
        type: 'Array<{ type: "invoice" | "credit-note" | "payment" | "refund" | "adjustment"; date: string; number: string; amount: number; currency?: string; reference?: string; description?: string }>',
        summary: "Every document, any order; only those inside the period print. Amounts positive except a credit adjustment. Missing or empty prints no transactions.",
      },
    ],
    files: node("statement-of-account", { deriver: true }),
    previewData: {
      statement: {
        currency: "GBP",
        periodStart: "2026-10-01",
        statementDate: "2026-10-31",
        openingBalance: 18452.75,
        closingBalance: 116828.33,
        dueBy: "2026-11-30",
        entries: [
          // Payment listed before the invoice it settles, on the same day, so the preview shows the ordering rule.
          { type: "payment", date: "2026-10-03", number: "PAY-2026-0098211", reference: "BACS 20261003-HSBC-77120", amount: 18452.75, description: "Payment received, thank you" },
          { type: "invoice", date: "2026-10-03", number: "INV-2026-104518", reference: "PO-4500098123-CC-EMEA-FIN", amount: 148926.40, description: "Managed infrastructure and 24/7 platform support, October 2026" },
          { type: "credit-note", date: "2026-10-14", number: "CN-2026-002217", reference: "INV-2026-104518", amount: 12450.00, description: "Service credit for the 6 October outage under the availability commitment" },
          { type: "adjustment", date: "2026-10-20", number: "ADJ-2026-000341", amount: 1286.33, description: "Late payment interest under the Late Payment of Commercial Debts (Interest) Act 1998" },
          { type: "payment", date: "2026-10-24", number: "PAY-2026-0101876", reference: "CHAPS 20261024-BARC-00481", amount: 120000.00 },
          { type: "refund", date: "2026-10-28", number: "RF-2026-000057", reference: "Overpayment on INV-2026-099874", amount: 2315.60 },
          { type: "invoice", date: "2026-10-28", number: "INV-2026-106002", amount: 96750.00, description: "Professional services: data migration and cutover, phase 2 of 3" },
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "bank-transfer-details",
    title: "Bank transfer details",
    summary: "The one way to pay that fits this payer, checked, with a reference their bank won't cut short.",
    detail:
      "Shows one set of bank details, picked from the payer's country and the " +
      "invoice currency: sort code and account for a UK payer in sterling, IBAN " +
      "and BIC for SEPA in euros, ABA routing and account for a US payer in " +
      "dollars, and IBAN and BIC with a SWIFT note for anyone else. Every rail at " +
      "once confuses payers. Each detail is checked (IBAN mod-97, the ABA 3-7-1 " +
      "checksum, sort code shape), and a failed one is withheld rather than " +
      "printed. The reference is fitted to the rail's limit and character set: 18 " +
      "for Bacs and Faster Payments, 80 for ACH, 140 for SEPA and SWIFT. A longer " +
      "one falls back to the customer reference rather than being cut short by " +
      "the payer's bank. The account name prints exactly as given, for Confirmation " +
      "of Payee. When the details have changed, a fixed warning to verify by phone " +
      "comes first.",
    category: "Closing",
    tags: ["invoice", "payment", "bank", "fraud", "deriver", "branching", "section"],
    exports: ["BankTransferDetails"],
    dataFields: [
      { path: "payee.accountName", type: "string", summary: "Exactly as the bank holds it, for Confirmation of Payee. Required." },
      { path: "payee.uk", type: "{ sortCode: string; accountNumber: string } | undefined", summary: "UK domestic details. Used for a UK payer in GBP; withheld when either part is invalid." },
      { path: "payee.iban", type: "string | undefined", summary: "Checked with mod-97, printed in groups of four. For SEPA and international payers." },
      { path: "payee.bic", type: "string | undefined", summary: "8 or 11 characters. Optional for SEPA, required for international." },
      { path: "payee.us", type: "{ achRouting: string; wireRouting?: string; accountNumber: string } | undefined", summary: "US details for a US payer in USD. Routing numbers are checked with the ABA checksum." },
      { path: "payee.detailsChangedOn", type: "string | undefined", summary: "ISO date the details changed. When given, a verify-by-phone warning prints first." },
      { path: "payer.country", type: "string | undefined", summary: "ISO 3166-1 alpha-2. With the currency, picks the rail. Missing means international." },
      { path: "invoice.currency", type: "string", summary: "ISO 4217 code. Required." },
      { path: "invoice.reference", type: "string", summary: "The reference to quote. Required; printed when it fits the rail." },
      { path: "invoice.customerReference", type: "string | undefined", summary: "A shorter reference printed when the invoice reference won't fit. Neither fitting fails the document." },
    ],
    arrangements: [
      { id: "stacked", title: "Stacked", summary: "One labelled line under another, down the text column." },
      { id: "table", title: "Table", summary: "Labels in a narrow column with the values beside them, so each detail is copied from one place." },
      { id: "split", title: "Split", summary: "Where to send the money on the left; the reference to quote ranged right." },
    ],
    files: node("bank-transfer-details", { deriver: true, arrangements: ["stacked", "table", "split"] }),
    previewData: {
      payee: {
        accountName: "Northern Facilities Management Services (Holdings) Ltd",
        uk: { sortCode: "60-16-13", accountNumber: "31926819" },
        iban: "GB29 NWBK 6016 1331 9268 19",
        bic: "NWBKGB2L",
        us: { achRouting: "121000358", wireRouting: "026009593", accountNumber: "000123456789" },
        detailsChangedOn: "2026-11-03",
      },
      payer: { country: "GB" },
      invoice: { currency: "GBP", reference: "INV-2026-11-004817/REV-2", customerReference: "NFMS-004817" },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "direct-debit-collection-notice",
    title: "Direct debit collection notice",
    summary: "The advance notice a direct debit needs, dated no earlier than its scheme allows.",
    detail:
      "Replaces the how-to-pay block on an invoice collected by direct debit, and " +
      "words the advance notice for SEPA Core, UK Bacs or US ACH. The notice " +
      "period is counted the way each scheme counts it: 14 calendar days under " +
      "SEPA unless a shorter one was agreed, 10 working days under Bacs (never " +
      "fewer than 3) skipping the bank holidays you supply, and 10 days under ACH " +
      "only when the amount differs from the last debit. A notice issued too late " +
      "says 'on or after' the earliest lawful date instead of the date asked for, " +
      "and the date moves to the next day the scheme settles on, with TARGET2's " +
      "Easter closing days worked out for the year. The SEPA creditor identifier's " +
      "check digits are verified without the business code, and a bad one stops " +
      "the document. The IBAN is shown by its last four characters only. SEPA " +
      "states the 8-week refund right. Bacs prints the Direct Debit Guarantee only " +
      "from the approved text you supply, and warns when it is missing. A zero or " +
      "negative amount says nothing will be collected. All of it is one deriver.",
    category: "Legal",
    tags: ["direct-debit", "sepa", "bacs", "ach", "payment", "invoice", "legal", "deriver", "branching", "section"],
    exports: ["DirectDebitCollectionNotice"],
    dataFields: [
      { path: "debit.scheme", type: '"sepa-core" | "bacs" | "ach"', summary: "Which scheme collects. Sets the currency, lead time and wording. Required." },
      { path: "debit.creditorName", type: "string", summary: "Who collects, as the payer's statement shows it. Required." },
      { path: "debit.creditorId", type: "string | undefined", summary: "SEPA creditor identifier. Required for SEPA; a failed check digit fails the document." },
      { path: "debit.sun", type: "string | undefined", summary: "Bacs six-digit Service User Number. Required for Bacs." },
      { path: "debit.mandateReference", type: "string", summary: "Mandate or Direct Debit reference. At most 35 characters (SEPA) or 18 (Bacs). Required." },
      { path: "debit.payerIban", type: "string | undefined", summary: "Payer's IBAN. Only the last four characters are printed." },
      { path: "debit.payerAccountLast4", type: "string | undefined", summary: "Last four of the payer's account, when no IBAN is given. Neither omits the line." },
      { path: "debit.requestedDate", type: "string", summary: "ISO 8601 date asked for. Moved later when the notice is short or the day is closed. Required." },
      { path: "debit.amount", type: "number", summary: "Amount in the scheme's currency. Zero or negative collects nothing. Required." },
      { path: "debit.previousAmount", type: "number | undefined", summary: "ACH only: the last debit. A different amount triggers the 10-day notice." },
      { path: "debit.first", type: "boolean | undefined", summary: "First collection under the mandate. Missing means recurring." },
      { path: "debit.agreedNoticeDays", type: "number | undefined", summary: "Agreed notice period: calendar days (SEPA), working days of at least 3 (Bacs). Missing means the default." },
      { path: "debit.bankHolidays", type: "string[] | undefined", summary: "ISO dates the payer's banks are closed. Required for Bacs; optional for ACH." },
      { path: "debit.guaranteeText", type: "string[] | undefined", summary: "Bacs only: the approved Direct Debit Guarantee, one entry per paragraph. Missing prints a warning." },
      { path: "document.issuedAt", type: "string", summary: "ISO 8601 date the notice is issued. The notice period counts from here. Required." },
    ],
    arrangements: [
      { id: "stacked", title: "Stacked", summary: "One labelled mandate detail under another, down the text column." },
      { id: "table", title: "Table", summary: "Labels in a narrow column with the mandate details beside them." },
    ],
    files: node("direct-debit-collection-notice", { deriver: true, arrangements: ["stacked", "table"] }),
    previewData: {
      debit: {
        scheme: "sepa-core",
        creditorName: "Northwind Facilities & Estates Management Services GmbH",
        creditorId: "IT58ZZZ0000012345678901",
        mandateReference: "MNDT-2026-NORTHFACILITIES-000048172",
        payerIban: "IT60 X054 2811 1010 0000 0123 456",
        requestedDate: "2027-03-26",
        amount: 1284750.0,
        first: true,
      },
      document: { issuedAt: "2027-03-08" },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "structured-payment-reference",
    title: "Structured payment reference",
    summary: "A bank payment reference with its check digits worked out, or the invoice number when none can be issued.",
    detail:
      "Generates or validates the structured reference a bank matches a payment " +
      "against: ISO 11649 RF, Norwegian KID (MOD10 or MOD11), the Finnish " +
      "viitenumero and its RF form, the Belgian OGM/VCS, the Swiss QR reference " +
      "and the Swedish OCR number with or without its length digit. Each is " +
      "printed in its own grouping. Bases are strings, because a number drops the " +
      "leading zeros OGM and QRR count and a 26-digit QR reference is past what a " +
      "JS number holds exactly. A base too long for its scheme is refused, never " +
      "truncated, because a truncated reference still validates and books the " +
      "payment against the wrong invoice. A KID whose MOD11 check is '-' is " +
      "treated as unissuable, since many banks reject it. Every failure falls " +
      "back to asking the payer to quote the invoice number. All of it is one deriver.",
    category: "Body",
    tags: ["payment", "invoice", "reference", "bank", "deriver", "branching"],
    exports: ["StructuredPaymentReference"],
    dataFields: [
      { path: "payment.scheme", type: '"rf" | "no-kid-mod10" | "no-kid-mod11" | "fi" | "fi-rf" | "be-ogm" | "ch-qrr" | "se-ocr" | "se-ocr-length"', summary: "Which scheme the reference follows. Required when payment is given; anything else fails the document." },
      { path: "payment.base", type: "string | undefined", summary: "What the check digits are generated from, as a string. Invalid or too long prints the invoice-number fallback, never a truncated reference." },
      { path: "payment.reference", type: "string | undefined", summary: "A reference the ERP already issued. Validated, not recalculated; takes precedence over base." },
      { path: "invoice.number", type: "string", summary: "Printed in the fallback line. Required." },
    ],
    files: node("structured-payment-reference", { deriver: true }),
    previewData: {
      payment: { scheme: "ch-qrr", base: "21000000000313947143000901" },
      invoice: { number: "SI-2026-11-004817/REV-2" },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "gst-inclusive-statement",
    title: "GST-inclusive statement (AU/NZ)",
    summary: "How GST shows on an Australian or New Zealand tax invoice, and who it must name at each threshold.",
    detail:
      "Prints the items and totals of an AU or NZ invoice and decides how GST appears. \"Total price includes GST.\" " +
      "replaces the GST figure only when every item is standard-rated, prices include GST and the GST is exactly 1/11 " +
      "(AU) or 3/23 (NZ) of the total. Otherwise the figure is shown, and on a mixed invoice each taxable item is marked " +
      "and a legend says so. GST is worked in cents under the ATO total-invoice rule, or the taxable-sale rule when " +
      "chosen. Thresholds use the GST-inclusive total: the supplier's ABN always and the buyer's name or ABN from " +
      "A$1,000 in Australia; in New Zealand, the GST number above NZ$200 and the buyer's name plus one identifier above " +
      "NZ$1,000. ABNs, IRD numbers and NZBNs are checksummed, and a failure or a missing identity fails the build. An " +
      "unregistered seller gets no GST wording. All of it is one deriver.",
    category: "Body",
    tags: ["invoice", "tax", "gst", "australia", "new-zealand", "deriver"],
    exports: ["GstInclusiveStatement"],
    dataFields: [
      { path: "gst.country", type: '"AU" | "NZ"', summary: "Whose GST rules apply. Required." },
      { path: "gst.sellerRegistered", type: "boolean", summary: "Missing means not registered: no GST wording at all." },
      { path: "gst.sellerTaxId", type: "string | undefined", summary: "ABN (AU) or GST number (NZ). Checksummed; required on an AU tax invoice and on NZ invoices above NZ$200." },
      { path: "gst.pricesIncludeGst", type: "boolean | undefined", summary: "Whether line amounts include GST. Required when any line is taxable." },
      { path: "gst.rounding", type: '"total-invoice" | "taxable-sale" | undefined', summary: "ATO rounding rule. Defaults to total-invoice." },
      { path: "gst.lines", type: 'Array<{ description: string; amountMinor: number; treatment: "taxable" | "gst-free" | "input-taxed" | "margin-scheme" | "zero-rated" | "exempt" }>', summary: "The items, amounts in integer cents. At least one." },
      { path: "gst.buyer", type: '{ name?: string; abn?: string; identifier?: { kind: "address" | "phone" | "email" | "nzbn" | "url" | "trading-name"; value: string } } | undefined', summary: "Required from A$1,000 (name or ABN) and above NZ$1,000 (name and one identifier)." },
    ],
    arrangements: [
      { id: "table", title: "Table", summary: "Items and totals in two columns, amounts right-aligned." },
      { id: "stacked", title: "Stacked", summary: "One line per item and per total, for an invoice laid out as text." },
    ],
    files: node("gst-inclusive-statement", { deriver: true, arrangements: ["table", "stacked"] }),
    previewData: {
      gst: {
        country: "AU",
        sellerRegistered: true,
        sellerTaxId: "51 824 753 556",
        pricesIncludeGst: true,
        rounding: "total-invoice",
        lines: [
          { description: "Commercial kitchen fit-out, Level 3, 1 Macquarie Place, Sydney: stage 2 of 3, labour and materials", amountMinor: 12950000, treatment: "taxable" },
          { description: "Fresh produce supplied to the staff canteen under the September 2026 catering agreement", amountMinor: 880000, treatment: "gst-free" },
          { description: "Residential lease, Unit 14/220 Sussex Street, Sydney: October 2026 rent", amountMinor: 429000, treatment: "input-taxed" },
          { description: "Early-payment rebate agreed 14 August 2026 on the fit-out contract", amountMinor: -259000, treatment: "taxable" },
        ],
        buyer: {
          name: "Harbourside Hospitality Group Holdings Pty Ltd as trustee for the Harbourside Unit Trust",
          abn: "98 765 432 122",
        },
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "exemption-certificate-check",
    title: "Exemption certificate check (US)",
    summary: "Which lines of a US sale may go out exempt, under which certificate, and why the rest stay taxable.",
    detail:
      "A certificate is tested line by line, never as a yes or no for the whole customer. It has to cover the " +
      "ship-to state, not the billing state, and be a form that state accepts. Streamlined Form F0003 counts only " +
      "in full member states, from a member list you supply with its as-of date, which is printed. The MTC " +
      "uniform certificate counts only where it is accepted, and only for resale where a state says so. The " +
      "certificate must be unexpired on the invoice date, and its reason must fit the line. Resale, " +
      "agricultural, industrial and direct-mail reasons cover only their own lines, so a resale certificate " +
      "never exempts the buyer's own-use items. Government, charity, diplomat, education and direct-pay reasons " +
      "cover the whole order. In a Streamlined state, a promised certificate keeps a line exempt until 90 days " +
      "after the sale, printed with that deadline, and the line turns taxable once the deadline passes. Every " +
      "taxable line prints the rule it failed. All of it is one deriver, whose lines match the " +
      "us-sales-tax-breakdown shape.",
    category: "Body",
    tags: ["tax", "sales-tax", "us", "invoice", "exemption", "table", "deriver", "branching", "section"],
    exports: ["ExemptionCertificateCheck"],
    dataFields: [
      { path: "exemption.invoiceDate", type: "string", summary: "ISO 8601 sale date. Expiry and the 90-day window run from it. Required." },
      { path: "exemption.checkedOn", type: "string | undefined", summary: "ISO 8601 date of the check. Decides whether a pending certificate has run out. Missing means the invoice date." },
      { path: "exemption.shipToState", type: "string", summary: "Two-letter ship-to state. Every certificate is judged against it. Required." },
      { path: "exemption.currency", type: "string | undefined", summary: "ISO 4217 code. Missing means USD." },
      { path: "exemption.sstMembers", type: "{ asOf: string; states: string[] }", summary: "Streamlined full-member states and their as-of date, which is printed. Required." },
      {
        path: "exemption.mtcAcceptingStates",
        type: "{ asOf: string; states: Array<{ code: string; resaleOnly: boolean }> } | undefined",
        summary: "States accepting the MTC uniform certificate. Missing means none do.",
      },
      {
        path: "exemption.certificates",
        type:
          'Array<{ form: "SST" | "MTC" | "state"; state?: string; states?: string[]; reason: "A" | "B" | "C" | "D" | "E" | "F" | "G" | "H" | "I" | "J" | "K" | "L" | "M"; number?: string; expires?: string; receivedOn?: string; pending?: boolean }> | undefined',
        summary: "The buyer's certificates, held or promised. Missing means none: every line is taxable.",
      },
      {
        path: "exemption.lines",
        type: "Array<{ description: string; amountMinor: number; category?: string }> | undefined",
        summary:
          "Invoice lines in minor units, at most 16. The category is the buyer's use: resale, agricultural, industrial, direct-mail. Missing or empty prints nothing.",
      },
    ],
    arrangements: [
      { id: "table", title: "Table", summary: "One row per line with its amount, status and reason, so the statuses line up." },
      { id: "stacked", title: "Stacked", summary: "One paragraph per line, for a narrow column or long reasons." },
    ],
    files: node("exemption-certificate-check", { deriver: true, arrangements: ["table", "stacked"] }),
    previewData: {
      exemption: {
        invoiceDate: "2026-11-30",
        checkedOn: "2026-12-14",
        shipToState: "NC",
        currency: "USD",
        sstMembers: {
          asOf: "2026-10-01",
          states: ["AR", "GA", "IN", "IA", "KS", "KY", "MI", "MN", "NE", "NV", "NJ", "NC", "ND", "OH", "OK", "RI", "SD", "UT", "VT", "WA", "WV", "WI", "WY"],
        },
        mtcAcceptingStates: {
          asOf: "2026-07-01",
          states: [{ code: "MD", resaleOnly: true }, { code: "NJ", resaleOnly: false }, { code: "GA", resaleOnly: true }],
        },
        certificates: [
          { form: "SST", states: ["NC", "GA", "TN", "OH"], reason: "G", number: "NC-SST-500194-7731-REV3", expires: "2028-12-31" },
          { form: "state", state: "VA", reason: "I", number: "VA-ST11-2026-0044170-MFG" },
          { form: "SST", states: ["NC"], reason: "H", pending: true },
        ],
        lines: [
          { description: "Industrial-grade stainless steel conveyor assemblies, model CX-4400 (for resale)", amountMinor: 184250000, category: "resale" },
          { description: "Replacement hydraulic press tooling, 48-station progressive die set", amountMinor: 96437550, category: "industrial" },
          { description: "Commercial grain-drying fans and controllers, farm installation kit", amountMinor: 12834975, category: "agricultural" },
          { description: "Annual on-site preventive maintenance and calibration service contract", amountMinor: 4875000, category: "own-use" },
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "japan-qualified-invoice",
    title: "Japan qualified invoice (適格請求書)",
    summary: "The tax part of a Japanese qualified invoice: lines by rate, ※ marks, tax once per rate, and a checked T-number.",
    detail:
      "Lines are grouped into the 10% and reduced 8% bands, every reduced line is " +
      "marked ※ with the legend 「※は軽減税率対象」, and each band prints its " +
      "total and its tax. Tax is worked once per rate per invoice, never summed " +
      "from rounded lines, which the NTA forbids, and is rounded the way the " +
      "issuer chooses: floor, ceil or half-up. Tax-inclusive prices are taxed at " +
      "10/110 and 8/108, not 10% and 8%. The registration number is T and 13 " +
      "digits with a check digit, and a wrong one fails the build. A simplified " +
      "qualified invoice drops only the buyer's name, never the T-number, the date " +
      "or the reduced-rate marker. A seller with no registration gets a plain " +
      "invoice with no T-number and no stated tax, plus a note of the share the " +
      "buyer may still deduct, taken from the transaction date under the FY2026 " +
      "schedule: 80%, then 70%, 50% and 30%, then nothing after September 2031. " +
      "All of it is one deriver.",
    category: "Body",
    tags: ["money", "invoice", "tax", "japan", "consumption-tax", "table", "deriver", "branching", "section"],
    exports: ["JapanQualifiedInvoice"],
    dataFields: [
      { path: "jpInvoice.registrationNumber", type: "string | undefined", summary: "T and 13 digits, check digit verified. Missing means not registered: a plain invoice." },
      { path: "jpInvoice.simplified", type: "boolean | undefined", summary: "A simplified qualified invoice, for retail, restaurants, taxis, parking and travel. Makes the buyer name optional." },
      { path: "jpInvoice.rounding", type: '"floor" | "ceil" | "half-up"', summary: "How each rate's tax is rounded, once per invoice. Required when registered." },
      { path: "jpInvoice.basis", type: '"exclusive" | "inclusive"', summary: "Whether amounts are before or including tax. Required when registered." },
      { path: "jpInvoice.transactionDate", type: "string", summary: "ISO 8601, on or after 2023-10-01. Required." },
      { path: "jpInvoice.buyerName", type: "string | undefined", summary: "Printed as given, honorific included. Required on a full qualified invoice." },
      { path: "jpInvoice.lines", type: 'Array<{ description: string; amountYen: number; rate: "10" | "8" }>', summary: "1 to 10 items in whole yen. Reduced-rate items are marked ※." },
    ],
    arrangements: [
      { id: "table", title: "Table", summary: "Items and band totals in one two-column table, amounts lined up on the right." },
      { id: "stacked", title: "Stacked", summary: "One line per item and per total down the column, for a receipt or a narrow page." },
    ],
    files: node("japan-qualified-invoice", { deriver: true, arrangements: ["table", "stacked"] }),
    previewData: {
      jpInvoice: {
        registrationNumber: "T9234567890123",
        simplified: false,
        rounding: "floor",
        basis: "exclusive",
        transactionDate: "2026-11-30",
        buyerName: "株式会社みちのく総合食品流通ホールディングス 仙台第二物流センター 御中",
        lines: [
          { description: "宮城県産ササニシキ 精米 30kg袋（特別栽培米・JAS有機認証）", amountYen: 46980000, rate: "8" },
          { description: "業務用冷凍ショーケース（多段式・幅3,600mm）据付工事一式 仙台第二物流センター北棟", amountYen: 128450000, rate: "10" },
          { description: "三陸産冷凍ほたて貝柱 Lサイズ 1kg×10袋入りケース", amountYen: 27315750, rate: "8" },
          { description: "配送用折りたたみコンテナ（大型・蓋付き）リース料 2026年11月分", amountYen: 3862500, rate: "10" },
          { description: "国産若鶏もも肉 正肉 2kg×6パック入りケース", amountYen: 18642300, rate: "8" },
          { description: "コールドチェーン温度管理システム保守料 2026年11月分", amountYen: 1275000, rate: "10" },
          { description: "業務用ミネラルウォーター 2L×6本入りケース", amountYen: 4215615, rate: "8" },
        ],
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "vat-in-local-currency",
    title: "VAT in local currency",
    summary: "A foreign-currency invoice's VAT bands restated in the currency the VAT is owed in, band by band.",
    detail:
      "An invoice in one currency for VAT owed in another restates its VAT analysis in that " +
      "currency. Each rate band converts its net and its VAT separately, rounded once in the " +
      "target currency's minor units, and the local totals are the band sums, never a converted " +
      "gross. The rate carries its direction, so an inverse quote divides instead of multiplying, " +
      "and is worked in exact integers. It is checked against the tax point under the " +
      "jurisdiction's rule: the UK market selling, HMRC period or agreed rate; the EU selling or " +
      "ECB rate; the German BMF monthly average; the Polish NBP or ECB rate of the last working " +
      "day before the tax point; Bulgaria's fixed euro rate, unrounded and never inverted. A wrong " +
      "source or date prints a check-before-issuing line. On a final invoice each advance keeps " +
      "the rate it was declared at. The rate, its source and its date print beside the figures.",
    category: "Body",
    tags: ["invoice", "vat", "tax", "currency", "deriver", "branching", "section"],
    exports: ["VatInLocalCurrency"],
    dataFields: [
      { path: "vatConversion.jurisdiction", type: '"UK" | "EU" | "DE" | "PL" | "BG"', summary: "Whose conversion rules apply." },
      { path: "vatConversion.invoiceCurrency", type: "string", summary: "ISO 4217 code the invoice is issued in." },
      { path: "vatConversion.vatCurrency", type: "string", summary: "ISO 4217 code the VAT is owed in. Must fit the jurisdiction." },
      { path: "vatConversion.taxPointDate", type: "string", summary: "ISO 8601 tax point. The rate is checked against it." },
      { path: "vatConversion.issueDate", type: "string | undefined", summary: "ISO 8601 issue date. Required in PL." },
      { path: "vatConversion.bands", type: "Array<{ ratePercent: string | number; net: string | number; vat: string | number }>", summary: "The VAT analysis per rate, as decimal strings in the invoice currency." },
      { path: "vatConversion.rate", type: "{ value: string | number; base: string; quote: string; source: string; rateDate?: string }", summary: "1 base = value quote, with its source and date." },
      { path: "vatConversion.advances", type: "Array<{ invoiceNumber: string; chargeableDate: string; bands: VatBand[]; rate: VatExchangeRate }> | undefined", summary: "Advances deducted here, each at its own rate." },
    ],
    arrangements: [
      { id: "columns", title: "Columns", summary: "Local-currency net and VAT beside each band, in one table." },
      { id: "block", title: "Block", summary: "A separate local-currency statement, one line per band, under the totals." },
    ],
    files: node("vat-in-local-currency", { deriver: true, arrangements: ["columns", "block"] }),
    previewData: {
      vatConversion: {
        jurisdiction: "PL",
        invoiceCurrency: "EUR",
        vatCurrency: "PLN",
        taxPointDate: "2026-03-31",
        issueDate: "2026-03-31",
        bands: [
          { ratePercent: "23", net: "1284750.00", vat: "295492.50" },
          { ratePercent: "8", net: "48210.40", vat: "3856.83" },
          { ratePercent: "0", net: "12500.00", vat: "0.00" },
        ],
        rate: { value: "4.2725", base: "EUR", quote: "PLN", source: "nbp", rateDate: "2026-03-30" },
        advances: [{
          invoiceNumber: "FV/ZAL/2026/02/0147",
          chargeableDate: "2026-02-16",
          bands: [{ ratePercent: "23", net: "400000.00", vat: "92000.00" }],
          rate: { value: "4.2187", base: "EUR", quote: "PLN", source: "nbp", rateDate: "2026-02-13" },
        }],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "payment-surcharge-line",
    title: "Payment method surcharge",
    summary: "A payment-method surcharge only where it is lawful, capped, taxed with the supply, and disclosed.",
    detail:
      "Charges a payment-method surcharge only where it is lawful, and says why when it is not. " +
      "No UK consumer is surcharged for any method; no consumer card (credit, debit or prepaid) is " +
      "surcharged in the UK or EU, and no SEPA transfer in the EU; in the US only credit cards are, " +
      "never debit or prepaid, and never in a state that bans it. What is allowed is capped at the " +
      "cost of acceptance, and in the US at the network's and the state's cap too. Every percentage " +
      "is applied to the net total, the smallest base a cap could mean, so it holds whichever one a " +
      "rule intends. In the UK and EU the surcharge is taxed as part of the main supply, spread " +
      "across the invoice's VAT rates so a zero-rated share carries no VAT, never as a flat 20% line " +
      "or as exempt. The disclosure prints with the rate actually charged. All of it is one deriver.",
    category: "Body",
    tags: ["money", "invoice", "payment", "surcharge", "vat", "tax", "deriver", "branching", "table"],
    exports: ["PaymentSurchargeLine"],
    dataFields: [
      { path: "surcharge.jurisdiction", type: '"GB" | "EU" | "US"', summary: "Whose surcharging rules apply. Anything else fails the document." },
      { path: "surcharge.usState", type: "string | undefined", summary: "US only: two-letter state code. Missing there prints a warning instead of a surcharge." },
      { path: "surcharge.customer", type: '"consumer" | "business" | undefined', summary: "UK only. Missing means consumer, which bans any surcharge." },
      { path: "surcharge.method.kind", type: '"credit-card" | "debit-card" | "prepaid-card" | "commercial-card" | "sepa" | "bank-transfer"', summary: "How the customer pays. Decides whether a surcharge is allowed." },
      { path: "surcharge.method.network", type: "string | undefined", summary: "US only: card network for its cap. Missing gets the lowest cap, 3%." },
      { path: "surcharge.percent", type: "number | undefined", summary: "Requested rate, percent of the net total. Missing or 0 prints nothing." },
      { path: "surcharge.costOfAcceptance", type: "number | undefined", summary: "Cost of accepting the method, percent. The cap everywhere; missing prints a warning." },
      { path: "surcharge.usTaxable", type: "boolean | undefined", summary: "US only: whether the state taxes the surcharge. Missing means it does." },
      { path: "surcharge.currency", type: "string", summary: "ISO 4217 code. Sets minor units. Required." },
      { path: "surcharge.bands", type: "Array<{ ratePercent: number; net: number; vat?: number }>", summary: "The invoice's pre-surcharge totals by tax rate. The surcharge is spread across them by net." },
    ],
    arrangements: [
      {
        id: "table",
        title: "Table",
        summary: "The surcharge, its tax and the new total as label and amount rows under the invoice totals, notes beneath.",
      },
      {
        id: "stacked",
        title: "Stacked",
        summary: "Each line as a sentence, one under another, down the text column.",
      },
    ],
    files: node("payment-surcharge-line", { deriver: true, arrangements: ["table", "stacked"] }),
    previewData: {
      surcharge: {
        jurisdiction: "GB",
        customer: "business",
        method: { kind: "commercial-card", network: "mastercard" },
        percent: 2.5,
        costOfAcceptance: 1.95,
        currency: "GBP",
        bands: [
          { ratePercent: 20, net: 1184250.0, vat: 236850.0 },
          { ratePercent: 5, net: 96400.0, vat: 4820.0 },
          { ratePercent: 0, net: 48975.5, vat: 0 },
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "voucher-redemption-lines",
    title: "Voucher and coupon redemptions",
    summary: "Vouchers, coupons and gift cards placed where their tax treatment puts them: above the tax or below the total.",
    detail:
      "The totals under an invoice's lines, with every voucher, coupon and gift card placed where its tax treatment puts it. " +
      "A retailer's coupon, and in the US a free promotional card, reduces the taxable amount: it is spread across the VAT or " +
      "sales-tax bands of the lines it applies to by largest remainder, so the parts add up to the coupon exactly, and it never " +
      "takes a line below zero. A single-purpose voucher had its VAT accounted for when it was sold, so the part it covers is " +
      "shown as paid by voucher and left out of the VAT analysis rather than taxed twice. A multi-purpose voucher and a gift " +
      "card are payments: tax is charged on the full price and they reduce only the amount due. A US manufacturer's coupon " +
      "leaves the pre-coupon price taxable unless the seller says its state treats it as a discount. Excess value on a voucher " +
      "or card is printed as a remaining balance, never as change or a negative amount due. All of it is one deriver.",
    category: "Body",
    tags: ["money", "invoice", "receipt", "vat", "sales-tax", "voucher", "coupon", "table", "deriver", "branching"],
    exports: ["VoucherRedemptionLines"],
    dataFields: [
      { path: "sale.currency", type: "string", summary: "ISO 4217 code. Sets the minor units. Required." },
      { path: "sale.jurisdiction", type: '"UK" | "EU" | "US"', summary: "Whose voucher and tax rules apply, and which kinds are allowed. Required." },
      { path: "sale.manufacturerCouponsReduceBase", type: "boolean | undefined", summary: "US: the state treats a manufacturer coupon as a discount. Missing taxes the pre-coupon price." },
      { path: "sale.lines", type: "Array<{ key: string; netMinor: number; ratePercent: number }>", summary: "Pre-tax line amounts in minor units and their rates. Not printed. Required, non-empty." },
      {
        path: "sale.redemptions",
        type:
          'Array<{ kind: "spv" | "mpv" | "discount-coupon" | "retailer-coupon" | "manufacturer-coupon" | "gift-card" | "promo-card"; amountMinor: number; appliesToLineKeys?: string[]; code?: string; spvRatePercent?: number }>',
        summary: "Every voucher, coupon or card used, in the order presented. Missing prints plain totals.",
      },
    ],
    files: node("voucher-redemption-lines", { deriver: true }),
    previewData: {
      sale: {
        currency: "GBP",
        jurisdiction: "UK",
        lines: [
          { key: "spa-weekend", netMinor: 124999, ratePercent: 20 },
          { key: "treatment-package", netMinor: 84950, ratePercent: 20 },
          { key: "childrens-car-seat", netMinor: 18999, ratePercent: 5 },
          { key: "printed-guidebooks", netMinor: 4500, ratePercent: 0 },
        ],
        redemptions: [
          {
            kind: "discount-coupon", amountMinor: 25000, code: "LOYALTY-ANNIVERSARY-25-OFF-2026",
            appliesToLineKeys: ["spa-weekend", "treatment-package", "childrens-car-seat"],
          },
          { kind: "spv", amountMinor: 60000, appliesToLineKeys: ["spa-weekend"], spvRatePercent: 20, code: "SPA-GIFT-VOUCHER-0004817-UK" },
          { kind: "mpv", amountMinor: 200000, code: "HIGH-STREET-GIFT-CARD-9921-3348-7710" },
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "financial-assistance-notice",
    title: "Financial assistance and collections timeline (501(r))",
    summary: "The financial assistance notice a US tax-exempt hospital's bill must carry, and the dates collections may start and applications close.",
    detail:
      "Every statement from a tax-exempt hospital carries the 501(r) notice: " +
      "assistance is available, and where to get the policy, the application and " +
      "the plain-language summary. A missing phone number or web address prints " +
      "an 'incomplete' warning instead. Collections may not start until 120 days " +
      "after the first post-discharge statement and 30 days after the written " +
      "collections notice, and the later date is printed. Until that notice is " +
      "sent, the statement says collection waits for it. Applications are " +
      "accepted for 240 days, or until the collections date if later. A pending " +
      "application suspends collections and the statement says so. A partial " +
      "approval caps the balance at the amount generally billed, worked out in " +
      "whole cents and rounded half-up, and the statement shows the working. " +
      "Free care says nothing is owed. A for-profit facility prints nothing " +
      "unless state law asks for the notice, and then gets the notice alone. One " +
      "deriver decides all of it.",
    category: "Legal",
    tags: ["healthcare", "legal", "us", "collections", "branching", "deriver"],
    exports: ["FinancialAssistanceNotice"],
    dataFields: [
      { path: "facility.taxExempt", type: "boolean", summary: "Whether 501(r) applies. Missing is read as true." },
      { path: "facility.stateNoticeRequired", type: "boolean | undefined", summary: "For a for-profit facility: print the notice alone because state law requires it." },
      { path: "facility.fapUrl", type: "string", summary: "Where the policy, form and plain-language summary are. Missing prints a warning." },
      { path: "facility.fapPhone", type: "string", summary: "Who to call about the policy. Missing prints a warning." },
      { path: "billing.currency", type: "string | undefined", summary: "ISO 4217 code. Defaults to USD." },
      { path: "billing.firstPostDischargeStatementDate", type: "string", summary: "ISO 8601. Both clocks start here. Required." },
      { path: "billing.ecaNoticeDate", type: "string | undefined", summary: "ISO 8601 date of the written collections notice. Missing means none sent." },
      {
        path: "billing.application",
        type: '{ status?: "none" | "pending" | "approved" | "denied"; discount?: "free" | "partial" } | undefined',
        summary: "The application. Missing means none.",
      },
      { path: "billing.grossCharges", type: "number | undefined", summary: "Gross charges. Required for a partial approval." },
      { path: "billing.agbPercent", type: "number | undefined", summary: "AGB percentage, such as 37.85. Required for a partial approval." },
      { path: "billing.balanceDue", type: "number | undefined", summary: "What would otherwise be billed. The patient owes the lower of this and AGB." },
    ],
    files: node("financial-assistance-notice", { deriver: true }),
    previewData: {
      facility: {
        taxExempt: true,
        fapUrl: "https://www.stmarysregionalmedicalcenter.org/patients/billing/financial-assistance-policy",
        fapPhone: "1-800-555-0142 (TTY 711), Monday to Friday 8am-6pm Central",
      },
      billing: {
        currency: "USD",
        firstPostDischargeStatementDate: "2026-11-02",
        ecaNoticeDate: "2027-02-25",
        application: { status: "approved", discount: "partial" },
        grossCharges: 1284567.89,
        agbPercent: 37.85,
        balanceDue: 812450.0,
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "prepaid-hours-drawdown",
    title: "Prepaid hours drawdown",
    summary: "Time entries drawn against prepaid hour blocks soonest-expiring first, with overage, rollover and lapsed hours in a ledger.",
    detail:
      "Applies a period's time entries to prepaid retainer hour blocks in date order, drawing first on the block that expires soonest, not the oldest purchase. A block counts if the entry's date falls inside it, whatever the invoice date. Each entry is rounded up to the billing increment before it draws hours, and an entry that runs past the hours it may use is split at the exact minute: the covered part draws on the block and the rest is overage at the engagement's overage rate, or the timekeeper's standard rate if none is set. Blocks can be limited to roles, so partner time never draws on an associate-only block. When a block ends, unused hours roll into its successor up to a cap in hours or as a percentage of the block, keeping a shorter expiry when rolled hours do not extend. Rolled hours never roll twice. Anything above the cap is shown as forfeited with its value at the block rate, never credited, and VAT charged on the purchase is not reversed. Prints a block ledger, the overage, any split entry, and a warning when the balance is low or exhausted or hours are about to expire. All of it is one deriver.",
    category: "Body",
    tags: ["money", "invoice", "time", "retainer", "hours", "professional-services", "table", "deriver", "branching", "section"],
    exports: ["PrepaidHoursDrawdown"],
    dataFields: [
      { path: "retainer.currency", type: "string", summary: "ISO 4217 code. Sets the minor units. Required." },
      { path: "retainer.periodEnd", type: "string", summary: "ISO 8601 last day billed. Blocks ending on or before it are closed. Required." },
      { path: "retainer.increment", type: "number", summary: "Billing increment in minutes; entries round up to it before drawing. Required." },
      { path: "retainer.overageRate", type: "number | undefined", summary: "Overage rate per hour, minor units. Missing means each entry's standardRate." },
      { path: "retainer.lowBalanceHours", type: "number | undefined", summary: "Warn when open hours fall to this. Missing means no low-balance warning." },
      { path: "retainer.expiringDays", type: "number | undefined", summary: "Days after periodEnd an expiry is warned about. Missing means 30." },
      { path: "retainer.vatOnPurchase", type: "boolean | undefined", summary: "VAT was charged on the blocks; says it is not adjusted on forfeiture." },
      { path: "retainer.blocks", type: "Array<{ id: string; name?: string; hours: number; opening?: number; rate: number; start: string; expiry: string; rolloverCap?: { hours?: number; percent?: number }; rolledExpiry?: string; allowedRoles?: string[]; rolledInFrom?: string }>", summary: "Blocks not yet closed, any order. Rates in minor units per hour." },
      { path: "retainer.entries", type: "Array<{ date: string; timekeeper: string; role?: string; minutes: number; standardRate?: number }>", summary: "Time entries, any order, minutes before rounding." },
    ],
    files: node("prepaid-hours-drawdown", { deriver: true }),
    previewData: {
      retainer: {
        currency: "GBP",
        periodEnd: "2026-03-31",
        increment: 6,
        overageRate: 145000,
        lowBalanceHours: 40,
        expiringDays: 61,
        vatOnPurchase: true,
        blocks: [
          { id: "q4-carried", name: "Hours carried from the Q4 2025 commercial litigation retainer", hours: 30, opening: 24, rate: 62500, start: "2025-10-01", expiry: "2026-03-15" },
          { id: "q1", name: "Q1 2026 retainer: commercial litigation and regulatory advisory", hours: 250, opening: 100.3, rate: 62500, start: "2026-01-01", expiry: "2026-03-31", rolloverCap: { percent: 10 }, rolledExpiry: "2026-04-30", allowedRoles: ["partner", "senior associate", "associate"] },
          { id: "q2", name: "Q2 2026 retainer: commercial litigation and regulatory advisory", hours: 250, rate: 62500, start: "2026-04-01", expiry: "2026-06-30", rolloverCap: { percent: 10 }, rolledInFrom: "q1" },
          { id: "paralegal-q1", name: "Q1 2026 document review and disclosure support (paralegal time only)", hours: 120, opening: 6, rate: 18500, start: "2026-01-01", expiry: "2026-05-31", allowedRoles: ["paralegal", "trainee solicitor"] },
        ],
        entries: [
          { date: "2026-02-27", timekeeper: "Alexandra Montgomery-Fitzwilliam", role: "partner", minutes: 184, standardRate: 145000 },
          { date: "2026-03-03", timekeeper: "Oluwaseun Adebayo-Richardson", role: "senior associate", minutes: 437, standardRate: 89500 },
          { date: "2026-03-17", timekeeper: "Krystyna Wiśniewska-Hargreaves", role: "paralegal", minutes: 1250, standardRate: 24500 },
          { date: "2026-03-12", timekeeper: "Alexandra Montgomery-Fitzwilliam", role: "partner", minutes: 611, standardRate: 145000 },
          { date: "2026-03-18", timekeeper: "Oluwaseun Adebayo-Richardson", role: "senior associate", minutes: 1904, standardRate: 89500 },
          { date: "2026-03-24", timekeeper: "Maximilian Featherstonehaugh", role: "associate", minutes: 1022, standardRate: 72500 },
          { date: "2026-03-27", timekeeper: "Alexandra Montgomery-Fitzwilliam", role: "partner", minutes: 377, standardRate: 145000 },
          { date: "2026-03-30", timekeeper: "Priyanka Venkataraman-Castellanos", role: "consultant", minutes: 248, standardRate: 115000 },
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "contingency-settlement-statement",
    title: "Contingency settlement statement",
    summary:
      "The closing statement for a contingent or damages-based fee: the fee by band, the statutory cap, liens, and the net to the client.",
    detail:
      "The closing statement a contingent or damages-based fee requires: the recovery, each deduction in the order the agreement and the law apply it, and the net to the client, in one column that adds up. " +
      "Tiered rates are charged marginally by band, never the top rate on the whole sum, at the rate for the stage the case reached. " +
      "The fee is worked out on gross or on recovery net of expenses as the agreement says, and the statement shows what the other base would have charged. " +
      "A statutory cap (New York §474-a, California §6146, the UK DBA Regulations or the CFA Order) is computed and compared with the agreed fee, and the lower is charged. " +
      "The DBA cap counts VAT and counsel's fees and leaves out future loss, and costs paid by the other side are credited. " +
      "Liens come off the client's share after the fee unless marked otherwise. " +
      "A recovery that does not cover everything prints as a shortfall with the agreement's terms on who carries it. " +
      "All of it is worked out by one deriver that ships with the component.",
    category: "Body",
    tags: ["legal", "settlement", "fees", "liens", "table", "deriver", "branching"],
    exports: ["ContingencySettlementStatement"],
    dataFields: [
      { path: "settlement.regime", type: '"contingency" | "DBA" | "CFA"', summary: "Which agreement the fee is paid under. Anything else throws." },
      { path: "settlement.currency", type: "string", summary: "ISO 4217 code all amounts are in." },
      { path: "settlement.grossRecovery", type: "number", summary: "Total awarded or agreed, before anything is taken off." },
      { path: "settlement.stage", type: '"pre-suit" | "filed" | "trial" | "appeal" | undefined', summary: "Picks the rate. Required when the schedule is staged or the California cap applies." },
      { path: "settlement.schedule", type: "Array<{ stage?: string; upTo?: number; percent: number }>", summary: "The agreement's rates: marginal bands, or the CFA uplift." },
      { path: "settlement.base", type: '"gross" | "net" | undefined', summary: "Whether the fee is on the gross recovery or net of expenses. Required for contingency and DBA." },
      { path: "settlement.statutoryCap", type: '"ny-medical-malpractice" | "ca-medical-negligence" | undefined', summary: "A statutory fee cap to compare against. Contingency only." },
      { path: "settlement.claimType", type: '"personal-injury" | "employment" | "other" | undefined', summary: "Which DBA or CFA cap applies. Required for DBA and CFA." },
      { path: "settlement.damagesBreakdown", type: "{ general: number; pastLoss: number; futureLoss?: number } | undefined", summary: "Heads of damage for the UK personal-injury caps. Future loss is excluded." },
      { path: "settlement.baseCosts", type: "number | undefined", summary: "CFA base costs before VAT. Required for CFA." },
      { path: "settlement.vatPercent", type: "number | undefined", summary: "VAT on the solicitor's charges, DBA and CFA. Missing is 0." },
      { path: "settlement.counselFees", type: "number | undefined", summary: "Counsel's fees including VAT, inside the DBA cap." },
      { path: "settlement.recoverableCostsPaid", type: "number | undefined", summary: "Costs paid by the other side, credited against what the client owes." },
      { path: "settlement.expenses", type: "Array<{ label: string; amount: number }> | undefined", summary: "Case expenses, one row each." },
      { path: "settlement.liens", type: "Array<{ holder: string; amount: number; reducedAmount?: number; beforeFee?: boolean; cru?: boolean }> | undefined", summary: "Liens paid from the client's share after the fee unless marked beforeFee." },
      { path: "settlement.shortfallTerms", type: "string | undefined", summary: "The agreement's words on who carries a shortfall. Missing there prints a warning." },
    ],
    files: node("contingency-settlement-statement", { deriver: true }),
    previewData: {
      settlement: {
        regime: "contingency",
        currency: "USD",
        grossRecovery: 4850000,
        stage: "trial",
        base: "net",
        statutoryCap: "ny-medical-malpractice",
        schedule: [
          { stage: "pre-suit", percent: 25 },
          { stage: "filed", percent: 30 },
          { stage: "trial", percent: 100 / 3 },
          { stage: "appeal", percent: 40 },
        ],
        expenses: [
          { label: "Expert witness fees: board-certified neurosurgeon, life-care planner and forensic economist", amount: 186450 },
          { label: "Court reporter and videographer charges for fourteen depositions", amount: 38912.75 },
          { label: "Medical records retrieval and certification, Mount Sinai and NYU Langone", amount: 6240.3 },
          { label: "Index number, request for judicial intervention and jury demand filing fees", amount: 1335 },
          { label: "Trial exhibits, demonstrative animations and courtroom technology", amount: 54780 },
        ],
        liens: [
          { holder: "Centers for Medicare & Medicaid Services (Medicare Secondary Payer Recovery)", amount: 312480.66, reducedAmount: 198214.09 },
          { holder: "Empire BlueCross BlueShield of New York, self-funded ERISA health plan", amount: 147905.12, reducedAmount: 98603.41 },
          { holder: "New York State Workers' Compensation Board lien", amount: 64220 },
        ],
        shortfallTerms:
          "If the recovery does not cover the fee, expenses and liens, the firm waives any unrecovered expenses and the client remains responsible for the liens.",
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "stage-payment-schedule",
    title: "Stage payment schedule",
    summary: "Stage payments that add up to the contract sum to the penny, with each stage's VAT tax point.",
    detail:
      "Turns a contract sum and its milestones into a schedule that adds up to the " +
      "contract sum exactly. Percentage stages are floored to the penny and the " +
      "pennies lost go on one stage, so the total is never a penny out; the last " +
      "stage can be the balance. The deposit is netted off by the contract's rule " +
      "(first stage, final stage or pro rata), and any part bigger than its stage " +
      "carries on rather than leaving a negative stage. Each stage is pending, " +
      "reached, invoiced or paid, and the totals split into paid, invoiced and not " +
      "yet invoiced. Under UK VAT regulation 93 a stage's tax point is the earlier " +
      "of its invoice and its payment, never the milestone, so a reached stage " +
      "shows no VAT due; and because the deposit was taxed when it was received, " +
      "stage VAT is worked on what is left after it. A single-payment contract " +
      "gets one line saying its tax point is completion. All of it is one deriver.",
    category: "Body",
    tags: ["money", "construction", "contract", "vat", "tax", "table", "deriver", "branching", "section"],
    exports: ["StagePaymentSchedule"],
    dataFields: [
      { path: "contract.currency", type: "string | undefined", summary: "ISO 4217 code. Missing means GBP." },
      { path: "contract.contractSumMinor", type: "number", summary: "Contract sum ex VAT, integer minor units. Required." },
      { path: "contract.vat", type: '{ jurisdiction: "UK"; ratePercent: number } | undefined', summary: "UK VAT and its rate. Missing means no VAT or tax point columns." },
      { path: "contract.singlePayment", type: "boolean | undefined", summary: "One payment on completion: exactly one stage, outside regulation 93." },
      { path: "contract.deposit", type: '{ amountMinor: number; receivedDate: string; application: "first" | "final" | "proRata" } | undefined', summary: "Deposit ex VAT, when received, and how it is netted off. Missing or 0 means none." },
      {
        path: "contract.stages",
        type: 'Array<{ label: string; percent?: number; amountMinor?: number; balance?: true; state: "pending" | "reached" | "invoiced" | "paid"; reachedDate?: string; invoiceDate?: string; paidDate?: string }>',
        summary: "1 to 8 stages in contract order; each a percent, a fixed amount or the balance (last only). Must add up to the contract sum.",
      },
    ],
    files: node("stage-payment-schedule", { deriver: true }),
    previewData: {
      contract: {
        currency: "GBP",
        contractSumMinor: 4875250033,
        vat: { jurisdiction: "UK", ratePercent: 20 },
        deposit: { amountMinor: 487525003, receivedDate: "2026-02-16", application: "first" },
        stages: [
          { label: "Site establishment, enabling works and temporary services", amountMinor: 350000000, state: "paid", reachedDate: "2026-03-02", invoiceDate: "2026-03-06", paidDate: "2026-03-27" },
          { label: "Substructure, piling and ground-floor slab complete", percent: 12.5, state: "paid", reachedDate: "2026-05-11", invoiceDate: "2026-05-15", paidDate: "2026-05-13" },
          { label: "Frame, upper floors and roof structure topped out", percent: 17.5, state: "invoiced", reachedDate: "2026-08-03", invoiceDate: "2026-08-07" },
          { label: "Building envelope weathertight, including curtain walling", percent: 15, state: "reached", reachedDate: "2026-10-19" },
          { label: "Mechanical, electrical and plumbing first fix complete", percent: 15, state: "pending" },
          { label: "Internal finishes, second fix and commissioning", percent: 12.5, state: "pending" },
          { label: "Practical completion and handover of the building", percent: 10, state: "pending" },
          { label: "Making good defects at the end of the rectification period", balance: true, state: "pending" },
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "pay-less-notice",
    title: "Pay less notice",
    summary: "A UK construction pay less notice: the sum due, every deduction itemised as its basis, and whether it was served in time.",
    detail:
      "Starts from the notified sum (the payment notice, or the payee's application or default notice) and lists every deduction under its own kind, so the itemised schedule is the basis of calculation section 111 requires, not one netted figure. " +
      "Liquidated damages are days of delay times the daily or weekly rate, pro rata by the day, capped at the LD cap less what earlier notices already deducted, and the line says when the cap bit. " +
      "A CIS deduction fails the document: tax withholding is not a pay-less ground. " +
      "When deductions exceed the notified sum the sum due is nil and a separate line says the excess is not recovered here, so no negative payment is printed. " +
      "A banner counts the pay-less deadline back from the final date for payment, skipping the days section 116 excludes, and says plainly when a late notice leaves the notified sum payable. " +
      "A generated paragraph explains the schedule; the schedule stands as the basis whatever it says. All the arithmetic is one deriver.",
    category: "Body",
    tags: ["construction", "payment", "notice", "legal", "money", "table", "deriver", "branching", "generated", "section"],
    exports: ["PayLessNotice"],
    dataFields: [
      { path: "payLess.contractRef", type: "string", summary: "The contract the notice is given under. Required." },
      { path: "payLess.payer", type: "string", summary: "Who gives the notice. Required." },
      { path: "payLess.payee", type: "string", summary: "Who receives it. Required." },
      {
        path: "payLess.notifiedSum",
        type: '{ amountMinor: number; source: "payment-notice" | "application" | "default-notice"; reference?: string }',
        summary: "The notified sum in pence and the notice that fixed it. Required.",
      },
      { path: "payLess.finalDateForPayment", type: "string", summary: "ISO 8601 final date for payment. Required." },
      { path: "payLess.payLessDeadline", type: "string | undefined", summary: "ISO 8601 last day to serve. Wins over the prescribed period." },
      { path: "payLess.prescribedPeriodDays", type: "number | undefined", summary: "Days before the final date for payment; 7 under the Scheme. One of this or the deadline is required." },
      { path: "payLess.excludedDays", type: "string[] | undefined", summary: "Christmas Day, Good Friday and bank holidays, skipped when counting back (s.116)." },
      { path: "payLess.servedDate", type: "string", summary: "ISO 8601 date of service. Required." },
      {
        path: "payLess.ld",
        type: '{ rateMinor: number; per: "day" | "week"; delayDays: number; capMinor?: number; previouslyDeductedMinor?: number } | undefined',
        summary: "Liquidated damages; the cap is net of earlier notices. Missing: no LD line.",
      },
      {
        path: "payLess.deductions",
        type: 'Array<{ kind: "contra" | "abatement" | "setoff" | "cis"; description: string; amountMinor: number; reference?: string }> | undefined',
        summary: "Each deduction in pence. CIS fails the document. At most 12 lines with LDs.",
      },
    ],
    files: node("pay-less-notice", { deriver: true }),
    previewData: {
      payLess: {
        contractRef: "JCT Design and Build Sub-Contract 2016, NWH-PH2-MEP-0047, North Wharf Hospital Phase 2 Main Works",
        payer: "Northgate Healthcare Infrastructure Construction (Phase 2) Limited",
        payee: "Brightwater Mechanical & Electrical Building Services Contractors Limited",
        notifiedSum: { amountMinor: 487231645, source: "payment-notice", reference: "Interim Payment Notice No. 14, valuation date 31 October 2026" },
        finalDateForPayment: "2026-11-30",
        prescribedPeriodDays: 7,
        excludedDays: [],
        servedDate: "2026-11-20",
        ld: { rateMinor: 3500000, per: "week", delayDays: 47, capMinor: 125000000, previouslyDeductedMinor: 105000000 },
        deductions: [
          { kind: "contra", description: "Replacement by others of damaged ductwork on the Level 4 ward block following a failed pressure test", amountMinor: 18644218, reference: "CC-0193" },
          { kind: "abatement", description: "Level 5 and 6 medical gas pipeline installation valued at 62% complete against 85% applied for", amountMinor: 31298050, reference: "Valuation 14, item 23.07" },
          { kind: "setoff", description: "Delay and disruption costs to Plant Room B under Sub-Contract Order SO-118", amountMinor: 9460000, reference: "Notice of set-off dated 3 November 2026" },
          { kind: "contra", description: "Out-of-hours security and site cleaning charged under clause 2.9 of the sub-contract", amountMinor: 4875000, reference: "CC-0201" },
        ],
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "deposit-cancellation-outcome",
    title: "Deposit outcome on cancellation",
    summary: "What a cancelled order refunds and keeps from money paid in advance, and the VAT on each part.",
    detail:
      "Security deposits, refundable advances and non-refundable advances look alike and are " +
      "taxed differently. A security deposit carries no VAT and comes back less itemised " +
      "deductions, each marked as a VAT-bearing supply or as compensation outside the scope. In " +
      "the UK since March 2019, VAT on a retained advance stays due with no credit note, and the " +
      "refunded part gets one at the original rate. In Germany the VAT is corrected only when the " +
      "refund is actually paid. In the US a state flag says whether a forfeit is taxed, rather " +
      "than a guess. A partly retained advance is split per VAT band, with VAT worked on the " +
      "retained part and the remainder refunded, so the halves always add up to the VAT charged. " +
      "No amount is printed negative. One deriver does all of it.",
    category: "Body",
    tags: ["cancellation", "deposit", "refund", "vat", "tax", "deriver", "branching", "section"],
    exports: ["DepositCancellationOutcome"],
    dataFields: [
      { path: "order.reference", type: "string", summary: "The cancelled order, printed as given. Required." },
      { path: "order.currency", type: "string", summary: "ISO 4217 code. Required." },
      { path: "order.jurisdiction", type: '"UK" | "DE" | "US"', summary: "Whose rules apply. Anything else fails the document." },
      { path: "order.usState", type: "string | undefined", summary: "US only: the state named in the sales-tax line. Required when an advance is retained." },
      { path: "order.usForfeitTaxable", type: "boolean | undefined", summary: "US only: whether that state taxes a forfeited deposit. Required when an advance is retained." },
      { path: "order.cancelledOn", type: "string", summary: "ISO 8601 cancellation date. Required; UK before 2019-03-01 fails." },
      { path: "order.payments", type: "Array<{ kind; gross; bands?; invoiceNumber? }> | undefined", summary: "Money taken in advance, with its VAT bands as invoiced. Missing means nothing was paid." },
      { path: "order.retention", type: "{ percent?; amount?; reason } | undefined", summary: "The cancellation fee kept from refundable advances: a percent or an amount, and why." },
      { path: "order.deductions", type: "Array<{ label; amount; isSupply?; vatRate? }> | undefined", summary: "Up to 6 itemised deductions from the security deposit." },
      { path: "order.creditNoteNumber", type: "string | undefined", summary: "The credit note for the refunded VAT. Missing says one will follow." },
      { path: "order.refundPaidOn", type: "string | undefined", summary: "When the refund was paid. Missing means not yet; in Germany it times the VAT correction." },
    ],
    arrangements: [
      { id: "table", title: "Table", summary: "Each label beside its amount, amounts right-aligned so they add down the column." },
      { id: "stacked", title: "Stacked", summary: "One 'Label: amount' line under another, down the text column." },
    ],
    files: node("deposit-cancellation-outcome", { deriver: true, arrangements: ["table", "stacked"] }),
    previewData: {
      order: {
        reference: "BK-2026-HALL-0048817/WEDDING-RECEPTION-REV-3",
        currency: "GBP",
        jurisdiction: "UK",
        cancelledOn: "2026-11-03",
        payments: [
          { kind: "security", gross: 25000.0 },
          {
            kind: "refundable-advance",
            gross: 1284750.0,
            bands: [
              { rate: 20, net: 1000000.0, vat: 200000.0 },
              { rate: 5, net: 80714.29, vat: 4035.71 },
            ],
            invoiceNumber: "SI-2026-07-004817/DEP-1",
          },
          {
            kind: "non-refundable-advance",
            gross: 48000.0,
            bands: [{ rate: 20, net: 40000.0, vat: 8000.0 }],
            invoiceNumber: "SI-2026-05-003302/RES",
          },
        ],
        retention: { percent: 25, reason: "cancellation less than 90 days before the event, clause 14.2" },
        deductions: [
          { label: "Specialist cleaning of the main hall carpets after the rehearsal dinner", amount: 1850.4, isSupply: true, vatRate: 20 },
          { label: "Replacement of two damaged antique Venetian glass chandeliers in the Long Gallery", amount: 12600.0, isSupply: false },
        ],
        creditNoteNumber: "CN-2026-11-000913/BK-0048817",
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "prompt-payment-discount",
    title: "Prompt payment discount terms",
    summary: "Early-payment discount terms, worked per VAT band under the rule where the invoice is issued.",
    detail:
      "Since April 2015 UK VAT is charged on the full price, so this never " +
      "recomputes the invoice's VAT on the discounted price and never takes the " +
      "discount off the gross total. A UK invoice gets, per band, the discounted " +
      "net, the VAT on it rounded once per band, and what is payable if paid " +
      "early, followed by HMRC's wording when no credit note will be issued, " +
      "without which the supplier has to issue one. A German invoice states the " +
      "Skonto with a fixed date and, on a mixed-rate invoice, the reduction per " +
      "rate, split to the cent by largest remainder; the invoice's own VAT stays " +
      "on the full price. Zero-rated, exempt and reverse-charge bands have their " +
      "net reduced and their VAT left at nothing. A US invoice gets the usual " +
      "'2/10, net 30' and the early-payment amount, with sales tax on the full or " +
      "discounted price looked up per state, never assumed. No discount prints " +
      "nothing; a discount deadline after the due date fails the document. All " +
      "of it is one deriver.",
    category: "Closing",
    tags: ["invoice", "vat", "tax", "payment", "discount", "table", "deriver", "branching", "section"],
    exports: ["PromptPaymentDiscount"],
    dataFields: [
      { path: "invoice.issueDate", type: "string", summary: "ISO 8601 date of issue; the deadline is counted from it. Required when a discount applies." },
      { path: "invoice.currency", type: "string", summary: "ISO 4217 code. Sets the minor units. Required when a discount applies." },
      { path: "invoice.jurisdiction", type: '"UK" | "EU" | "DE" | "FR" | "US"', summary: "Whose rules apply, and which language the block is written in." },
      { path: "invoice.usState", type: "string | undefined", summary: "US only: decides whether sales tax follows the discount. Missing or unknown fails the document." },
      { path: "invoice.dueDays", type: "number | undefined", summary: "The payment term in days. A discount deadline after it fails the document. Missing skips the check." },
      { path: "invoice.bands", type: 'Array<{ rate: number | null; treatment: "standard" | "reduced" | "zero" | "exempt" | "reverse-charge"; net: number; vat: number }>', summary: "The VAT analysis at full price, major units, at most six non-zero bands. Required when a discount applies." },
      { path: "invoice.promptPaymentDiscount", type: "{ percent: number; days: number; issuesCreditNote?: boolean } | undefined", summary: "The terms. Missing, 0% or no deadline prints nothing. issuesCreditNote is UK only; missing prints HMRC's no-credit-note wording." },
    ],
    files: node("prompt-payment-discount", { deriver: true }),
    previewData: {
      invoice: {
        issueDate: "2026-11-30",
        currency: "GBP",
        jurisdiction: "UK",
        dueDays: 30,
        bands: [
          { rate: 20, treatment: "standard", net: 1284750.0, vat: 256950.0 },
          { rate: 5, treatment: "reduced", net: 48215.4, vat: 2410.77 },
          { rate: 0, treatment: "zero", net: 12480.0, vat: 0 },
          { rate: null, treatment: "reverse-charge", net: 96300.0, vat: 0 },
        ],
        promptPaymentDiscount: { percent: 2.5, days: 14, issuesCreditNote: false },
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "credits-applied-to-invoice",
    title: "Credits applied to invoice",
    summary: "Open credit notes and payments on account set against an invoice, from total to amount due.",
    detail:
      "Credits reduce the amount due, never the total or its VAT: a credit note " +
      "adjusted VAT when it was issued and a payment on account is a means of " +
      "payment, so the section sits below the VAT analysis. Credits are taken " +
      "oldest first, or by your allocation, each capped at what is still owed, " +
      "and what is left of each is printed as carried forward. Only credits in " +
      "the invoice currency are applied; others, expired credits and disputed " +
      "ones are listed with the reason. Credit notes and payments on account are " +
      "grouped apart. The amount due never goes negative: when credits cover it, " +
      "the section says there is nothing to pay. Integer minor units throughout, " +
      "so the figures tie out, worked by a deriver that ships with it.",
    category: "Body",
    tags: ["invoice", "credit", "payment", "vat", "table", "deriver", "branching", "section"],
    exports: ["CreditsAppliedToInvoice"],
    dataFields: [
      { path: "invoice.currency", type: "string", summary: "ISO 4217 code. Only credits in this currency are applied. Required." },
      { path: "invoice.issueDate", type: "string", summary: "ISO 8601 issue date. Expiry is judged against it. Required." },
      { path: "invoice.totalMinor", type: "number", summary: "VAT-inclusive total in minor units, as the VAT analysis printed it. Never reduced. Required." },
      { path: "invoice.vatTotalMinor", type: "number | undefined", summary: "VAT within the total, in minor units. When given, a sentence says credits leave it unchanged." },
      { path: "credits.open", type: "Array<{ id; kind: \"credit-note\" | \"payment-on-account\"; issuedOn; currency; remainingMinor; expiresOn?; disputed? }> | undefined", summary: "Open credits in any order. Missing or empty prints nothing." },
      { path: "credits.allocation", type: "Array<{ creditId: string; amountMinor: number }> | undefined", summary: "Your ledger's allocation, used in place of oldest first. Each amount is capped." },
    ],
    files: node("credits-applied-to-invoice", { deriver: true }),
    previewData: {
      invoice: { currency: "GBP", issueDate: "2026-11-30", totalMinor: 128475000, vatTotalMinor: 21412500 },
      credits: {
        open: [
          { id: "CN-2026-000418/REV-1-FACILITIES-NORTH", kind: "credit-note", issuedOn: "2026-03-14", currency: "GBP", remainingMinor: 4250000 },
          { id: "CN-2026-000771-SERVICE-LEVEL-REBATE-Q2", kind: "credit-note", issuedOn: "2026-05-30", currency: "GBP", remainingMinor: 1892745 },
          { id: "POA-2026-07-0093-ADVANCE-MOBILISATION", kind: "payment-on-account", issuedOn: "2026-07-02", currency: "GBP", remainingMinor: 125000000 },
          { id: "CN-2025-001207-GOODWILL-ADJUSTMENT", kind: "credit-note", issuedOn: "2025-10-31", currency: "GBP", remainingMinor: 975000, expiresOn: "2026-10-31" },
          { id: "POA-2026-09-0117-RETENTION-RELEASE", kind: "payment-on-account", issuedOn: "2026-09-15", currency: "GBP", remainingMinor: 3500000, disputed: true },
          { id: "CN-US-2026-000052-FREIGHT-OVERCHARGE", kind: "credit-note", issuedOn: "2026-06-01", currency: "USD", remainingMinor: 1250000 },
        ],
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "canadian-sales-tax",
    title: "Canadian sales tax",
    summary: "GST/HST and provincial sales tax for the province of supply, each on its own base, with CRA's invoice checks.",
    detail:
      "Tax follows the province of supply, where goods are delivered or the " +
      "customer is, never the seller's province. HST provinces get one line. BC " +
      "and SK get GST plus PST, MB GST plus RST, Quebec GST plus QST, and Alberta " +
      "and the territories GST alone. Each tax is worked on its own base and " +
      "rounded once, so PST and QST exemptions are honoured. QST is never charged " +
      "on GST. Rates are effective-dated rows picked by the tax point date, because " +
      "they change mid-year. The GST/HST number must pass its check digit to be " +
      "printed, and a seller that is not registered charges none and says so. " +
      "CRA's $100 and $500 information tiers are applied to the tax-included " +
      "total, and anything missing is printed as a warning. All of it is one deriver.",
    category: "Body",
    tags: ["tax", "invoice", "canada", "gst", "money", "table", "deriver", "section"],
    exports: ["CanadianSalesTax"],
    arrangements: [
      { id: "table", title: "Table", summary: "Labels on the left, amounts right-aligned beside them, under the line items." },
      { id: "stacked", title: "Stacked", summary: "One labelled line under another, for a narrow column or a plain letter." },
    ],
    dataFields: [
      { path: "salesTax.supplyProvince", type: '"AB" | "BC" | "MB" | "NB" | "NL" | "NS" | "NT" | "NU" | "ON" | "PE" | "QC" | "SK" | "YT"', summary: "Province of supply: delivery province for goods, customer's for services. Not the seller's. Required." },
      { path: "salesTax.taxPointDate", type: "string", summary: "ISO 8601 day tax became payable; picks the rate rows. Required; before 2013-04-01 fails." },
      { path: "salesTax.lines", type: '{ description?: string; amount: number; gstCategory?: "taxable" | "zero-rated" | "exempt"; provincialCategory?: "taxable" | "exempt" }[]', summary: "Pre-tax lines. GST status defaults to taxable; provincial status follows it." },
      { path: "salesTax.rates", type: '{ tax: "GST" | "HST" | "PST" | "RST" | "QST"; province?: string; rate: number; from: string; to?: string }[] | undefined', summary: "Effective-dated rates in percent, both ends inclusive. Missing uses the built-in table; given, replaces it." },
      { path: "salesTax.paymentTerms", type: "string | undefined", summary: "Not printed; missing warns from $500." },
      { path: "seller.name", type: "string | undefined", summary: "Not printed; missing warns at every tier." },
      { path: "seller.gstRegistered", type: "boolean | undefined", summary: "Missing means a small supplier: no GST/HST, and a note." },
      { path: "seller.bn", type: "string | undefined", summary: "GST/HST number, 9-digit BN + RT + 4 digits. Printed only when it passes the check digit." },
      { path: "seller.provincialRegistered", type: "boolean | undefined", summary: "Registered for the province's PST, RST or QST. Missing charges none and warns." },
      { path: "seller.qst", type: "string | undefined", summary: "QST number, 10 digits + TQ + 4 digits. Printed when QST is charged and it is well-formed." },
      { path: "customer.name", type: "string | undefined", summary: "Not printed; missing warns from $500." },
    ],
    files: node("canadian-sales-tax", { deriver: true, arrangements: ["table", "stacked"] }),
    previewData: {
      salesTax: {
        supplyProvince: "QC",
        taxPointDate: "2026-11-30",
        paymentTerms: "Net 45 days from invoice date; 2% per month on overdue balances",
        lines: [
          { description: "Annual preventive maintenance contract for HVAC and building automation systems, Tour Centre-Ville, 1250 boulevard René-Lévesque Ouest, floors 1 to 42", amount: 1284750.0, gstCategory: "taxable", provincialCategory: "taxable" },
          { description: "Printed operating manuals and technical reference books, bilingual edition, 12 volumes", amount: 1840.0, gstCategory: "taxable", provincialCategory: "exempt" },
          { description: "Long-term residential accommodation for relocated caretaker, 1 to 30 November 2026", amount: 2450.0, gstCategory: "exempt", provincialCategory: "exempt" },
          { description: "Emergency out-of-hours call-out, chilled water pump failure, 14 and 15 November", amount: 18425.5, gstCategory: "taxable", provincialCategory: "taxable" },
        ],
      },
      seller: {
        name: "Groupe Maintenance Technique Saint-Laurent-Outaouais Inc.",
        gstRegistered: true,
        bn: "843219767RT0001",
        provincialRegistered: true,
        qst: "1219876543TQ0001",
      },
      customer: { name: "Fiducie de placement immobilier Centre-Ville Montréal-Laval S.E.C." },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "net-payable-withholdings",
    title: "Net payable after withholdings",
    summary: "From a Spanish or Italian invoice total to what the customer actually transfers, after IRPF, ritenuta and split payment.",
    detail:
      "Bridges a Spanish or Italian invoice total to what the customer actually transfers, each row labelled with its legal basis. IRPF (15%, or 7% for a new professional) comes off the base imponible and the ritenuta d'acconto (20%) off the fee plus rivalsa INPS, never off IVA. A cassa contribution carries IVA but stays out of the ritenuta base. Consumers are never withheld from. In Italy only a sostituto d'imposta withholds, and never from a forfettario supplier. The recargo de equivalencia follows the IVA rate. Split payment shows the IVA and then takes it off the net payable, and is refused alongside a ritenuta. The 2.00 € imposta di bollo applies over 77.47 € carrying no IVA and is added only when re-charged. Every amount is rounded to the cent in integer cents, in legal order, so the rows add up. Unclear data fails the document instead of guessing.",
    category: "Closing",
    tags: ["invoice", "withholding", "irpf", "ritenuta", "vat", "tax", "spain", "italy", "deriver", "table", "branching"],
    exports: ["NetPayableWithholdings"],
    dataFields: [
      { path: "invoice.country", type: '"ES" | "IT"', summary: "Whose rules apply. Required." },
      { path: "invoice.base", type: "string | number", summary: "Price or fee before contributions, VAT and withholding, in euros. Required." },
      { path: "invoice.vatPercent", type: "string | number | undefined", summary: "VAT rate as a percentage. Required except for an Italian forfettario supplier." },
      { path: "invoice.exemptAmount", type: "string | number | undefined", summary: "Expenses re-charged at cost outside VAT and withholding. Missing means none." },
      { path: "invoice.splitPayment", type: "boolean | undefined", summary: "Italy: IVA shown, then deducted (art. 17-ter). Refused with a ritenuta." },
      { path: "invoice.stampRecharged", type: "boolean | undefined", summary: "Italy: add the 2.00 € bollo to the total. Otherwise it is a note." },
      { path: "supplier.professional", type: "boolean | undefined", summary: "Fees carry IRPF or ritenuta. Missing means nothing is withheld." },
      { path: "supplier.newProfessional", type: "boolean | undefined", summary: "Spain: IRPF at 7% instead of 15%." },
      { path: "supplier.forfettario", type: "boolean | undefined", summary: "Italy: regime forfettario, so no IVA and no ritenuta." },
      { path: "supplier.rivalsaInps", type: "boolean | undefined", summary: "Italy: 4% rivalsa INPS, in the ritenuta base." },
      { path: "supplier.cassaPercent", type: "string | number | undefined", summary: "Italy: cassa contributo integrativo, subject to IVA, not to ritenuta." },
      { path: "supplier.cassaName", type: "string | undefined", summary: "The cassa's name, printed in its row." },
      { path: "customer.type", type: '"business" | "consumer" | "public-administration"', summary: "A consumer is never withheld from. Required." },
      { path: "customer.reRegime", type: "boolean | undefined", summary: "Spain: retailer in the recargo de equivalencia, goods only." },
      { path: "customer.withholdingAgent", type: "boolean | undefined", summary: "Italy: sostituto d'imposta. Required when a ritenuta could apply." },
    ],
    files: node("net-payable-withholdings", { deriver: true, arrangements: ["table", "stacked"] }),
    previewData: {
      invoice: { country: "IT", base: "48650.00", vatPercent: "22", exemptAmount: "1285.40", stampRecharged: true },
      supplier: { professional: true, cassaPercent: "4", cassaName: "Cassa Nazionale di Previdenza e Assistenza Forense" },
      customer: { type: "business", withholdingAgent: true },
    },
    arrangements: [
      { id: "table", title: "Table", summary: "Labels in a column with the figures right-aligned beside them, so the bridge can be added down the page." },
      { id: "stacked", title: "Stacked", summary: "One labelled figure under another, for a narrow column." },
    ],
    requires: [],
  },
  {
    kind: "component",
    id: "invoice-change-explainer",
    title: "Invoice change explainer",
    summary: "Why this invoice differs from the last one, in one sentence, or nothing when it barely does.",
    detail:
      "Explains why this invoice differs from the last one, in one generated sentence, and says nothing when the difference is too small to matter. Lines are matched by a stable key, never by description. Each difference becomes one structured cause: a line added or removed, a price or quantity change, a discount starting or ending, a tax rate or tax base change, a longer or shorter period, or a one-off charge. Price and volume effects use one convention that adds up exactly to the line's change: volume is the change in quantity at the old price, and price is the rest. When usage per day is flat, a longer period is the cause, not 'usage rose'. It names the fewest causes that cover 90% of the gross movement, at most three, so offsetting changes cannot hide behind a small net. A balance brought forward is never part of the change. It is silent below a floor (the larger of an absolute amount and a share of the previous total), on a first invoice, after a currency change, and when either invoice is a credit note. The model is limited to the chosen causes and their figures and may not guess at reasons. The placeholder is the same explanation written deterministically. All of it is one deriver.",
    category: "Generated",
    tags: ["dynamic", "prompts", "ai", "invoice", "money", "deriver", "branching"],
    exports: ["InvoiceChangeExplainer"],
    dataFields: [
      { path: "invoiceChange.currency", type: "string", summary: "ISO 4217 code of this invoice. Sets the minor units and symbol. Required." },
      { path: "invoiceChange.current", type: '{ total: number; broughtForward?: number; periodDays?: number; documentType?: "invoice" | "credit-note"; lines: Array<{ key: string; description: string; kind: "recurring" | "usage" | "one-off" | "discount" | "tax"; quantity?: number; unitPrice?: number; amount: number; taxRate?: number }> }', summary: "This invoice. Lines must add up to total less broughtForward, or it throws. Keys are stable across invoices and unique within one." },
      { path: "invoiceChange.previous", type: "same shape plus currency?: string, or null", summary: "The invoice before. Missing means a first invoice and prints nothing; a credit note or a different currency also prints nothing." },
      { path: "invoiceChange.threshold", type: "{ absolute?: number; pct?: number }", summary: "Silent below max(absolute, pct × previous total). Defaults 1 and 0.02." },
      { path: "invoiceChange.periodLabel", type: "string", summary: "What the previous invoice is called after 'than'. Missing means 'your last invoice'." },
    ],
    files: node("invoice-change-explainer", { deriver: true }),
    previewData: {
      invoiceChange: {
        currency: "GBP",
        periodLabel: "last month",
        threshold: { absolute: 5, pct: 0.02 },
        previous: {
          total: 106320, periodDays: 30, documentType: "invoice",
          lines: [
            { key: "platform", description: "Enterprise Data Platform licence (annual commitment, billed monthly)", kind: "recurring", quantity: 1, unitPrice: 18500, amount: 18500 },
            { key: "api-calls", description: "Metered API requests over the included allowance", kind: "usage", quantity: 12400000, unitPrice: 0.00085, amount: 10540 },
            { key: "seats", description: "Named analyst seats with Extended Data Retention", kind: "recurring", quantity: 1180, unitPrice: 46, amount: 54280 },
            { key: "loyalty", description: "Multi-year loyalty discount (12% of platform licence)", kind: "discount", amount: -2220 },
            { key: "migration", description: "Data migration and onboarding services (one-off)", kind: "one-off", amount: 7500 },
            { key: "vat", description: "VAT", kind: "tax", taxRate: 0.2, amount: 17720 },
          ],
        },
        current: {
          total: 234360, broughtForward: 106320, periodDays: 31, documentType: "invoice",
          lines: [
            { key: "platform", description: "Enterprise Data Platform licence (annual commitment, billed monthly)", kind: "recurring", quantity: 1, unitPrice: 18500, amount: 18500 },
            { key: "api-calls", description: "Metered API requests over the included allowance", kind: "usage", quantity: 15500000, unitPrice: 0.00085, amount: 13175 },
            { key: "seats", description: "Named analyst seats with Extended Data Retention", kind: "recurring", quantity: 1450, unitPrice: 49.5, amount: 71775 },
            { key: "support", description: "Dedicated Support (Premier tier, 24×7 with named engineer)", kind: "recurring", quantity: 1, unitPrice: 3250, amount: 3250 },
            { key: "vat", description: "VAT", kind: "tax", taxRate: 0.2, amount: 21340 },
          ],
        },
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "renewal-and-price-change-notice",
    title: "Renewal and price change notice",
    summary: "The cancel-by date and what a renewal or price-change notice must say, measured from the date the law uses.",
    detail:
      "Works out the last day to cancel as the renewal date less the notice period, clamping month " +
      "arithmetic so 31 March less a month is the end of February, and prints it as 'on or before'. " +
      "Each law's window is measured from its own anchor: California's 2025 rules from the renewal " +
      "date, the trial end and the price change's effective date; New York §5-903 from the cancellation " +
      "deadline, for business service contracts as well as consumers. A late notice drops the deadline " +
      "and says what the law gives the customer, including that a late New York notice makes the " +
      "renewal unenforceable. A German consumer contract past its minimum term gets 'cancel at any time " +
      "with one month's notice' instead of a date, and a longer notice period is flagged. UK subscription " +
      "rules apply only to contracts started on or after a commencement date you can move. A price change " +
      "shows old and new amounts, the difference and the percentage to one decimal place. All of it is one deriver.",
    category: "Legal",
    tags: ["subscription", "renewal", "price-change", "legal", "consumer", "dates", "deriver", "branching", "section"],
    exports: ["RenewalAndPriceChangeNotice"],
    dataFields: [
      { path: "renewal.jurisdiction", type: '"US-CA" | "US-NY" | "DE" | "UK" | "other"', summary: "Whose rules apply. Other is contractual wording only." },
      { path: "renewal.customerType", type: '"consumer" | "business"', summary: "CA, DE and UK rules are consumer only; NY covers both." },
      { path: "renewal.contractStart", type: "string", summary: "ISO 8601 date the contract was entered. Required." },
      { path: "renewal.amendedOn", type: "string | undefined", summary: "Latest amendment or extension, read for California only." },
      { path: "renewal.initialTermMonths", type: "number", summary: "First term in whole months. Required." },
      { path: "renewal.renewalTermMonths", type: "number", summary: "Each renewal term in whole months; printed as the frequency. Required." },
      { path: "renewal.renewalDate", type: "string", summary: "ISO 8601 date the next term starts. Required." },
      { path: "renewal.noticePeriod", type: '{ value: number; unit: "days" | "months" }', summary: "Notice the customer must give. Zero means the day before renewal." },
      { path: "renewal.sendDate", type: "string", summary: "ISO 8601 date the notice is sent; every window is checked against it. Required." },
      { path: "renewal.currency", type: "string", summary: "ISO 4217 code. Required." },
      { path: "renewal.currentAmount", type: "number", summary: "Price per renewal term now, major units. Required." },
      { path: "renewal.previousAmount", type: "number | undefined", summary: "Last payment. Printed when given; a UK reminder warns without it." },
      { path: "renewal.priceChange", type: "{ newAmount: number; effectiveDate: string } | undefined", summary: "A price change. Missing or unchanged prints no price line." },
      { path: "renewal.trialEnds", type: "string | undefined", summary: "When a trial or introductory price ends. Missing or past prints nothing." },
      { path: "renewal.cancelInstructions", type: "string", summary: "How to cancel, as printed. Required." },
      { path: "renewal.commencementDate", type: "string | undefined", summary: "When UK subscription rules start. Missing means 2027-01-01." },
    ],
    arrangements: [
      { id: "stacked", title: "Stacked", summary: "One labelled line under another, down the text column." },
      { id: "table", title: "Table", summary: "Labels in a narrow column with the values beside them, like a key-facts box." },
    ],
    files: node("renewal-and-price-change-notice", { deriver: true, arrangements: ["stacked", "table"] }),
    previewData: {
      renewal: {
        jurisdiction: "US-NY",
        customerType: "business",
        contractStart: "2024-05-31",
        initialTermMonths: 36,
        renewalTermMonths: 12,
        renewalDate: "2027-05-31",
        noticePeriod: { value: 3, unit: "months" },
        sendDate: "2027-02-01",
        currency: "USD",
        currentAmount: 1284750.0,
        previousAmount: 1168000.0,
        priceChange: { newAmount: 1413225.0, effectiveDate: "2027-05-31" },
        cancelInstructions:
          "Write to Contract Renewals at the address on your master services agreement, or sign in to the customer portal, open Agreements, choose the maintenance agreement and select Cancel automatic renewal, at least by the date above",
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "commitment-drawdown",
    title: "Commitment drawdown",
    summary: "Prepaid commitment and credit drawn down line by line before tax, with what is left and any true-up.",
    detail:
      "Applies prepaid commitment and credit grants to each usage line in invoice " +
      "order, before tax, and splits the line where a balance runs out into a " +
      "covered part and an overage part, exact in minor units. A grant counts if " +
      "it was live on the usage date, not the invoice date, so usage from the last " +
      "days of a period still draws on a grant that expired before the invoice. " +
      "Rollover burns first, then paid commitment, then promotional credit, and " +
      "within each by priority then earliest expiry, so the credit closest to " +
      "lapsing is used first. Unit grants such as minutes cover only the prices " +
      "they name, and their overage is priced at the line's unit price. On the " +
      "invoice that ends the term, the minimum commitment is trued up against " +
      "eligible spend, which excludes promotional credit, less true-ups already " +
      "billed. Credits expiring soon are named, with what is projected to go " +
      "unused. The tax basis decides whether only the overage carries VAT or the " +
      "full usage does, with the credit deducted after tax. All of it is one deriver.",
    category: "Body",
    tags: ["money", "invoice", "subscription", "usage", "credits", "commitment", "tax", "table", "deriver", "branching", "section"],
    exports: ["CommitmentDrawdown"],
    dataFields: [
      { path: "drawdown.currency", type: "string", summary: "ISO 4217 code. Sets the minor units. Required." },
      { path: "drawdown.invoiceDate", type: "string", summary: "ISO 8601 issue date. Drives expiry warnings only. Required." },
      { path: "drawdown.period", type: "{ start: string; end: string }", summary: "ISO 8601 usage period, end inclusive. Required." },
      { path: "drawdown.term", type: "{ start: string; end: string; commitment: number; eligibleSpendToDate?: number; trueUpsBilled?: number } | undefined", summary: "Minimum-commitment term in minor units. Missing means no true-up." },
      { path: "drawdown.lines", type: "Array<{ key?: string; description: string; usageDate?: string; quantity: number; quantityDecimals?: number; unit?: string; unitPrice?: number; amount: number; priceId?: string }>", summary: "Priced usage lines in invoice order, amounts in integer minor units." },
      { path: "drawdown.grants", type: "Array<{ id: string; name: string; kind: \"rollover\" | \"commit\" | \"promotional\"; denomination: \"money\" | \"units\"; unit?: string; balance: number; priority?: number; effective: string; expires?: string; created?: string; priceIds?: string[]; dailyBurn?: number }>", summary: "Grants with a balance, any order. Units grants need priceIds." },
      { path: "drawdown.creditTaxBasis", type: "\"taxed-at-purchase\" | \"multi-purpose\"", summary: "How paid credit was taxed. Decides the VAT base. Required." },
      { path: "drawdown.warnDays", type: "number | undefined", summary: "Days ahead an expiring credit is warned about. Missing means 30." },
    ],
    files: node("commitment-drawdown", { deriver: true }),
    previewData: {
      drawdown: {
        currency: "USD",
        invoiceDate: "2026-07-02",
        period: { start: "2026-06-01", end: "2026-06-30" },
        term: { start: "2025-07-01", end: "2026-06-30", commitment: 300000000, eligibleSpendToDate: 198450000, trueUpsBilled: 2500000 },
        creditTaxBasis: "multi-purpose",
        warnDays: 30,
        lines: [
          { key: "L1", description: "Inference API: frontier model, input tokens (EU data residency)", usageDate: "2026-06-30", quantity: 18400000000, unit: "tokens", unitPrice: 0.0003, amount: 5520000, priceId: "price_frontier_input_eu" },
          { key: "L2", description: "Inference API: frontier model, output tokens (EU data residency)", usageDate: "2026-06-30", quantity: 2950000000, unit: "tokens", unitPrice: 0.0015, amount: 4425000, priceId: "price_frontier_output_eu" },
          { key: "L3", description: "Speech-to-text transcription, enhanced accuracy tier", usageDate: "2026-06-30", quantity: 1284750.5, unit: "minutes", unitPrice: 2.4, amount: 3083401, priceId: "price_stt_enhanced" },
          { key: "L4", description: "Vector storage, GB-months (multi-region replicated)", usageDate: "2026-06-30", quantity: 48250.75, unit: "GB-months", unitPrice: 23, amount: 1109767, priceId: "price_vector_storage" },
          { key: "L5", description: "Dedicated capacity reservation: 8 accelerator nodes, eu-west", usageDate: "2026-06-30", quantity: 720, unit: "hours", unitPrice: 98000, amount: 70560000, priceId: "price_dedicated_capacity" },
          { key: "L6", description: "Batch inference: frontier model, input tokens (batch discount)", usageDate: "2026-06-30", quantity: 9200000000, unit: "tokens", unitPrice: 0.00015, amount: 1380000, priceId: "price_frontier_batch_input" },
          { key: "L7", description: "Premium support: 24/7 enterprise response (monthly)", usageDate: "2026-06-30", quantity: 1, unitPrice: 4500000, amount: 4500000, priceId: "price_support_enterprise" },
        ],
        grants: [
          { id: "g-roll", name: "Rollover from the 2024–25 annual commitment", kind: "rollover", denomination: "money", balance: 12500000, priority: 1, effective: "2025-07-01", expires: "2026-07-01" },
          { id: "g-commit", name: "Annual prepaid commitment 2025–26 (enterprise platform agreement)", kind: "commit", denomination: "money", balance: 48000000, priority: 1, effective: "2025-07-01", expires: "2026-07-31", dailyBurn: 1600000 },
          { id: "g-minutes", name: "Prepaid transcription minutes bundle, enhanced tier", kind: "commit", denomination: "units", unit: "minutes", balance: 1000000, priority: 0, effective: "2025-07-01", expires: "2026-07-15", priceIds: ["price_stt_enhanced"], dailyBurn: 9500 },
          { id: "g-promo", name: "Launch partner promotional credit (frontier models)", kind: "promotional", denomination: "money", balance: 2500000, effective: "2026-01-01", expires: "2026-12-31", priceIds: ["price_frontier_input_eu", "price_frontier_output_eu"] },
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "timesheet-fee-summary",
    title: "Timesheet fee summary",
    summary: "Billed fees per timekeeper from raw time entries, rounded per entry and rated by date.",
    detail:
      "Each entry is rounded to the billing increment on its own (up or to nearest, " +
      "minimum one increment), because per-entry, per-day and per-total rounding " +
      "bill different hours and the client's guidelines name one: three 7-minute " +
      "entries are 0.60 hours, not 0.40. Each entry is rated by its date, so a rate " +
      "change mid-invoice splits the timekeeper into two rows instead of applying the " +
      "new rate backwards. A client-approved rate below standard never replaces the " +
      "rate; the row keeps the engagement-letter rate and a 'reduction per billing " +
      "guidelines' line shows the difference. Travel is its own row at the agreed " +
      "factor, no-charge time is listed at 0.00 with its hours, and a write-down is " +
      "its own line with its reason. Hours always print with two decimals and every " +
      "row's hours times rate equals its amount, counted in integer minor units. " +
      "All of it is one deriver.",
    category: "Body",
    tags: ["invoice", "legal", "time", "billing", "table", "deriver"],
    exports: ["TimesheetFeeSummary"],
    dataFields: [
      { path: "time.entries", type: 'Array<{ date: string; timekeeperId: string; minutes: number; kind?: "work" | "travel" | "no-charge" }>', summary: "Raw time entries, each rounded on its own. Empty prints 'No time was recorded'." },
      { path: "time.timekeepers", type: "Array<{ id: string; name: string; role: string; rates: Array<{ from: string; rate: number }>; approvedRate?: number }>", summary: "In print order; dated standard rates, and an optional client-approved rate." },
      { path: "time.billing.incrementMinutes", type: "number | undefined", summary: "Billing increment in minutes, a multiple of 3. Missing means 6 (0.1 hours)." },
      { path: "time.billing.rounding", type: '"up" | "nearest" | undefined', summary: "How each entry rounds to the increment. Missing means up." },
      { path: "time.billing.travelFactor", type: "number | undefined", summary: "Share of the rate travel bills at, 0 to 1. Missing means the full rate." },
      { path: "time.billing.writeDown", type: "{ amount: number; reason: string } | undefined", summary: "An agreed write-down, printed as its own line with its reason." },
      { path: "time.currency", type: "string | undefined", summary: "ISO 4217 code; sets the minor unit. Missing means GBP." },
    ],
    files: node("timesheet-fee-summary", { deriver: true }),
    previewData: {
      time: {
        currency: "GBP",
        billing: {
          incrementMinutes: 6,
          rounding: "up",
          travelFactor: 0.5,
          writeDown: { amount: 4750, reason: "Agreed reduction for duplicated research on the limitation point" },
        },
        timekeepers: [
          { id: "AMF", name: "Alexandra Montgomery-Fitzwilliam KC", role: "Partner, Head of Commercial Litigation",
            rates: [{ from: "2025-04-01", rate: 1150 }, { from: "2026-10-01", rate: 1295 }], approvedRate: 1200 },
          { id: "COH", name: "Chukwuemeka Okonkwo-Hargreaves", role: "Senior Associate (Solicitor-Advocate)",
            rates: [{ from: "2025-04-01", rate: 685 }] },
          { id: "PSL", name: "Priyanka Subramaniam-Lindqvist", role: "Trainee Solicitor",
            rates: [{ from: "2025-09-01", rate: 245 }] },
        ],
        entries: [
          { date: "2026-09-14", timekeeperId: "AMF", minutes: 547 },
          { date: "2026-09-29", timekeeperId: "AMF", minutes: 412 },
          { date: "2026-10-02", timekeeperId: "AMF", minutes: 583 },
          { date: "2026-10-02", timekeeperId: "AMF", minutes: 185, kind: "travel" },
          { date: "2026-09-15", timekeeperId: "COH", minutes: 598 },
          { date: "2026-09-16", timekeeperId: "COH", minutes: 601 },
          { date: "2026-10-02", timekeeperId: "COH", minutes: 185, kind: "travel" },
          { date: "2026-10-05", timekeeperId: "COH", minutes: 44, kind: "no-charge" },
          { date: "2026-09-17", timekeeperId: "PSL", minutes: 589 },
          { date: "2026-09-18", timekeeperId: "PSL", minutes: 7 },
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "disbursements-and-recharges",
    title: "Disbursements and recharges",
    summary: "Expenses split into recharges that carry your fees' VAT and disbursements passed on at cost outside it.",
    detail:
      "Splits a professional-services invoice's expenses into the two VAT treatments. " +
      "A recharge is a cost of your own service: it takes the main supply's VAT rate " +
      "on what you are out of pocket, net of input VAT you reclaimed and including tax " +
      "you could not, so a zero-rated fare or exempt postage recharged becomes " +
      "standard-rated. A disbursement is paid as the client's agent at exact cost, " +
      "outside the scope of VAT, and listed after the recharges' VAT total. An item " +
      "marked as paid as agent is moved to the recharges with a printed note, or " +
      "refused in strict mode, when it carries a markup, input VAT you reclaimed, or a " +
      "category HMRC treats as your own cost such as travel or electronic search fees. " +
      "Foreign receipts are converted at the rate for the expense date and print the " +
      "original amount, rate and converted figure. VAT is worked once on the " +
      "recharges' net, half away from zero. US invoices skip VAT and split on a " +
      "per-line taxable flag. All of it is one deriver.",
    category: "Body",
    tags: ["money", "invoice", "vat", "sales-tax", "expenses", "table", "deriver", "branching", "section"],
    exports: ["DisbursementsAndRecharges"],
    dataFields: [
      {
        path: "invoice.jurisdiction",
        type: '"UK" | "EU" | "DE" | "FR" | "US"',
        summary: "Whose rules apply. DE and FR read as EU; US skips VAT. Required with expenses.",
      },
      { path: "invoice.currency", type: "string", summary: "ISO 4217 code every figure is printed in. Required with expenses." },
      {
        path: "invoice.mainSupplyVatRate",
        type: "number | null",
        summary: "Fees' VAT rate as a percentage; every recharge takes it. null: fees carry no VAT, so neither do recharges.",
      },
      {
        path: "invoice.expenses",
        type:
          "Array<{ date: string; description: string; category?: string; net: number; inputVatReclaimed?: number; irrecoverableTax?: number; currency?: string; fxRate?: number; markupPercent?: number; paidAsAgent?: boolean; supplierName?: string; taxable?: boolean }>",
        summary: "Each expense in major units before tax on it. Missing or empty prints nothing.",
      },
    ],
    files: node("disbursements-and-recharges", { deriver: true }),
    previewData: {
      invoice: {
        jurisdiction: "UK",
        currency: "GBP",
        mainSupplyVatRate: 20,
        expenses: [
          { date: "2026-10-20", description: "Stamp Duty Land Tax on the purchase of Unit 4, Riverside Commercial Park", net: 487250, supplierName: "HM Revenue & Customs", paidAsAgent: true },
          { date: "2026-10-07", description: "Return rail fare, London Euston to Manchester Piccadilly, for the site inspection", category: "travel", net: 382.4, paidAsAgent: true },
          { date: "2026-10-02", description: "Court issue fee, claim form, Business and Property Courts", category: "court-fee", net: 10000, supplierName: "HM Courts & Tribunals Service", paidAsAgent: true },
          { date: "2026-10-08", description: "Hotel, two nights, Frankfurt am Main, for the witness interviews", category: "accommodation", currency: "EUR", net: 1250, irrecoverableTax: 87.5, fxRate: 0.8571 },
          { date: "2026-10-13", description: "Electronic local authority and drainage searches, 14 registered titles", category: "electronic-search", net: 2940, supplierName: "Northgate Property Searches Ltd", paidAsAgent: true },
          { date: "2026-10-06", description: "HM Land Registry official copies and title plans, 14 registered titles", category: "land-registry", net: 147, supplierName: "HM Land Registry", paidAsAgent: true },
          { date: "2026-10-15", description: "Printing and binding of eight lever-arch trial bundles", category: "printing", net: 1184.6, inputVatReclaimed: 236.92 },
          { date: "2026-10-21", description: "Counsel's fee for advice on the lease and rent review", net: 18750, supplierName: "Three Crowns Lane Chambers", paidAsAgent: true },
          { date: "2026-10-17", description: "Translation of German witness statements into English", currency: "EUR", net: 4800, fxRate: 0.8569, markupPercent: 10 },
        ],
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "fee-cap-status",
    title: "Fee cap status",
    summary: "What a capped, ceiling or estimated engagement may charge on this invoice, and what the invoice has to say about it.",
    detail:
      "A cap is clipped, not printed: a hard cap or ceiling charges only what " +
      "is left under it, and the excess is printed as fees above the cap, not " +
      "charged, so the client sees the value written off. A cap already used up " +
      "bills the period at nil. An estimate is not a cap: fees are charged in " +
      "full, and once they pass it the invoice prints the revised estimate and " +
      "when the client was told, or an 'incomplete' warning when nobody was, " +
      "because an untold overrun can be challenged. A ceiling past its notice " +
      "threshold (85% unless set) prints the percentage used and the notice. " +
      "Optional withholding stops at its limit. Four states, four ids: under, " +
      "approaching, reached, exceeded, all decided by one deriver.",
    category: "Body",
    tags: ["invoice", "fees", "cap", "estimate", "legal", "deriver", "branching", "section"],
    exports: ["FeeCapStatus"],
    dataFields: [
      { path: "feeCap.kind", type: '"hard" | "estimate" | "ceiling"', summary: "What kind of limit. Anything else fails the document." },
      { path: "feeCap.amount", type: "number", summary: "The cap, ceiling or original estimate. Required, above 0." },
      { path: "feeCap.billedToDate", type: "number", summary: "Charged on earlier invoices under this limit, before withholding. Required." },
      { path: "feeCap.noticeThreshold", type: "number | undefined", summary: "Percent of the limit that counts as approaching. Missing means 85 for a ceiling, none otherwise." },
      { path: "feeCap.revisedEstimate", type: "{ amount: number; notifiedOn: string } | undefined", summary: "Estimate only: the revised figure and the ISO date the client was told. Missing on an overrun prints a warning." },
      { path: "feeCap.withholding", type: "{ percent: number; limit: number; withheldToDate: number } | undefined", summary: "A retention that stops at its limit. Missing means nothing is withheld." },
      { path: "feeCap.currency", type: "string | undefined", summary: "ISO 4217 code. Defaults to GBP." },
      { path: "fees.thisPeriod", type: "number", summary: "Fees for this period, before the cap and withholding. Required, 0 or more." },
    ],
    arrangements: [
      { id: "table", title: "Table", summary: "Labels in a column with the amounts right-aligned beside them, so the figures check against each other." },
      { id: "stacked", title: "Stacked", summary: "One labelled figure under another, down the text column." },
    ],
    files: node("fee-cap-status", { deriver: true, arrangements: ["table", "stacked"] }),
    previewData: {
      feeCap: {
        kind: "ceiling",
        amount: 2500000,
        billedToDate: 2312480.55,
        noticeThreshold: 85,
        withholding: { percent: 10, limit: 250000, withheldToDate: 231248.06 },
        currency: "GBP",
      },
      fees: { thisPeriod: 418960.4 },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "final-invoice-settlement",
    title: "Final invoice settlement",
    summary: "The final invoice after advances: the whole supply by VAT rate, each advance deducted, and the balance.",
    detail:
      "A final invoice restates the whole supply, then takes off what the advance " +
      "invoices already covered. This one prints the full supply net and VAT per " +
      "rate band, deducts each advance on its own row with the net and VAT printed " +
      "on that advance invoice, never re-derived from its gross, and ends on the " +
      "balance, so VAT already declared is not charged twice. Only advances " +
      "received are deducted: one invoiced but unpaid is listed as not yet " +
      "received and stays in the amount due, except in the UK, where an issued " +
      "VAT invoice is already a tax point. When the rate changed in between, the " +
      "UK keeps each deposit at its own rate and prints two bands; Germany and " +
      "France tax the whole supply at the rate on completion and show the " +
      "re-rating as an 'of which' line. Advances that exceed the total print as " +
      "overpaid, to be refunded or credited, with a note that a credit note is " +
      "needed, never as a negative amount due. A small business prints no VAT " +
      "columns. Up to four rates and eight deducted advances; more fails the " +
      "document. All of it is one deriver, in integer minor units.",
    category: "Body",
    tags: ["invoice", "vat", "tax", "money", "advance-payment", "table", "deriver", "branching", "section"],
    exports: ["FinalInvoiceSettlement"],
    dataFields: [
      { path: "invoice.jurisdiction", type: '"UK" | "EU" | "DE" | "FR" | "US"', summary: "Whose rate-change and deduction rules apply. Anything else fails the document." },
      { path: "invoice.currency", type: "string", summary: "ISO 4217 code. Sets the minor units. Required." },
      { path: "invoice.lines", type: "Array<{ net: number; vatRate?: number }>", summary: "The whole supply in integer minor units, each line at its rate on completion. Required." },
      { path: "invoice.advances", type: "Array<{ number: string; issueDate: string; net: number; vat?: number; vatRate?: number; received: boolean; supplyRate?: number }> | undefined", summary: "Every advance invoiced, amounts exactly as printed on it, in minor units. Missing means none." },
      { path: "seller.smallBusinessExempt", type: "boolean | undefined", summary: "True for a seller who charges no VAT: no VAT columns. Missing means VAT is shown." },
    ],
    files: node("final-invoice-settlement", { deriver: true }),
    previewData: {
      invoice: {
        jurisdiction: "DE",
        currency: "EUR",
        lines: [
          { net: 487654321, vatRate: 16 },
          { net: 12345678, vatRate: 5 },
        ],
        advances: [
          { number: "AR-2020-000147/BAU-NORD-01", issueDate: "2020-05-29", net: 150000000, vat: 28500000, vatRate: 19, received: true, supplyRate: 16 },
          { number: "AR-2020-000163/BAU-NORD-02", issueDate: "2020-08-31", net: 120000000, vat: 19200000, vatRate: 16, received: true },
          { number: "AR-2020-000171/BAU-NORD-03", issueDate: "2020-09-30", net: 80000000, vat: 12800000, vatRate: 16, received: false },
        ],
      },
      seller: { smallBusinessExempt: false },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "advance-payment-invoice",
    title: "Advance payment invoice",
    summary: "A deposit or prepayment against an order: the share requested, what's already been asked for, and what's left.",
    detail:
      "A deposit or prepayment invoice against an order: the order total, what has already been requested, the share requested now and what is left to invoice. A percentage is taken of the order gross, never of the net plus VAT, because the customer expects the figure on the quote and the two routes differ by a cent. On a mixed-rate order the deposit is spread across the rates by largest remainder in minor units, so the shares add up to the deposit and deposits plus balance add up to the order, exactly. Requesting more than 100% of the order fails the document. UK and EU prepayments are charged VAT now, because the payment is a tax point, and the date of payment prints when it differs from the issue date. Germany says it is an Anzahlungsrechnung. A US deposit carries no sales tax until the final invoice, and a refundable security deposit carries no VAT and says so. Every figure comes from one deriver.",
    category: "Body",
    tags: ["invoice", "deposit", "vat", "tax", "deriver", "branching", "section"],
    exports: ["AdvancePaymentInvoice"],
    dataFields: [
      { path: "order.reference", type: "string", summary: "The order the deposit is against. Required." },
      { path: "order.date", type: "string | undefined", summary: "ISO 8601 order date. Printed when given." },
      { path: "order.currency", type: "string", summary: "ISO 4217 code. Required; sets the decimal places." },
      { path: "order.lines", type: "Array<{ gross?: number; net?: number; vatRate?: number }>", summary: "Enough of the order to total it per rate. One of gross or net per line. Required, non-empty." },
      { path: "deposit.percent", type: "number | undefined", summary: "Share of the order gross, over 0 and up to 100. Give this or amount." },
      { path: "deposit.amount", type: "number | undefined", summary: "A fixed gross amount. Give this or percent." },
      { path: "deposit.kind", type: '"prepayment" | "security"', summary: "Part of the price, or a refundable security deposit with no VAT. Required." },
      { path: "deposit.receivedDate", type: "string | undefined", summary: "ISO 8601. Printed on a prepayment when it differs from the issue date." },
      { path: "previousDeposits", type: "Array<{ number: string; amount: number }> | undefined", summary: "Prepayments already requested. Over 100% of the order fails the document." },
      { path: "invoice.jurisdiction", type: '"UK" | "EU" | "DE" | "FR" | "US"', summary: "Whose tax wording applies. US charges no tax on the deposit." },
      { path: "invoice.issueDate", type: "string", summary: "ISO 8601 date of issue. Required; compared with the payment date." },
    ],
    arrangements: [
      { id: "table", title: "Table", summary: "Labels in a column with the figures right-aligned beside them, so the sum can be checked." },
      { id: "stacked", title: "Stacked", summary: "One labelled line under another, down the text column." },
    ],
    files: node("advance-payment-invoice", { deriver: true, arrangements: ["table", "stacked"] }),
    previewData: {
      order: {
        reference: "SO-2026-11-004817/FACILITIES-NORTH-REFIT/REV-3",
        date: "2026-10-14",
        currency: "GBP",
        lines: [
          { gross: 1284750.0, vatRate: 20 },
          { net: 412333.33, vatRate: 5 },
          { gross: 96410.07, vatRate: 0 },
        ],
      },
      deposit: { percent: 30, kind: "prepayment", receivedDate: "2026-11-28" },
      previousDeposits: [
        { number: "ADV-2026-11-000417-A", amount: 250000.0 },
        { number: "ADV-2026-11-000418-B", amount: 125000.0 },
      ],
      invoice: { jurisdiction: "UK", issueDate: "2026-11-30" },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "credit-note-totals",
    title: "Credit note totals",
    summary: "What a credit note credits, by VAT band, worked from the original invoice so it nets to zero.",
    detail:
      "Credits are worked from the original invoice, never from today's settings. " +
      "Each credited line keeps the VAT rate the invoice charged, and a foreign-currency " +
      "invoice keeps its sterling exchange rate, so a rate change since cannot credit the " +
      "wrong VAT. A full cancellation copies the original band figures exactly rather than " +
      "recomputing them, so the two documents net to zero. A partial return credits quantity " +
      "times the original unit price, with VAT rounded per band. A goodwill amount is spread " +
      "across the original VAT bands pro rata by largest remainder, so no band is credited " +
      "more VAT than it charged. Earlier credit notes are subtracted first, and a credit " +
      "larger than what remains in any band fails the document. Amounts print positive under " +
      "a Credit heading, as Peppol 381 expects, never as negative invoice lines.",
    category: "Body",
    tags: ["credit-note", "invoice", "vat", "totals", "deriver", "table", "section"],
    exports: ["CreditNoteTotals"],
    dataFields: [
      { path: "creditNote.original.number", type: "string", summary: "The invoice credited. Required." },
      { path: "creditNote.original.issueDate", type: "string", summary: "ISO 8601 issue date of that invoice. Required." },
      { path: "creditNote.original.currency", type: "string", summary: "ISO 4217 code of that invoice. The credit is in the same currency. Required." },
      { path: "creditNote.original.fxToGBP", type: "number | undefined", summary: "Pounds per 1 unit of currency, as on the original. Given: VAT is also stated in sterling." },
      { path: "creditNote.original.vatGBP", type: "number | undefined", summary: "The original's sterling VAT total, copied on a first full cancellation." },
      {
        path: "creditNote.original.bands",
        type: "Array<{ vatRate: number | null; label?: string; net: number; vat: number }>",
        summary: "The original VAT analysis, in printed order. Required.",
      },
      {
        path: "creditNote.original.lines",
        type: "Array<{ id: string; description: string; qty: number; unitNet: number; vatRate: number | null }> | undefined",
        summary: "The original lines. Required for a lines credit.",
      },
      {
        path: "creditNote.previousCredits",
        type: "Array<{ bands: Array<{ vatRate: number | null; net: number; vat: number }> }> | undefined",
        summary: "Earlier credit notes on this invoice, subtracted before the limit check.",
      },
      { path: "creditNote.credit.mode", type: '"full" | "lines" | "amount"', summary: "Cancel the remainder, return lines, or credit a lump sum." },
      { path: "creditNote.credit.lines", type: "Array<{ id: string; qty: number }> | undefined", summary: "What is returned, for a lines credit." },
      { path: "creditNote.credit.amount", type: "number | undefined", summary: "The VAT-inclusive sum, for an amount credit." },
    ],
    files: node("credit-note-totals", { deriver: true }),
    previewData: {
      creditNote: {
        original: {
          number: "SI-2026-09-004817/REV-2",
          issueDate: "2026-09-30",
          currency: "EUR",
          fxToGBP: 0.8567,
          vatGBP: 223790.6,
          bands: [
            { vatRate: 20, net: 1283650.0, vat: 256730.0 },
            { vatRate: 5, net: 89880.0, vat: 4494.0 },
            { vatRate: 0, net: 18500.0, vat: 0 },
          ],
          lines: [
            { id: "L1", description: "Annual enterprise licence, Docxcelerate Document Engine, 2,500 seats, unlimited recipients", qty: 2500, unitNet: 412.75, vatRate: 20 },
            { id: "L2", description: "Implementation and data migration services, phase two (statement and arrears templates)", qty: 186.5, unitNet: 1350, vatRate: 20 },
            { id: "L3", description: "Printed welcome packs on recycled board, posted to every registered recipient", qty: 48000, unitNet: 1.8725, vatRate: 5 },
            { id: "L4", description: "Accessibility audit of published templates by an accredited third party", qty: 1, unitNet: 18500, vatRate: 0 },
          ],
        },
        previousCredits: [{ bands: [{ vatRate: 20, net: 41275.0, vat: 8255.0 }] }],
        credit: { mode: "lines", lines: [{ id: "L1", qty: 300 }, { id: "L3", qty: 12000 }, { id: "L2", qty: 12.5 }] },
      },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "construction-payment-timetable",
    title: "Construction Act payment timetable",
    summary: "The Construction Act payment dates for one interim payment, counted the way section 116 counts them.",
    detail:
      "Works out when an interim payment falls due, the payment notice deadline five " +
      "days later, the final date for payment, and the pay less deadline counted back " +
      "from that final date, never forward from the due date. Days are counted under " +
      "section 116: Christmas Day, Good Friday and the contract region's own bank " +
      "holidays are skipped, weekends are not, so a Scottish job never uses " +
      "England's list. With no valid payment notice, the payee's claim stands as the " +
      "notified sum, or the payee's default notice does and moves the final date back " +
      "by its delay. A contract with a residential occupier gets one sentence saying " +
      "the Act does not apply. All of it is one deriver.",
    category: "Body",
    tags: ["construction", "payment", "legal", "dates", "uk", "deriver"],
    exports: ["ConstructionPaymentTimetable"],
    dataFields: [
      { path: "contract.region", type: '"EW" | "SC" | "NI"', summary: "Whose Act and whose bank holidays. Required." },
      { path: "contract.residentialOccupier", type: "boolean | undefined", summary: "True prints only that the Act does not apply." },
      { path: "contract.dueDateDays", type: "number | undefined", summary: "Days after the valuation period that payment falls due. Missing means the Scheme." },
      { path: "contract.finalDateDays", type: "number | undefined", summary: "Days after the due date of the final date for payment. Missing means 17." },
      { path: "contract.payLessDays", type: "number | undefined", summary: "Days before the final date a pay less notice is due. Missing means 7." },
      { path: "payment.valuationPeriodEnd", type: "string", summary: "ISO 8601 last day of the valuation period. Required." },
      { path: "payment.claimReceived", type: "string | undefined", summary: "ISO 8601 date the payment claim arrived. Required under the Scheme." },
      { path: "payment.paymentNoticeIssued", type: "string | undefined", summary: "ISO 8601 date of the payer's notice. Late counts as none." },
      { path: "payment.defaultNoticeIssued", type: "string | undefined", summary: "ISO 8601 date of the payee's default notice. Moves the final date back." },
      { path: "bankHolidays", type: "{ EW?: string[]; SC?: string[]; NI?: string[] }", summary: "Bank holidays by region, covering every year the timetable runs through." },
    ],
    arrangements: [
      { id: "table", title: "Table", summary: "The four dates in a label and date table, for a timetable read at a glance." },
      { id: "stacked", title: "Stacked", summary: "One line per date, for a letter that reads as prose." },
    ],
    files: node("construction-payment-timetable", { deriver: true, arrangements: ["table", "stacked"] }),
    previewData: {
      contract: { region: "SC", dueDateDays: 14, finalDateDays: 30, payLessDays: 7 },
      payment: {
        valuationPeriodEnd: "2026-12-04",
        claimReceived: "2027-01-12",
        paymentNoticeIssued: "2027-01-06",
        defaultNoticeIssued: "2027-01-14",
      },
      bankHolidays: {
        SC: [
          "2026-01-01", "2026-01-02", "2026-04-03", "2026-05-04", "2026-05-25", "2026-08-03", "2026-11-30", "2026-12-25", "2026-12-28",
          "2027-01-01", "2027-01-04", "2027-03-26", "2027-05-03", "2027-05-31", "2027-08-02", "2027-11-30", "2027-12-27", "2027-12-28",
        ],
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "cis-invoice-deduction",
    title: "CIS invoice deduction",
    summary: "What a construction subcontractor is paid on an invoice, after CIS and VAT.",
    detail:
      "CIS comes off the labour only: the ex-VAT gross less materials at cost, " +
      "plant hire and consumables, with any markup left in labour. VAT is worked " +
      "out on the whole gross and is never reduced by CIS. The rate follows HMRC " +
      "verification: nothing under gross payment status, 20% registered, 30% " +
      "unverified with a warning to verify. The deduction is floored to the " +
      "penny, so it never over-deducts. The domestic reverse charge applies when " +
      "both parties are VAT-registered, the customer is CIS-registered and has not " +
      "notified end-user or intermediary status: silence means reverse charge. " +
      "The 5% disregard and zero-rated work stay standard. Under the reverse " +
      "charge no VAT is added, and the notional VAT is returned for " +
      "tax-treatment-notice. Unclear data fails the document instead of guessing.",
    category: "Body",
    tags: ["invoice", "cis", "construction", "vat", "tax", "deriver", "table"],
    exports: ["CisInvoiceDeduction"],
    dataFields: [
      { path: "invoice.grossExVat", type: "number", summary: "Invoice value before VAT, materials included. Required." },
      { path: "invoice.materialsCost", type: "number", summary: "Materials at cost, net of reclaimable VAT. Required; 0 when none." },
      { path: "invoice.plantHire", type: "number | undefined", summary: "Plant hire at cost. Missing means 0." },
      { path: "invoice.consumables", type: "number | undefined", summary: "Consumables and fuel at cost. Missing means 0." },
      { path: "invoice.vatRate", type: "0.2 | 0.05 | 0", summary: "The supply's VAT rate. Zero-rated is never reverse-charged." },
      { path: "invoice.reverseChargeShare", type: "number | undefined", summary: "Fraction that is reverse-charge services. At 5% or less the whole invoice is standard. Missing means 1." },
      { path: "supplier.vatRegistered", type: "boolean", summary: "Whether the subcontractor charges VAT. Required." },
      { path: "supplier.cisStatus", type: '"gross" | "net" | "unverified"', summary: "HMRC verification: 0%, 20% or 30%. Anything else fails the document." },
      { path: "supplier.verificationNumber", type: "string | undefined", summary: "HMRC verification number, printed when given." },
      { path: "customer.vatRegistered", type: "boolean", summary: "Required when VAT is charged at 20% or 5%." },
      { path: "customer.cisRegistered", type: "boolean", summary: "Required when VAT is charged at 20% or 5%." },
      { path: "customer.endUserNotified", type: '"end-user" | "intermediary" | undefined', summary: "Written notice that takes the invoice out of the reverse charge. Missing means reverse charge." },
    ],
    files: node("cis-invoice-deduction", { deriver: true, arrangements: ["table", "stacked"] }),
    previewData: {
      invoice: {
        grossExVat: 486250.75,
        materialsCost: 172430.18,
        plantHire: 38615.4,
        consumables: 9872.65,
        vatRate: 0.2,
      },
      supplier: { vatRegistered: true, cisStatus: "unverified", verificationNumber: "V1234567890AB" },
      customer: { vatRegistered: true, cisRegistered: true },
    },
    arrangements: [
      { id: "table", title: "Table", summary: "Labels in a column with the figures right-aligned beside them, so the sums can be checked." },
      { id: "stacked", title: "Stacked", summary: "One labelled figure under another, for a narrow column." },
    ],
    requires: [],
  },

  {
    kind: "component",
    id: "us-sales-tax-breakdown",
    title: "US sales tax breakdown",
    summary: "Sales tax split into state, county, city and district rows that add up to the tax charged.",
    detail:
      "Splits the sales tax into one row per jurisdiction, state first, each on its " +
      "own taxable amount, because a city can exempt clothing the state taxes and a " +
      "county can cap the taxable price of a single item. Each row is rounded on its " +
      "own base and the cent left over against the total is put on the state row, so " +
      "the rows always add up to the tax charged, which multiplying a combined rate " +
      "by the subtotal does not give you. Rates are decimal strings worked in whole " +
      "numbers, never floats. Shipping is taxed only where the jurisdiction taxes " +
      "delivery and only for the taxable share of the order, and a charge sent as " +
      "shipping and handling together prints a warning, because that makes delivery " +
      "taxable where it would not be. Exempt lines are listed under their reason. " +
      "When a marketplace facilitator collected the tax, the rows are replaced by a " +
      "statement saying so. Everything is worked out by a deriver that ships with it.",
    category: "Body",
    tags: ["tax", "sales-tax", "us", "invoice", "table", "deriver"],
    exports: ["UsSalesTaxBreakdown"],
    dataFields: [
      { path: "salesTax.currency", type: "string | undefined", summary: "ISO 4217 code for every amount. Defaults to USD." },
      {
        path: "salesTax.lines",
        type:
          "Array<{ description: string; amountMinor: number; taxable: boolean; category?: string; exemptReason?: string }> | undefined",
        summary: "Net line prices in cents. Taxable is required; non-taxable lines are listed under their reason.",
      },
      {
        path: "salesTax.shipping",
        type: "{ amountMinor: number; separatelyStated?: boolean; combinedWithHandling?: boolean } | undefined",
        summary: "The delivery charge. Combined with handling, it warns in states that exempt delivery stated on its own.",
      },
      {
        path: "salesTax.jurisdictions",
        type:
          '{ level: "state" | "county" | "city" | "special"; name: string; rate: string; taxesShipping?: boolean; exemptCategories?: string[]; perItemCapMinor?: number }[] | undefined',
        summary: 'Every taxing jurisdiction, exactly one state. Rates are decimal strings such as "0.0625".',
      },
      {
        path: "salesTax.collectedBy",
        type: '"seller" | "marketplace" | undefined',
        summary: "Marketplace replaces the rows with the facilitator statement. Defaults to seller.",
      },
      {
        path: "salesTax.marketplaceName",
        type: "string | undefined",
        summary: "Named in the facilitator statement when the marketplace collected the tax.",
      },
    ],
    files: [...node("us-sales-tax-breakdown", { deriver: true })],
    previewData: {
      salesTax: {
        currency: "USD",
        lines: [
          {
            description: "Ride-on industrial floor scrubber, 40-inch cleaning path, lithium-ion battery (x12)",
            amountMinor: 148750000, taxable: true, category: "manufacturing-machinery",
          },
          { description: "Replacement squeegee blades and brush decks, assorted", amountMinor: 2486475, taxable: true },
          {
            description: "Extended service agreement, 60 months", amountMinor: 18750000, taxable: false,
            exemptReason: "Optional service contract, not tangible personal property",
          },
        ],
        shipping: { amountMinor: 2485000, separatelyStated: true, combinedWithHandling: true },
        jurisdictions: [
          { level: "state", name: "Colorado", rate: "0.029", exemptCategories: ["manufacturing-machinery"] },
          { level: "county", name: "Arapahoe County", rate: "0.0025", exemptCategories: ["manufacturing-machinery"] },
          { level: "city", name: "City of Greenwood Village", rate: "0.03" },
          { level: "special", name: "Regional Transportation District", rate: "0.01", exemptCategories: ["manufacturing-machinery"] },
          { level: "special", name: "Scientific and Cultural Facilities District", rate: "0.001", exemptCategories: ["manufacturing-machinery"] },
        ],
        collectedBy: "seller",
      },
    },
    requires: [],
  },
  {
    kind: "component",
    id: "vat-place-of-supply",
    title: "VAT place of supply",
    summary: "Decides whose VAT applies to a sale and at what rate, and prints the one line that says so.",
    detail:
      "Decides the VAT treatment before anything is printed: domestic, reverse " +
      "charge, intra-Community supply, zero-rated export, outside the scope, or " +
      "the customer's country's rate under the OSS. The EUR 10,000 distance-" +
      "selling threshold is tested EU-wide on the current and the preceding " +
      "year, and either one over means destination VAT. A VAT number only " +
      "counts once it is marked validated, so an unproven business claim is " +
      "charged VAT, with a note saying why, and a consumer never gets reverse " +
      "charge. Northern Ireland counts as the EU for goods. Rates come from your " +
      "data, never from a table that goes stale. The decision is made by a " +
      "deriver that ships with it, and its treatment and rate are readable as " +
      "derived values for the totals.",
    category: "Legal",
    tags: ["tax", "vat", "invoice", "legal", "oss", "deriver"],
    exports: ["VatPlaceOfSupply"],
    dataFields: [
      { path: "seller.country", type: "string", summary: "Where the seller is established: an EU member state, GB or XI. Anything else throws." },
      { path: "seller.ossRegistered", type: "boolean | undefined", summary: "True when the seller charges destination VAT under the Union OSS even under the threshold." },
      { path: "seller.euDistanceSales", type: "{ currentYear: number; priorYear: number } | undefined", summary: "EU-wide cross-border B2C sales in EUR, this year and last. Required for a B2C EU distance sale unless OSS-registered." },
      { path: "customer.country", type: "string", summary: "Where the customer is, as an ISO country code. Outside the EU and UK means export or outside the scope." },
      { path: "customer.isBusiness", type: "boolean", summary: "Whether the customer buys as a business. Missing means a consumer." },
      { path: "customer.vatNumber", type: "string | undefined", summary: "Needed with vatNumberValidated for reverse charge or an intra-Community supply." },
      { path: "customer.vatNumberValidated", type: "boolean | undefined", summary: "True only once the number has been checked. Anything else means VAT is charged." },
      { path: "supply.kind", type: '"goods" | "services" | "digital"', summary: "What is sold. Digital means telecoms, broadcasting and electronic services." },
      { path: "supply.shipsFrom", type: "string | undefined", summary: "Goods only: where they leave from. Defaults to the seller's country." },
      { path: "supply.shipsTo", type: "string | undefined", summary: "Goods only: where they arrive. Defaults to the customer's country." },
      { path: "vatRates", type: "Record<string, number>", summary: "Current standard rate per country, as a percentage. Only the rate needed is read; a missing one throws." },
    ],
    files: [...node("vat-place-of-supply", { deriver: true })],
    previewData: {
      seller: { country: "DE", ossRegistered: false, euDistanceSales: { currentYear: 4850000, priorYear: 3920000 } },
      customer: { country: "FI", isBusiness: true, vatNumber: "FI12345678", vatNumberValidated: false },
      supply: { kind: "digital", shipsFrom: "DE", shipsTo: "FI" },
      vatRates: { DE: 19, FI: 25.5, NL: 21 },
    },
    requires: [],
  },

  {
    kind: "component",
    id: "india-gst-split",
    title: "India GST split",
    summary: "CGST and SGST, or IGST, decided from the place of supply against the GSTIN's state code, never the billing address.",
    detail:
      "Picks the tax head from the place of supply against the state code in the " +
      "supplier's GSTIN, never from the billing address or a typed state name: the " +
      "same state gets CGST plus SGST (UTGST in a Union Territory without a " +
      "legislature) at half the rate each, a different state gets IGST at the full " +
      "rate. Supplies to an SEZ and exports are always inter-state; under a letter " +
      "of undertaking they carry IGST at nil and the Rule 46 mention. Each half is " +
      "rounded to the paisa on its own, so an odd paisa can make CGST plus SGST " +
      "differ from IGST on the same base, which is correct. Prints place of supply " +
      "as 'State (code)', the mandatory 'Reverse charge: Yes/No', an HSN/SAC-wise " +
      "summary, the round-off to the nearest rupee, and the total in words in lakh " +
      "and crore. Under reverse charge the tax is shown but not added to the amount " +
      "payable. All of it is worked out by a deriver that ships with it.",
    category: "Body",
    tags: ["tax", "gst", "india", "invoice", "table", "branching", "deriver"],
    exports: ["IndiaGstSplit"],
    dataFields: [
      { path: "seller.gstin", type: "string", summary: "The supplier's GSTIN. Its first two digits are the supplier's state, the only one believed." },
      { path: "seller.stateCode", type: "string | undefined", summary: "Optional cross-check. Differing from the GSTIN's prefix throws." },
      { path: "placeOfSupply.stateCode", type: "string", summary: "Two-digit GST state code, or 96 for an export. Decides the tax head with the GSTIN prefix." },
      { path: "placeOfSupply.stateName", type: "string", summary: "Printed as 'Name (code)'. Never used to decide anything." },
      { path: "customer.isSez", type: "boolean | undefined", summary: "A supply to an SEZ unit or developer, always inter-state. Missing means no." },
      { path: "supply.exportUnderLut", type: "boolean | undefined", summary: "An export or SEZ supply under a letter of undertaking: IGST at nil with the Rule 46 mention." },
      { path: "supply.reverseCharge", type: "boolean", summary: "Printed as 'Reverse charge: Yes/No'. Required. Yes leaves the tax out of the amount payable." },
      { path: "lines", type: "Array<{ hsnSac: string; taxableMinor: number; rate: number }>", summary: "Taxable value in paise and the full GST rate per line, summed per HSN/SAC and rate." },
    ],
    files: node("india-gst-split", { deriver: true }),
    previewData: {
      seller: { gstin: "26AAACR5055K1Z5", stateCode: "26" },
      placeOfSupply: { stateCode: "26", stateName: "Dadra and Nagar Haveli and Daman and Diu" },
      customer: { isSez: false },
      supply: { exportUnderLut: false, reverseCharge: false },
      lines: [
        { hsnSac: "84714190", taxableMinor: 12345678950, rate: 18 },
        { hsnSac: "998314", taxableMinor: 4567890125, rate: 18 },
        { hsnSac: "84714190", taxableMinor: 9999999, rate: 18 },
        { hsnSac: "30049099", taxableMinor: 234567891, rate: 5 },
        { hsnSac: "71023910", taxableMinor: 9876543, rate: 0.25 },
      ],
    },
    requires: [],
  },
];

/** Kept in this order everywhere the catalog is listed. */
export const COMPONENT_CATEGORIES: ComponentCategory[] = [
  "Opening",
  "Body",
  "Generated",
  "Closing",
  "Legal",
];
