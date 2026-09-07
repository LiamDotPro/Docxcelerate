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
 */
function node(id: string): RegistryFile[] {
  return [{ source: `components/${id}/${id}.node.tsx`, target: `nodes/${id}.node.tsx` }];
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
    tags: ["money", "branching", "section"],
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
];

/** Kept in this order everywhere the catalog is listed. */
export const COMPONENT_CATEGORIES: ComponentCategory[] = [
  "Opening",
  "Body",
  "Generated",
  "Closing",
  "Legal",
];
