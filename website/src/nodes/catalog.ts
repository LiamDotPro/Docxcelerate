/**
 * The node catalog — one description of every node type.
 *
 * scripts/build-node-previews.mjs walks this file, resolves each variant
 * through the real framework, and writes the previews plus
 * src/generated/node-catalog.json. The docs pages and the homepage read that
 * JSON, so nothing about a node type is written down twice.
 *
 * Adding a type: a directory under src/nodes/ with one file per variant, an
 * entry here, and — for a shipped type — an MDX page that names it.
 */
import type { Component } from "docxcelerate/template";
import type { SampleData } from "./sample-data.ts";

import { Greeting } from "./paragraph/static.node.tsx";
import { PriceChange } from "./paragraph/conditional.node.tsx";
import { NextSteps } from "./paragraph/dynamic.node.tsx";
import { Apology } from "./paragraph/prompted.node.tsx";
import { Opening } from "./section/basic.node.tsx";
import { YourYear } from "./section/nested.node.tsx";
import { Signature } from "./image/static.node.tsx";
import { CentrePhoto } from "./image/dynamic.node.tsx";
import { VisitsByMonth } from "./graph/bar.node.tsx";
import { CumulativeVisits } from "./graph/line.node.tsx";
import { ClassMix } from "./graph/pie.node.tsx";
import { PeakTimes } from "./graph/dynamic.node.tsx";
import { VisitLog } from "./table/basic.node.tsx";
import { PriceSummary } from "./table/totals.node.tsx";
import { Contents } from "./table-of-contents/basic.node.tsx";

/** Where a node type stands: shipped, authorable only by hand, or not yet built. */
export type NodeStatus = "stable" | "no-helper" | "planned";

export type NodeCategory = "Structure" | "Text" | "Media" | "Data";

export interface NodeOption {
  name: string;
  type: string;
  required?: boolean;
  summary: string;
}

export interface NodeVariant {
  /**
   * Doubles as the file name (src/nodes/<type>/<id>.node.tsx) and the preview
   * name (public/demo/nodes/<type>/<id>.html). The build script checks the
   * source file exists, so a rename that misses one half fails the build.
   */
  id: string;
  title: string;
  summary: string;
  component: Component;
}

export interface NodeTypeEntry {
  /** Directory under src/nodes/, and the docs slug under /docs/nodes/. */
  id: string;
  title: string;
  /** The NodeKind it resolves to in a DocumentModel. */
  kind: string;
  category: NodeCategory;
  status: NodeStatus;
  /** Elements and hooks exported from `docxcelerate/template`. */
  helpers: string[];
  /** One line, for lists and cards. */
  summary: string;
  /** A paragraph, for the type's own page and the homepage disclosure. */
  detail: string;
  /** What this node may contain. */
  children: string;
  /** Where its content is decided. */
  resolves: "Locally" | "At request time" | "Both" | "By the renderer";
  /**
   * Set where the shipped renderers do less than the node describes. Node kinds
   * land in the schema before the renderers catch up, and a docs page that
   * showed the model without saying so would be selling something that is not
   * there yet.
   */
  renderNote?: string;
  options: NodeOption[];
  variants: NodeVariant[];
}

/** Every node option set accepts this, so it is listed once rather than seven times. */
const DERIVERS: NodeOption = {
  name: "derivers",
  type: "DeriverInvocation[]",
  summary:
    "Values the engine computes before the node resolves, written to " +
    "`derived.*` and readable from a template token. Built with `derive()`. " +
    "These survive publishing and run per document — use them for anything " +
    "computed from request data. `useDeriver` runs one during the build instead.",
};

const VARIANT: NodeOption = {
  name: "variant",
  type: "string",
  summary:
    "A block style the theme looks up — `\"band\"`, `\"badge\"`, `\"panel\"`. Names " +
    "what the node is, never what it looks like: the appearance lives in the " +
    "style's `blocks`, so a document restyles without a node changing. A name " +
    "the theme has not heard of draws as an ordinary block rather than failing.",
};

const ID: NodeOption = {
  name: "id",
  type: "string",
  summary:
    "Stable address for the node. Engines target it and build artifacts diff " +
    "on it, so treat a rename as a breaking change. Optional: a node without " +
    "one takes an id from where it sits, which is what keeps branches and " +
    "loops from forcing you to invent names. Two nodes claiming one id is an " +
    "error rather than a race.",
};

/**
 * Prompts, settable either way.
 *
 * As props they sit on the element. Through `useSetPrompts` they are set by the
 * component, which is what lets a shared hook add house style to a node it does
 * not own. Props win, so a caller can override what a hook set.
 */
const PROMPT_OPTIONS: NodeOption[] = [
  {
    name: "placeholder",
    type: "string",
    summary:
      "What previews show in place of generated content. Also settable with " +
      "`useSetPlaceholders`. Optional, but a document that reads badly without " +
      "one cannot be reviewed.",
  },
  {
    name: "generalPrompt",
    type: "string",
    required: true,
    summary: "What this node should say.",
  },
  {
    name: "infoPrompt",
    type: "string",
    summary: "Context the model should have but should not restate.",
  },
  {
    name: "negativePrompt",
    type: "string",
    summary: "What to avoid — claims, tones, or facts it must not invent.",
  },
  {
    name: "systemPrompt",
    type: "string",
    summary: "Role and voice, applied ahead of the other prompts.",
  },
  {
    name: "examplePrompt",
    type: "string",
    summary:
      "What a good answer looks like, written out as finished text. Shown last, " +
      "because it is what the answer gets measured against.",
  },
];

export const NODE_TYPES: NodeTypeEntry[] = [
  {
    id: "section",
    title: "Section",
    kind: "section",
    category: "Structure",
    status: "stable",
    helpers: ["Section"],
    summary: "Groups nodes under a titled heading.",
    detail:
      "The only construct that nests today. Its title carries into the document " +
      "outline, so the structure you write is the structure the reader sees.",
    children: "Any node, including other sections. No depth limit.",
    resolves: "Locally",
    options: [
      ID,
      {
        name: "title",
        type: "string",
        required: true,
        summary: "The heading printed above the children, and the outline entry.",
      },
      {
        name: "children",
        type: "Yield",
        summary:
          "The children, as JSX children of `<Section>`. Components, elements, " +
          "arrays and conditionals all count; anything falsy is skipped.",
      },
      DERIVERS,
    ],
    variants: [
      {
        id: "basic",
        title: "A titled group",
        summary: "Two paragraphs under one heading — the common case.",
        component: Opening,
      },
      {
        id: "nested",
        title: "Mixed and nested children",
        summary:
          "A graph, then a nested section holding a graph and a dynamic paragraph.",
        component: YourYear,
      },
    ],
  },
  {
    id: "paragraph",
    title: "Paragraph",
    kind: "paragraph",
    category: "Text",
    status: "stable",
    helpers: ["Paragraph", "useSetPrompts", "useSetPlaceholders"],
    summary: "A block of prose, written from your data or generated from prompts.",
    detail:
      "The workhorse. A static paragraph holds the text as its children; a " +
      "dynamic one carries prompts and a placeholder, and is filled at request " +
      "time. Both land as the same node kind, differing only by `mode` — which " +
      "is inferred from what you supply, never declared. Supplying both text " +
      "and a prompt on one element is an error rather than a precedence rule.",
    children: "None. Paragraphs are leaves.",
    resolves: "Both",
    options: [
      ID,
      {
        name: "children",
        type: "string",
        summary:
          "Static only. The text, interpolated the way any JSX children are. " +
          "A node given its own text is static, whatever prompts a hook set " +
          "around it.",
      },
      {
        name: "text",
        type: "string",
        summary:
          "Static only. The same thing as children, for when a computed string " +
          "reads better as a prop than as a body.",
      },
      {
        name: "align",
        type: `"left" | "center" | "right" | "justify"`,
        summary:
          "How the lines sit in the text column. Say it here when the " +
          "alignment is what the paragraph *is* — a date ranged right, a " +
          "standfirst centred. Leave it out and let a `variant` carry it when " +
          "the alignment is what the theme thinks that kind of block looks " +
          "like. A node that states both wins over its block, the same way a " +
          "cell wins over its column.",
      },
      ...PROMPT_OPTIONS.map((option) => ({
        ...option,
        summary: `Dynamic only. ${option.summary}`,
      })),
      DERIVERS,
    ],
    variants: [
      {
        id: "static",
        title: "Static",
        summary: "Data in, a line of text out.",
        component: Greeting,
      },
      {
        id: "conditional",
        title: "Static, with branching",
        summary: "One node, several outcomes — and the id stays put.",
        component: PriceChange,
      },
      {
        id: "dynamic",
        title: "Dynamic",
        summary: "A prompt and a placeholder. Previews show the placeholder, labelled.",
        component: NextSteps,
      },
      {
        id: "prompted",
        title: "Dynamic, all four prompts",
        summary: "System, general, info and negative, each doing one job.",
        component: Apology,
      },
    ],
  },
  {
    id: "image",
    title: "Image",
    kind: "image",
    category: "Media",
    status: "stable",
    helpers: ["Image"],
    summary: "A picture resolved from your data or described by a prompt.",
    detail:
      "A static image points at something you hold — a signature, a logo, a " +
      "site photograph — with every field able to vary per recipient. A " +
      "dynamic image describes what is wanted and leaves the endpoint to make it.",
    children: "None.",
    resolves: "Both",
    renderNote:
      "Both shipped renderers print a labelled frame in place of the picture — " +
      "`[image: <alt>]` in the DOCX, a dashed box in the browser. The node " +
      "carries its path, alt and size through unchanged.",
    options: [
      ID,
      {
        name: "src",
        type: "string",
        required: true,
        summary:
          "Static only. Lands on the node as `path`. A `data:` URI carries the " +
          "bytes and is the only form that survives into a DOCX; a path or URL " +
          "draws on screen only.",
      },
      {
        name: "fallbackSrc",
        type: "string",
        summary:
          "Static only. A raster to embed in place of an SVG, which Word will " +
          "not take alone. Screen renderers ignore it and draw the SVG.",
      },
      {
        name: "alt",
        type: "string",
        summary:
          "Static only. Alternative text, carried into the DOCX — and the words " +
          "printed in the frame when there is no picture yet.",
      },
      {
        name: "width",
        type: "number",
        summary: "Static only. Rendered width in points, honoured by both renderers.",
      },
      {
        name: "height",
        type: "number",
        summary: "Static only. Rendered height in points, on the same terms as `width`.",
      },
      ...PROMPT_OPTIONS.map((option) => ({
        ...option,
        summary: `Dynamic only. ${option.summary}`,
      })),
      DERIVERS,
    ],
    variants: [
      {
        id: "static",
        title: "Static",
        summary: "A signature whose file and caption both follow the data.",
        component: Signature,
      },
      {
        id: "dynamic",
        title: "Dynamic",
        summary: "Described rather than supplied, with the failure modes fenced off.",
        component: CentrePhoto,
      },
    ],
  },
  {
    id: "graph",
    title: "Graph",
    kind: "graph",
    category: "Data",
    status: "stable",
    helpers: ["Graph"],
    summary: "A bar, line or pie chart declared as data.",
    detail:
      "Charts are declared, never drawn: `graphType` fixes the form, `data` " +
      "returns the payload. Holding numbers rather than an image means one " +
      "declaration serves every renderer and stays diffable in the artifact.",
    children: "None.",
    resolves: "Both",
    renderNote:
      "Both shipped renderers print `[<type> graph: <caption>]` rather than " +
      "plotting anything. `data` passes through untouched, so its shape is a " +
      "contract between you and whichever renderer eventually draws it.",
    options: [
      ID,
      {
        name: "graphType",
        type: '"bar" | "line" | "pie"',
        summary: "The form of the chart. Defaults to `bar`.",
      },
      {
        name: "data",
        type: "JsonObject",
        required: true,
        summary:
          "Static only. The plot payload, as plain JSON. Any shape you like — " +
          "string values in it are run through the template renderer, so " +
          "`{{derived.total}}` resolves inside the payload as it would in prose.",
      },
      {
        name: "caption",
        type: "string",
        summary:
          "Printed beneath the chart, and the words the placeholder frame " +
          "shows today. Optional on both modes.",
      },
      ...PROMPT_OPTIONS.map((option) => ({
        ...option,
        summary: `Dynamic only. ${option.summary}`,
      })),
      DERIVERS,
    ],
    variants: [
      {
        id: "bar",
        title: "Bar",
        summary: "Discrete values across a handful of buckets.",
        component: VisitsByMonth,
      },
      {
        id: "line",
        title: "Line",
        summary: "The same data as a running total — derived in `data`, not upstream.",
        component: CumulativeVisits,
      },
      {
        id: "pie",
        title: "Pie",
        summary: "Shares of a whole.",
        component: ClassMix,
      },
      {
        id: "dynamic",
        title: "Dynamic",
        summary: "Figures the endpoint derives; the form still fixed locally.",
        component: PeakTimes,
      },
    ],
  },
  {
    id: "table",
    title: "Table",
    kind: "table",
    category: "Data",
    status: "stable",
    helpers: ["Table", "Row", "Cell"],
    summary: "A grid of cells, with the columns declared once.",
    detail:
      "The columns belong to the table, because every row shares them — a " +
      "table whose columns do not line up is not a table. Everything else is " +
      "an ordinary node: a `.map()` produces rows, a condition drops one, and " +
      "each names itself. That is what lets a published invoice carry one loop " +
      "the engine walks rather than a table full of special cases.",
    children: "`Row`s, and any `.map()` producing them. A `Row` holds `Cell`s.",
    resolves: "Both",
    options: [
      ID,
      {
        name: "columns",
        type: "TableColumn[]",
        required: true,
        summary:
          "The columns, left to right. Each takes a `width` in millimetres or " +
          "`\"auto\"` to share what the fixed ones leave, and an `align` of " +
          "`left`, `center` or `right`.",
      },
      VARIANT,
      DERIVERS,
    ],
    variants: [
      {
        id: "basic",
        title: "Rows from data",
        summary: "A header row, then one row per entry from a `.map()`.",
        component: VisitLog as Component,
      },
      {
        id: "totals",
        title: "A closing row",
        summary: "A cell holding two paragraphs, and a row marked as a heading.",
        component: PriceSummary as Component,
      },
    ],
  },
  {
    id: "table-of-contents",
    title: "Table of contents",
    kind: "tableOfContents",
    category: "Structure",
    status: "stable",
    helpers: ["TableOfContents"],
    summary: "A marker for a contents list, ahead of the renderers that build one.",
    detail:
      "The kind is part of the document schema and both renderers accept it. " +
      "What is missing is downstream rather than in the authoring: the element " +
      "places the marker and carries its title, and building the entries from " +
      "the surrounding section titles is a renderer's job.",
    children: "None. The entries would come from the sections beside it.",
    resolves: "By the renderer",
    renderNote:
      "Both shipped renderers print the title and nothing else. Building the " +
      "entries from the surrounding section titles is the intent, not yet " +
      "implemented.",
    options: [
      ID,
      {
        name: "title",
        type: "string",
        summary: "Heading for the list. Renderers fall back to their own wording.",
      },
      DERIVERS,
    ],
    variants: [
      {
        id: "basic",
        title: "A titled marker",
        summary: "The element, and the title it carries into the document.",
        component: Contents,
      },
    ],
  },
  {
    id: "clip-art",
    title: "Clip art",
    kind: "clipArt",
    category: "Media",
    status: "planned",
    helpers: [],
    summary: "Named visual blocks — rules, marks, callouts — drawn by the renderer.",
    detail:
      "Not built yet. Nothing is fetched: the node names a shape and the " +
      "renderer draws it, so it stays sharp in the DOCX and ships no asset.",
    children: "Planned: none for marks, a paragraph for callouts.",
    resolves: "By the renderer",
    options: [],
    variants: [],
  },
];

/** Kept in this order everywhere a catalog is listed. */
export const NODE_CATEGORIES: NodeCategory[] = ["Structure", "Text", "Media", "Data"];
