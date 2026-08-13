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
import type { NodeComponent } from "docxcelerate";
import type { SampleData } from "./sample-data.ts";

import { Greeting } from "./paragraph/static.node.ts";
import { PriceChange } from "./paragraph/conditional.node.ts";
import { NextSteps } from "./paragraph/dynamic.node.ts";
import { Apology } from "./paragraph/prompted.node.ts";
import { Opening } from "./section/basic.node.ts";
import { YourYear } from "./section/nested.node.ts";
import { Signature } from "./image/static.node.ts";
import { CentrePhoto } from "./image/dynamic.node.ts";
import { VisitsByMonth } from "./graph/bar.node.ts";
import { CumulativeVisits } from "./graph/line.node.ts";
import { ClassMix } from "./graph/pie.node.ts";
import { PeakTimes } from "./graph/dynamic.node.ts";
import { Contents } from "./table-of-contents/basic.node.ts";

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
   * Doubles as the file name (src/nodes/<type>/<id>.node.ts) and the preview
   * name (public/demo/nodes/<type>/<id>.html). The build script checks the
   * source file exists, so a rename that misses one half fails the build.
   */
  id: string;
  title: string;
  summary: string;
  component: NodeComponent<SampleData>;
}

export interface NodeTypeEntry {
  /** Directory under src/nodes/, and the docs slug under /docs/nodes/. */
  id: string;
  title: string;
  /** The NodeKind it resolves to in a DocumentModel. */
  kind: string;
  category: NodeCategory;
  status: NodeStatus;
  /** Authoring helpers exported from `docxcelerate`. */
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
    "Values computed before the node resolves, written to `derived.*` and " +
    "readable from a template token. Built with `derive()`.",
};

const ID: NodeOption = {
  name: "id",
  type: "string",
  required: true,
  summary:
    "Stable address for the node. Generation endpoints target it and build " +
    "artifacts diff on it, so treat a rename as a breaking change.",
};

const PROMPT_OPTIONS: NodeOption[] = [
  {
    name: "placeholder",
    type: "(data, availableTokens) => string",
    summary:
      "What previews show in place of generated content. Optional, but a letter " +
      "that reads badly without one cannot be reviewed.",
  },
  {
    name: "generalPrompt",
    type: "(data, availableTokens) => string",
    required: true,
    summary: "What this node should say.",
  },
  {
    name: "infoPrompt",
    type: "(data, availableTokens) => string",
    summary: "Context the model should have but should not restate.",
  },
  {
    name: "negativePrompt",
    type: "(data, availableTokens) => string",
    summary: "What to avoid — claims, tones, or facts it must not invent.",
  },
  {
    name: "systemPrompt",
    type: "(data, availableTokens) => string",
    summary: "Role and voice, applied ahead of the other prompts.",
  },
];

export const NODE_TYPES: NodeTypeEntry[] = [
  {
    id: "section",
    title: "Section",
    kind: "section",
    category: "Structure",
    status: "stable",
    helpers: ["section", "Section"],
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
        name: "nodes",
        type: "NodeComponent[]",
        summary:
          "The children. Passed as the second argument, as the `nodes` option, " +
          "or as JSX children of `<Section>` — the three are the same call.",
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
    helpers: ["paragraph"],
    summary: "A block of prose, rendered from your data or from prompts.",
    detail:
      "The workhorse. A static paragraph returns a string from your typed data; " +
      "a dynamic one carries prompts and a placeholder, and is filled at " +
      "request time. Both land as the same node kind, differing only by `mode`.",
    children: "None. Paragraphs are leaves.",
    resolves: "Both",
    options: [
      ID,
      {
        name: "render",
        type: "(data, availableTokens) => string",
        required: true,
        summary:
          "Static only. Receives your typed data and the token budget, returns " +
          "the text. May be async.",
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
    helpers: ["image"],
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
        type: "string | (data) => string",
        required: true,
        summary: "Static only. Path or URL to the image; lands on the node as `path`.",
      },
      {
        name: "alt",
        type: "string | (data) => string",
        summary:
          "Static only. Alternative text — and, while the renderers draw a " +
          "frame rather than the picture, the words printed inside it.",
      },
      {
        name: "width",
        type: "number | (data) => number",
        summary:
          "Static only. Intended width. Carried onto the node untouched; no " +
          "shipped renderer reads it yet.",
      },
      {
        name: "height",
        type: "number | (data) => number",
        summary: "Static only. Intended height, on the same terms as `width`.",
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
    helpers: ["graph"],
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
        type: "JsonObject | (data) => JsonObject",
        required: true,
        summary:
          "Static only. The plot payload, as plain JSON. Any shape you like — " +
          "string values in it are run through the template renderer, so " +
          "`{{derived.total}}` resolves inside the payload as it would in prose.",
      },
      {
        name: "caption",
        type: "string | (data) => string",
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
    id: "table-of-contents",
    title: "Table of contents",
    kind: "tableOfContents",
    category: "Structure",
    status: "no-helper",
    helpers: [],
    summary: "A marker for a contents list, ahead of the renderers that build one.",
    detail:
      "The kind is part of the letter schema and both renderers accept it, but " +
      "no authoring helper is exported yet. Writing the component by hand " +
      "works — a node component is a function returning a definition, and the " +
      "helpers are conveniences over exactly that shape.",
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
        title: "Written by hand",
        summary: "The definition shape the helpers would produce.",
        component: Contents,
      },
    ],
  },
  {
    id: "table",
    title: "Table",
    kind: "table",
    category: "Data",
    status: "planned",
    helpers: [],
    summary: "Rows and columns, with cells that are themselves nodes.",
    detail:
      "Not built yet. The intent is a node whose cells hold other nodes, so a " +
      "table composes the way a section does rather than becoming a second " +
      "content model beside it.",
    children: "Planned: paragraphs, images and graphs, per cell.",
    resolves: "Locally",
    options: [],
    variants: [],
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
