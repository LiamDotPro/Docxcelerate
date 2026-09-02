# API reference

Every importable path, element prop, hook and build function. Node 20 or newer;
nothing else is required for authoring, preview or DOCX packing.

## Entrypoints

| Import | What it gives you |
| --- | --- |
| `docxcelerate/template` | The authoring surface: elements, hooks, `template`, branch helpers |
| `docxcelerate/document` | `defineDocumentProject`, `DocumentStyle`, `cleanMinimalDocumentStyle`, `DeriverDefinitions` — plus a re-export of everything in `/template` |
| `docxcelerate` | `buildDocument`, the `buildProject*` functions, artifact helpers, deriver helpers, every domain type |
| `docxcelerate/docx` | `createDocxDocument`, `createDocxBlob`, `renderDocxBytes` — the one renderer |
| `docxcelerate/scaffold`, `/cli` | The workspace generator and its command-line front end |
| `docxcelerate/template/jsx-runtime` | What `jsxImportSource` resolves to; never imported by hand |

Component files import from `/template`. `defineDocumentProject` comes from
`/document`, which is why `document.project.ts` imports from a different path
than the nodes beside it.

## Elements

Every element accepts `id` and `variant`, and — written by the branch compiler
rather than by hand — `when` and `derivers`.

```ts
interface CommonElementProps {
  id?: string;
  variant?: string;                   // a block style the theme looks up
  when?: Condition;                   // a decision left to the engine
  derivers?: DeriverInvocation[];     // derivers the engine runs before this node
}
```

`variant` names what a node *is* — `"band"`, `"badge"`, `"panel"` — never what
it looks like. The appearance lives in the style's `blocks`, so a document
restyles without a node changing, and a name the theme has not heard of draws as
an ordinary block rather than failing. See **Block styles** below.

`Paragraph`, `Image` and `Graph` also accept the prompt slots, which are what
make a node dynamic:

```ts
interface PromptProps {
  systemPrompt?: string;      // role and tone
  generalPrompt?: string;     // what the node should say — the only required one
  infoPrompt?: string;        // context the model should have but not restate
  negativePrompt?: string;    // what to avoid
  examplePrompt?: string;     // what a good answer looks like, written out
  placeholder?: string;       // what previews show in place of generated content
}
```

| Element | Own props |
| --- | --- |
| `Document` | `id` (required), `title` (required), `metadata?`, `header?`, `footer?`, `firstHeader?`, `firstFooter?`, `evenHeader?`, `evenFooter?`, `children?` |
| `Section` | `title` (required), `children?` |
| `Paragraph` | `text?`, `children?`, `align?` (`"left" \| "center" \| "right" \| "justify"`) |
| `Image` | `src?`, `fallbackSrc?`, `alt?`, `width?`, `height?` |
| `Graph` | `graphType?` (`"bar" \| "line" \| "pie"`, default `bar`), `data?`, `caption?` |
| `Table` | `columns` (required), `children?` |
| `Row` | `header?`, `children?` |
| `Cell` | `span?`, `align?`, `children?` |
| `TableOfContents` | `title?` |
| `PageBreak` | nothing beyond the common props |
| `PageNumber` | `format?` (`"current" \| "total" \| "currentOfTotal"`, default the last), `separator?` (default `" / "`) |

Notes worth knowing:

- `Section` is the only *titled* container. Depth of a document is depth of its
  sections — but both shipped renderers print every section title at one level,
  so a three-deep document does not yet *look* three-deep. `Table`, `Row` and
  `Cell` hold nodes too; they just do not head anything.
- **A table's columns are declared once, on the table**, in millimetres or
  `"auto"`, each with an optional `align`. Rows and cells are ordinary nodes, so
  a `.map()` produces rows, a condition drops one, and every id names itself —
  none of it is a special case. `header` on a row draws it as a heading; only
  the rows a table *opens* with repeat across pages, so a totals row marked
  `header` stays where it is rather than being lifted above its own figures.
- **`Cell` takes text directly** — `<Cell>{line.qty}</Cell>` is the common case.
  Give it paragraphs when one line is not enough.
- `Image` takes `src` and resolves to a node holding `path`. **Only a `data:`
  URI travels**: it carries the bytes, so an engine writing the document
  elsewhere still has the picture. A path or URL draws on screen, where a
  browser can fetch it, but packs into Word as a note — nothing is ever fetched
  or read from disk while packing. Word will not embed an SVG alone, so give one
  a `fallbackSrc` raster: the screen draws the SVG, the `.docx` gets the raster.
- `PageBreak` is for a break that is part of what the document *is* — payment
  details on their own page. Nudging a paragraph off the bottom of a page is the
  margins' job.
- `Document.header` and `.footer` are running furniture, drawn on every page and
  sitting outside the margins — how far out is `page.headerMm` / `page.footerMm`,
  measured from the paper rather than the margin. A letterhead meant to appear
  once goes in the body. `PageNumber` belongs here; a build cannot know the
  count, so Word gets a field it recounts and the preview counts its own pages.
- `firstHeader` / `firstFooter` are page one's own, for a letter whose letterhead
  *is* the top of the first page. `false` means the first page shows nothing
  where the others show the strip; absent means it is like every other page.
- `evenHeader` / `evenFooter` are a left-hand page's, for a document printed on
  both sides and bound — the reference belongs at the *outside* edge of each, so
  a folio in the same place on every sheet sits in the gutter on half of them.
  Naming either makes `header` and `footer` the right-hand page's.
- `Graph.data` is a `JsonObject` the framework never looks inside.
  `{ labels, series: [{ name, values }] }` is a convention, not a schema — pick
  one shape and keep it consistent across a project. String values inside the
  payload *are* run through the template renderer, so a label containing
  `{{derived.total}}` resolves.
- A paragraph given an empty string is an empty paragraph, not an absent node.
  To drop a node, return `false`, `null` or `undefined`.
- Renderers escape paragraph text. A paragraph cannot smuggle markup into a page.
- There is no dynamic `Section`. An engine fills nodes in; it does not decide
  what a document contains.

### What the renderer rejects

Each of these is an error naming the position, not a quiet reinterpretation:

| Written | Why it fails |
| --- | --- |
| `<Section>Some text</Section>` | Text only lives inside a `<Paragraph>` |
| `<Paragraph><Image /></Paragraph>` | A paragraph holds text; put elements beside it |
| `<Section><Document …/></Section>` | `<Document>` is only ever the root of a template |
| Two nodes with the same `id` | Ids address a node for the engine, so they must be unique |
| Text **and** a prompt on one node | A node is written or it is generated — supply one |
| `<Image>` with neither `src` nor a prompt | Nothing says what it shows |
| `<Graph>` with neither `data` nor a prompt | Nothing says what it plots |
| A dynamic node in a build with no `aiClient` | Only a preview build may resolve those to placeholders |

## Component types

Each element name is also a component type, so a value and a type share a name:

```tsx
export const Greeting: Paragraph = () => <Paragraph id="greeting">Hello.</Paragraph>;
export const Opening: Section = () => <Section id="opening" title="Opening">…</Section>;
export const Letterhead: Nodes = () => …;   // free to yield whatever fits
```

`Document`, `Section`, `Paragraph`, `Image`, `Graph` and `TableOfContents` are
all available as types. `Nodes` is the loose one, for wrappers and
layout pieces. Each takes an optional props parameter:

```tsx
export const Arrears: Paragraph<{ amount: number }> = ({ amount }) => {
  const { currency } = useFormat();
  return <Paragraph id="arrears">You owe {currency(amount)}.</Paragraph>;
};
```

Something has to read the data first, so a props-only component sits under a
parent that took it into state. Use props for components reused against
different values, state for components that know which document they belong to.

## Hooks

All of them must be called before the first `await`, before any branch, and
before any `return`.

| Hook | Signature | Purpose |
| --- | --- | --- |
| `useState` | `(initial: TState \| ((data: TData) => TState)) => [TState, setter]` | Data, taken in once and kept. The one door data comes through. |
| `useShared` | `(key: string, initial) => [TValue, setter]` | A value left for the components rendered *after* this one, in document order. |
| `useAi` | `(config: AiConfig) => PromptDraft` | Everything a generated node needs, in one call. Calling it is what makes the node dynamic. |
| `useSetPrompts` | `(prompts: PromptDraft) => PromptDraft` | The same prompts under their downstream names, for a hook adding house style to a node it does not own. |
| `useSetPlaceholders` | `(placeholder: string \| { placeholder?: string })` | What previews show in place of generated content. |
| `usePlaceholderData` | `() => PlaceholderData` | Stand-in values, seeded from where the component sits. |
| `useFormat` | `(locale?: string) => Formatters` | Locale-aware formatting; defaults to the build locale (`en-GB`). |
| `useAvailableTokens` | `() => number` | The token budget this build allotted — `2000` by default. |
| `useDeriver` | `(name: string) => (output, ...inputs) => Promise<unknown>` | Runs a registered deriver **now**, during the build. |

`useAi` is the one to reach for. It says in one call what a generated node is:

```ts
interface AiConfig {
  ask: string;                  // what the node should say — required
  placeholder: string;          // what stands in its place until written — required
  voice?: string;               // how it should sound
  from?: JsonObject | string;   // the facts to write from, as data or as prose
  avoid?: string;               // what it must not say
  example?: string | string[];  // what a good answer looks like, written out
}
```

`placeholder` is required because a blank in a preview reads as a finished
document. `example` is the field that buys the most: a described format is
something a model interprets, a shown one is something it matches, so an example
holds still the opening, the order, the length and the register and leaves only
the per-document parts to be written. Give it as finished text, not a form with
blanks in it. Pass an array when the shape legitimately varies by case — they
travel numbered, and read as a pattern rather than a template.

The `useState` setter is not a re-render request — a build is a single pass. It
updates the value this component reads later, and the value anything sharing it
through `useShared` will read.

There is deliberately no `useMemo`. A build renders each component once and keeps
no instances between builds, so a memo could never return a cached value, and
caching across builds would be wrong anyway. Compute in the `useState`
initializer.

```ts
interface PlaceholderData {
  name(): string;
  city(): string;
  date(offsetDays?: number): string;
  currency(amount?: number): string;
  sentence(words?: number): string;
  paragraph(sentences?: number): string;
  pick<T>(values: readonly T[]): T;
}

interface Formatters {
  currency(amount: number, currency?: string): string;      // defaults to GBP
  number(value: number, options?: Intl.NumberFormatOptions): string;
  date(value: Date | string | number, options?: Intl.DateTimeFormatOptions): string;
  list(values: readonly string[], type?: "conjunction" | "disjunction"): string;
  plural(count: number, singular: string, plural?: string): string;
}
```

`usePlaceholderData` is seeded from the component's position, so the same node
shows the same name and the same figures on every build. A preview that
reshuffles itself is one nobody can proofread.

## `template`

```tsx
import { Document, Section, template } from "docxcelerate/template";

export const documentTemplate = template<DocumentData>(
  <Document id="tenancy-renewal" title="Tenancy Renewal">…</Document>,
);
```

Takes exactly one `<Document>` element, which must have both `id` and `title`.
Nothing is rendered — evaluating the JSX only builds elements, which is what
lets a component decide what it is later, with data in hand. A template lives at
module scope where no data exists, so structure that depends on data belongs in
a component.

## `defineDocumentProject`

```ts
import { defineDocumentProject } from "docxcelerate/document";

export default defineDocumentProject<DocumentData>({
  id: "tenancy-renewal",       // how the document is addressed, in artifacts and by an engine
  name: "Tenancy Renewal",     // what a person sees in the preview app
  version: "0.1.0",            // stamped into every artifact this project builds
  template: documentTemplate,
  previewData,                 // what the preview resolves against
  derivers,                    // values computed per document rather than per build
  style: documentStyle,        // applied when packing
  previewOptions: { availableTokens: 800 },
  buildOptions: {},            // runtime options for engine and final builds
  metadata: {},
});
```

`version` defaults to `"0.1.0"`. The preview app discovers projects by globbing
for `document.project.ts` — a new document appears in the picker as soon as it
exists, and nothing registers it.

## Building

```ts
import {
  buildDocument,
  buildProjectPreviewDocument,
  buildProjectEngineDocument,
  buildProjectFinalDocument,
  createDocumentProjectArtifact,
} from "docxcelerate";
```

| Function | Builds |
| --- | --- |
| `buildDocument(template, data, options?)` | The raw model from a template and data you hold |
| `buildProjectPreviewDocument(project, options?)` | What the preview app shows: dynamic nodes → placeholders, derivers resolved, project style applied |
| `buildProjectEngineDocument(project, options?)` | The upload artifact: values stay tokens, branches keep both arms, loops stay loops |
| `buildProjectFinalDocument(project, { data, aiClient, … })` | A finished document at request time, dynamic nodes run through the AI client |
| `createDocumentProjectArtifact(project, options?)` | The manifest, both documents and the deriver bundle together |

For a project, prefer `buildProjectPreviewDocument` over `buildDocument` — it is
exactly what the preview app calls, so what you build in code matches what you
saw in the browser.

```ts
interface ComponentRuntimeOptions {
  availableTokens?: number;        // default 2000
  aiClient?: AiClient;             // required to resolve dynamic nodes
  dynamicMode?: "resolve" | "placeholder";
  derivers?: DeriverDefinitions | DeriverRegistry;
  deriverMode?: "resolve" | "preserve";
  branchMode?: "decide" | "publish";   // default "decide"
  branchLimit?: number;                // default 32
  locale?: string;                     // default "en-GB"
}
```

`branchMode: "decide"` is what a preview, a local pack and a live generation all
want. `"publish"` is for the artifact that goes to an engine, where the decision
belongs to a request that has not happened yet.

## The built document

```ts
interface DocumentModel {
  schemaVersion: "docxcelerate.letter/v0";
  id: string;
  title: string;
  style?: DocumentStyle;
  metadata?: JsonObject;
  nodes: DocumentNode[];
  header?: DocumentNode[];   // running furniture, drawn on every page
  footer?: DocumentNode[];
}
```

A resolved paragraph:

```json
{ "id": "greeting", "kind": "paragraph", "mode": "static", "text": "Dear Adaeze Nkemelu," }
```

`kind` is `"section" | "paragraph" | "image" | "graph" | "table" | "tableRow" |
"tableCell" | "tableOfContents" | "pageBreak" | "pageNumber" | "repeat"`.
`mode` is `"static" | "dynamic"` and is **output, not input** — the build derives
it from what the component supplied. Both modes resolve to the same `kind`; a
renderer reads `mode` if it cares at all, and never branches on how the node was
written.

The model carries no styling and no layout. Those belong to the renderer.

## Style

```ts
import { cleanMinimalDocumentStyle, type DocumentStyle } from "docxcelerate/document";

export const documentStyle: DocumentStyle = {
  ...cleanMinimalDocumentStyle,
  page: {
    ...cleanMinimalDocumentStyle.page,
    margins: { topMm: 25.4, rightMm: 25.4, bottomMm: 25.4, leftMm: 25.4 },
  },
};
```

`DocumentStyle` covers `preset`, `page` (size `A4` or `LETTER`, orientation,
margins in mm, and `headerMm` / `footerMm` — how far the running strips stand
from the *paper's* edge rather than from the margin, 12.5mm unless said),
`typography` (body and heading font, body size in pt, line
height, colour), `paragraph` (spacing after, in pt), and `title` and
`sectionHeading` text blocks (font size, weight, spacing before and after).
`cleanMinimalDocumentStyle` is the one shipped preset — spread it and override
what you need. `showTitle: false` stops a renderer printing the document's title
above the body, for a document whose own letterhead already carries it; the
title stays the document's name for everything reading the model.

### Block styles

`blocks` is what a node's `variant` looks up. The node names what it is; the
theme decides what that means.

```ts
blocks: {
  band:  { fill: "F4F6FD", bleed: true, border: "E3E7F5", borderSides: ["bottom"], paddingPt: 10 },
  badge: { fill: "FBF0DC", border: "E5C78A", color: "8A5A06", fontSizePt: 7,
           weight: "bold", transform: "uppercase", letterSpacingEm: 0.1, paddingPt: 5 },
  quote: { indentMm: 10, indentRightMm: 10 },
  heading: { keepWithNext: true, spacingBeforePt: 18, spacingAfterPt: 4 },
  contentsLine: { tabStopsMm: [{ at: 170, align: "right", leader: "dot" }] },
}
```

| Field | Means |
| --- | --- |
| `fill` | background, hex without the `#` |
| `color` | text colour |
| `border`, `borderWidthPt`, `borderSides` | a border, on all four edges unless sides are named |
| `paddingPt` | the gap between a drawn border and the text. Word records it on the border, so a block with a fill and no border cannot be padded — see below |
| `fontSizePt`, `weight`, `transform`, `letterSpacingEm`, `font`, `lineHeight` | how the text is set |
| `align` | `"left" \| "center" \| "right" \| "justify"`. A node's own `align` wins |
| `spacingBeforePt`, `spacingAfterPt` | space above and below, in points |
| `indentMm`, `indentRightMm` | inset from the left and right margins |
| `firstLineIndentMm` | the first line only, for a book's paragraph mark |
| `hangingIndentMm` | the first line pulled back from the rest. Word writes one or the other, and a hang wins over a first-line indent |
| `keepWithNext`, `keepLines` | refuse to be the last thing on a page, or to be split across one |
| `tabStopsMm` | `[{ at, align?, leader? }]`, measured from the left margin. Put a `\t` in the text to reach one |
| `bleed` | the block runs the full width of the page rather than the text |
| `maxWidthMm` | a measure, narrowing the column from the right |
| `valign`, `heightPt`, `paddingSidesPt` | for a block that is a strip or a cell rather than prose |

**Every one of these means the same thing on screen and in the `.docx`** — a
fill is shading, a border is a real border, a bleed is a negative indent past
the margin, letter spacing is character spacing, a tab stop is `w:tabs`. If a
property cannot be expressed in Word it is not offered: a style that quietly did
nothing in the format the framework produces would be worse than one that never
existed. That is why there is no corner rounding — Word has no rounded blocks,
so neither does this.

Two consequences of that rule are worth knowing before you reach for them:

- **Padding needs a border.** Word has no padding on a paragraph. What it has is
  `w:pBdr`'s `w:space`, the gap between a rule and the text — so `paddingPt` on
  a block with a `fill` and no `border` does nothing, in Word *and* on screen. A
  filled panel that must hold its words off its own edge has to be a table cell,
  which is what a picture's variant already becomes.
- **Padding is not symmetric.** Measured against Word: a top border pushes the
  text *down*, and a left border does not push it *in* — the rule steps out into
  the margin and the words stay on the column. The preview draws both the way
  Word does.

The paragraph properties above are held to Word by
`conformance/cases/text/`, which packs a document for each, opens it in Word,
and compares the two against each other to the millimetre.

A cell takes its block from its own `variant`, then its row's, then its table's
— the narrower statement wins. A paragraph inside a cell is set by its own
variant over the cell's, which is how a muted note sits inside a filled row.
