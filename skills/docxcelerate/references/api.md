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

`variant` is a published token like any text. `variant={line.status}` inside a
`.map()` publishes as `"{{ctx.lines.status}}"`, and the engine settles it per
entry. A template literal (`` `status-${line.status}` ``) keeps the token inside
a fixed prefix. The resolved name is looked up in `blocks` like a written one.
If the theme has no block of that name, or the field is missing, the node draws
as an ordinary block.

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
| `Document` | `id?`, `title` (required), `metadata?`, `properties?`, `header?`, `footer?`, `firstHeader?`, `firstFooter?`, `evenHeader?`, `evenFooter?`, `trackChanges?`, `children?` |
| `Section` | `title` (required), `showTitle?` (default `true`), `children?` |
| `Paragraph` | `text?`, `children?`, `align?` (`"left" \| "center" \| "right" \| "justify"`), `list?` (`"bullet" \| "number"`), `listLevel?` (0–8, default 0), `listStyle?` (`"decimal" \| "lowerLetter" \| "upperLetter" \| "lowerRoman" \| "upperRoman"`, or any of them with `Paren` appended: `"lowerLetterParen"` is (a); default by depth 1. a. i.), `listStart?` (a whole number from 1, default 1; numbered items only) |
| `Text` | `bold?`, `italic?`, `underline?`, `strike?`, `script?` (`"superscript" \| "subscript"`), `href?` (`https:`, `http:`, `mailto:` or `tel:`), `variant?` (a block name; only its character settings apply), `children?`. Not a node: no `id`, no common props. Only among the words of a `Paragraph`, `Cell` or `Shape` |
| `Image` | `src?`, `fallbackSrc?`, `alt?`, `width?`, `height?` (points; give one and the other keeps the picture's proportions, neither is 120pt wide), `align?` (`"left" \| "center" \| "right"`), `href?` (`https:`, `http:`, `mailto:` or `tel:`) |
| `Graph` | `graphType?` (`"bar" \| "barHorizontal" \| "line" \| "area" \| "pie" \| "doughnut" \| "scatter" \| "radar" \| "bubble"`, default `bar`), `data?`, `title?`, `alt?`, `caption?`, `width?`, `height?` (points), `align?` (`"left" \| "center" \| "right"`), `legend?`, `stacked?` (`true` stacks to a total, `"percent"` to a share), `numberFormat?`, `categoryAxisTitle?`, `valueAxisTitle?`, `valueAxisMin?`, `valueAxisMax?`, `valueAxisStep?` (numbers; fix the value axis, Word picks any left out), `dataLabels?` |
| `Table` | `columns` (required), `align?` (`"left" \| "center" \| "right"`, default left), `children?` |
| `Row` | `header?`, `keepTogether?`, `height?`, `heightRule?`, `children?` |
| `Cell` | `span?`, `align?`, `children?` |
| `Shape` | `width?`, `height?` (points), `children?` |
| `TableOfContents` | `title?` |
| `PageBreak` | nothing beyond the common props |
| `PageNumber` | `format?` (`"current" \| "total" \| "currentOfTotal"`, default the last), `prefix?` (printed before the number as written, e.g. `"Page "`; default none), `separator?` (default `" / "`), `align?` (`"left" \| "center" \| "right" \| "justify"`; default none, so it inherits) |

Notes worth knowing:

- A `variant` on `PageNumber` draws like the same variant on a `Paragraph`: the
  block's type on every run, both numbers included, and its fill, rules and
  spacing on the line. `align` beats the block's `align`. Inside a `Cell` the
  column's alignment wins, as it does for a paragraph.
- A list is paragraphs with `list="bullet"` or `list="number"`. Items written
  one after another are one list and the renderer counts them; anything that is
  not an item ends it, and the next numbered list starts at one. Never type a
  bullet or a number into the text. `listLevel={1}` nests an item under the one
  above. `listStyle` on a numbered item picks the numerals for that item's level in its list — `listStyle="lowerLetterParen"` on the first sub-clause makes every sub-clause (a), (b); the first item at a level to say decides it, so say it once. Without it a level counts 1., a., i. by depth. Ignored on bullets. `listStart={4}` on the item that opens a numbered list (or a level of one) makes it count from 4: a list that resumes after a table or a quote is a new list, so say where it picks up. It is ignored on any later item and on bullets, and anything but a whole number from 1 is refused at build. Not drawn on a `<Shape>`.
- `<Text>` sets some of a paragraph's words differently:
  `Due within <Text bold>thirty days</Text>.` One inside another takes both
  settings. m<Text script="superscript">2</Text>, CO<Text script="subscript">2</Text>
  and 1<Text script="superscript">st</Text> are raised or lowered by Word — never
  type `²` or `₂` in their place. It carries no colour or size. A link is drawn in its block's `color` when the
  paragraph's or cell's variant sets one, and in a header cell's ink. Otherwise it
  is drawn in the theme's accent. Either way it is underlined, so a link on a
  dark band stays readable. To make a phrase stand out, name a theme block:
  `The amount due is <Text variant="amountDue">£1,240.00</Text>.` The block's
  `color`, `fill` (shading behind the words), `fontSizePt`, `weight`, `font`,
  `transform` and `letterSpacingEm` apply. Borders, padding and spacing do not.
  An unknown name prints plain, and an inner `variant` replaces an outer one.
  A line end (`\n`) in a paragraph's text is a line break, not a new
  paragraph, so an address is one `<Paragraph>` with `{"Flat 2\n14 Example Road"}`.
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
  `keepTogether` on a row stops Word breaking it across a page — the whole row
  moves over instead. Put it on the row a `.map()` produces and every line of an
  invoice or schedule gets it; leave it off a long table's short rows, where a
  break between them is fine. To number the rows, pass the map's index to
  `position`: `lines.map((line, i) => <Row id="line"><Cell id="no">{position(i)}</Cell>…</Row>)`
  prints 1, 2, 3 in a preview and in every published document, counting only the
  rows a `.filter()` kept. The bare index counts from zero and, in a published
  filtered loop, counts the entries dropped too; `i + 1` throws while publishing.
- **A row can be given a height**, in points: `<Row height={56}>` is a floor the
  row grows past if its cells need more, which is what a box to sign in wants.
  `heightRule="exact"` makes it a fixed depth instead — a label sheet, a form
  grid — and Word clips whatever does not fit, so keep an exact row's contents
  inside it (the preview grows the row rather than hiding the overflow). This
  is not a block's `heightPt`, which sets a paragraph's leading, not a row's depth.
- **`align` on a `<Table>` moves the whole table**, not its words: a totals
  block of fixed-width columns with `align="right"` sits against the right
  margin, with no empty spacer column to hold it there. It only moves a table
  narrower than the text column — any `"auto"` column fills the column, and a
  `bleed` table spans the paper and ignores it. Column and cell `align` still
  set the text inside.
- **`Cell` takes text directly** — `<Cell>{line.qty}</Cell>` is the common case.
  Give it paragraphs when one line is not enough.
- **`Shape` is a drawn box the words sit on top of**, and takes text directly
  the way a `Cell` does. Reach for it when the box has to be a *stated* size —
  a status banner the same depth on every document, whatever it says. When the
  box should grow with its text, that is a `Table` with one cell in it, not a
  shape. `width` defaults to the text column; `height` falls back to the block
  style's `heightPt`. What it looks like is the theme's, through `variant`.
- **`align` on `Image` and `Graph` places the picture, not text around it.**
  Centred or ranged right in the column, nothing floats or wraps. A chart with
  no `width` already fills the column, so give it one. Inside a `Cell` or a
  `Shape` a picture or chart takes the column's or the shape's alignment, not
  its own — set it there.
- `Image` takes `src` and resolves to a node holding `path`. **Only a `data:`
  URI travels**: it carries the bytes, so an engine writing the document
  elsewhere still has the picture. A path or URL draws on screen, where a
  browser can fetch it, but packs into Word as a note — nothing is ever fetched
  or read from disk while packing. Word will not embed an SVG alone, so give one
  a `fallbackSrc` raster: the screen draws the SVG, the `.docx` gets the raster.
  A side left out follows the picture's own proportions, read from the PNG,
  JPEG, GIF or BMP header (for an SVG, from its `fallbackSrc` raster); both
  given is drawn as given. A picture with no bytes yet, such as a URL or an
  unproduced prompt, reserves a square.
  `href` links the picture — block, inline, in a cell or in the furniture — and
  lands on the node as `href`, settled like a run's. A refused scheme or an
  unsettled token draws the picture alone.
- `PageBreak` is for a break that is part of what the document *is* — payment
  details on their own page. Nudging a paragraph off the bottom of a page is the
  margins' job.
- `Document.header` and `.footer` are running furniture, drawn on every page and
  sitting outside the margins — how far out is `page.headerMm` / `page.footerMm`,
  measured from the paper rather than the margin. A letterhead meant to appear
  once goes in the body. `PageNumber` belongs here; a build cannot know the
  count, so Word gets a field it recounts and the preview counts its own pages.
- `PageNumber` counts in the page style's numbering: `page.firstPageNumber: 37`
  makes the first page 37, `page.numberStyle: "lowerRoman"` writes i, ii, iii.
  Only the current number moves — `total` is still the count of pages, as Word
  prints it. The whole document shares one numbering; roman front matter then a
  body from 1 is two documents.
- `firstHeader` / `firstFooter` are page one's own, for a letter whose letterhead
  *is* the top of the first page. `false` means the first page shows nothing
  where the others show the strip; absent means it is like every other page.
- `evenHeader` / `evenFooter` are a left-hand page's, for a document printed on
  both sides and bound — the reference belongs at the *outside* edge of each, so
  a folio in the same place on every sheet sits in the gutter on half of them.
  Naming either makes `header` and `footer` the right-hand page's.
  Like every strip, they publish: a `{{data.…}}` token or a deriver in one is settled per document by the engine.
- `trackChanges` opens the document in Word with Track Changes on — for a
  contract, NDA or agreement sent out to be negotiated, so every edit the
  other side makes is recorded. It is a setting, not a lock: Word lets the
  reviewer switch it off. It applies to the whole template, not per
  recipient, and the preview looks the same because nothing has been edited
  yet.
- `Graph.data` is `{ categories?: string[], series: { label?, values: (number
  | null)[], color?, sizes? }[] }`. A `null` value is a gap in the plot, not a
  zero — a month nobody measured is not a month that measured nothing. A series
  shorter than the categories is padded with gaps; one with no categories is
  counted by position. Strings inside the payload are run through the template
  renderer, so a label containing `{{derived.total}}` resolves.
- `scatter` and `bubble` read the categories as their x values, so those are
  numbers written as text; anything that will not parse counts as its position.
  `sizes` is the third figure a `bubble` carries and no other type reads.
- **A generated chart's answer is checked before it is drawn.** What
  `aiClient.generateGraph` returns is read the way model JSON is written: a
  number written as text (`"12"`, `" 3.5 "`) is a number, a numeric category
  (`2024`) is its text, and `null` is the same as leaving a field out. An absent
  `graphType`, `data` or `caption` keeps the node's own, so a `caption` written
  on a prompted `<Graph>` survives a model that only sends numbers. Anything else
  that is not a chart — a `graphType` that is not one of the nine, a `series`
  that is not an array, a series without a `values` array, a value such as
  `"twelve"` — fails with an error naming the node and the field, rather than
  packing a chart that is empty in Word.
- **A chart is packed as a real Word chart**, not as a picture of one: a
  `c:chartSpace` part with every value cached in it, plus the workbook "Edit
  Data" opens. The reader can select it, restyle it, change its type and open
  its numbers, and it prints as vectors at any size. Nothing is rasterised, so
  charts need no extra dependency and no native module.
- Series take the theme's `palette.series` in order, which is what keeps a
  chart re-themeable; a series naming its own `color` overrides it. The
  gridlines take `palette.rule` and the axis and key text `palette.muted`, so a
  chart reads as part of the document rather than as something pasted into it.
- `numberFormat` is an OOXML format code — `"#,##0"`, `"0.0%"`, `"£#,##0"` —
  handed to the value axis and to the data labels verbatim.
- `valueAxisMin` / `valueAxisMax` fix the ends of the value axis and
  `valueAxisStep` the distance between its gridlines — for data whose movement
  is the point, like uptime between 99% and 100%, which an axis starting at zero
  flattens into one line. Each is optional and Word picks whatever is left out.
  They set the y axis on `scatter`/`bubble`, the bottom axis on `barHorizontal`,
  are ignored on a pie, and on a percent stack are shares (`0.5` is 50%). A
  value outside the bounds is clipped. `min` must be below `max` and `step`
  above zero, or packing throws. A workspace scaffolded before 0.9.0 has an
  older `preview/charts.ts` that ignores the bounds — the `.docx` is right
  regardless.
- `legend` defaults to a key under the plot for more than one series and none
  for one, because a key naming the only thing on the chart repeats the title.
  A pie always keeps one: its key names the slices.
- `barHorizontal` reads top-down, in the order the categories were written.
  OOXML's default is bottom-up — which is what Word's own bar charts do, and
  not what a document handing over an array means by it.
- `title` is drawn *on* the chart and travels with it; `caption` is a paragraph
  printed underneath, styled by a `chartCaption` block where the theme has one.
- `alt` is the chart's description for screen readers and Word's
  accessibility checker. Say what the numbers show — "Revenue rose 40% from
  Q1 to Q4" — not what the chart is called. Left out, the title is read
  instead, then the caption, then the id. Nothing prints it.
- The local preview draws charts with ECharts, from the packed file — see
  `preview/charts.ts` in a scaffolded workspace. The frame comes from the file
  and is exact; the plot inside it is another renderer's drawing of the same
  data. Open the `.docx` when the plot itself has to be right.
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
| `<Paragraph><Section …/></Paragraph>` | A paragraph holds words; only a picture or a `<Text>` may sit among them |
| `<Section><Text bold>…</Text></Section>` | `<Text>` sets words inside a `<Paragraph>`, `<Cell>` or `<Shape>`; on its own it is a `<Paragraph>` |
| `<Section><Document …/></Section>` | `<Document>` is only ever the root of a template |
| Two nodes with the same `id` | Ids address a node for the engine, so they must be unique |
| Text **and** a prompt on one node | A node is written or it is generated — supply one |
| `<Image>` with neither `src` nor a prompt | Nothing says what it shows |
| `<Graph>` with neither `data` nor a prompt | Nothing says what it plots |
| A dynamic node in a build with no `aiClient` | Only a preview build may resolve those to placeholders |
| A prompted `<Image>` among a paragraph's words, published to an engine | An engine draws block nodes only — give the `<Image>` its own place beside the paragraph |
| A generated chart whose answer is not a chart (`"twelve"` as a value, an unknown `graphType`) | Word would draw it empty; the error names the node and the field |
| A generated paragraph whose answer is not a string (a provider's response object, `undefined`) | It would pack as `[object Object]` or nothing; the error names the node |

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
| `useDeriver` | `(deriver, inputs) => Promise<result>` | Runs a `deriver({ name, run })` module **now** when the data is in hand. While publishing it travels on the node and the engine runs it per document, and what comes back is a stand-in. |

`arrange(arrangement, { name: () => … })` is not a hook, but it is how a
component offers more than one placement of the same lines: it builds the named
arrangement, the first when none is named, and throws on a name it does not
have. See "The same lines, placed differently" in patterns.md.

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

What `aiClient.generateParagraph` returns is the paragraph itself, so return the
text, not the provider's response object: anything but a string fails with an
error naming the node. The answer is trimmed, so a model's trailing newline does
not become a line break, and characters XML cannot hold — control characters
other than tab, line feed and carriage return — are removed, so the `.docx`
still opens. A line break or tab inside the answer is kept. An empty answer is
an empty paragraph, not an error.

`from` takes request data as it is and survives publishing: write
`from: { period: invoice.period, lines: invoice.lines }`, not a sentence built
from them. Each value read from the data publishes as the token the engine fills
per document, so the model is sent that recipient's period and every one of
their lines. Once filled, a string arrives without quotes and a list or object as
compact JSON; literals you write beside them stay as they are. Name the fields —
handing over the whole request data, or the result of a `.map()`, is refused at
publish.

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

Takes exactly one `<Document>` element, which must have a `title`; `id` is taken
from it unless written. Nothing is rendered — evaluating the JSX only builds elements, which is what
lets a component decide what it is later, with data in hand. A template lives at
module scope where no data exists, so structure that depends on data belongs in
a component.

`properties` is what the file says about itself — never printed, read by Word's
File > Info, Explorer, Finder and DMS indexes. The title is the document's own
`title`. Give an object, or a function of the data when a value differs per
document:

```ts
interface DocumentProperties {
  author?: string;       // Author and Last Modified By; absent writes none
  subject?: string;
  keywords?: string;     // one string — Word shows it as Tags
  description?: string;  // Word's Comments
  custom?: Record<string, string>;  // text only; an empty value is left out
}
```

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
| `buildProjectEngineDocument(project, options?)` | The upload artifact: values stay tokens, branches keep both arms, loops stay loops, and every generated paragraph, image and chart keeps its prompts for the engine's client to answer |
| `buildProjectFinalDocument(project, { data, aiClient, … })` | A finished document at request time, dynamic nodes run through the AI client |
| `createDocumentProjectArtifact(project, options?)` | The manifest, both documents and the deriver bundle together |

A published document is answered by the engine's `AiClient`, not the one you
build with. A generated `<Graph>` needs the engine's client to implement
`generateGraph`, a generated `<Image>` `generateImage`; resolving a node whose
method is missing fails, naming the node. A generated image keeps the `alt`,
`width` and `height` the template gave it unless the client returns its own.

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

## Resolving a published document

An engine holding the artifact fills its `engineDocument` at request time:

```ts
import { createRuntimeState, resolveDocument } from "docxcelerate";

const doc = await resolveDocument(
  artifact.engineDocument,
  createRuntimeState({ data, aiClient }),
);
```

| Option | Default | Meaning |
| --- | --- | --- |
| `data` | required | What `{{data.…}}` reads; copied into `ctx`, never written to |
| `aiClient` | none | Writes dynamic nodes; without one, the first dynamic node reached throws |
| `locale` | the document's | A BCP 47 tag; an invalid one throws |
| `availableTokens` | `2000` | What `{{ctx.availableTokens}}` prints — what `useAvailableTokens()` returned in the build |

Build the state with `createRuntimeState`, not by hand. It is seeded exactly as
`buildDocument` seeds its own, and a hand-built `ctx` without `availableTokens`
silently prints nothing wherever a template read `useAvailableTokens()`.

### Stand-in AI clients

Both come from `docxcelerate` and need no network. Use them in tests and offline previews.

```ts
import { EchoAiClient, ScriptedAiClient } from "docxcelerate";

new EchoAiClient();
// paragraph → "[AI draft for <id>] <prompt>"
// graph     → three sample points, captioned with the echoed prompt; the node keeps its graphType
// image     → a one-pixel grey PNG at the node's width/height, alt = the echoed prompt

const ai = new ScriptedAiClient({
  "tutor-note": "A warm note.",                                   // paragraph: a string
  revenue: { data: { series: [{ values: [1, 2] }] }, caption: "" }, // graph: AiGraphResult
  portrait: { path: "data:image/png;base64,…" },                  // image: AiImageResult
});
ai.requests; // ScriptedAiRequest[]: { id, kind, prompt, prompts }, in the order asked
```

`ScriptedAiClient` matches an exact id first, then the id with each trailing
`-<index>` a repeat added taken off in turn (`note-1-2` → `note-1` → `note`). An
id with no answer, or an answer of the wrong shape for the node, throws and the
build or resolve rejects. Previews (`dynamicMode: "placeholder"`) never call it.

## The built document

```ts
interface DocumentModel {
  schemaVersion: "docxcelerate.letter/v0";
  id: string;
  title: string;
  style?: DocumentStyle;
  metadata?: JsonObject;
  properties?: DocumentProperties;   // docProps; tokens settled by resolveDocument
  trackChanges?: boolean;   // Word opens it with Track Changes on
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
"tableCell" | "shape" | "tableOfContents" | "pageBreak" | "pageNumber" |
"repeat"`.
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

`DocumentStyle` covers `preset`, `page` (size `A4`, `LETTER`, `LEGAL`, `A5`, `A3`, or `{ widthMm, heightMm }` for any other sheet given upright; `orientation` `portrait` or `landscape` turns either kind, and a table, shape or chart with no width fills the column of the page as turned,
margins in mm, `headerMm` / `footerMm` — how far the running strips stand
from the *paper's* edge rather than from the margin, 12.5mm unless said —
`valign`, `"top"` unless said, `"center"` or `"bottom"`, which is where the body
sits between the top and bottom margins: Word's own vertical alignment, for a
certificate or cover sheet set mid-page; it applies to every page, so a short
last page is centred too — and `firstPageNumber` / `numberStyle`: where page numbering starts, a whole number
from 0 (default 1), and how it is written, `"decimal"` (default),
`"lowerRoman"`, `"upperRoman"`, `"lowerLetter"` or `"upperLetter"`, for an
appendix that continues its parent's numbering or front matter numbered
i, ii, iii),
`typography` (body and heading font, body size in pt, line
height, colour), `paragraph` (spacing after, in pt), and `title` and
`sectionHeading` text blocks (font size, weight, spacing before and after, and
`keepWithNext`). A heading is kept on the page with whatever follows it and never
split across a page; that is on without asking, and
`sectionHeading: { keepWithNext: false }` turns the first half off for a theme
that wants headings to fall where they fall.
`cleanMinimalDocumentStyle` is the one shipped preset — spread it and override
what you need. `showTitle: false` stops a renderer printing the document's title
above the body, for a document whose own letterhead already carries it; the
title stays the document's name for everything reading the model.

### Block styles

`blocks` is what a node's `variant` looks up. The node names what it is; the
theme decides what that means.

```ts
blocks: {
  grandTotal: { border: "1C2340", borderSides: ["bottom"], borderStyle: "double", weight: "bold" },
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
| `color` | text colour, and the colour of any link in the block |
| `border`, `borderWidthPt`, `borderSides`, `borderStyle` | a border, on all four edges unless sides are named. `borderStyle` is `"single"` (default) \| `"double"` \| `"dashed"` \| `"dotted"` and applies to every drawn edge — a double rule under a total, a dashed tear-off line. Paragraphs and cells only; a `<Shape>` outline is always solid |
| `paddingPt` | the gap between a drawn border and the text. Word records it on the border, so a block with a fill and no border cannot be padded — see below |
| `fontSizePt`, `weight`, `transform`, `letterSpacingEm`, `font`, `lineHeight` | how the text is set. `transform: "none"` turns capitals off where a heading or header row would otherwise use them |
| `align` | `"left" \| "center" \| "right" \| "justify"`. A node's own `align` wins |
| `spacingBeforePt`, `spacingAfterPt` | space above and below, in points |
| `indentMm`, `indentRightMm` | inset from the left and right margins |
| `firstLineIndentMm` | the first line only, for a book's paragraph mark |
| `hangingIndentMm` | the first line pulled back from the rest. Word writes one or the other, and a hang wins over a first-line indent |
| `keepWithNext`, `keepLines` | refuse to be the last thing on a page, or to be split across one |
| `tabStopsMm` | `[{ at, align?, leader? }]`, measured from the left margin — or, inside a `Cell` or `Shape`, from its inner left edge. Put a `\t` in the text to reach one |
| `bleed` | the block runs the full width of the page rather than the text. Ignored inside a `Cell` or `Shape` |
| `maxWidthMm` | a measure, narrowing the column from the right. Ignored inside a `Cell` or `Shape` |
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
— the narrower statement wins. A paragraph inside a cell or shape is set by its
own `align` and variant over the container's: its alignment, indents, tab stops,
border and fill all apply, so a signature cell can carry a dotted leader and a
card can hold a ruled quote. That is how a muted note sits inside a filled row.
Only the paragraph's own variant draws a border or fill on it; the cell's stays
on the cell. A tab leader inside a `Shape` draws in Word; the preview shows the
gap without the dots.

A section's `variant` styles its printed heading, not its children. The heading takes the block's `fontSizePt`, `weight`, `transform`, `letterSpacingEm`, `color`, `font`, `align`, `border`/`borderWidthPt`/`borderSides`/`paddingPt`, `spacingBeforePt`, `spacingAfterPt`, `keepWithNext` and `keepLines`. Anything the block leaves out comes from `sectionHeading`, and it stays a Heading 1, so the contents still lists it. Fill, indents, bleed, measure, leading and tab stops are not applied to a heading. With `showTitle={false}` there is no heading and the variant does nothing.

```tsx
<Section id="schedule-1" title="Schedule 1: Services" variant="schedule">…</Section>
```
