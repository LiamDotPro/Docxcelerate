# API reference

Every importable path, element prop, hook and build function. Node 20 or newer;
nothing else is required for authoring, preview or DOCX packing.

## Entrypoints

| Import | What it gives you |
| --- | --- |
| `docxcelerate/template` | The authoring surface: elements, hooks, `template`, branch helpers |
| `docxcelerate/document` | `defineDocumentProject`, `DocumentStyle`, `cleanMinimalDocumentStyle`, `DeriverDefinitions` — plus a re-export of everything in `/template` |
| `docxcelerate` | `buildDocument`, the `buildProject*` functions, artifact helpers, deriver helpers, every domain type |
| `docxcelerate/renderer` | `renderDocumentWebsite` — returns a **complete standalone HTML document**, not a fragment |
| `docxcelerate/docx` | `createDocxDocument` |
| `docxcelerate/scaffold`, `/cli` | The workspace generator and its command-line front end |
| `docxcelerate/template/jsx-runtime` | What `jsxImportSource` resolves to; never imported by hand |

Component files import from `/template`. `defineDocumentProject` comes from
`/document`, which is why `document.project.ts` imports from a different path
than the nodes beside it.

## Elements

Every element accepts `id`, and — written by the branch compiler rather than by
hand — `when` and `derivers`.

```ts
interface CommonElementProps {
  id?: string;
  when?: Condition;                  // a decision left to the engine
  derivers?: DeriverInvocation[];     // derivers the engine runs before this node
}
```

`Paragraph`, `Image` and `Graph` also accept the prompt slots, which are what
make a node dynamic:

```ts
interface PromptProps {
  systemPrompt?: string;      // role and tone
  generalPrompt?: string;     // what the node should say — the only required one
  infoPrompt?: string;        // context the model should have but not restate
  negativePrompt?: string;    // what to avoid
  placeholder?: string;       // what previews show in place of generated content
}
```

| Element | Own props |
| --- | --- |
| `Document` | `id` (required), `title` (required), `metadata?`, `children?` |
| `Section` | `title` (required), `children?` |
| `Paragraph` | `text?`, `children?` |
| `Image` | `src?`, `alt?`, `width?`, `height?` |
| `Graph` | `graphType?` (`"bar" \| "line" \| "pie"`, default `bar`), `data?`, `caption?` |
| `TableOfContents` | `title?` |
| `Repeat` | `over` (required), `as?` (default `item`), `indexAs?` (default `index`), `children?` |

Notes worth knowing:

- `Section` is the only element with children. Depth of a document is depth of
  its sections — but both shipped renderers print every section title at one
  level, so a three-deep document does not yet *look* three-deep.
- `Image` takes a `src` prop and resolves to a node holding `path`. The node
  holds a path, never bytes; nothing is fetched or validated during a build.
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

`Document`, `Section`, `Paragraph`, `Image`, `Graph`, `TableOfContents` and
`Repeat` are all available as types. `Nodes` is the loose one, for wrappers and
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
| `useSetPrompts` | `(prompts: PromptDraft) => PromptDraft` | Prompts for the node this component yields. Calling it is what makes the node dynamic. |
| `useSetPlaceholders` | `(placeholder: string \| { placeholder?: string })` | What previews show in place of generated content. |
| `usePlaceholderData` | `() => PlaceholderData` | Stand-in values, seeded from where the component sits. |
| `useFormat` | `(locale?: string) => Formatters` | Locale-aware formatting; defaults to the build locale (`en-GB`). |
| `useAvailableTokens` | `() => number` | The token budget this build allotted — `2000` by default. |
| `useDeriver` | `(name: string) => (output, ...inputs) => Promise<unknown>` | Runs a registered deriver **now**, during the build. |

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
}
```

A resolved paragraph:

```json
{ "id": "greeting", "kind": "paragraph", "mode": "static", "text": "Dear Adaeze Nkemelu," }
```

`kind` is `"section" | "paragraph" | "image" | "graph" | "tableOfContents" | "repeat"`.
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
margins in mm), `typography` (body and heading font, body size in pt, line
height, colour), `paragraph` (spacing after, in pt), and `title` and
`sectionHeading` text blocks (font size, weight, spacing before and after).
`cleanMinimalDocumentStyle` is the one shipped preset — spread it and override
what you need.
