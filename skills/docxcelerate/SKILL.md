---
name: docxcelerate
description: Write and maintain Docxcelerate documents — DOCX letters composed from typed JSX components, with prose an engine generates per recipient. Use when a workspace has docxcelerate.config.json, documents/*/document.project.ts or *.node.tsx files, when code imports from docxcelerate, docxcelerate/template or docxcelerate/document, or when asked to create a document, add or edit a node, write prompts for generated prose, style the packed .docx, or publish a document to the engine.
---

# Docxcelerate

A document is **a JSX tree plus a data type**. Components return nodes; building
the tree against data produces a `DocumentModel` — plain JSON, no styling and no
layout. Renderers turn that JSON into a `.docx` or a preview page.

If you know a frontend framework, the shape maps over: `document.tsx` is the
entrypoint, `nodes/*.node.tsx` are components, `useState` is where data enters,
and everything else is ordinary TypeScript.

The word "template" here means a document tree, not a string-substitution
language. There is no template language — an `if` is an `if`, `.map()` is
`.map()`, and formatting is a function call.

## Read this before writing code

Four rules cause nearly every mistake an agent makes in this framework.

1. **`useState` is the only door data comes through.** Its initializer receives
   the document data. Nothing else reaches for it.
2. **Every hook runs before the first `await` and before any branch or return.**
   Same rule as React, same reason. There is no `useMemo` — compute in the
   `useState` initializer, which runs once by construction.
3. **Static or dynamic is inferred, never declared.** A node given text (or an
   `Image` a `src`, or a `Graph` its `data`) resolves locally. A node given
   prompts is filled in by the engine. Supplying both on one element is an error.
4. **Never add a `@jsxImportSource` pragma comment.** The workspace
   `tsconfig.json` already sets `jsxImportSource: "docxcelerate/template"` for
   every file. A pragma is only for a foreign project that points
   `jsxImportSource` somewhere else.

## Where things live

```text
my-documents/                        # dxcl init writes this — an ordinary Vite project
  docxcelerate.config.json           # build + upload presets, workspace-wide
  documents/
    tenancy-renewal/
      document.project.ts            # the entrypoint; ties everything below together
      document.tsx                   # structure only — which nodes, which sections, what order
      types.ts                       # the data contract
      preview-data.ts                # one realistic instance of that contract
      document-style.ts              # fonts, spacing, margins for the packed .docx
      nodes/
        greeting.node.tsx            # one node per file
        index.ts                     # re-exports every node
      derivers/index.ts              # named functions the engine runs per document
```

The split is the point — each file answers one question. Keep prose out of
`document.tsx`; a template that inlines its text stops being readable about
halfway down.

Imports: **`docxcelerate/template`** for authoring (elements, hooks, `template`),
**`docxcelerate/document`** for `defineDocumentProject` and style types,
**`docxcelerate`** for `buildDocument` and the domain types.

## A node

```tsx
import { Paragraph, useFormat, useState } from "docxcelerate/template";
import type { TenancyData } from "../types.ts";

export const Balance: Paragraph = () => {
  const { currency } = useFormat();
  const [state] = useState((data: TenancyData) => ({
    name: data.recipientName,
    due: data.balanceDue,
  }));

  if (state.due === 0) {
    return <Paragraph id="balance-settled">Nothing outstanding, {state.name}.</Paragraph>;
  }

  return <Paragraph id="balance-arrears">You owe {currency(state.due)}.</Paragraph>;
};
```

`Paragraph` is both the element and the component type, so
`const Balance: Paragraph` declares what this yields — returning a `<Section>`
from it is a compile error. Give each branch arm **its own id**: that is what
lets a resolved document record which one this recipient got.

The `if` publishes: the build compiles it into a condition, so both arms travel
to the engine and it decides per recipient. The `currency()` call does not —
computing on request data needs a deriver. See
[references/publishing.md](references/publishing.md).

## A node whose prose is generated

Set prompts instead of text, and a placeholder so previews stay readable.

```tsx
import { Paragraph, useSetPlaceholders, useSetPrompts, useState } from "docxcelerate/template";
import type { OfferData } from "../types.ts";

export const TutorNote: Paragraph = () => {
  const [state] = useState((data: OfferData) => ({
    applicant: data.applicantName,
    interviewer: data.interviewer,
  }));

  useSetPrompts({
    systemPrompt: "You are an admissions tutor. Warm, never effusive. Promise nothing.",
    generalPrompt: `Write two specific sentences from ${state.interviewer} about ` +
      `${state.applicant}'s interview.`,
    negativePrompt: "Do not restate the offer, the conditions, or the reply deadline.",
  });
  useSetPlaceholders(`A short note from ${state.interviewer}.`);

  return <Paragraph id="tutor-note" />;
};
```

Only `generalPrompt` is required. `infoPrompt` is context the model should have
but not restate; `negativePrompt` is what to avoid; `systemPrompt` is role and
tone; `examplePrompt` is a finished answer to match rather than a description of
one, which is the cheapest way to pin down an opening, an order and a length. All
six slots (those five plus `placeholder`) can also be given as props, which reads
better when they are short — and props win over the hook, so a caller can
override what a shared hook set around it.

Previews resolve dynamic nodes to their placeholder, never to generated prose.
Nothing leaves the machine, and the same build gives the same page every time.

## The elements

| Element | Holds | Notes |
| --- | --- | --- |
| `Document` | sections and nodes | `id` and `title` both required; one per template; `header`/`footer` take running furniture |
| `Section` | any nodes, including sections | the **only** container; `title` required and becomes a heading |
| `Paragraph` | text children, or prompts | `text` prop says the same as children |
| `Image` | `src`, `alt`, `width`, `height`, or prompts | `src` becomes `path`; only a `data:` URI travels — see below |
| `Graph` | `graphType`, `data`, `caption`, or prompts | figures, never a picture |
| `Table` | `Row`s, and any `.map()` producing them | `columns` declared once, in mm or `"auto"`, with `align` |
| `Row` | `Cell`s | `header` marks a heading row; only *leading* ones repeat across pages |
| `Cell` | text, or paragraphs when a line is not enough | `span`, and `align` when it departs from its column |
| `TableOfContents` | nothing | a marker; renderers print the title and stop |
| `PageBreak` | nothing | for a break that is part of what the document *is* |
| `PageNumber` | nothing | `format` (`current`/`total`/`currentOfTotal`) and `separator`; counted by the renderer |

**An `<Image>` only travels if it carries its bytes.** A `data:` URI does; a
path or a URL draws on screen, where a browser can fetch it, but packs into Word
as a note rather than a picture — the packer never reaches for a file, because
the engine writing the document is not on the machine the file was on. Word will
not embed an SVG alone either, so give one a `fallbackSrc` raster: the screen
draws the SVG and the Word file gets the raster.

Every element also takes **`variant`** — a name the theme looks up, never an
appearance: `<Cell variant="badge">`, `<Paragraph variant="band">`. The colours
live in the style's `blocks`, so a document restyles without a node changing,
and a name the theme has not heard of draws as an ordinary block rather than
failing. Never write a colour into a component.

A block style says `fill`, `color`, `border` (with `borderWidthPt`,
`borderSides`), `paddingPt`, `fontSizePt`, `weight`, `transform`,
`letterSpacingEm` and `bleed`. **All of them mean the same thing on screen and
in the `.docx`** — a fill is shading, a border is a real border, a bleed is a
negative indent past the margin. If a property cannot be expressed in Word it
does not exist here, because a style that quietly did nothing in the format the
framework produces is worse than one that was never offered.

Ids are addresses: an engine targets a node by id and two build artifacts line
up in a diff by id, so **treat a rename as a breaking change**. **Do not write
ids by default.** A node without one is named after its heading, or after the
component that yielded it — `<Greeting />` becomes `greeting`, a section titled
"Fees and funding" becomes `fees-and-funding` — and repeats are numbered
(`greeting-2`). Those names come from what a node is rather than where it sits,
so they survive insertion and reordering. Write one only to pin an address a
request asks for by name. This also keeps
`.map()` and branches from demanding names you do not have. Reusing an id is an
error reported with both positions.

`{condition && <Node />}` reads the way it does everywhere else, and publishes
the way an `if` does — the build compiles it into a condition rather than
deciding once. Falsy children are skipped, so a `&&` yielding anything other
than a node still just drops out. There is no `key` prop — an element accepts
only its own props, so `key={…}` is a type error. Text lives only inside a
`<Paragraph>`, and a paragraph holds text rather than elements.

## Commands

```sh
npx docxcelerate init my-documents      # scaffold a workspace and npm install it
npm run dev                             # preview on 127.0.0.1:4507
dxcl document new tenancy-renewal --title "Tenancy Renewal"
dxcl document node documents/tenancy-renewal next-steps --type paragraph
```

`npx docxcelerate`, not `npx dxcl` — npx resolves the package name and the
binary inside is `dxcl`. Any command run with no arguments asks for what it
needs instead of failing.

`dxcl document node` writes `nodes/<name>.node.tsx` and updates
`nodes/index.ts`. It deliberately does **not** place the node in the template —
after generating one, add it to `document.tsx` where it belongs.

Every flag is in [references/cli.md](references/cli.md).

## Before writing a node from scratch

The package ships a small registry of themes and prebuilt nodes. `dxcl list`
prints it, `dxcl show <id>` prints one entry in full, and `dxcl add <id>`
installs it.

```sh
dxcl list                               # 5 themes, 6 components
dxcl show payment-summary               # what it does and what it reads
dxcl add slate-report letterhead        # a theme and a node, into this project
```

A component is **copied in as source** — `nodes/<name>.node.tsx`, re-exported
from `nodes/index.ts`. It has no version and is never upgraded behind you, so
editing it afterwards is the expected next step rather than a fork. Two things
it leaves to you, both printed as follow-up when it installs: the fields it
reads have to be added to `types.ts` and `preview-data.ts`, and the node itself
has to be placed in `document.tsx`.

A theme is written out as `document-style.ts`, which `document.project.ts`
already passes through — so the next preview is themed. It replaces a
`document-style.ts` nothing has touched; once you have edited that file,
replacing it needs `--force`.

Reach for the registry first when a request names something ordinary — a
letterhead, an address block, a signature, small print. Every entry is listed in
[references/cli.md](references/cli.md) and browsable at
[docxcelerate.com/components](https://docxcelerate.com/components/).

## Publishing changes the rules

Everything above assumes you hold the data. Publishing to the engine builds the
artifact **once**, against stand-ins for a request nobody has made yet, so a
decision that depends on request data has to travel to the engine instead:

- **Interpolating** a value publishes fine — it becomes a `{{data.x}}` token.
  **Computing** on one does not; use a **deriver**, which the engine runs per
  document. Never hand-write a `{{data.…}}` token; interpolation produces it.
- **A preview never waits.** It is rebuilt on every save, so generated nodes show
  the placeholder `useAi` required and derivers that declared a `placeholder`
  stand in rather than run. Cheap derivers still run, so the figures are real.
  Give any deriver that renders, reads or fetches a `placeholder`; leave it off
  for a total or a currency format.
- **`.map()` over request data is how a loop is written.** It walks the list when
  the data is real and is published as a loop the engine walks when it is not.
  Never hand-write a `{{ctx.…}}` token inside one — the entry writes its own
  references. Anything needing the entries first (`.filter`, `.length`, `for…of`)
  belongs in a deriver.
- **A decision is written as an ordinary conditional.** An `if` that returns, a
  ternary, and `cond && <Node />` are all compiled into the condition the engine
  evaluates per document, so both arms travel with the test that selects them.
  A conditional picking a *value* rather than a node is a deriver's job. This
  needs the transform in the build — `docxcelerateTransform()` for Vite,
  `docxcelerateEsbuildTransform()` for esbuild, from `docxcelerate/transform`.
  Without it the decision is made once, at build time, for every recipient.

Read [references/publishing.md](references/publishing.md) before touching a
document that ships to an engine, or when a document is right in preview and
wrong in production.

## Before you call it done

- Each branch arm has its own id; no id is used twice; no `key` props.
- Hooks all called before any `await`, branch or `return`.
- No node carries both text and prompts.
- Every dynamic node has a placeholder, and the document still reads with
  placeholders in place. If it does not, the structure is doing too little work.
- `preview-data.ts` uses the longest name and largest figure you actually
  expect — short names and placeholder cities hide layout problems.
- New nodes are exported from `nodes/index.ts` **and** placed in `document.tsx`.
- `npm run documents:check` type-checks every document.

## Going deeper

- [references/api.md](references/api.md) — every entrypoint, element prop, hook and build function
- [references/patterns.md](references/patterns.md) — copyable recipes: repeats, graphs, house style, shared state, computed sections
- [references/publishing.md](references/publishing.md) — derivers, build artifacts, preview vs engine vs final
- [references/cli.md](references/cli.md) — every `dxcl` command and flag
- [docxcelerate.com/docs](https://docxcelerate.com/docs/start-here/) — the full documentation
