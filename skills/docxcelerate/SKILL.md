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

That node is written for data you hold. Both the `if` and the `currency()` call
work locally and neither survives publishing to an engine — see
[references/publishing.md](references/publishing.md) for the versions that do.

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
tone. All five slots (those four plus `placeholder`) can also be given as props,
which reads better when they are short — and props win over the hook, so a
caller can override what a shared hook set around it.

Previews resolve dynamic nodes to their placeholder, never to generated prose.
Nothing leaves the machine, and the same build gives the same page every time.

## The elements

| Element | Holds | Notes |
| --- | --- | --- |
| `Document` | sections and nodes | `id` and `title` both required; one per template |
| `Section` | any nodes, including sections | the **only** container; `title` required and becomes a heading |
| `Paragraph` | text children, or prompts | `text` prop says the same as children |
| `Image` | `src`, `alt`, `width`, `height`, or prompts | `src` becomes `path` in the built JSON |
| `Graph` | `graphType`, `data`, `caption`, or prompts | figures, never a picture |
| `TableOfContents` | nothing | a marker; renderers print the title and stop |
| `Repeat` | a body walked per entry | `over`, `as` (default `item`), `indexAs` (default `index`) |

Ids are addresses: an engine targets a node by id and two build artifacts line
up in a diff by id, so **treat a rename as a breaking change**. You may omit
one — a node without an id takes one from where it sits, which is what keeps
`.map()` and branches from demanding names you do not have. Reusing an id is an
error reported with both positions.

Falsy children are skipped, so `{condition && <Node />}` reads the way it does
everywhere else. There is no `key` prop — an element accepts only its own props,
so `key={…}` from React habit is a type error. Text lives only inside a
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

## Publishing changes the rules

Everything above assumes you hold the data. Publishing to the engine builds the
artifact **once**, against stand-ins for a request nobody has made yet, so a
decision that depends on request data has to travel to the engine instead:

- **Interpolating** a value publishes fine — it becomes a `{{data.x}}` token.
  **Computing** on one does not; use a **deriver**, which the engine runs per
  document. Never hand-write a `{{data.…}}` token; interpolation produces it.
- **`.map()` over request data is an error while publishing.** Write `<Repeat>`,
  which is published as a loop the engine walks.
- **A decision must be written as a condition.** A plain `if` decides at build
  time; in a published document it silently takes one arm. Use `branch(...)` or a
  `when` prop so both arms travel with the condition that selects them.

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
