<!-- Absolute URLs, and markdown rather than <img>: this file is the package
     page on npm and on JSR as well as the front page here, and neither registry
     has the repository around it to resolve a relative path against. The banner
     carries its own dark ground, so it reads the same on a light theme and a
     dark one. Rebuild it, and every icon beside it, with `npm run brand` in
     website/. -->

![Docxcelerate — Documents as components. DOCX as output.](https://raw.githubusercontent.com/LiamDotPro/Docxcelerate/main/website/public/brand/banner.png)

[![npm](https://img.shields.io/npm/v/docxcelerate?label=npm&color=17624a&labelColor=16211c)](https://www.npmjs.com/package/docxcelerate)
[![JSR](https://img.shields.io/jsr/v/%40docxcelerate/docxcelerate?label=jsr&color=17624a&labelColor=16211c)](https://jsr.io/@docxcelerate/docxcelerate)
[![JSR score](https://jsr.io/badges/@docxcelerate/docxcelerate/score)](https://jsr.io/@docxcelerate/docxcelerate)
[![node](https://img.shields.io/node/v/docxcelerate?label=node&color=17624a&labelColor=16211c)](https://nodejs.org)

Compose DOCX documents from small typed components, using the JSX you already write.
Write and preview them on your machine, then generate finished documents through
the engine.

**[Documentation](https://docxcelerate.com/docs/start-here/)** · [docxcelerate.com](https://docxcelerate.com)

## Requirements

Node.js 20 or newer. Nothing else — authoring, preview and DOCX packing all run locally.

## Quick start

```sh
npx docxcelerate init my-documents
cd my-documents
npm run dev
```

This creates a workspace and opens the preview app, where you can write documents,
edit and add nodes, and pack the result into a finished DOCX.

To get the `dxcl` binary on your path:

```sh
npm install -g docxcelerate
```

## What a document looks like

A document is a small project of its own, under `documents/<name>/`. The tree
lives in **document.tsx**, and **document.project.ts** beside it is the
entrypoint that pairs the template with the data, style and derivers it is built
with.

```tsx
// documents/tenancy-renewal/document.tsx
import { Document, Section, template } from "docxcelerate/template";
import * as Nodes from "./nodes/index.ts";
import type { DocumentData } from "./types.ts";

export const documentTemplate = template<DocumentData>(
  <Document title="Tenancy Renewal">
    <Section id="opening" title="Opening">
      <Nodes.Greeting />
      <Nodes.Balance />
    </Section>
  </Document>,
);
```

If you are familiar with frontend frameworks, the idea of having an entrypoint for
your application maps well onto the shape of how documents are structured with
Docxcelerate. The document's id is taken from its title when you do not write one.

We build components which represent the contents of a document: a paragraph, a
table, an image, a graph, a page break, a table of contents. Each one declares
what it yields, and reads data through `useState` — the only way data enters a
component.

```tsx
// documents/tenancy-renewal/nodes/balance.node.tsx
import { Paragraph, useState } from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const Balance: Paragraph = () => {
  const [state] = useState((data: DocumentData) => ({
    name: data.recipientName,
    settled: data.balanceDue === 0,
  }));

  if (state.settled) {
    return <Paragraph id="settled">Nothing outstanding, {state.name}.</Paragraph>;
  }

  return <Paragraph id="arrears">There is a balance left to settle, {state.name}.</Paragraph>;
};
```

## Nodes an engine writes

Some nodes are written per document rather than at build time. `useAi` is what
says so — one call carrying what to ask for, and what stands in the node's place
until something has written it.

```tsx
import { Paragraph, useAi, useState } from "docxcelerate/template";
import type { DocumentData } from "../types.ts";

export const Balance: Paragraph = () => {
  const [state] = useState((data: DocumentData) => data.account);

  useAi({
    ask: "Explain the balance on this account and what to do about it.",
    placeholder: "A short note about the balance left on this account.",
    voice: "A housing officer writing to a tenant. Plain, no sales tone.",
    from: { balanceDue: state.balanceDue, dueBy: state.dueBy },
    avoid: "Do not invent a payment method or a phone number.",
    example:
      "Your account is £248.50 in arrears, and the balance is due by 14 March. " +
      "Please get in touch before then so we can agree how to clear it.",
  });

  return <Paragraph id="balance" />;
};
```

You can imagine that if the account is a long way into arrears the engine may
choose to write something firmer than it would at a few pounds. The `placeholder`
is required, and deliberately so: a preview is how a document gets proofread, and
a document proofread with a blank in it is a document nobody read.

`example` is the one that does the most for the least. Describing a format is
something a model has to interpret; showing it one is something it can match, so
an example holds still the parts of a node that were never meant to vary — the
opening, the order, the length, the register — and leaves only what genuinely
differs per document. Write it as finished text rather than a form with blanks in
it, and pass an array of them when the shape legitimately varies by case.

Jump into our documentation to understand how you can structure documents with AI
and create more complex documents for your own needs.

## Components you can install

The registry ships nodes that are already written — a letterhead, a recipient
block, a signature block, a payment summary that branches on whether an account is
in credit, clear or owing. `dxcl add` copies one into your document project as
source you own and can edit, and wires it into `nodes/index.ts`.

```sh
dxcl list                   # every component and theme
dxcl show payment-summary   # what it renders, and the fields it reads
dxcl add payment-summary    # copy it into documents/<name>/nodes/
```

Themes install the same way and write the project's `document-style.ts`:

```sh
dxcl add slate-report
```

Browse them all at [docxcelerate.com/registry](https://docxcelerate.com/registry/).

## Documentation

Everything is on [docxcelerate.com](https://docxcelerate.com):

- [Start here](https://docxcelerate.com/docs/start-here/) — install the CLI and scaffold a workspace
- [Writing nodes](https://docxcelerate.com/docs/writing-nodes/) — the pieces a document is made of, and where AI fits
- [Document projects](https://docxcelerate.com/docs/document-projects/) — the files in a document, and what configures them
- [Documents and nodes](https://docxcelerate.com/docs/essentials/documents-and-nodes/) — the model
- [Templates](https://docxcelerate.com/docs/essentials/templates/) — composing with JSX
- [The node reference](https://docxcelerate.com/docs/nodes/overview/) — every node type, rendered
- [The registry](https://docxcelerate.com/registry/) — components and themes to install
- [CLI commands](https://docxcelerate.com/docs/cli/commands/) — every `dxcl` command
- [The engine](https://docxcelerate.com/docs/generation/endpoint/) — generating at scale
- [Package entrypoints](https://docxcelerate.com/docs/reference/entrypoints/) — what each import gives you

## Development

```sh
npm install       # install dependencies
npm run build     # compile src/ to dist/ with type declarations
npm test          # build, then run the Node test suite
npm run typecheck # type-check sources and tests
npm run jsr:doc   # check every entrypoint and exported symbol is documented
```

`jsr:doc` is what keeps the package's JSR score at nine out of nine. It needs
[Deno](https://deno.com) on your path, and runs on every pull request that
touches `src/`.

The published package ships compiled output from `dist/` plus the `dxcl` binary.
Pushing a change to `src/` on `main` publishes a new version.

## License

MIT
