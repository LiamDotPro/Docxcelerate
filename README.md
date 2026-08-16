# Docxcelerate

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

This creates a workspace and opens the preview app, where you can begin to create documents
and edit or add nodes and finally package it into a reusable format.

To get the `dxcl` binary on your path:

```sh
npm install -g docxcelerate
```

## What a document looks like

Below is the **index.ts** which is the root of a document inside of a workspace. All documents use the Document component as the parent
and allow us to configure the documents settings and stylings.

```tsx
import { Document, Section, template } from "docxcelerate/template";
import { Balance, Greeting } from "./nodes/index.ts";
import type { TenancyData } from "./types.ts";

export const documentTemplate = template<TenancyData>(
  <Document id="tenancy-renewal" title="Tenancy Renewal">
    <Section id="opening" title="Opening">
      <Greeting />
      <Balance />
    </Section>
  </Document>,
);
```

If you are fammilar with frontend frameworks the idea of having an entrypoint for your application
maps well into the shape of how we structure documents with docxcerlate.

We build components which represent the contents of a document, these are basic elements
of a word document such as a paragraph, image, clipart or something else.

```tsx
import { Paragraph, useSetPrompts, useState } from "docxcelerate/template";
import type { TenancyData } from "../types.ts";

export const Balance: Paragraph = () => {
  const [state] = useState((data: TenancyData) => ({
    name: data.recipientName,
    settled: data.balanceDue === 0,
  }));

  useSetPrompts({ generalPrompt: `Explain the balance to ${state.name}.` });

  if (state.settled) {
    return <Paragraph id="settled">Nothing outstanding, {state.name}.</Paragraph>;
  }

  return <Paragraph id="arrears" />;
};
```

We structure documents with hooks that allow us to bring in ai features, at the most basic level
this is just stuff like making a paragraph say X or Y or write something unique.

In the example above we are simply using the generalPrompt to get some text out about the balance a user has left,
you can imagine that if the user has a lot of money the AI may choose to write "You've got a substantial amount of money William".

Jump into our documentation to understand how you can structure documents with AI and create more complex documents for your own needs.

## Documentation

Everything is on [docxcelerate.com](https://docxcelerate.com):

- [Start here](https://docxcelerate.com/docs/start-here/) — install the CLI and scaffold a workspace
- [Writing nodes](https://docxcelerate.com/docs/writing-nodes/) — the pieces a document is made of, and where AI fits
- [Document projects](https://docxcelerate.com/docs/document-projects/) — the files in a document, and what configures them
- [Documents and nodes](https://docxcelerate.com/docs/essentials/documents-and-nodes/) — the model
- [Templates](https://docxcelerate.com/docs/essentials/templates/) — composing with JSX
- [The node reference](https://docxcelerate.com/docs/nodes/overview/) — every node type, rendered
- [CLI commands](https://docxcelerate.com/docs/cli/commands/) — every `dxcl` command
- [The engine](https://docxcelerate.com/docs/generation/endpoint/) — generating at scale
- [Package entrypoints](https://docxcelerate.com/docs/reference/entrypoints/) — what each import gives you

## Development

```sh
npm install       # install dependencies
npm run build     # compile src/ to dist/ with type declarations
npm test          # build, then run the Node test suite
npm run typecheck # type-check sources and tests
```

The published package ships compiled output from `dist/` plus the `dxcl` binary.
Pushing a change to `src/` on `main` publishes a new version.

## License

MIT
