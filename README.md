# Docxcelerate

Compose DOCX documents from small typed components, using the JSX you already write.
Author and preview them locally, then generate finished documents through the engine.

**[Documentation](https://docxcelerate.com/docs/start-here/)** · [docxcelerate.com](https://docxcelerate.com)

## Requirements

Node.js 20 or newer. Nothing else — authoring, preview and DOCX packing all run locally.

## Quick start

```sh
npx docxcelerate init my-documents
cd my-documents
npm run dev
```

This scaffolds a workspace and opens the preview app, where you can pick a document,
resolve it against preview data, and pack it into a real `.docx` in the browser.

To get the `dxcl` binary on your path:

```sh
npm install -g docxcelerate
```

## What a document looks like

Start at the **template**. It is the shape of the whole document in one file: the
sections it has, and the components that fill them.

```tsx
/** @jsxImportSource docxcelerate/template */
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

Nothing runs when that is written. Evaluating the JSX only builds elements, which
is what lets each component decide what it is later, once there is data.

`Greeting` and `Balance` are those **components**. Each lives in its own file, takes
its data as state, and returns the **node** it wants:

```tsx
/** @jsxImportSource docxcelerate/template */
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

Data enters through `useState` and nothing else reaches for it, so what a component
depends on is written down in one place. Branching is an ordinary `if`.

The two arms show where AI fits. The settled arm has its own text, so it is produced
on your machine. The arrears arm has none, so the model writes it from the prompt.
You never declare which is which — a component decides by what it supplies.

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
