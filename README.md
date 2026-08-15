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

This creates a workspace and opens the preview app, where you can pick a document,
fill it with sample data, and pack it into a real `.docx` in the browser.

To get the `dxcl` binary on your path:

```sh
npm install -g docxcelerate
```

## What a document looks like

Start at the **template**. It is the shape of the whole document in one file: the
sections it has, and the components that fill them.

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

That file is structure and nothing else. `Greeting` and `Balance` do not run until
you build the document with real data.

They are the **components**. Each one lives in its own file, reads the data it
needs, and returns a piece of the document:

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

`useState` is the only way data gets in, so everything a component depends on is
listed in one place. Deciding what to say is an ordinary `if`.

The two paragraphs it can return are where AI comes in. If the balance is settled,
the paragraph already has its words and your machine produces it. If it is not, the
paragraph is left empty and the model writes it using the prompt above. You never
mark a paragraph as AI-written: leaving the words out is what asks for that.

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
