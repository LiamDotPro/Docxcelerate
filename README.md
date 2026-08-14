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

A document is a tree of components. Each one takes its data as state, decides what
it is with ordinary JavaScript, and returns the nodes it wants.

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
depends on is written down in one place. Setting prompts is what makes a node
dynamic — the mode is inferred, never declared.

Components compose into a template:

```tsx
export const documentTemplate = template<TenancyData>(
  <Document id="tenancy-renewal" title="Tenancy Renewal">
    <Section id="opening" title="Opening">
      <Greeting />
      <Balance />
    </Section>
  </Document>,
);
```

## Documentation

Everything is on [docxcelerate.com](https://docxcelerate.com):

- [Start here](https://docxcelerate.com/docs/start-here/) — install, scaffold, first document
- [Documents and nodes](https://docxcelerate.com/docs/essentials/documents-and-nodes/) — the model
- [Templates](https://docxcelerate.com/docs/essentials/templates/) — composing with JSX
- [Static and dynamic](https://docxcelerate.com/docs/essentials/static-and-dynamic/) — where AI fits
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
