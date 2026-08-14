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

A document is a tree of typed components. Static nodes render locally; dynamic ones
carry prompts and a placeholder, and the engine resolves them at request time.

```tsx
/** @jsxImportSource docxcelerate/template */
import { Document, Section, template } from "docxcelerate/template";
import type { TenancyData } from "./types.ts";
import { Greeting } from "./nodes/index.ts";

export const documentTemplate = template<TenancyData>(
  <Document id="tenancy-renewal" title="Tenancy Renewal">
    <Section id="opening" title="Opening">
      <Greeting />
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
