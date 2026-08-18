<!-- Absolute URLs, and markdown rather than <img>: this file is the package
     page on npm and on JSR as well as the front page here, and neither registry
     has the repository around it to resolve a relative path against. The banner
     carries its own dark ground, so it reads the same on a light theme and a
     dark one. Rebuild it, and everything beside it, with `npm run brand` in
     website/ — see Brand assets below. -->

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
npm run jsr:doc   # check every entrypoint and exported symbol is documented
```

`jsr:doc` is what keeps the package's JSR score at nine out of nine. It needs
[Deno](https://deno.com) on your path, and runs on every pull request that
touches `src/`.

The published package ships compiled output from `dist/` plus the `dxcl` binary.
Pushing a change to `src/` on `main` publishes a new version.

## Brand assets

Everything below is cut from one piece of geometry by
`website/scripts/build-brand-assets.mjs`, so no two of them can drift apart. Run
`npm run brand` in `website/` to rebuild the set — the outputs are committed, so
that is only needed when the identity itself changes. Anything carrying type is
rendered by headless Chrome, because the wordmark is a serif from a system stack
that only a browser resolves faithfully.

| Asset | Size | Used for |
| --- | --- | --- |
| [`brand/banner.png`][banner] | 1280×320 | the masthead at the top of this file |
| [`brand/github-social.png`][social] | 1280×640 | this repository's social preview card |
| [`og.png`][og] | 1200×630 | the link card the site announces |
| [`social-square.png`][square] | 1200×1200 | a square post, or a timeline card |
| [`brand/avatar.png`][avatar] | 1024×1024 | a profile picture, wherever one is cropped to a circle |
| [`brand/mark.svg`][mark] | vector | the mark alone, following the reader's colour scheme |
| [`brand/lockup-light.svg`][lockup-light] · [`-dark`][lockup-dark] | vector | mark and wordmark together, on paper or on ink |

The social preview is the one asset no build installs: upload
`brand/github-social.png` under Settings › General › Social preview to set what
a link to this repository unfurls into.

[banner]: https://raw.githubusercontent.com/LiamDotPro/Docxcelerate/main/website/public/brand/banner.png
[social]: https://raw.githubusercontent.com/LiamDotPro/Docxcelerate/main/website/public/brand/github-social.png
[og]: https://raw.githubusercontent.com/LiamDotPro/Docxcelerate/main/website/public/og.png
[square]: https://raw.githubusercontent.com/LiamDotPro/Docxcelerate/main/website/public/social-square.png
[avatar]: https://raw.githubusercontent.com/LiamDotPro/Docxcelerate/main/website/public/brand/avatar.png
[mark]: https://raw.githubusercontent.com/LiamDotPro/Docxcelerate/main/website/public/brand/mark.svg
[lockup-light]: https://raw.githubusercontent.com/LiamDotPro/Docxcelerate/main/website/public/brand/lockup-light.svg
[lockup-dark]: https://raw.githubusercontent.com/LiamDotPro/Docxcelerate/main/website/public/brand/lockup-dark.svg

## License

MIT
