# Docxcelerate

Docxcelerate is a TypeScript toolkit and CLI for building DOCX letter projects. It scaffolds a
workspace with a browser preview, lets you compose letters from reusable nodes, and produces plain
JSON artifacts that a document generation endpoint turns into finished documents.

Everything in this repository is free and open source, and runs entirely on Node.js. Previewing and
packing DOCX happens locally in your browser — no server required.

## Requirements

- Node.js 20 or newer

## Getting Started

Clone the repository and build the CLI:

```sh
git clone https://github.com/LiamDotPro/Docxcelerate.git
cd Docxcelerate
npm install
npm run build
```

Create a workspace and open the preview app:

```sh
node ./bin/dxcl.mjs init housing-letters
cd housing-letters
npm run dev
```

`dxcl init` scaffolds a self-contained Vite workspace and installs its dependencies. `npm run dev`
opens the preview site, where you can pick a letter, resolve it with preview data, pack it into a
real `.docx` in the browser, and view it with `docx-preview`.

To use `dxcl` from anywhere, link it once:

```sh
npm link
dxcl init housing-letters
```

## Workspace Layout

```text
housing-letters/
  docxcelerate.config.json
  index.html
  package.json
  preview/
    main.ts
    styles.css
  vite.config.ts
  tsconfig.json
  letters/
```

## CLI

Create a new workspace:

```sh
dxcl init my-letters
dxcl init my-letters --dir workspaces
dxcl init my-letters --blank
```

Run `dxcl init` with no arguments for a guided setup that asks for the template and API endpoint.

Create a letter inside a workspace:

```sh
dxcl letter new arrears-notice --title "Arrears Notice"
```

Add nodes to a letter:

```sh
dxcl letter node letters/arrears-notice repayment-summary --type paragraph --mode dynamic
dxcl letter node letters/arrears-notice customer-photo --type image --mode static
dxcl letter node letters/arrears-notice balance-trend --type graph --mode static
```

The generator writes `nodes/<name>.node.ts` and updates `nodes/index.ts`. Add the generated
component to `letter.tsx` where it should appear in the document.

## Letter Projects

A generated letter project has one entrypoint, `letter.project.ts`, and keeps template, preview
data, styles, types, and reusable nodes separate:

```text
letters/arrears-notice/
  letter.project.ts
  letter-style.ts
  letter.tsx
  preview-data.ts
  types.ts
  nodes/
    index.ts
```

The entrypoint exports `defineLetterProject(...)`:

```tsx
import { defineLetterProject } from "docxcelerate/letter";
import { letterTemplate } from "./letter.tsx";
import { letterStyle } from "./letter-style.ts";
import type { LetterData } from "./types.ts";

export default defineLetterProject<LetterData>({
  id: "arrears-notice",
  name: "Arrears Notice",
  version: "0.1.0",
  template: letterTemplate,
  style: letterStyle,
  previewData: {
    recipientName: "Avery",
    city: "Berlin",
  },
});
```

Templates can be written with TSX:

```tsx
/** @jsxImportSource docxcelerate/template */
import { Letter, Section, template } from "docxcelerate/template";
import type { LetterData } from "./types.ts";
import { Greeting } from "./nodes/index.ts";

export const letterTemplate = template<LetterData>(
  <Letter id="arrears-notice" title="Arrears Notice">
    <Section id="opening" title="Opening">
      <Greeting />
    </Section>
  </Letter>,
);
```

Static and dynamic nodes live in normal TypeScript modules:

```ts
import { dynamicParagraph } from "docxcelerate/letter";
import type { LetterData } from "../types.ts";

export const RepaymentSummary = dynamicParagraph<LetterData>({
  id: "repayment-summary",
  placeholder(data) {
    return `${data.recipientName} has an active repayment plan.`;
  },
  generalPrompt(data) {
    return `Write a concise repayment summary for ${data.recipientName}.`;
  },
});
```

Static nodes render their text locally. Dynamic nodes keep a placeholder for preview and carry
prompts that a document generation endpoint resolves at request time.

## Project Config

Every workspace includes `docxcelerate.config.json`, storing named presets for build and upload
behavior:

```json
{
  "schemaVersion": "docxcelerate.config/v0",
  "activePreset": "local",
  "presets": {
    "local": {
      "build": { "outDir": "build" },
      "upload": {
        "endpoint": "",
        "method": "POST",
        "headers": {},
        "body": "letter"
      }
    }
  }
}
```

Build writes artifacts to `build.outDir` inside the active letter folder, such as
`letters/arrears-notice/build`. Build & upload is enabled when the active preset has an
`upload.endpoint`.

## Build Artifacts

The preview UI build action writes:

```text
letters/<letter-id>/build/
  manifest.json
  preview.json
  letter.json
```

`preview.json` is for local development and includes preview data. `letter.json` is the upload
artifact — it keeps request-time values as placeholders such as `{{data.recipientName}}` so the
endpoint can resolve final data later.

## Document Generation Endpoint

Authoring, previewing, and building are entirely local and free. Turning a `letter.json` artifact
into a finished document — resolving dynamic nodes, applying request-time data, and returning the
DOCX — is done by a document generation endpoint.

A hosted service is available at `https://docxcelerate.thoughtup.deno.net/`, and its server
implementation is not part of this repository. Point a workspace at it during setup:

```sh
dxcl init my-letters --official-server
```

Or configure any endpoint of your own:

```sh
dxcl init my-letters --api-endpoint https://letters.example.com/api/letters
dxcl init my-letters --no-api-endpoint
```

You can also edit `upload.endpoint` in `docxcelerate.config.json` at any time.

## Package Entrypoints

```ts
import { buildLetterDocument, createLetterProjectArtifact } from "docxcelerate";
import { createDocxDocument } from "docxcelerate/docx";
import { defineLetterProject, dynamicParagraph, staticParagraph } from "docxcelerate/letter";
import { Letter, Section, template } from "docxcelerate/template";
import { renderLetterWebsite } from "docxcelerate/renderer";
```

## Development

```sh
npm install       # install dependencies
npm run build     # compile src/ to dist/ with type declarations
npm test          # build, then run the Node test suite
npm run typecheck # type-check sources and tests
```

The published package ships compiled output from `dist/` plus the `dxcl` binary.

## License

MIT
