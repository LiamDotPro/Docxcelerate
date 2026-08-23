# CLI reference

The CLI is `dxcl`, shipped inside the `docxcelerate` package. Run it one-off with
`npx docxcelerate <command>`, or install the package and call `dxcl` directly.

```sh
npx docxcelerate init my-documents      # one-off, no install
npm install -g docxcelerate             # then: dxcl …
npm install --save-dev docxcelerate     # then: npx dxcl …
```

`npx docxcelerate`, not `npx dxcl` — npx resolves the *package* name, and the
binary inside it is called `dxcl`.

Any command run with **no arguments** switches to a guided setup and asks for what
it needs, so a missing argument is a prompt rather than a failure. Every command
also takes `--force` to overwrite existing files, and `--help`.

## `dxcl init [name]`

Creates a workspace and runs `npm install` in it.

```sh
dxcl init my-documents
dxcl init my-documents --dir workspaces
dxcl init my-documents --blank
dxcl init my-documents --api-endpoint https://documents.example.com/api/letters
```

| Flag | Effect |
| --- | --- |
| `--dir <path>` | Create the workspace inside an existing directory |
| `--sample` | Include a working document to read (the default) |
| `--blank` | Leave `documents/` empty |
| `--official-server` | Point `upload.endpoint` at the official engine |
| `--api-endpoint <url>` | Point `upload.endpoint` at your own service |
| `--no-api-endpoint` | Leave `upload.endpoint` empty; fill it in later |

`--sample` and `--blank` are mutually exclusive, as are the three endpoint flags.
An endpoint given as a bare server root gains `/api/letters`. The endpoint is only
needed for documents with generated prose and is editable in
`docxcelerate.config.json` at any time.

What you get:

```text
my-documents/
  docxcelerate.config.json
  index.html
  package.json
  tsconfig.json          # jsx: react-jsx, jsxImportSource: docxcelerate/template
  vite.config.ts
  .gitignore
  preview/
    main.ts
    styles.css
  documents/
    welcome/             # with --sample
```

Everything above `documents/` is the preview app. `documents/` is your work.

## `dxcl document new [name]`

Creates a document project inside a workspace.

```sh
dxcl document new tenancy-renewal --title "Tenancy Renewal"
```

| Flag | Effect |
| --- | --- |
| `--title <title>` | The document title; defaults to a title-cased name |
| `--dir <path>` | The documents directory; defaults to `documents` |

Writes `document.project.ts`, `document.tsx`, `document-style.ts`,
`preview-data.ts`, `types.ts`, and the `nodes/` and `derivers/` directories.

## `dxcl document node [project-dir] [name]`

Generates a node component and registers it.

```sh
dxcl document node documents/tenancy-renewal next-steps --type paragraph
dxcl document node documents/tenancy-renewal signature  --type image
dxcl document node documents/tenancy-renewal rent-trend --type graph
```

| Flag | Values |
| --- | --- |
| `--type` | `paragraph` (default), `image`, `graph` |

Writes `nodes/<name>.node.tsx` and updates `nodes/index.ts`. It does **not** place
the node in your template — add it to `document.tsx` at the point in the document
where it belongs. The node it writes starts with its content; what makes a node
generated is setting prompts on it, which a component decides with data in hand,
so there is nothing to choose here.

## `dxcl list [themes|components]`

Prints the registry — every theme and prebuilt component the package ships.

```sh
dxcl list
dxcl list themes
dxcl list components
```

| Theme | What it is |
| --- | --- |
| `clean-minimal` | A4, one-inch margins, Aptos over Cambria. The default. |
| `slate-report` | Dense, sober, built to be skimmed. Small-cap headings on slate. |
| `warm-letter` | Georgia on wide margins. Correspondence that reads like a person. |
| `legal-serif` | Times New Roman at 12pt, all black. Convention, on purpose. |
| `bold-brief` | US Letter, an outsized accent title, one page and one point. |

| Component | What it is |
| --- | --- |
| `letterhead` | Who sent this, from where, and when. |
| `recipient-block` | The address, and a greeting that survives a missing name. |
| `payment-summary` | What is owed, by when — and what to say when nothing is. |
| `next-steps` | A generated paragraph with all four prompts already fenced off. |
| `signature-block` | A closing, a signature image, and who signed it. |
| `terms-notice` | Numbered small print, one paragraph and one stable id per clause. |

## `dxcl show <id>`

Prints one entry in full: what it is, the data fields it reads, and where its
files land. Read this before `dxcl add` — the data fields are what you will be
adding to `types.ts` afterwards.

```sh
dxcl show slate-report
dxcl show payment-summary
```

Ids are unique across both kinds, so the bare id is enough. `theme:legal-serif`
and `component:terms-notice` are accepted where you want to be explicit.

## `dxcl add [ids...]`

Installs themes and components into a document project.

```sh
dxcl add slate-report
dxcl add letterhead signature-block
dxcl add payment-summary --project documents/arrears-notice
```

| Flag | Effect |
| --- | --- |
| `--project <dir>` | The document project to install into |
| `--force` | Overwrite files that are already there |

With no `--project`, the command uses the document project you are standing in,
or the only one under `documents/`. Several with none implied is an error rather
than a guess.

**A component** is copied in as source — `nodes/<name>.node.tsx` — and
re-exported from `nodes/index.ts`. It is then your file: no version, never
upgraded behind you, and editing it is the expected next step rather than a
fork. Two things it deliberately leaves to you, and both are follow-up work
printed by the command:

- **The data fields.** Nothing here edits `types.ts`, so the paths the component
  reads (`sender.addressLines`, `account.balanceDue`, …) are printed for you to
  add to the data type and to `preview-data.ts`.
- **Placing it.** Same rule as `dxcl document node` — add `<Letterhead />` to
  `document.tsx` at the point in the document where it belongs.

**A theme** is written out as the project's `document-style.ts`, which
`document.project.ts` already passes through, so the next preview is themed. The
file imports the theme and merges over it one group at a time, so overriding a
single margin keeps the other three:

```ts
import type { DocumentStyle } from "docxcelerate/document";
import { slateReportTheme, themeStyle } from "docxcelerate/themes";

export const documentStyle: DocumentStyle = themeStyle(slateReportTheme, {
  page: { margins: { topMm: 20 } },
});
```

Installing a second theme replaces a `document-style.ts` nothing has touched —
the scaffold's default, or a previous theme install. Once you have written into
it by hand, replacing it needs `--force`.

Installing the same component twice is not an error and does not double the
export line. Installing over a file you have edited is, unless you pass
`--force`.

To install from a script rather than a shell, `installRegistryEntry` and
`resolveInstallOrder` from `docxcelerate/registry/install` do the same work, and
`docxcelerate/registry` is the catalog itself.

## Aliases

```text
dxcl new         -> dxcl document new
dxcl node        -> dxcl document node
dxcl project     -> dxcl init
dxcl registry    -> dxcl list
dxcl scaffold    -> dxcl document new
dxcl letter …    -> dxcl document …     (the old spelling, still accepted)
```

## Workspace scripts

`dxcl init` writes these into the workspace `package.json`:

| Script | Runs |
| --- | --- |
| `npm run dev` | Vite on `127.0.0.1:4507` — the preview app |
| `npm run document:new` | `dxcl document new` |
| `npm run document:node` | `dxcl document node` |
| `npm run documents:check` | `tsc` over every document |

The preview picks a document, resolves it against its preview data, packs it into
a real `.docx` in the browser and reads it back with `docx-preview` — so what you
are looking at is the file itself, not an approximation. Edit anything under
`documents/` and the page reloads with the change in place.
