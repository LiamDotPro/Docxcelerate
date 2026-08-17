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

## Aliases

```text
dxcl new         -> dxcl document new
dxcl node        -> dxcl document node
dxcl project     -> dxcl init
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
