# Docxcelerate

A toolkit for composing DOCX documents from typed JSX components. Published to
npm as `docxcelerate` and to JSR as `@docxcelerate/docxcelerate` from the same
source tree.

## Layout

```
src/          the framework; every published entrypoint is declared in deno.json
templates/    the files a scaffolded project is written from, as real source
tests/        node:test suites, compiled to dist-tests/ before they run
registry/     component sources `dxcl add` copies into a project, read from disk
skills/       the agent skill published alongside the package
scripts/      repository tooling (JSR doc check, prepare hook, launch copy)
website/      a separate private Astro package with its own deps and deploy
conformance/  the conformance suite
```

`src/` is grouped by what a thing does, not by what it is:

- `domain/` — the document model and the types everything else speaks in.
- `template/` — evaluating JSX into elements, and turning elements into nodes.
  Nothing here renders; a build calls the components later.
- `runtime/` — settling a document against data: conditions, derivers, templates,
  the resolver an engine runs.
- `render/` — a settled model out to DOCX bytes.
- `project/` — defining, building, scaffolding and shipping a document project.
- `registry/` — the catalog of installable components and themes, and the installer.
- `theme/`, `transform/` — themes, and the compile step for control flow.

## Commands

```sh
npm test          # clean, build, compile tests, run them
npm run typecheck # build, then type-check tests/ and registry/ too
npm run lint      # deno lint; rules and file list live in deno.json
npm run templates # type-check the files a scaffolded project is written from
npm run build     # tsc -p tsconfig.build.json into dist/
npm run jsr:doc   # every entrypoint and exported symbol is documented
npm run version:guard # a version bump has to ship something
```

Run `npm test` before proposing a change. The suite is the safety net for
anything structural, and it is fast.

`deno publish --dry-run --allow-dirty` is what the JSR score check runs; it
catches slow types and unresolvable imports that `tsc` accepts.

## Conventions

**`letter` is either the engine's wire format or a typographic term — never a
leftover rename.** The vocabulary is "document" everywhere the toolkit owns the
name. What is left is deliberate and must not be "finished":

- `schemaVersion: "docxcelerate.letter/v0"` on every DocumentModel
- `previewLetter` / `engineLetter` on an artifact manifest
- `/api/letters`, and the `stored-letter` upload body
- `letterSpacing`, `Letterhead`, `warm-letter`, the US `LETTER` page size

The first three are what the generation engine reads. Renaming them here does
not finish anything; it moves the mismatch onto the wire. The engine changes
first, or they stay.

**`graph` is the API's name for a chart, and it is not a leftover either.**
`kind: "graph"`, `graphType`, `GraphNode` and `<Graph>` are all on the wire and
in the published types, so renaming them to `chart` would break every document
already written and every model already stored. What they *produce* is an
OOXML chart part, and Word calls it a chart — so the prose says "chart", the
same way the prose says "document" over `letter`. Both names are deliberate:

- `graph`, `graphType`, `GraphNode`, `GraphData`, `<Graph>` — the model and the
  authoring surface, unchanged
- "chart" — what it is called in documentation, in a placeholder a reader sees,
  and in every name under `src/render/` that touches the OOXML part
  (`chart_part.ts`, `docx_charts.ts`, `readPackedCharts`)

**Every module has a `@module` doc comment**, and every exported symbol has a
doc comment. This is not decoration — `npm run jsr:doc` fails the build without
them, and the JSR score depends on it.

**Comments say why, not what.** The code already says what it does. A comment
earns its place by explaining a decision a reader would otherwise have to
reconstruct — why the body renders before the furniture, why a stand-in is
passed through whole rather than copied.

**Export only what callers need.** A helper used inside one module is not
exported from it. Every export in `src/` is either reachable from an entrypoint
in `deno.json` or has no reason to exist.

**Generated project files are real files under `templates/`, never strings.**
If you are about to write TypeScript inside a template literal, put it in a
file instead. `npm run templates` type-checks the whole directory against this
package, which is how a scaffolded workspace type-checks against the published
one.

- Placeholders are `__UPPER_SNAKE__`, written so the file stays valid source —
  an identifier where an identifier goes, a string where a string goes. That is
  what lets them be compiled rather than only pattern-matched. `readTemplate`
  throws on a placeholder nothing was given for.
- A template whose own name would make a tool act on it carries a `.template`
  suffix: `package.json.template`, `tsconfig.json.template`,
  `gitignore.template`. A real `package.json` in there becomes the enclosing
  package for every file beside it and `docxcelerate` stops resolving.
- `templates/document/` is the blank project, `templates/sample/` the worked
  example a new workspace comes with, `templates/workspace/` the workspace
  around them. The node templates live in `templates/document/nodes/` so their
  `../types.ts` import resolves like it will where they land.

**A version bump ships something, or it does not happen.** Merging a version
change to main publishes it to npm and JSR both, so a bump on a branch that
only touched the website or the conformance suite would cut a release
identical to the one before it. `npm run version:guard` refuses that, and CI
runs it on every pull request. The other way round is fine: a change can land
without a bump, and several of them can ship under one number.

**Entrypoints are declared in `deno.json` and mirrored in `package.json`.**
The two export maps have to agree, and every npm target has to exist after a
build. `tests/package.test.ts` fails when they drift, and when the version in
`src/version.ts` disagrees with either of them.

## Things worth knowing

**Packing happens twice for a document with a chart in it.** `docx` owns the
zip and offers no way to be handed a part, so `renderDocxBytes` packs, then
`addChartsToPackage` opens the result and adds `word/charts/chartN.xml`, its
relationships, the workbook behind it and the content-type overrides. The join
is a token: the renderer writes `r:id="dxclChart_<node id>"` into the drawing,
because the real relationship ids are not known until the file exists — an
image or a running strip may have taken the number first. A document with no
charts is returned byte for byte, untouched.

Two failures found there the hard way, both of which produce a file Word
refuses to open outright rather than one that draws wrongly:

- `ImportedXmlComponent.fromXmlString` returns a component *wrapping* what it
  parsed, not the element. Its `w:r` came out inside a tag named `undefined`.
  Build drawings element by element, as `renderShape` does.
- An empty relationships part is written `<Relationships …/>`, self-closing. An
  append that only knew how to insert before `</Relationships>` silently did
  nothing, and the chart part was written with nothing pointing at it.
  `withChild` handles both forms and throws rather than no-op.

`core.autocrlf` is irrelevant here — `.gitattributes` pins the repository to LF.

`deno lint` runs with two React rules off, in deno.json: `jsx-key`, because
this JSX is not React's and a key means something else here, and
`jsx-curly-braces`, because an expression child is the only way to write a
string containing an escape such as a tab.

There is no formatter configured. `deno fmt` would rewrite most of the tree,
and there are a lot of open branches; if you want to adopt it, do it as one
isolated commit and add the hash to `.git-blame-ignore-revs`.

`bin/dxcl.mjs` shells into `dist/`, so the CLI needs a build before it runs.
