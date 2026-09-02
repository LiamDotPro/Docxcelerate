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

**Entrypoints are declared in `deno.json` and mirrored in `package.json`.**
The two export maps have to agree, and every npm target has to exist after a
build. `tests/package.test.ts` fails when they drift, and when the version in
`src/version.ts` disagrees with either of them.

## Things worth knowing

`core.autocrlf` is irrelevant here — `.gitattributes` pins the repository to LF.

There is no formatter configured. `deno fmt` would rewrite most of the tree,
and there are a lot of open branches; if you want to adopt it, do it as one
isolated commit and add the hash to `.git-blame-ignore-revs`.

`bin/dxcl.mjs` shells into `dist/`, so the CLI needs a build before it runs.
