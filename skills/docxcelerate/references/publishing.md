# Publishing, derivers and the engine

Local work — authoring, preview, packing a `.docx` — needs nothing but Node. The
**engine** is a separate service, and it is what turns one stored package into a
document per recipient.

## Three steps, two of them not on your machine

1. **Build.** The framework renders the document and packages it. The only local step.
2. **Publish.** The package goes to an engine, which stores it and gives it an address.
3. **Write.** Your application calls the API with a set of data; the engine writes
   the document and returns it.

Publishing and writing are separate. You publish when the document's *wording*
changes — a deploy-shaped event. You call the API every time somebody needs a
document, with no build step and no workspace involved.

## What changes when a document is published

A local build has the data in hand: an `if` means what it says and a `.map()`
walks the list it was given. A published artifact is built **once**, against
stand-in values for a request nobody has made yet. Three rules follow, and the
build enforces all three rather than letting a wrong document reach a recipient.

### 1. Interpolate, don't compute

Interpolating a value works — it becomes the token the engine substitutes:

```tsx
<Paragraph id="greeting">Hello {state.name},</Paragraph>
// published as: "Hello {{data.name}},"
```

Computing on one does not, because the value does not exist yet. Formatting a
currency, totalling a list, comparing a date — those need a **deriver**:

```tsx
import { dataRef, derive, Paragraph, useState } from "docxcelerate/document";

<Paragraph
  id="balance"
  derivers={[derive("currencyLabel", { output: "balanceLabel", inputs: [dataRef("balanceDue")] })]}
>
  Your balance is {"{{derived.balanceLabel}}"}.
</Paragraph>
```

`useFormat` and `useDeriver` compute **during the build**, which is right whenever
the value is already known. Reach for a `derivers` prop when it is not.

### 2. Loops are published as loops

A branch has two arms and both can be published. A loop has as many arms as the
request has entries, and nobody knows that number until a document is written —
so the loop itself is what gets stored:

```tsx
<Repeat over="visits" as="visit">
  <Paragraph>Visit: {"{{ctx.visit.label}}"}</Paragraph>
</Repeat>
```

With real data this walks the collection immediately, so previews show the
repetition rather than a description of it. Published, it stays a loop: the
engine binds each entry under `as`, its position under `indexAs`, and suffixes
child ids with the index so one pass stays distinguishable from the next.

**Calling `.map()` on request data while publishing is an error** that names the
`Repeat` to write instead.

### 3. A decision has to be written as a condition

A published branch stores **both arms**, each carrying the condition that selects
one, and the engine decides per document:

```json
{ "id": "settled", "text": "Nothing outstanding.",
  "when": { "type": "truthy", "ref": { "scope": "data", "path": "settled" } } }
{ "id": "arrears", "text": "A balance remains.",
  "when": { "type": "not", "ref": { "scope": "data", "path": "settled" } } }
```

**A plain `if` does not produce that in the published package.** The intended
authoring surface is an `if` in a component, compiled into a condition — but that
compiler is not in the shipped package yet. An `if` decides at build time, and
while publishing the stand-in is truthy for any path, so a truthy `if` silently
publishes its **true** arm for every recipient.

So for a document that ships to an engine, emit the condition — either with
`branch`, which decides locally and publishes both arms:

```tsx
import { branch, dataPath, Paragraph, truthy } from "docxcelerate/template";

export const Balance: Paragraph = () =>
  branch(
    truthy(dataPath("settled")),
    () => <Paragraph id="balance-settled">Nothing outstanding.</Paragraph>,
    () => <Paragraph id="balance-arrears">A balance remains.</Paragraph>,
  );
```

…or with a `when` prop on the element, when only one arm exists:

```tsx
import { compare, dataPath, literal, Paragraph, refValue } from "docxcelerate/template";

<Paragraph id="arrears" when={compare(refValue(dataPath("balanceDue")), "gt", literal(0))}>
  A balance remains.
</Paragraph>
```

Condition builders, all from `docxcelerate/template`: `truthy(ref)`,
`compare(left, operator, right)` with `"eq" | "ne" | "gt" | "gte" | "lt" | "lte"`,
`and(...)`, `or(...)`. Value expressions: `literal(value)` and `refValue(ref)`.
References: `dataPath(path)`, `ctxPath(path)`, `derivedPath(path)`.

Both forms work locally too — with real data the condition is evaluated and one
arm is taken, exactly as the `if` would have. **A document built only from data you
hold can use a plain `if`;** reach for `branch` when the document will be
published.

Branches multiply: nesting them multiplies what a document carries, so
`branchLimit` (32 by default) fails the build rather than producing a quietly
enormous artifact. Lift the decision into a deriver that yields one value, or
raise the limit if the document really is that conditional.

### And never hand-write a `{{data.…}}` token

Interpolation produces it for you. A `{{data.x}}` written literally into text is
*resolved* during the publish build — against a stand-in that has no value — so it
comes out empty. `{{derived.x}}` is the exception: derivers are preserved, not run,
so that token survives.

### What the publish stand-in refuses

While publishing, the data is a stand-in that answers **any path, however deep**,
and interpolates as the token the engine substitutes. What it refuses is the
operation whose answer it cannot honestly give — because quietly returning `false`
or `NaN` would publish a document that is wrong for every recipient:

| Operation on request data | Result while publishing |
| --- | --- |
| `` `${data.name}` `` or `{state.name}` in text | fine — becomes `{{data.name}}` |
| `data.total > 0`, `Number(data.total)`, arithmetic | throws — branch with an `if`, or use a deriver |
| `.map`, `.filter`, `.length`, `.join`, `.at`, spread, `for…of` | throws — write `<Repeat over="…">` |
| `data.something()` | throws — request data holds values, not behaviour |
| `await data.x` | resolves to `undefined` rather than hanging |

Each message names the fix, so read it rather than working around it. An `if` in a
component is the authoring surface for a condition the engine evaluates per
document; you never write `when` or `branch()` by hand.

## Derivers

A deriver is a named function that computes a value the document then refers to.
`derivers/index.ts` holds them:

```ts
import type { DeriverDefinitions } from "docxcelerate/document";

export const derivers = {
  currencyLabel: ([amount]) =>
    new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" })
      .format(Number(amount ?? 0)),
} satisfies DeriverDefinitions;

export default derivers;
```

A deriver takes `(inputs: unknown[], state: RuntimeState)` and returns anything.
Its result lands in `derived` under the invocation's `output`, so
`{{derived.<output>}}` in any text resolves to it — including string values
inside a `Graph` payload.

Three are registered by default: **`sum`**, **`join`** (non-empty inputs joined
with a space) and **`count`**.

Two ways to invoke one, and the difference is *when*:

| Route | Runs | Use for |
| --- | --- | --- |
| `useDeriver("name")(output, ...inputs)` | now, during this build | a value already known |
| `derivers={[derive("name", { output, inputs })]}` prop | per document, on the engine | a value the request supplies |

Value expressions for inputs: `dataRef(path)`, `ctxRef(path)`,
`derivedRef(path)`, `literalValue(value)` — all from `docxcelerate` (and
re-exported from `docxcelerate/document`).

Referencing a deriver the project has not registered fails the artifact build by
name, rather than at request time.

## Build artifacts

A build writes plain JSON into `build.outDir` inside the document's own
directory — `documents/tenancy-renewal/build/` with the default config.

```text
manifest.json     the index: id, name, version, title, builtAt, style, where the rest is
preview.json      resolved against previewData, dynamic nodes → placeholders
document.json     the upload artifact: request-time values kept as tokens
derivers.js       only when the document references derivers, and only the ones it uses
```

`preview.json` is the local development artifact — it is what the preview app
shows and it never needs a service. `document.json` is the same tree with the
decisions still in it:

```json
{ "kind": "paragraph", "mode": "static", "id": "greeting", "text": "Dear {{data.applicantName}}," }
```

Dynamic nodes carry prompts instead of text:

```json
{ "kind": "paragraph", "mode": "dynamic", "id": "tutor-note",
  "prompts": [
    { "kind": "general",  "text": "Write two warm, specific sentences…" },
    { "kind": "negative", "text": "Do not restate the offer…" },
    { "kind": "system",   "text": "You are an admissions tutor…" }
  ] }
```

**Debugging tip that pays for itself:** if a document is right in preview and
wrong in production, diff `preview.json` against `document.json`. The difference
is exactly the set of values resolved somewhere else.

Because the engine resolves nodes remotely, deriver functions have to travel with
the document — `derivers.js` is that bundle, containing only the derivers the
document actually references. Documents that use none do not get the file.

## Workspace configuration

`docxcelerate.config.json` sits at the top of the workspace and covers every
document in it:

```json
{
  "schemaVersion": "docxcelerate.config/v0",
  "activePreset": "local",
  "presets": {
    "local": {
      "build": { "outDir": "build" },
      "upload": { "endpoint": "", "method": "POST", "headers": {}, "body": "document" }
    },
    "staging": {
      "build": { "outDir": "build" },
      "upload": {
        "endpoint": "https://documents.staging.example.com/api/letters",
        "method": "POST",
        "headers": { "Authorization": "Bearer ${LETTERS_TOKEN}" },
        "body": "document"
      }
    }
  }
}
```

| Field | Meaning |
| --- | --- |
| `build.outDir` | Where artifacts are written, relative to the document directory |
| `upload.endpoint` | The generation endpoint; empty disables upload |
| `upload.method` | HTTP method, normally `POST` |
| `upload.headers` | Sent with the upload request |
| `upload.body` | Which artifact to send — `document` is the request-time one |

`activePreset` selects which one applies, which is how a local setup and a staging
engine sit side by side without editing config between runs. Build & upload
appears in the preview UI only when the active preset has a non-empty
`upload.endpoint`; with it empty everything still builds and you get artifacts on
disk instead of a finished document.

## Editions

**Self-hosted (free)** — a stripped-down engine you run yourself: store a
package, write documents from it, with a reduced feature set.

**Managed cloud** — the full engine, hosted. Not open yet; when it opens it will
have a free tier needing only a sign-up.

They are not the same build, so choose on features rather than hosting preference.

## Working without one

You can go a long way before an engine is needed. Authoring, preview and packing
to `.docx` all run locally, so a document can be written, reviewed and produced
on your machine without publishing anything. An engine is what you add when
documents must be produced by something other than a person at a keyboard — on a
schedule, from a queue, or in response to something in your own system.
