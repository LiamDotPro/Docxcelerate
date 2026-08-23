# Patterns

Recipes for the shapes that come up. Each one is a whole file or a whole
component, so it can be copied and edited rather than assembled.

## Add a document to a workspace

```sh
dxcl document new tenancy-renewal --title "Tenancy Renewal"
```

Then work outward from the contract, in this order:

1. `types.ts` — what does a document need in order to be written?
2. `preview-data.ts` — one realistic instance. Longest name, largest figure.
3. `nodes/*.node.tsx` — one component per node.
4. `nodes/index.ts` — re-export each one.
5. `document.tsx` — place them in sections, in reading order.
6. `document.project.ts` — already written by the generator; check `id` and `name`.

`npm run dev` and the document appears in the picker. Nothing registers it.

## Add a node to an existing document

```sh
dxcl document node documents/tenancy-renewal next-steps --type paragraph
```

That writes `nodes/next-steps.node.tsx` and adds it to `nodes/index.ts`. It does
**not** place it — open `document.tsx` and add `<Nodes.NextSteps />` at the point
in the document where it belongs. A node that exists but is not placed silently
does nothing.

## The structure file

```tsx
import { Document, Section, template } from "docxcelerate/template";
import * as Nodes from "./nodes/index.ts";
import type { TenancyData } from "./types.ts";

export const documentTemplate = template<TenancyData>(
  <Document id="tenancy-renewal" title="Tenancy Renewal">
    <Section id="opening" title="Opening">
      <Nodes.Greeting />
      <Nodes.Balance />
    </Section>
    <Section id="next-steps" title="Next steps">
      <Nodes.NextSteps />
      <Nodes.SignOff />
    </Section>
  </Document>,
);
```

Structure only. No prose, no conditionals, no data — the template is evaluated at
module scope where no data exists yet.

## Branch on data

Two arms, two ids:

```tsx
export const Balance: Paragraph = () => {
  const [state] = useState((data: TenancyData) => ({ settled: data.balanceDue === 0 }));

  if (state.settled) {
    return <Paragraph id="balance-settled">Nothing outstanding.</Paragraph>;
  }

  return <Paragraph id="balance-arrears">A balance remains.</Paragraph>;
};
```

When the arms differ only in wording, compute the string and keep one id — the
older habit, and still a good one:

```tsx
const [state] = useState((data: TenancyData) => ({
  line: data.settled ? "Nothing outstanding." : "A balance remains.",
}));

return <Paragraph id="balance">{state.line}</Paragraph>;
```

Distinct ids are what let a resolved document record which arm this recipient
got, and they are *required* once a branch is published — an engine stores both
arms, and an id addresses exactly one node.

The **first** form publishes: the build compiles the `if` into a condition, so
both arms travel and the engine picks one per recipient. A ternary and a
`cond && <Node />` compile the same way. The **second** does not — it picks a
*string*, and a value that varies per document is a deriver's job. That is the
line: an arm has to yield a node. See
[publishing.md](publishing.md#3-a-decision-has-to-be-written-as-a-condition).

## Structure computed from data

A template cannot do this; a component can:

```tsx
export const Enclosures: Section = () => {
  const [state] = useState((data: TenancyData) => ({ items: data.enclosures }));

  return (
    <Section id="enclosures" title="Enclosed">
      {state.items.map((item) => <Paragraph id={item.id}>{item.label}</Paragraph>)}
    </Section>
  );
};
```

There is **no `key` prop** — unlike React, elements accept only their own props,
so `key={…}` is a type error. Either give each node a real id, as above, or leave
it out: a node without an id takes one from where it sits, which is what stops a
`map` from demanding names you do not have. Falsy children are skipped, so
`{state.overdue && <Reminder />}` works as written.

The same `.map()` is correct whether the list is known now or only per request.
It walks the collection here, and is published as a loop the engine walks — see
[publishing.md](publishing.md).

## A table of line items

The columns are declared once; the rows are the same `.map()` as anywhere else,
so an invoice with three lines and one with thirty are the same document:

```tsx
export const Charges: Section = () => {
  const { currency, number } = useFormat("en-GB");
  const [state] = useState((data: InvoiceData) => ({ lines: data.lines }));

  return (
    <Section id="charges" title="Charges">
      <Table
        id="lines"
        columns={[
          { width: "auto" },
          { width: 16, align: "right" },
          { width: 26, align: "right" },
        ]}
      >
        <Row header>
          <Cell>Description</Cell>
          <Cell>Qty</Cell>
          <Cell>Amount</Cell>
        </Row>
        {state.lines.map((line) => (
          <Row>
            <Cell>
              <Paragraph>{line.desc}</Paragraph>
              <Paragraph variant="muted">{line.meta}</Paragraph>
            </Cell>
            <Cell>{number(line.qty, { minimumFractionDigits: 1 })}</Cell>
            <Cell>{currency(line.qty * line.rate)}</Cell>
          </Row>
        ))}
      </Table>
    </Section>
  );
};
```

A cell takes text directly; give it paragraphs only when one line is not enough,
as the description does above. Mark a totals row `header` to draw it as one — it
stays where it is, because only the rows a table *opens* with repeat onto a new
page.

## Running header, footer and page numbers

Furniture goes on the `<Document>`, not in the body, because it repeats and sits
outside the margins:

```tsx
export const documentTemplate = template<InvoiceData>(
  <Document id="invoice" title="Invoice" header={<RunningHeader />} footer={<RunningFooter />}>
    <Letterhead />
    <Charges />
    <PageBreak id="to-payment" />
    <Payment />
  </Document>,
);
```

`<PageNumber />` goes in the footer and prints `1 / 2` by default. A build cannot
know the count — an engine writing a longer paragraph changes it — so Word gets
a field it recounts and the preview counts its own pages. Reach for `<PageBreak>`
only when the break is part of what the document *is*; nudging a paragraph off
the bottom of a page is the margins' job.

## Interpolate an engine value into prose

Braces are JSX expressions, so a literal token is written as a string:

```tsx
<Paragraph id="balance">Your balance is {"{{derived.balanceLabel}}"}.</Paragraph>
```

`{{data.x}}`, `{{ctx.x}}` and `{{derived.x}}` are the three scopes.

## Size a prompt to the budget

```tsx
const availableTokens = useAvailableTokens();

useSetPrompts({
  generalPrompt: `Explain the change. At most ${Math.floor(availableTokens / 4)} words.`,
});
```

Static nodes can ignore the budget. Dynamic ones should spend it.

## House style applied to nodes it does not own

Because `useSetPrompts` sets prompts on whatever the *calling* component yields,
a shared hook can carry house voice into any node:

```tsx
export function useHouseVoice() {
  useSetPrompts({
    systemPrompt: "You write for a housing association. Plain English, second person.",
  });
}
```

Any component that calls `useHouseVoice()` gets that system prompt on the node it
returns — and can still override it with a prop, because props win over the hook.

## A wrapper component

A component is a plain function, so a layout piece is a plain function that takes
children:

```tsx
import { Image, Section, type Yield } from "docxcelerate/template";

export const Letterhead: Section<{ children?: Yield }> = ({ children }) => (
  <Section id="letterhead" title="">
    <Image id="logo" src="assets/logo.png" alt="Ashcroft Housing" />
    {children}
  </Section>
);
```

## Carry a value forward through the document

`useShared` reads whatever the components rendered *before* this one left behind.
Order is document order, so a later section can react to an earlier one:

```tsx
export const Total: Paragraph = () => {
  const [charges] = useShared<number[]>("charges", []);
  const { currency } = useFormat();

  return <Paragraph id="total">Total: {currency(charges.reduce((a, b) => a + b, 0))}.</Paragraph>;
};
```

Whatever writes to `"charges"` has to sit above this in the tree.

## An image chosen per recipient

Every field is a prop, computed from state before the element is returned — so a
signature per office or a logo per brand stays inside the component instead of
becoming a branch in the template:

```tsx
export const Signature: Image = () => {
  const [state] = useState((data: TenancyData) => ({
    src: data.signatureUrl,
    manager: data.managerName,
  }));

  return <Image id="signature" src={state.src} alt={`Signed by ${state.manager}`} width={180} />;
};
```

**A `src` that is a URL draws on screen but does not pack.** Packing means
embedding bytes, and nothing is fetched or read from disk while it happens — the
engine writing the document is not on the machine the file was on. A picture
that has to survive into the `.docx` travels as a `data:` URI:

```tsx
import { markPng, markSvg } from "../assets.ts";

<Image id="mark" src={markSvg} fallbackSrc={markPng} alt="" width={8} height={8} />;
```

Word will not embed an SVG without a raster beside it, which is what
`fallbackSrc` is for: the screen draws the sharp SVG, the `.docx` gets the PNG.

A dynamic image takes the same four prompt slots. `negativePrompt` earns its
place here more than anywhere: generated imagery fails in predictable ways —
baked-in text, invented logos, recognisable faces — and the slot rules them out
once, in the node.

## A graph

A graph node holds numbers and a form, never a picture:

```tsx
export const VisitsByMonth: Graph = () => {
  const [state] = useState((data: TenancyData) => ({
    labels: data.visitsByMonth.map((entry) => entry.month),
    visits: data.visitsByMonth.map((entry) => entry.visits),
  }));

  return (
    <Graph
      id="visits-by-month"
      graphType="bar"
      data={{ labels: state.labels, series: [{ name: "Visits", values: state.visits }] }}
      caption="Your visits, last six months"
    />
  );
};
```

Do the running totals, percentages and rebasing in the state initializer. The
chart and the prose beside it are then computed from one source, so they cannot
disagree.

## Placeholder values for a preview

```tsx
export const Example: Paragraph = () => {
  const fake = usePlaceholderData();

  return (
    <Paragraph id="example">
      {fake.name()} of {fake.city()} owes {fake.currency(1284.5)} by {fake.date(30)}.
    </Paragraph>
  );
};
```

Seeded from where the component sits, so the values do not move between builds.

## An async component

A component may be `async`, but one that fetches is one that can fail mid-build —
prefer putting the value in your data first. If you do await, **every hook must
be called before the first `await`**:

```tsx
import { dataRef, Paragraph, useDeriver, useState } from "docxcelerate/document";
import type { TenancyData } from "../types.ts";

export const Summary: Paragraph = async () => {
  const [state] = useState((data: TenancyData) => ({ ref: data.propertyRef }));
  const runDeriver = useDeriver("monthsBetween");     // last hook

  const months = await runDeriver("monthsRemaining", dataRef("startDate"));

  return <Paragraph id="summary">{state.ref}: {String(months)} months remain.</Paragraph>;
};
```

`dataRef` and `derive` live on `docxcelerate` and are re-exported from
`docxcelerate/document`, not from `/template` — which is why a node that uses
them imports its elements from `/document` too.

## Build a document in code

```ts
import { buildProjectPreviewDocument } from "docxcelerate";
import project from "./documents/tenancy-renewal/document.project.ts";

const doc = await buildProjectPreviewDocument(project);
```

That is exactly what the preview app calls. For a template without a project,
`buildDocument(documentTemplate, data)` gives the raw model.

## Render one

```ts
import { renderDocumentWebsite } from "docxcelerate/renderer";

const html = renderDocumentWebsite(doc);   // a complete standalone HTML document
```

It returns its own `<html>` and `<style>`, not a fragment — embed it in an iframe
if it needs to sit inside another page. For a `.docx`, use `createDocxDocument`
from `docxcelerate/docx`.
