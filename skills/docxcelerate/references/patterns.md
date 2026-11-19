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
          <Row keepTogether>
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
as the description does above. `keepTogether` keeps a wrapped line whole on one
page. Mark a totals row `header` to draw it as one — it
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

## A list

Items are paragraphs that say so. Write them one after another and the renderer
counts — never type a bullet or a number:

```tsx
export const Steps: Nodes = () => {
  const [state] = useState((data: DocumentData) => ({ steps: data.steps }));

  return (
    <>
      <Paragraph id="steps-intro">What to do next:</Paragraph>
      {state.steps.map((step) => (
        <Paragraph id="step" list="number">{step.text}</Paragraph>
      ))}
    </>
  );
};
```

`list="bullet"` for a marked list, `listLevel={1}` for a point under an item.
Anything between two items ends the list, and a numbered list after it starts
again at one. Published, the `.map()` is a loop the engine walks, and the steps
it writes are one list however many there are. An item on a `<Shape>` is drawn
as a plain paragraph.

## Stress a word, or link a phrase

`<Text>` goes among a paragraph's words. It is not a node, so it takes no `id`:

```tsx
<Paragraph id="terms">
  Payment is due within <Text bold>{state.termDays} days</Text>. You can{" "}
  <Text href={state.payUrl}>pay online</Text>.
</Paragraph>
```

`bold`, `italic`, `underline`, `strike`, `script` and `href`; one inside another takes
both. An interpolated value inside it publishes as a token and is settled per
document like any other. `href` must be `https:`, `http:`, `mailto:` or `tel:`
— anything else prints as plain words. There is no colour or size: a link is
drawn in the theme's accent, and a phrase or a whole paragraph is restyled by
naming a theme block with `variant`.
`underline="double"` draws the double rule accounting puts under a final total — `Total due <Text bold underline="double">{state.total}</Text>` — under the figure alone, where a paragraph border would rule the whole line.

The same works in a cell: `<Cell>Total <Text bold>due</Text></Cell>`.

## Several lines in one paragraph

A line end in the text is a line break, not a new paragraph — one node, one id,
and none of the space a paragraph leaves under itself:

```tsx
<Paragraph id="address">{`${state.line1}\n${state.line2}\n${state.postcode}`}</Paragraph>
```

It has to be a string expression: JSX collapses a line end written between
tags into a space. Do not write three `<Paragraph>`s for one address.

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

## The same lines, placed differently

Where something sits is a separate choice from what it says. A component that
can be placed more than one way takes an `arrangement` prop and picks with
`arrange`, which builds only the arrangement asked for and treats the first one
written as the default.

Keep each arrangement in its own file, and let the node decide every line and
hand them over. A registry component installs this way:

```
nodes/reference.node.tsx
arrangements/reference/lines.ts              // what the node hands each arrangement
arrangements/reference/stacked.arrangement.tsx
arrangements/reference/table.arrangement.tsx
```

```tsx
// arrangements/reference/table.arrangement.tsx
import { Cell, Row, Table, type Yield } from "docxcelerate/template";
import type { ReferenceLines } from "./lines.ts";

export const table = (lines: ReferenceLines): Yield => (
  <Table id="reference-table" columns={[{ width: 45 }, { width: "auto" }]}>
    <Row id="reference-number">
      <Cell>Invoice number</Cell>
      <Cell>{lines.number}</Cell>
    </Row>
    {lines.purchaseOrder && (
      <Row id="reference-purchase-order">
        <Cell>Purchase order</Cell>
        <Cell>{lines.purchaseOrder}</Cell>
      </Row>
    )}
  </Table>
);
```

```tsx
// nodes/reference.node.tsx
import { arrange, Section, useState } from "docxcelerate/template";
import type { ReferenceLines } from "../arrangements/reference/lines.ts";
import { stacked } from "../arrangements/reference/stacked.arrangement.tsx";
import { table } from "../arrangements/reference/table.arrangement.tsx";

export const Reference: Section<{ arrangement?: "stacked" | "table" }> = ({ arrangement }) => {
  const [invoice] = useState((data: Data) => data.invoice);
  const lines: ReferenceLines = { number: invoice.number, purchaseOrder: invoice.purchaseOrder };

  return (
    <Section id="reference" title="Reference">
      {arrange(arrangement, { stacked: () => stacked(lines), table: () => table(lines) })}
    </Section>
  );
};

<Reference arrangement="table" />;
```

Do not write `arrangement === "table" ? … : …`. The compiler publishes a
comparison even of two literals, so every arrangement would ship to the engine
as a branch and count against the branch limit. An unknown name throws, so a
typo fails the build instead of printing the default. Test an optional line
inside the arrangement's JSX (`{lines.purchaseOrder && …}`), never in a
variable above it: outside JSX a publish stand-in is always truthy. Keep a
line's id the same in every arrangement. Registry components that offer
arrangements list them under "Arrangements" in `dxcl show`.

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

A dynamic image takes the same prompt slots. `negativePrompt` earns its
place here more than anywhere: generated imagery fails in predictable ways —
baked-in text, invented logos, recognisable faces — and the slot rules them out
once, in the node.

## A graph

A graph node holds numbers and a form, never a picture. It is packed as a real
Word chart — the reader can select it, restyle it and open its data:

```tsx
export const VisitsByMonth: Graph = () => {
  const [state] = useState((data: TenancyData) => ({
    months: data.visitsByMonth.map((entry) => entry.month),
    visits: data.visitsByMonth.map((entry) => entry.visits),
  }));

  return (
    <Graph
      id="visits-by-month"
      title="Visits by month"
      graphType="bar"
      numberFormat="#,##0"
      valueAxisTitle="Visits"
      data={{
        categories: state.months,
        series: [{ label: "Visits", values: state.visits }],
      }}
      caption="Your visits, last six months"
    />
  );
};
```

Do the running totals, percentages and rebasing in the state initializer. The
chart and the prose beside it are then computed from one source, so they cannot
disagree.

**Leave the colours to the theme.** A series takes its place in the theme's
`palette.series`, which is what lets a document be re-themed without its charts
going out of step with it. Name a `color` only when the colour is what the
series *means* — a red line for the limit it must stay under — and never to
pick a nicer blue.

**Say `null`, not `0`, for a reading nobody took.** They draw differently and
they mean different things: a gap is a month not yet measured, a zero is a
month that measured nothing.

```tsx
values: [12, 18, null, 22]   // March has not been counted yet
```

**Two series of different scales are two charts.** A chart with one axis is a
chart a reader can trust; giving revenue and headcount the same axis makes one
of them a flat line, and giving them two makes the crossing point meaningless.

**Comparing categories is a `bar`; change over time is a `line`; a share of a
whole is a `pie`, and only when there are few enough slices to tell apart.**
More than about six and a bar chart is easier to read than a pie of the same
data.

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

## A component in more than one language

```tsx
const messages = defineMessages("en", {
  due: (amount: string) => `The balance is ${amount}.`,
}, {
  fr: { due: (amount) => `Le solde s’élève à ${amount}.` },
});

export const Due: Paragraph = () => {
  const t = useMessages(messages);
  const format = useFormat();
  const [amount] = useState((data: { amount: number }) => data.amount);

  return <Paragraph id="due">{t.due(format.currency(amount))}</Paragraph>;
};
```

The build's `locale` picks the messages (`fr-CA`, then `fr`, then any `fr-…`,
then the source language), and `useFormat` follows it. A project lists its
languages as `locales: ["en-GB", "fr-FR"]`, default first.

A deriver only sees its inputs, so pass it the locale as one:
`useDeriver(money, [amount, useLocale()])`. Keep the text in the component.
Don't hard-code `"en-GB"` in an `Intl` call.

## An async component

A component may be `async`, but one that fetches is one that can fail mid-build —
prefer putting the value in your data first. If you do await, **every hook must
be called before the first `await`**:

```tsx
import { deriver, Paragraph, useDeriver, useState } from "docxcelerate/document";
import type { TenancyData } from "../types.ts";

const monthsBetween = deriver({
  name: "monthsBetween",
  run: (end: string) => Math.round((Date.parse(end) - Date.now()) / 2_629_800_000),
});

export const Summary: Paragraph = async () => {
  const [state] = useState((data: TenancyData) => ({
    ref: data.propertyRef,
    end: data.endDate,
  }));
  const months = await useDeriver(monthsBetween, [state.end]);

  return <Paragraph id="summary">{state.ref}: {String(months)} months remain.</Paragraph>;
};
```

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
import { createDocxBlob } from "docxcelerate/docx";

const blob = await createDocxBlob(doc);   // the .docx itself
```

There is one renderer, and it writes Word files. To show a document on a screen,
pack it and read the file back with a DOCX viewer — the preview app a workspace
is scaffolded with does exactly that, over `docx-preview`. A second renderer
laying the model out in HTML would be a second answer to what the document looks
like, and the one nobody opens in Word is the one that drifts.
