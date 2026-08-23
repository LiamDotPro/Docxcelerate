# How the docs are written

The reference for prose on the docs site. The model is the Vue guide: plain,
second person, and unhurried about explaining *why* before showing *how*.

The old voice was compressed and aphoristic. It read as though it were being
transcribed rather than said to anyone — clever, and hard work. Everything below
is a correction of a real sentence from our own pages.

## The rules

### 1. Write to a reader, in second person

The reader is a person with a job to do. Address them.

> **Before.** A component that is given what it needs does not have to reach for
> it.
>
> **After.** If a component is handed what it needs, it doesn't have to go and
> fetch it.

Use `you` and `your`. Use `we` when walking through something together — "here
we define the columns once". Avoid the passive and the impersonal third person.

### 2. Answer the question the reader actually has

Do not assert a principle and move on. Name the question, then answer it.

> **Before.** That is a rule with a reason: because every dependency passes
> through one declaration, what a component needs is written down rather than
> scattered through the code that uses it.
>
> **After.** Why route everything through one call? Because it puts everything a
> component depends on in one place. You can read the top of the file and know
> what it needs, instead of hunting for it.

"Why does it work this way?" is a good heading. So is "When would you use this?"

### 3. Explain the concept before the code

A code block is evidence, not an introduction. One or two sentences of what this
is and why you'd want it, then the code, then a note on the part that isn't
obvious. Never open a section with a bare fence.

### 4. Say it plainly. Cut the aphorism

The old voice liked maxims. They sound authoritative and teach nothing.

> **Before.** A table whose columns do not line up is not a table.
>
> **After.** Columns are declared on the table rather than per row, so every row
> lines up automatically.

> **Before.** Nudging a paragraph off the bottom of a page is the margins' job.
>
> **After.** Only use `PageBreak` when the break belongs to the document itself —
> like payment details that should start their own page. If you just want to move
> a paragraph down, adjust the margins instead.

### 5. One idea per sentence

Break the em-dash chains. If a sentence has two dashes and a colon, it is at
least three sentences.

> **Before.** A published branch stores both arms, each carrying the condition
> that selects one, and the engine decides per document.
>
> **After.** When you publish, both arms travel with the document. Each one
> carries the condition that selects it, and the engine picks per recipient.

### 6. Frame limits as choices, with a way forward

Vue does this well: a constraint is context, not a scolding.

> **Before.** Text only lives inside a `<Paragraph>`.
>
> **After.** Text needs to live inside a `<Paragraph>`. If you want text next to
> an element, put it in its own paragraph beside it.

When something genuinely isn't supported, say what to do instead.

### 7. Signpost, and let people skip

"For now", "we'll come back to this", "if you're just getting started, skip
this". Tell the reader what they can safely ignore. Link forward rather than
explaining everything at once.

### 8. Reassure where it's earned

If a concept looks intimidating, say so and say it gets easier. Do not be
relentlessly brisk.

## What does not change

- **Be accurate.** The old voice was precise, and that is worth keeping. Plain
  does not mean vague.
- **No marketing.** Do not sell. Explain.
- **Keep the concrete examples.** Real invoices, real tenancy letters, real
  figures. They are the best thing about these docs.
- **Code comments and JSDoc keep the dense voice.** They are read next to the
  code they describe, by someone already in context. This guide is about the
  docs site.

## Quick checklist

- [ ] Does the page say `you` at least once in the first paragraph?
- [ ] Does every code block have a sentence before it saying what it is?
- [ ] Any sentence with two em-dashes? Split it.
- [ ] Any maxim? Replace it with the reason behind it.
- [ ] Any limit stated without a way forward?
- [ ] Read it aloud. Does it sound like a person explaining?
