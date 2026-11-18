/**
 * The words a component prints, in every language it ships in.
 *
 * A component's strings are a plain object, one per locale, and the build
 * picks the one that fits the document. Nothing is looked up by key at render
 * time and nothing is extracted by a tool: a message is a property, so a typo
 * is a type error, and a sentence with a value in it is a function, so the
 * value lands where that language's grammar puts it rather than where English
 * does.
 *
 * @example
 * ```tsx
 * import { defineMessages, Paragraph, useMessages } from "docxcelerate/template";
 *
 * const messages = defineMessages("en", {
 *   greeting: (name: string) => `Dear ${name},`,
 * }, {
 *   fr: { greeting: (name) => `Chère ${name},` },
 * });
 *
 * export const Greeting: Paragraph = () => {
 *   const t = useMessages(messages);
 *   return <Paragraph id="greeting">{t.greeting("Avery")}</Paragraph>;
 * };
 * ```
 *
 * @module
 */

import { requireContext } from "./context.ts";

/**
 * A component's messages, keyed by locale. The first key is the language the
 * component was written in, and what a document in any other language falls
 * back to.
 *
 * @typeParam M The shape of one locale's messages.
 */
export type Messages<M> = Readonly<Record<string, M>>;

/**
 * Declares a component's messages: the ones it was written in, and its
 * translations.
 *
 * The source decides the shape. Every translation is checked against it, so a
 * translation missing a message, or one whose function takes different
 * arguments, fails to type-check instead of printing English in the middle of a
 * French document.
 *
 * @typeParam M The shape of one locale's messages, taken from the source.
 * @param sourceLocale The locale the source messages are written in, such as `en`.
 * @param source The messages as the component was written.
 * @param translations The same messages in other locales.
 * @returns The messages, ready for {@linkcode useMessages}.
 */
export function defineMessages<M>(
  sourceLocale: string,
  source: M,
  translations: Readonly<Record<string, NoInfer<M>>> = {},
): Messages<M> {
  return { [sourceLocale]: source, ...translations };
}

/**
 * The messages that fit this build's locale.
 *
 * @typeParam M The shape of one locale's messages.
 * @param messages What {@linkcode defineMessages} returned.
 * @returns One locale's messages.
 */
export function useMessages<M>(messages: Messages<M>): M {
  return messagesFor(messages, requireContext("useMessages").locale);
}

/**
 * The locale this document is being built in, as a BCP 47 tag.
 *
 * For the decision messages cannot make on their own — an address block laid
 * out differently, a clause only one jurisdiction needs.
 */
export function useLocale(): string {
  return requireContext("useLocale").locale;
}

/**
 * The messages that fit a locale, outside a build.
 *
 * What a deriver calls, since it runs in the engine where there are no hooks;
 * it is handed the locale as `state.locale`.
 *
 * @typeParam M The shape of one locale's messages.
 * @param messages What {@linkcode defineMessages} returned.
 * @param locale The locale wanted. The source is used when this is absent.
 * @returns One locale's messages.
 */
export function messagesFor<M>(messages: Messages<M>, locale: string | undefined): M {
  const available = Object.keys(messages);

  if (available.length === 0) {
    throw new Error("A component's messages have to include at least one locale.");
  }

  return messages[matchLocale(available, locale) ?? available[0]];
}

/**
 * Picks the available locale closest to the one asked for.
 *
 * `fr-CA` takes `fr-CA`, then `fr`, then any other `fr-…`. Tags compare without
 * regard to case, as BCP 47 says they do.
 *
 * @param available The locales on offer.
 * @param requested The locale wanted.
 * @returns The best of `available`, or `undefined` when none shares a language.
 */
export function matchLocale(
  available: readonly string[],
  requested: string | undefined,
): string | undefined {
  if (!requested) {
    return undefined;
  }

  const lower = available.map((tag) => tag.toLowerCase());
  const subtags = requested.toLowerCase().split("-");

  // Most specific first: drop a subtag at a time, so a script or region that
  // nothing offers falls back to the language rather than to the source.
  for (let length = subtags.length; length > 0; length -= 1) {
    const index = lower.indexOf(subtags.slice(0, length).join("-"));
    if (index !== -1) {
      return available[index];
    }
  }

  const index = lower.findIndex((tag) => tag.split("-")[0] === subtags[0]);
  return index === -1 ? undefined : available[index];
}
