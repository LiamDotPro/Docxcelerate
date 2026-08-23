/**
 * The registry: everything `dxcl add` can install, themes and components alike.
 *
 * Two kinds of thing sit behind one lookup because that is how they are asked
 * for. Somebody adding a payment summary and somebody setting a document in
 * Slate Report are both saying "give me the version of this that already works",
 * and neither should have to say which drawer it lives in.
 *
 * ```sh
 * dxcl list
 * dxcl add slate-report letterhead signature-block
 * ```
 *
 * @module
 */

import { THEMES } from "../theme/mod.ts";
import type { Theme } from "../theme/theme.ts";
import { COMPONENT_CATEGORIES, COMPONENTS, type RegistryComponent } from "./catalog.ts";

export * from "./catalog.ts";
export { COMPONENT_CATEGORIES, COMPONENTS };

/** Which drawer an entry came out of. */
export type RegistryKind = "theme" | "component";

/** A theme, wearing the same shape as everything else the registry hands back. */
export interface RegistryTheme {
  /** Discriminator, so themes and components can share one lookup. */
  kind: "theme";
  /** The id: the URL slug, and what `dxcl add` is given. */
  id: string;
  /** The theme itself. */
  theme: Theme;
}

/** Anything the registry can hand back. */
export type RegistryEntry = RegistryTheme | RegistryComponent;

/** Every theme, as registry entries. */
export const REGISTRY_THEMES: RegistryTheme[] = THEMES.map((theme) => ({
  kind: "theme",
  id: theme.id,
  theme,
}));

/** Everything installable, themes first. */
export const REGISTRY: RegistryEntry[] = [...REGISTRY_THEMES, ...COMPONENTS];

/**
 * Finds an entry by id, optionally narrowed to one kind.
 *
 * Ids are unique across both kinds — a test enforces it — so the bare form is
 * unambiguous. The kind is accepted anyway because `theme:slate-report` is what
 * somebody types when they want to be sure, and refusing it would be pedantry.
 *
 * @param id The entry's id, with or without a `theme:` or `component:` prefix.
 * @returns The entry, or `undefined` when nothing carries that id.
 */
export function findRegistryEntry(id: string): RegistryEntry | undefined {
  const { kind, bare } = splitRegistryRef(id);

  return REGISTRY.find(
    (entry) => entry.id === bare && (kind === undefined || entry.kind === kind),
  );
}

/**
 * Finds an entry by id, insisting there is one.
 *
 * @param id The entry's id, with or without a kind prefix.
 * @returns The entry.
 * @throws If nothing carries that id, listing what does.
 */
export function registryEntry(id: string): RegistryEntry {
  const entry = findRegistryEntry(id);

  if (!entry) {
    throw new Error(
      `Nothing in the registry is called "${id}".\n` +
        `Themes: ${REGISTRY_THEMES.map((theme) => theme.id).join(", ")}\n` +
        `Components: ${COMPONENTS.map((component) => component.id).join(", ")}`,
    );
  }

  return entry;
}

/**
 * Splits `component:letterhead` into its parts. An unprefixed id, or one
 * carrying a prefix that names no kind, comes back whole and unnarrowed — so a
 * future id with a colon in it is not silently mistaken for a prefix.
 */
export function splitRegistryRef(ref: string): { kind?: RegistryKind; bare: string } {
  const separator = ref.indexOf(":");

  if (separator === -1) {
    return { bare: ref };
  }

  const prefix = ref.slice(0, separator);

  if (prefix !== "theme" && prefix !== "component") {
    return { bare: ref };
  }

  return { kind: prefix, bare: ref.slice(separator + 1) };
}

/** One line per entry, for `dxcl list` and anything else printing a catalog. */
export function registrySummary(entry: RegistryEntry): string {
  return entry.kind === "theme" ? entry.theme.summary : entry.summary;
}

/** The title an entry is printed under. */
export function registryTitle(entry: RegistryEntry): string {
  return entry.kind === "theme" ? entry.theme.title : entry.title;
}
