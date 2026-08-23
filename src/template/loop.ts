/**
 * The loop a published document carries.
 *
 * Nothing here is written by hand, and none of it is exported from the package.
 * A loop in a document is an ordinary `.map()` in a component; this is only
 * where that lands when the build is publishing rather than deciding.
 *
 * The distinction is the whole reason the element exists. Building against real
 * data walks the collection now, so `.map()` is the one in the standard library
 * and the entries become ordinary nodes. Publishing cannot do that — the number
 * of entries belongs to a request nobody has made — so the stand-in intercepts
 * `.map()`, runs the body once, and hands back one of these for the renderer to
 * turn into the loop the engine walks.
 *
 * @module
 */

import type { Condition, ReferenceScope } from "../domain/types.ts";
import { type CommonElementProps, host, type HostElementType, type Yield } from "./element.ts";

/** What the stand-in records when it intercepts a `.map()`. */
export interface LoopProps extends CommonElementProps {
  /** Path to the collection, within `overScope`. */
  over: string;
  /** Which bag `over` is read from. */
  overScope: ReferenceScope;
  /** The `ctx` key each entry is bound to. */
  as: string;
  /** The `ctx` key each entry's position is bound to. */
  indexAs: string;
  /** The test a `.filter()` left behind, applied per entry by the engine. */
  where?: Condition;
  /** The body, walked once against a stand-in for one entry. */
  children?: Yield;
}

/** The element a publish-time `.map()` produces. Never written in a template. */
export const Loop: HostElementType<LoopProps, "repeat"> = host("repeat", "Loop");
