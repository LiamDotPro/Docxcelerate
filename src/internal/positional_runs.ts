/**
 * The function each `deriver()` wrapper unpacks its inputs into.
 *
 * `deriver()` takes a computation written with positional arguments and hands
 * back a `run` that takes the array the engine calls with. That wrapper closes
 * over the options it was given, and a bundle is shipped as source text, so the
 * wrapper's text names a variable the engine never has. The bundler looks the
 * wrapper up here and ships the computation itself.
 *
 * Kept out of both modules' public exports: it is how two of them agree, not
 * something a caller has a use for.
 *
 * @module
 */

import type { DeriverFunction } from "../runtime/derivers.ts";

/** Each `deriver()` wrapper, to the positional function it unpacks into. */
export const positionalRuns: WeakMap<DeriverFunction, (...inputs: never[]) => unknown> = new WeakMap();
