/**
 * The development JSX runtime, which TypeScript reaches for under
 * `"jsx": "react-jsxdev"`.
 *
 * Identical to the production one: building an element is already a pure
 * recording step, so there is no extra checking a dev build could add. It
 * exists so the dev setting resolves rather than failing.
 *
 * @module
 */

export * from "./jsx-runtime.ts";
export { jsx as jsxDEV } from "./jsx-runtime.ts";
