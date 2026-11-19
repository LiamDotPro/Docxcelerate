/**
 * Headless Chrome over the DevTools protocol, with nothing but Node's own
 * WebSocket.
 *
 * The other probes drive Chrome through `--dump-dom` and `--screenshot`, one
 * process per picture. A thirteen-page document needs a page loaded once and
 * photographed thirteen times, element by element, which only the protocol
 * offers.
 *
 * @module
 */

import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { setTimeout as sleep } from "node:timers/promises";

const CANDIDATES = [
  process.env.CHROME_PATH,
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  "/usr/bin/google-chrome",
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
].filter(Boolean);

/**
 * Starts Chrome and opens one page in it.
 *
 * @param {{ width?: number, height?: number, scale?: number }} viewport
 * @returns {Promise<{ send: (method: string, params?: object) => Promise<any>, goto: (url: string) => Promise<void>, evaluate: (expression: string) => Promise<any>, close: () => Promise<void> }>}
 */
export async function openChrome({ width = 1000, height = 1200, scale = 1 } = {}) {
  const binary = CANDIDATES.find((candidate) => existsSync(candidate));
  if (!binary) throw new Error("No Chrome found; set CHROME_PATH.");

  const profile = await mkdtemp(join(tmpdir(), "dxcl-cdp-"));
  const chrome = spawn(binary, [
    "--headless=new",
    "--remote-debugging-port=0",
    `--user-data-dir=${profile}`,
    "--no-first-run",
    "--hide-scrollbars",
    "--font-render-hinting=none",
    "about:blank",
  ], { stdio: "ignore" });

  // Chrome writes the port it picked here once it is listening.
  const portFile = join(profile, "DevToolsActivePort");
  for (let i = 0; i < 100 && !existsSync(portFile); i++) await sleep(100);
  const [port] = (await readFile(portFile, "utf8")).split("\n");

  const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
  const target = targets.find((t) => t.type === "page");
  const socket = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => {
    socket.onopen = resolve;
    socket.onerror = reject;
  });

  let nextId = 1;
  const pending = new Map();
  const listeners = new Set();
  socket.onmessage = ({ data }) => {
    const message = JSON.parse(data);
    if (message.id && pending.has(message.id)) {
      const { resolve, reject } = pending.get(message.id);
      pending.delete(message.id);
      if (message.error) reject(new Error(`${message.error.message} (${message.error.data ?? ""})`));
      else resolve(message.result);
    } else if (message.method) {
      for (const listener of listeners) listener(message);
    }
  };

  const send = (method, params = {}) =>
    new Promise((resolve, reject) => {
      const id = nextId++;
      pending.set(id, { resolve, reject });
      socket.send(JSON.stringify({ id, method, params }));
    });

  await send("Page.enable");
  await send("Runtime.enable");
  await send("Emulation.setDeviceMetricsOverride", {
    width,
    height,
    deviceScaleFactor: scale,
    mobile: false,
  });

  const evaluate = async (expression) => {
    const { result, exceptionDetails } = await send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (exceptionDetails) throw new Error(exceptionDetails.exception?.description ?? exceptionDetails.text);
    return result.value;
  };

  const goto = async (url) => {
    const loaded = new Promise((resolve) => {
      const listener = (message) => {
        if (message.method === "Page.loadEventFired") {
          listeners.delete(listener);
          resolve();
        }
      };
      listeners.add(listener);
    });
    await send("Page.navigate", { url });
    await loaded;
  };

  const close = async () => {
    socket.close();
    chrome.kill();
    await sleep(200);
    await rm(profile, { recursive: true, force: true }).catch(() => {});
  };

  return { send, goto, evaluate, close };
}

/**
 * Screenshots one rectangle of the page, in CSS pixels, as PNG bytes.
 *
 * @param {{ send: Function }} page
 * @param {{ x: number, y: number, width: number, height: number }} rect
 * @param {number} scale device pixels per CSS pixel in the result
 */
export async function capture(page, rect, scale = 1) {
  const { data } = await page.send("Page.captureScreenshot", {
    format: "png",
    captureBeyondViewport: true,
    clip: { ...rect, scale },
  });
  return Buffer.from(data, "base64");
}
