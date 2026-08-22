/**
 * The agent skill, at an address you can paste into an agent.
 *
 * The homepage prints this file so a reader can copy it; this serves the same
 * bytes so nobody has to. "Read https://docxcelerate.com/skill.md" is a whole
 * instruction, and it works in any agent that can fetch a URL.
 *
 * Same source as the homepage block — the file that ships in the repository.
 */
import type { APIRoute } from "astro";
import { agentSkill } from "../skill";
import { MARKDOWN_HEADERS } from "../markdown";

export const prerender = true;

export const GET: APIRoute = () =>
  new Response(`${agentSkill()}\n`, { headers: MARKDOWN_HEADERS });
