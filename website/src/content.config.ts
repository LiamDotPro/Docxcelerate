import { defineCollection, z } from "astro:content";
import { glob } from "astro/loaders";
import { DOC_GROUPS } from "./site";

const docs = defineCollection({
  loader: glob({ pattern: "**/*.{md,mdx}", base: "./src/content/docs" }),
  schema: z.object({
    title: z.string(),
    description: z.string(),
    /** Sidebar grouping. Order comes from DOC_GROUPS in site.ts. */
    group: z.enum(DOC_GROUPS),
    /**
     * Optional second level inside the group, for groups big enough to need
     * one. Subgroups appear in the order their first page does.
     */
    subgroup: z.string().optional(),
    /** Position within the group; lower sorts first. */
    order: z.number().default(100),
  }),
});

export const collections = { docs };
