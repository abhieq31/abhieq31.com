import { defineCollection, z } from "astro:content";
import { glob } from "astro/loaders";
import { SITE } from "@/config";

export const BLOG_PATH = "src/content/blog";

// Build a clean URL slug from the post title — max two meaningful words.
const STOP_WORDS = new Set([
  "a", "an", "and", "are", "as", "at", "be", "by", "for", "from", "how", "i", "in",
  "into", "is", "it", "its", "of", "on", "or", "that", "the", "to", "we", "what",
  "when", "why", "with", "you", "your", "my", "this",
]);

function slugFromTitle(title: string): string {
  const words = title
    .toLowerCase()
    .replace(/['’]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .split(/\s+/)
    .filter(Boolean);
  const meaningful = words.filter(w => w.length > 1 && !STOP_WORDS.has(w));
  return (meaningful.length ? meaningful : words).slice(0, 2).join("-");
}

const blog = defineCollection({
  loader: glob({
    pattern: "**/[^_]*.{md,mdx}",
    base: `./${BLOG_PATH}`,
    // URL = two meaningful words from the title (falls back to the filename).
    generateId: ({ entry, data }) => {
      const slug = typeof data?.title === "string" ? slugFromTitle(data.title) : "";
      return slug || entry.replace(/\.(mdx?|markdown)$/i, "").replace(/\//g, "-");
    },
  }),
  schema: ({ image }) =>
    z.object({
      author: z.string().default(SITE.author),
      pubDatetime: z.coerce.date().catch(() => new Date()),
      modDatetime: z.date().optional().nullable(),
      title: z.string(),
      subtitle: z.string().optional(),
      featured: z.boolean().optional(),
      draft: z.boolean().optional(),
      unlisted: z.boolean().optional(),
      // Resilient: blank/null/single-string tags never break the build.
      tags: z.preprocess(v => {
        if (Array.isArray(v)) {
          const f = v.filter(x => x != null && String(x).trim() !== "");
          return f.length ? f : ["others"];
        }
        if (typeof v === "string" && v.trim() !== "") return [v.trim()];
        return ["others"];
      }, z.array(z.string())),
      ogImage: image().or(z.string()).optional(),
      heroImage: z.string().optional(),
      description: z.string().optional().default(""),
      canonicalURL: z.string().optional(),
      hideEditPost: z.boolean().optional(),
      timezone: z.string().optional(),
      // Additional fields from existing posts
      source: z.string().optional(),
      AIDescription: z.boolean().optional(),
    }),
});

export const collections = { blog };
