import { defineCollection, z } from 'astro:content';
import { glob } from 'astro/loaders';

// Each post is a Markdown file in src/content/posts/.
// To publish: add a .md file with the frontmatter below and push to git.
const posts = defineCollection({
  loader: glob({ pattern: '**/*.md', base: './src/content/posts' }),
  schema: z.object({
    title: z.string(),
    subtitle: z.string().optional(),
    date: z.coerce.date(),
    // Show this post under "Top Posts" on the archive page.
    topPost: z.boolean().default(false),
    // Audio post (podcast episode): URL to the audio file + a display duration.
    audio: z.string().optional(),
    duration: z.string().optional(),
    // Set true to keep a draft out of the build.
    draft: z.boolean().default(false),
  }),
});

export const collections = { posts };
