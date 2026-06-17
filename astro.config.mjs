import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';

// Change `site` to your real domain once you have it (used for RSS, sitemap, OG tags).
export default defineConfig({
  site: 'https://abhieq.vercel.app',
  integrations: [sitemap()],
  markdown: {
    shikiConfig: { theme: 'github-dark' },
  },
});
