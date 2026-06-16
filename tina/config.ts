import { defineConfig } from 'tinacms';

export default defineConfig({
  branch: process.env.TINA_BRANCH || process.env.HEAD || 'main',
  clientId: process.env.TINA_CLIENT_ID || null,
  token: process.env.TINA_TOKEN || null,

  build: {
    outputFolder: 'admin',
    publicFolder: 'public',
  },
  media: {
    tina: {
      mediaRoot: 'media',
      publicFolder: 'public',
    },
  },

  schema: {
    collections: [
      {
        name: 'post',
        label: 'Posts',
        path: 'src/content/posts',
        format: 'md',
        ui: {
          filename: {
            // Use the last significant word of the title as the filename/slug.
            // Mirrors the logic in scripts/new-post.mjs.
            slugify: (values) => {
              const stopWords = new Set([
                'a','an','and','are','as','at','be','by','for','from','how',
                'in','into','is','it','its','new','of','on','or','part','the',
                'to','with','your',
              ]);
              const words = (values.title || '')
                .toLowerCase()
                .replace(/[^a-z0-9\s]/g, '')
                .split(/\s+/)
                .filter((w: string) => w.length > 2 && !stopWords.has(w));
              return words[words.length - 1] || 'post';
            },
          },
        },
        fields: [
          {
            type: 'string',
            name: 'title',
            label: 'Title',
            isTitle: true,
            required: true,
          },
          {
            type: 'string',
            name: 'subtitle',
            label: 'Subtitle',
          },
          {
            type: 'datetime',
            name: 'date',
            label: 'Date',
            required: true,
            ui: {
              dateFormat: 'YYYY-MM-DD',
              timeFormat: undefined,
            },
          },
          {
            type: 'boolean',
            name: 'topPost',
            label: 'Top Post',
            description: 'Show under "Top Posts" on the archive page',
          },
          {
            type: 'string',
            name: 'audio',
            label: 'Audio URL',
            description: 'Path to audio file, e.g. /media/episode-1.mp3',
          },
          {
            type: 'string',
            name: 'duration',
            label: 'Duration',
            description: 'Display length, e.g. 24:10',
          },
          {
            type: 'boolean',
            name: 'draft',
            label: 'Draft',
            description: 'Keep this post hidden from the live site',
          },
          {
            type: 'rich-text',
            name: 'body',
            label: 'Body',
            isBody: true,
          },
        ],
      },
    ],
  },
});
