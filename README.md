# Hi, I'm Abhi 👋

I write essays, keep notes, and run a podcast — all from one place I own and control.
No algorithms. No feed. Just ideas, posted when they're worth sharing.

**→ [abhieq31.com](https://abhieq31.com)**

[![X](https://img.shields.io/badge/X-%40abhieq31-black?logo=x&logoColor=white)](https://x.com/abhieq31)
[![Instagram](https://img.shields.io/badge/Instagram-%40abhieq31-E4405F?logo=instagram&logoColor=white)](https://instagram.com/abhieq31)
[![YouTube](https://img.shields.io/badge/YouTube-%40abhieq31-FF0000?logo=youtube&logoColor=white)](https://youtube.com/@abhieq31)

---

## This repo — abhieq31.com

My personal site — essays, notes, and a podcast. A minimal, content-first site
built with [Astro](https://astro.build). Dark, editorial theme — a single red
accent on charcoal, inspired by [nav.al](https://nav.al).

Each post is a Markdown file. Write one, `git push`, and the site rebuilds and
deploys itself.

## Quick start

```bash
npm install      # once
npm run dev      # local preview at http://localhost:4321
npm run build    # production build into dist/
npm run preview  # preview the production build
```

## Write a new post

Fastest — let the scaffold create the file and frontmatter for you:

```bash
npm run new -- "My Post Title"
# podcast episode:
npm run new -- "Episode 2: The Topic" --audio episode-2.mp3 --duration "24:10"
```

Then write, `git push`, and it's live. Full options and templates are in
[`POSTING.md`](./POSTING.md).

Or do it by hand — add a Markdown file in `src/content/posts/`. The filename
becomes the URL (`my-post.md` → `/posts/my-post/`).

```markdown
---
title: My Post Title
subtitle: An optional one-line description
date: 2026-06-16
topPost: false   # set true to feature it under "Top Posts" in the archive
draft: false     # set true to keep it out of the build
---

Write the post body in Markdown here.
```

### Podcast / audio episodes

Add `audio` and `duration` to the frontmatter and the post gets a play button
plus a "Get podcast" menu (just like the homepage feed):

```markdown
---
title: "Episode 2: The Topic"
subtitle: Short show description
date: 2026-06-16
audio: /media/episode-2.mp3   # drop the file in public/media/
duration: "24:10"
---
```

## Make it yours

Almost everything lives in **`src/config.ts`**:

- **Name / brand** — the logo text and titles (`name`, `title`).
- **Socials** — your X, Instagram, YouTube (already set to `abhieq31`).
- **Podcast links** — the row on the homepage and the "Get podcast" menu.
- **Newsletter** — paste your provider's form endpoint into `newsletter.action`
  to make the Subscribe form live (Buttondown, Substack, ConvertKit, Formspree…).

**Theme color:** change `--accent` (and `--accent-hover`) at the top of
`src/styles/global.css` to re-skin the whole site.

**Domain:** set `site` in `astro.config.mjs` to your real domain (used for RSS,
sitemap, and social preview tags).

## Deploy

This builds to plain static files, so it runs anywhere.

- **Vercel** — Import the repo. Astro is auto-detected (build `npm run build`,
  output `dist`). Add your custom domain in the project settings.
- **Netlify** — Import the repo. `netlify.toml` already sets build = `npm run build`
  and publish = `dist`. Add your custom domain in Site settings → Domain.

Both give a free HTTPS subdomain immediately and let you attach a custom domain.

## What's included

- Homepage feed of all posts (essays + audio).
- `/archive` with live search, **Top Posts**, and posts grouped by year.
- Individual post pages with a minimal audio player for episodes.
- `/subscribe` page wired to your newsletter provider.
- `/rss.xml` feed and a sitemap.
