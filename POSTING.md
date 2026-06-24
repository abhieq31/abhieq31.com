# How to post on abhieq31.com

Every post is **one Markdown (`.md`) or MDX (`.mdx`) file** in `src/content/blog/`.
Add a file, push to GitHub `main`, and Vercel rebuilds and publishes automatically
(usually within ~1–2 minutes). The repo is public, so deploys are not blocked.

---

## Frontmatter (the block at the top of every post)

```markdown
---
title: "Your Post Title"
pubDatetime: 2026-06-23T10:00:00Z
description: "One or two sentences shown in lists, search, and link previews."
tags:
  - writing
  - tech
featured: false
draft: false
---

Write your post body here in Markdown.

## A section heading

- bullet points
- **bold**, _italic_, [links](https://example.com)

> Quotes look great in the dark theme.
```

| Field         | Required | What it does                                                        |
| ------------- | -------- | ------------------------------------------------------------------- |
| `title`       | yes      | Post title (page, lists, browser tab, OG image).                    |
| `pubDatetime` | yes      | ISO-8601 date/time. Controls order; must be in the past to show.    |
| `description` | yes      | Summary for lists, search, and social/link previews.               |
| `tags`        | no       | List of tags (defaults to `others`). Creates `/tags/<tag>` pages.   |
| `featured`    | no       | `true` pins it to the **Featured** section on the homepage.         |
| `draft`       | no       | `true` keeps it out of the published site while you write.          |
| `modDatetime` | no       | Last-updated date, if you revise a published post.                  |
| `heroImage` / `ogImage` | no | Custom image; otherwise an OG image is generated automatically. |

- **The filename becomes the URL.** `my-first-post.md` → `/posts/my-first-post`.
- You can group posts in year folders (e.g. `src/content/blog/2026/...`) — optional.

---

## Option A — Post from GitHub (no tools, works from phone or laptop)

1. Go to **github.com/abhieq31/abhieq31.com** → `src/content/blog/`.
2. Click **Add file → Create new file**.
3. Name it `your-slug.md`.
4. Paste the template above, write your post.
5. Scroll down, **Commit changes** (commit straight to `main`).
6. Vercel deploys automatically — live in ~1–2 minutes.

> Tip: each post page also has an **"Edit on GitHub"** link that opens the file directly.

---

## Option B — Post locally (preview before publishing)

```bash
pnpm install          # first time only
pnpm run dev          # preview at http://localhost:4321 (do NOT leave running in agents)

# create src/content/blog/your-slug.md and write it, then:
git add -A
git commit -m "Post: Your Post Title"
git push              # Vercel redeploys main automatically
```

Useful commands:

```bash
pnpm run build        # production build (what Vercel runs)
pnpm run check        # format check + lint
```

---

## Images

Put image files in `public/` (e.g. `public/img/photo.jpg`) and reference them with a
leading slash: `![alt text](/img/photo.jpg)`.

---

## Housekeeping

- Delete the seed posts in `src/content/blog/` (`welcome`-style / the sample ones) when
  you publish your own.
- Site name, description, and social links live in `src/consts.ts` and `src/constants.ts`.
- Site logo/avatar is `public/icon-512.png` and `public/favicon.svg`.
