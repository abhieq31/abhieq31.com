# How to post on abhieq31.com

Every post is **one Markdown (`.md`) or MDX (`.mdx`) file** in `src/content/blog/`.
Add a file, push to GitHub `main`, and Vercel rebuilds and publishes automatically
(usually within ~1–2 minutes).

The live site is styled like [nav.al](https://nav.al): dark surface, red accent,
home list with **More ›**, Archive with search + Top Posts + years.

---

## Frontmatter (required block at the top of every post)

```markdown
---
title: "Your Post Title"
subtitle: "Optional one-line deck under the title"
pubDatetime: 2026-07-27T10:00:00Z
description: "One or two sentences for lists, search, and link previews."
tags:
  - writing
  - tech
featured: false
draft: false
---

Write your post body here in Markdown.

## A section heading

Section headings (`##`) render in the red accent, like nav.al.

- bullet points
- **bold**, _italic_, [links](https://example.com)

> Quotes look great in the dark theme.
```

| Field | Required | What it does |
| --- | --- | --- |
| `title` | yes | Post title (page, lists, tab, OG image). Also drives the URL slug. |
| `pubDatetime` | yes | ISO-8601 date. Controls order; must be in the past to show. |
| `description` | yes | Summary for home excerpt, search, and social previews. |
| `subtitle` | no | Deck under the title on home + post page (nav.al style). |
| `tags` | no | Tags (default `others`). Related posts prefer shared tags. |
| `featured` | no | `true` → shows under **Top Posts** on Archive. |
| `draft` | no | `true` keeps it off the published site while you write. |
| `modDatetime` | no | Last-updated date if you revise a published post. |
| `heroImage` / `ogImage` | no | Custom image; otherwise OG is generated. |

### URL slug

The URL is built from the **title** (first two meaningful words), not the filename.

Example: title `"Managers: The Real Bottleneck…"` → `/posts/managers-real`

You can still put files in year folders, e.g. `src/content/blog/2026/my-file.md`.

---

## Option A — Post from GitHub (phone or laptop, no tools)

1. Open [github.com/abhieq31/abhieq31.com](https://github.com/abhieq31/abhieq31.com) → `src/content/blog/`.
2. **Add file → Create new file**.
3. Name it something like `2026/my-topic.md` (folder optional).
4. Paste the template above, write the post.
5. Set `draft: false` when ready (or omit `draft`).
6. **Commit changes** straight to `main`.
7. Vercel deploys — live in about 1–2 minutes at [abhieq31.com](https://abhieq31.com).

---

## Option B — Post locally (preview first)

```bash
cd path/to/abhieq31.com
pnpm install          # first time only
pnpm run dev          # http://localhost:4321

# create src/content/blog/2026/your-slug.md, write it, then:
git add src/content/blog/
git commit -m "Post: Your Post Title"
git push              # Vercel redeploys main
```

Useful commands:

```bash
pnpm run build        # production build (what Vercel runs)
pnpm run preview      # serve the build locally
pnpm run check        # format + lint
```

---

## Option C — Ask me (in this chat)

Say something like:

> new blog post: *How to Think About Leverage*

I’ll scaffold `src/content/blog/<year>/<slug>.md` as a draft with your title,
then you fill the body (or ask me to draft it). When you say **publish**, we
set `draft: false`, commit, and push.

---

## Images

Put files under `public/` (e.g. `public/img/photo.jpg`) and reference with a
leading slash:

```markdown
![alt text](/img/photo.jpg)
```

---

## Checklist before you push

1. `title` + `pubDatetime` + `description` filled in  
2. `draft` is `false` or omitted  
3. `pubDatetime` is not in the future  
4. Optional: `subtitle` for the deck line, `featured: true` for Top Posts  

---

## Housekeeping

- Site name, description, and social links: `src/consts.ts`
- Favicon / OG fallback: `public/favicon.svg`, `public/icon-512.png`
- Full design notes live in the components under `src/components` and `src/styles`
