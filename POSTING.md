# How to post

Your publishing cheatsheet. Every post is **one Markdown file** in
`src/content/posts/`. Add a file, push to git, and the site rebuilds and goes
live on its own — the same write-a-file-and-publish flow nav.al uses.

## The 3-step workflow

1. **Create a file:** `src/content/posts/my-new-post.md`
   (the filename becomes the URL → `/posts/my-new-post/`)
2. **Paste a template** from below and write.
3. **Publish:**
   ```bash
   git add src/content/posts/my-new-post.md
   git commit -m "Post: My New Post"
   git push
   ```
   Vercel/Netlify redeploys in ~30 seconds.

## Frontmatter reference

The block between the `---` lines at the very top of each file:

| Field      | Required | What it does                                              |
| ---------- | -------- | --------------------------------------------------------- |
| `title`    | yes      | Post title (feed, archive, page, browser tab).            |
| `subtitle` | no       | One-line description shown under the title.               |
| `date`     | yes      | `YYYY-MM-DD`. Controls feed order and the archive year.   |
| `topPost`  | no       | `true` features it under **Top Posts** in the archive.    |
| `audio`    | no       | Path to an audio file → turns the post into an episode.   |
| `duration` | no       | Display length, e.g. `"24:10"` (keep the quotes).         |
| `draft`    | no       | `true` keeps it out of the build while you're writing.    |

---

## Template — written post (essay / note)

```markdown
---
title: Your Title Here
subtitle: One line that hooks the reader
date: 2026-06-16
topPost: false
---

Open with your strongest sentence.

Write in Markdown — **bold**, *italic*, [links](https://example.com), and:

- bullet points
- like this

> Pull-quotes look great in the olive theme.

## A section heading

More text. Push when you're happy with it.
```

---

## Template — podcast episode

1. Drop your audio file in `public/media/` — e.g. `public/media/episode-2.mp3`.
2. Create the post:

```markdown
---
title: "Episode 2: The Topic"
subtitle: Short description of the conversation
date: 2026-06-16
audio: /media/episode-2.mp3
duration: "24:10"
topPost: false
---

Show notes, links, and timestamps go here.
```

The post automatically gets the play button and the **Get podcast** menu — both
in the homepage feed and on its own page.

---

## Tips

- **Preview first:** run `npm run dev` and open <http://localhost:4321> before pushing.
- **Drafts:** set `draft: true` to keep something unpublished; remove it to go live.
- **Feature your best:** set `topPost: true` to pin a post to the top of the archive.
- **Images:** put them in `public/` and reference them as `![alt text](/my-image.jpg)`.
- **Everything else** (name, socials, podcast links, newsletter) lives in
  `src/config.ts`; the theme color is `--accent` in `src/styles/global.css`.
