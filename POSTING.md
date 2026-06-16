# How to post

Your publishing cheatsheet. Every post is **one Markdown file** in
`src/content/posts/`. Add a file, push to git, and the site rebuilds and goes
live on its own — the same write-a-file-and-publish flow nav.al uses.

## The fastest way — one command

Let the scaffold write the file and frontmatter for you:

```bash
npm run new -- "My New Post"
```

That creates a one-word URL file chosen from the title, such as `src/content/posts/post.md` for "My New Post" or `src/content/posts/hardware.md` for "Vibe Coding Hardware" (dated today, ready to write).
Options:

```bash
npm run new -- "My New Post" --subtitle "One line that hooks the reader"
npm run new -- "Episode 2: The Topic" --audio episode-2.mp3 --duration "24:10"
npm run new -- "Half-finished idea" --draft     # keep it out of the build
npm run new -- "My best essay" --top            # feature under Top Posts
```

The scaffold chooses the last distinctive unused title word because that is usually the concrete word readers remember. For example, "Vibe Coding Hardware" becomes `/posts/hardware/`. If the chosen word is already used by another post, the script tries another meaningful word from the title; if none are unique, add a more specific word to the title.

Then write, and publish:

```bash
git add -A
git commit -m "Post: My New Post"
git push          # Vercel/Netlify redeploys in ~30 seconds
```


## From your phone — no terminal

Use the GitHub mobile app or github.com in your mobile browser:

1. Open `src/content/posts/` in the repo.
2. Tap **Add file** → **Create new file**.
3. Name it with one unique word from the title, e.g. `hardware.md` for
   "Vibe Coding Hardware".
4. Paste the template below, write the post, then commit directly to the branch.

```markdown
---
title: Vibe Coding Hardware
subtitle: The new vertical integration
date: 2026-06-16
topPost: false
draft: false
---

Write from your phone here.
```

That is the mobile equivalent of `npm run new`: one Markdown file, one-word
filename, commit, and the deploy rebuilds automatically.

## The manual way — 3 steps

1. **Create a file:** `src/content/posts/keyword.md`
   (use one meaningful word from the title; the filename becomes the URL → `/posts/keyword/`)
2. **Paste a template** from below and write.
3. **Publish:**
   ```bash
   git add src/content/posts/keyword.md
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

> Pull-quotes look great in the dark theme.

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
- **URLs:** keep each post filename to one unique word from the title, e.g. `hardware.md` → `/posts/hardware/`.
- **Everything else** (name, socials, podcast links, newsletter) lives in
  `src/config.ts`; the theme color is `--accent` in `src/styles/global.css`.

## Decap CMS on Netlify — easiest mobile editor

This repo now includes a Decap CMS admin at `/admin/`. After deploying on Netlify:

1. In Netlify, enable **Identity** for the site.
2. Set registration to **Invite only**.
3. Enable **Git Gateway** under Identity services.
4. Invite yourself as a user.
5. Open `/admin/` on desktop or mobile, sign in, and create posts from the form.

For the URL field, enter one lowercase word from the title, e.g. `hardware` for
"Vibe Coding Hardware". Decap saves posts into `src/content/posts/`, commits to
Git, and Netlify rebuilds the site.
