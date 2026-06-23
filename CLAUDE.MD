# Instructions for Claude

Personal website for Abhi, built on Astro (AstroPaper theme) + Tailwind + React + MDX, deployed on Vercel. Tooling: pnpm + oxlint/oxfmt, Node 24+.

## Guidelines

- **Never create blog posts without explicit consent.** Only edit or fix existing posts when asked.
- **Always use the latest stable dependencies. Never downgrade.** Check versions with tooling (`pnpm outdated`, `pnpm view <pkg> version`) — never guess.
- **Do not run `pnpm run dev` in agent mode** (it never exits). Use `pnpm run build` to verify changes.

## Common tasks

1. **New blog post:** add a Markdown/MDX file under `src/content/blog/` (organized by year). Frontmatter needs `title`, `pubDatetime` (ISO-8601: `YYYY-MM-DDTHH:MM:SS+HH:MM`), `description`, and `tags` (array).
2. **Site config:** `src/consts.ts` for global constants; `src/constants.ts` for social/share links.
3. **Styling:** `src/styles/` plus Tailwind classes in components.
4. **Layouts:** `src/layouts/`.

## Verify

```bash
pnpm run build       # build the production site
pnpm run check       # format:check + lint
```
