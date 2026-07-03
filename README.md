# abhieq31.com

**Writing it down is part of the work.**

The source for [abhieq31.com](https://abhieq31.com) — my personal site and
blog. If you can't say it in one line, you don't understand it yet; the posts
here are where I find out which of my ideas survive being written down.

Built with [Astro](https://astro.build), deployed on [Vercel](https://vercel.com).
Static-first, so the fastest possible page — no server between a reader and a
post — is also the cheapest one to run.

## First principles

- **Content is Markdown, everything else is machinery.** Posts live in
  `src/content/blog/`, organized by year. The rest of the repo exists to render
  them fast and stay out of the way.
- **The reader's device does the least work possible.** Astro ships zero JS by
  default; islands only where interaction earns them.
- **Own the pipes.** Search is local ([Pagefind](https://pagefind.app)), fonts
  are self-hosted, analytics stay minimal. Fewer third parties, fewer failure
  modes, fewer people between the writing and the reading.

## Structure

```text
├── public/               # Static assets (images, fonts, favicon)
├── src/
│   ├── components/       # Reusable UI (React islands under ui/)
│   ├── content/blog/     # The actual writing — Markdown, by year
│   ├── layouts/          # Page templates
│   ├── pages/            # Routes
│   ├── styles/           # Global CSS (Tailwind)
│   └── utils/            # Helpers
├── astro.config.mjs      # Astro configuration
├── vercel.json           # Redirects, headers, CSP
└── LICENSE               # Dual license (CC BY 4.0 + MIT)
```

## Run it

| Command | Action |
| :--- | :--- |
| `pnpm install` | Install dependencies |
| `pnpm run dev` | Dev server at `localhost:4321` |
| `pnpm run build` | Production build to `./dist/` (includes Pagefind index) |
| `pnpm run preview` | Preview the build locally |

## License

Dual-licensed: **writing** under [CC BY 4.0](http://creativecommons.org/licenses/by/4.0/),
**code** under [MIT](LICENSE). See [LICENSE](LICENSE) for details.

## Credit where it's owed

Built on [Sat Naing](https://github.com/satnaing)'s excellent
[AstroPaper](https://astro-paper.pages.dev/) theme — clean architecture is a
gift to everyone who builds on it. Thank you.
