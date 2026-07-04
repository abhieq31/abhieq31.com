# abhieq31.com

Personal blog and writing site — built to stay out of the way of the words.

**[abhieq31.com →](https://abhieq31.com)**

[![Astro](https://img.shields.io/badge/Astro-5-BC52EE.svg)](https://astro.build)
[![Hosted on Vercel](https://img.shields.io/badge/hosted-Vercel-000000.svg)](https://vercel.com)
[![License: MIT](https://img.shields.io/badge/code-MIT-blue.svg)](LICENSE)

## What it is

I write to find out what I actually think. If an idea can't survive being written down in plain words, it wasn't an idea — it was a mood.

The site matches that philosophy: posts are plain text files in `src/content/blog/`, organized by year. Everything else exists to show them fast — no page heavier than it needs to be, search runs on-device, fonts are self-hosted.

## Engineering choices

| Choice | Why |
| --- | --- |
| **Astro** | Static-first — ships zero JS unless a page needs it |
| **Plain-text content** | No CMS, no database — posts are files in git |
| **Self-hosted fonts** | No third-party font CDN, faster and more private |
| **On-device search** | No search API, no latency, works offline |

Built on [Sat Naing](https://github.com/satnaing)'s excellent [AstroPaper](https://astro-paper.pages.dev/) theme.

## Run locally

| Command | What it does |
| :--- | :--- |
| `pnpm install` | Install dependencies |
| `pnpm run dev` | Dev server at `localhost:4321` |
| `pnpm run build` | Production build to `./dist/` |
| `pnpm run preview` | Preview the build locally |

## License

- **Writing** — [CC BY 4.0](http://creativecommons.org/licenses/by/4.0/)
- **Code** — [MIT](LICENSE)