import type { APIRoute } from "astro";

export const GET: APIRoute = async () => {
  const markdownContent = `# Abhi (@abhieq31)

AI-powered tools from Swift roots to web frontiers. Every commit lands on GitHub for you to fork & remix.

## Navigation

- [About](/about.md)
- [Recent Posts](/posts.md)
- [Archives](/archives.md)
- [RSS Feed](/rss.xml)

## Links

- Twitter: [@abhieq31](https://twitter.com/abhieq31)
- GitHub: [@abhieq31](https://github.com/abhieq31)
- Email: abhi@abhieq31.com

---

*This is the markdown-only version of abhieq31.com. Visit [abhieq31.com](https://abhieq31.com) for the full experience.*`;

  return new Response(markdownContent, {
    status: 200,
    headers: {
      "Content-Type": "text/markdown; charset=utf-8",
      "Cache-Control": "public, max-age=3600",
    },
  });
};
