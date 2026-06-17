import rss from '@astrojs/rss';
import { getPublishedPosts, sortPostsNewestFirst } from '../lib/posts';
import { site } from '../config';

export async function GET(context) {
  const posts = sortPostsNewestFirst(await getPublishedPosts());
  return rss({
    title: site.title,
    description: site.description,
    site: context.site,
    items: posts.map((post) => ({
      title: post.title,
      pubDate: post.date,
      description: post.excerpt,
      link: `/posts/${post.id}/`,
    })),
  });
}
