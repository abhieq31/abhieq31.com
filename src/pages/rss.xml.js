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
      title: post.data.title,
      pubDate: post.data.date,
      description: post.data.subtitle ?? '',
      link: `/posts/${post.id}/`,
    })),
  });
}
