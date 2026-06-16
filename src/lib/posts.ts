import { getCollection, type CollectionEntry } from 'astro:content';

export type Post = CollectionEntry<'posts'>;

export async function getPublishedPosts(): Promise<Post[]> {
  try {
    return await getCollection('posts', ({ data }) => !data.draft);
  } catch (error) {
    if (error instanceof Error && /does not exist or is empty/i.test(error.message)) {
      return [];
    }
    throw error;
  }
}

export function sortPostsNewestFirst(posts: Post[]): Post[] {
  return posts.sort((a, b) => b.data.date.valueOf() - a.data.date.valueOf());
}
