import Parser from 'rss-parser';

export interface SubstackPost {
  id: string;
  title: string;
  subtitle: string;
  date: Date;
  link: string;
  excerpt: string;
  content: string;
  audio?: string;
  duration?: string;
  topPost: false;
}

type CustomItem = {
  'content:encoded': string;
  enclosure?: { url: string; type: string; length: string };
  'itunes:duration'?: string;
};

const parser = new Parser<{}, CustomItem>({
  customFields: {
    item: [
      ['content:encoded', 'content:encoded'],
      'enclosure',
      ['itunes:duration', 'itunes:duration'],
    ],
  },
});

let _cache: SubstackPost[] | null = null;

export async function getPublishedPosts(): Promise<SubstackPost[]> {
  if (_cache) return _cache;
  try {
    const feed = await parser.parseURL('https://abhieq31.substack.com/feed');
    _cache = feed.items.map((item) => {
      const link = item.link ?? '';
      const id = link.split('/p/').at(1)?.split('/')[0] ?? link.split('/').pop() ?? crypto.randomUUID();
      const raw = item['content:encoded'] ?? item.content ?? '';
      const excerpt = (item.contentSnippet ?? raw.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim())
        .slice(0, 260).replace(/\s+\S*$/, '') + '…';
      return {
        id,
        title: item.title ?? '',
        subtitle: item.contentSnippet?.split('\n')[0] ?? '',
        date: new Date(item.pubDate ?? Date.now()),
        link,
        excerpt,
        content: raw,
        audio: item.enclosure?.url,
        duration: item['itunes:duration'],
        topPost: false as const,
      };
    });
    return _cache;
  } catch {
    return [];
  }
}

export function sortPostsNewestFirst(posts: SubstackPost[]): SubstackPost[] {
  return [...posts].sort((a, b) => b.date.valueOf() - a.date.valueOf());
}
