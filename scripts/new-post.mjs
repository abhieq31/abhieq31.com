#!/usr/bin/env node
// ─────────────────────────────────────────────────────────────────────────────
// new-post — scaffold a post so you never hand-write frontmatter.
//
//   npm run new -- "My Post Title"
//   npm run new -- "Episode 2: The Topic" --audio episode-2.mp3 --duration 24:10
//
// URL slugs are one meaningful, unique word chosen from the title. For example,
// "Vibe Coding Hardware" becomes /posts/hardware/ when that word is available.
//
// Options:
//   --subtitle "..."   one-line description under the title
//   --audio <file>     turns the post into a podcast episode (file lives in
//                      public/media/); pass just the filename or a /media/ path
//   --duration "24:10" display length for the episode
//   --top              feature it under "Top Posts" in the archive
//   --draft            create it as a draft (kept out of the build)
// ─────────────────────────────────────────────────────────────────────────────
import { writeFile, mkdir, access, readdir } from 'node:fs/promises';
import { join, dirname, basename } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const POSTS_DIR = join(ROOT, 'src', 'content', 'posts');
const STOP_WORDS = new Set([
  'a',
  'an',
  'and',
  'are',
  'as',
  'at',
  'be',
  'by',
  'for',
  'from',
  'how',
  'in',
  'into',
  'is',
  'it',
  'its',
  'new',
  'of',
  'on',
  'or',
  'part',
  'the',
  'to',
  'with',
  'your',
]);

// ── Parse args: first non-flag is the title; --flags may take a value. ──────
const argv = process.argv.slice(2);
const opts = { _: [] };
for (let i = 0; i < argv.length; i++) {
  const a = argv[i];
  if (a.startsWith('--')) {
    const key = a.slice(2);
    const next = argv[i + 1];
    if (key === 'top' || key === 'draft') opts[key] = true;
    else { opts[key] = next; i++; }
  } else {
    opts._.push(a);
  }
}

const title = opts._.join(' ').trim();
if (!title) {
  console.error('Usage: npm run new -- "My Post Title" [--subtitle "..."] [--audio file.mp3] [--duration "24:10"] [--top] [--draft]');
  process.exit(1);
}

const toWord = (word) => word
  .toLowerCase()
  .replace(/['’]/g, '')
  .replace(/[^a-z0-9]+/g, '')
  .trim();

const titleWords = title
  .split(/\s+/)
  .map(toWord)
  .filter(Boolean);

const uniqueWords = [...new Set(titleWords)];
const candidates = uniqueWords
  .filter((word) => word.length > 2 && !STOP_WORDS.has(word))
  // Prefer the last distinctive word: titles usually end on the concrete noun
  // readers remember, e.g. "Vibe Coding Hardware" → "hardware".
  .sort((a, b) => titleWords.lastIndexOf(b) - titleWords.lastIndexOf(a));

if (candidates.length === 0) {
  console.error('✗ Could not find a meaningful one-word slug in the title. Add one distinctive word to the title.');
  process.exit(1);
}

await mkdir(POSTS_DIR, { recursive: true });

const existingSlugs = new Set(
  (await readdir(POSTS_DIR, { withFileTypes: true }))
    .filter((entry) => entry.isFile() && entry.name.endsWith('.md'))
    .map((entry) => basename(entry.name, '.md'))
);
const slug = candidates.find((word) => !existingSlugs.has(word));

if (!slug) {
  console.error(`✗ No unique one-word slug is available from this title. Tried: ${candidates.join(', ')}`);
  console.error('  Change the title by adding a distinctive word that has not been used as a post URL.');
  process.exit(1);
}

const today = new Date().toISOString().slice(0, 10); // YYYY-MM-DD

// YAML needs quotes around titles with a leading-colon-ish value; quote when in doubt.
const needsQuotes = (s) => /[:#"']/.test(s) || /^\s|\s$/.test(s);
const yamlStr = (s) => (needsQuotes(s) ? JSON.stringify(s) : s);

const audio = opts.audio
  ? (opts.audio.startsWith('/') ? opts.audio : `/media/${opts.audio}`)
  : null;

const fm = [
  '---',
  `title: ${yamlStr(title)}`,
  `subtitle: ${opts.subtitle ? yamlStr(opts.subtitle) : ''}`,
  `date: ${today}`,
  ...(audio ? [`audio: ${audio}`, `duration: ${yamlStr(opts.duration || '00:00')}`] : []),
  `topPost: ${opts.top ? 'true' : 'false'}`,
  `draft: ${opts.draft ? 'true' : 'false'}`,
  '---',
  '',
  audio
    ? 'Show notes, links, and timestamps go here.\n'
    : 'Open with your strongest sentence.\n',
].join('\n');

const file = join(POSTS_DIR, `${slug}.md`);

// Don't clobber an existing post, even if the directory listing changed.
try {
  await access(file);
  console.error(`✗ A post already exists at src/content/posts/${slug}.md — add another distinctive word to the title.`);
  process.exit(1);
} catch { /* file doesn't exist — good */ }

await writeFile(file, fm, 'utf8');

console.log(`✓ Created src/content/posts/${slug}.md`);
console.log(`  URL when live:  /posts/${slug}/`);
if (audio) console.log(`  Audio:          drop the file at public${audio}`);
console.log('');
console.log('Next:');
console.log(`  1. Write the post in src/content/posts/${slug}.md`);
console.log('  2. Preview:  npm run dev   (http://localhost:4321)');
console.log(`  3. Publish:  git add -A && git commit -m "Post: ${title}" && git push`);
