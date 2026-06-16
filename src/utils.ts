// Format a date like "Jun 1 2026" (no comma), matching the minimal feed style.
export function formatDate(date: Date): string {
  return new Intl.DateTimeFormat('en-US', {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
  })
    .format(date)
    .replace(',', '');
}

// Build a short plain-text excerpt from raw markdown for the homepage feed.
export function excerpt(body: string, len = 260): string {
  const text = (body || '')
    .replace(/^---[\s\S]*?---/, '')        // strip frontmatter if present
    .replace(/```[\s\S]*?```/g, ' ')        // code blocks
    .replace(/!\[[^\]]*\]\([^)]*\)/g, ' ')  // images
    .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1') // links -> link text
    .replace(/[#>*_`~]/g, ' ')              // markdown symbols
    .replace(/\s+/g, ' ')
    .trim();
  if (text.length <= len) return text;
  return text.slice(0, len).replace(/\s+\S*$/, '') + '…';
}
