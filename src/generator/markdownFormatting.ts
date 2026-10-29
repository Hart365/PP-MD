/**
 * @file markdownFormatting.ts
 * @description Low-level Markdown formatting primitives shared by the documentation generator
 * and its section renderers (escaping, headings, anchors, blockquotes, and sorting).
 */

/**
 * Escapes characters special to Markdown tables (pipe and backslash) and
 * converts line breaks to `<br>` so multi-line text cannot split a table row.
 *
 * @param text - Raw text to escape
 * @returns Escaped text safe for use in a Markdown table cell
 */
export function mdEscape(text: string | undefined | null): string {
  if (!text) return '';
  return text
    .replace(/\\/g, '\\\\')
    .replace(/\|/g, '\\|')
    .trim()
    .replace(/\s*[\r\n]+\s*/g, '<br>');
}

/** Formats text as a GFM blockquote, prefixing every line with `> `. */
export function blockquote(text: string): string {
  return text
    .trim()
    .split(/\r?\n/)
    .map((line) => (line.trim() ? `> ${line.trimEnd()}` : '>'))
    .join('\n');
}

/**
 * Returns a Markdown heading at the requested level.
 *
 * @param level - Heading level 1–6
 * @param text  - Heading text
 */
export function heading(level: number, text: string): string {
  return `${'#'.repeat(Math.min(6, Math.max(1, level)))} ${text}`;
}

/** Returns the GitHub-style anchor (`#slug`) generated for a heading. */
export function headingAnchor(text: string): string {
  return `#${text
    .toLowerCase()
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-')}`;
}

/** Sorts a copy of the items by a case-insensitive label. */
export function sortByLabel<T>(items: T[], labelFn: (item: T) => string): T[] {
  return [...items].sort((a, b) => labelFn(a).localeCompare(labelFn(b), undefined, { sensitivity: 'base' }));
}

/** Builds a Markdown table from headers and pre-escaped cell values. */
export function markdownTable(headers: string[], rows: string[][]): string[] {
  return [
    `| ${headers.join(' | ')} |`,
    `|${headers.map((header) => '-'.repeat(Math.max(3, header.length + 2))).join('|')}|`,
    ...rows.map((row) => `| ${row.map((cell) => cell || '–').join(' | ')} |`),
  ];
}
