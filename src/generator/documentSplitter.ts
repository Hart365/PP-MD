/**
 * @file documentSplitter.ts
 * @description Post-processing for generated Markdown: Back to Top links, per-category
 * document splitting, and separation of Mermaid diagrams into a companion document.
 */

import { heading, headingAnchor } from './markdownFormatting';
import { unified } from 'unified';
import remarkParse from 'remark-parse';

function sectionHeadings(markdown: string): Array<{ index: number; title: string }> {
  return unified().use(remarkParse).parse(markdown).children.flatMap((node) => {
    if (node.type !== 'heading' || node.depth !== 2 || node.position?.start.offset === undefined || node.position.end.offset === undefined) return [];
    return [{ index: node.position.start.offset, title: markdown.slice(node.position.start.offset, node.position.end.offset).replace(/^##\s+/, '').trim() }];
  });
}

function categoryDocument(header: string, sections: Array<{ heading: string; body: string }>): string {
  const contents = [heading(2, 'Table of Contents'), '', ...sections.map((section) => `- [${section.heading}](${headingAnchor(section.heading)})`)];
  return appendBackToTopLinks([header, contents.join('\n'), ...sections.map((section) => section.body)].join('\n\n'));
}

/**
 * Appends a Back to Top link at the end of each level-2 section.
 *
 * The link targets the generated Table of Contents heading anchor.
 */
export function appendBackToTopLinks(markdown: string): string {
  const lines = markdown.split('\n');
  const sectionStarts: number[] = [];
  let inCodeFence = false;

  lines.forEach((line, index) => {
    if (line.startsWith('```')) {
      inCodeFence = !inCodeFence;
      return;
    }
    if (!inCodeFence && /^##\s+/.test(line)) {
      sectionStarts.push(index);
    }
  });

  const inserts: Array<{ index: number }> = [];

  sectionStarts.forEach((start, idx) => {
    const title = lines[start].replace(/^##\s+/, '').trim();
    if (title.toLowerCase() === 'table of contents') return;

    const nextStart = idx < sectionStarts.length - 1 ? sectionStarts[idx + 1] : lines.length;
    const sectionText = lines.slice(start, nextStart).join('\n');
    if (sectionText.includes('[Back to Top](#table-of-contents)')) return;

    let insertAt = nextStart;
    while (insertAt > start + 1 && lines[insertAt - 1].trim() === '') {
      insertAt -= 1;
    }

    inserts.push({ index: insertAt });
  });

  inserts
    .sort((a, b) => b.index - a.index)
    .forEach(({ index }) => {
      lines.splice(index, 0, '', '[Back to Top](#table-of-contents)', '');
    });

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Category-grouped export (splits an already-generated document into logical files)
// ---------------------------------------------------------------------------

export interface MarkdownCategoryFile {
  /** Stable, filename-safe category key, e.g. 'tables-and-relationships' */
  key: string;
  /** Human-readable category title */
  title: string;
  /** Self-contained Markdown for this category, including a shared solution header */
  markdown: string;
}

/** Logical groupings of generated H2 sections used by the "export by category" feature. */
const CATEGORY_DEFINITIONS: Array<{ key: string; title: string; headings: string[] }> = [
  { key: 'tables-and-relationships', title: 'Tables & Relationships', headings: ['Entity Relationship Diagram', 'Tables & Columns', 'Global Option Sets', 'Forms & Views'] },
  { key: 'apps-and-agents', title: 'Apps & Agents', headings: ['Power Apps', 'Copilot Studio Agents', 'AI Models'] },
  { key: 'automation-and-processes', title: 'Automation & Processes', headings: ['Processes & Automation', 'Desktop Flows', 'Dataflows', 'Custom APIs', 'Offline Profiles'] },
  { key: 'plugins', title: 'Plugins & Plugin Steps', headings: ['Plugin Assemblies & Steps'] },
  { key: 'security-and-integration', title: 'Security & Integration', headings: ['Security Roles', 'Column Level Security Profiles', 'Connection References', 'Connections', 'Environment Variables', 'Email Templates', 'SharePoint Data Sources'] },
  { key: 'reports-and-insights', title: 'Web Resources, Reports & Insights', headings: ['Web Resources', 'Reports & Dashboards', 'Dataverse Deep Insights', 'Additional Solution Components', 'Other Solution Components', 'Missing Dependencies'] },
];

/**
 * Splits a fully generated Markdown document into a set of self-contained,
 * logically-grouped files (e.g. Tables & Relationships, Apps, Automation &
 * Processes, Plugins). Each file repeats the solution header/overview so it
 * can be read independently of the others.
 */
export function splitMarkdownByCategory(markdown: string): MarkdownCategoryFile[] {
  const matches = sectionHeadings(markdown);
  if (matches.length === 0) return [];

  const preamble = markdown.slice(0, matches[0].index).trimEnd();

  const sections = matches.map((match, idx) => {
    const start = match.index ?? 0;
    const end = idx + 1 < matches.length ? (matches[idx + 1].index ?? markdown.length) : markdown.length;
    return { heading: match.title, body: markdown.slice(start, end).trimEnd() };
  });

  const overview = sections.find((s) => s.heading === 'Solution Overview');
  const sharedHeader = [preamble, overview?.body ?? ''].filter(Boolean).join('\n\n');

  const usedHeadings = new Set(CATEGORY_DEFINITIONS.flatMap((c) => c.headings));
  usedHeadings.add('Solution Overview');
  usedHeadings.add('Table of Contents');

  const files: MarkdownCategoryFile[] = [];
  CATEGORY_DEFINITIONS.forEach((category) => {
    const matched = sections.filter((s) => category.headings.includes(s.heading));
    if (matched.length === 0) return;
    files.push({
      key: category.key,
      title: category.title,
      markdown: categoryDocument(sharedHeader, matched),
    });
  });

  const leftover = sections.filter((s) => !usedHeadings.has(s.heading));
  if (leftover.length > 0) {
    files.push({
      key: 'other',
      title: 'Other',
      markdown: categoryDocument(sharedHeader, leftover),
    });
  }

  return files;
}

/**
 * Builds a standalone document containing only Mermaid diagram sections from
 * an already-generated Markdown document.
 */
export function splitDiagramsDocument(markdown: string, title = 'Diagrams'): { mainMarkdown: string; diagramsMarkdown: string } {
  const sectionRegex = /^## (.+)$/gm;
  const matches = [...markdown.matchAll(sectionRegex)];
  const preamble = markdown.slice(0, matches[0]?.index ?? markdown.length).trim();
  const sections = matches.map((match, index) => {
    const start = match.index ?? 0;
    const end = index + 1 < matches.length ? (matches[index + 1].index ?? markdown.length) : markdown.length;
    return markdown.slice(start, end).trim();
  }).filter((section) => section.includes('```mermaid'));

  if (sections.length === 0) return { mainMarkdown: markdown, diagramsMarkdown: '' };

  const withoutDiagrams = (section: string) => section
    .replace(/```mermaid[^\n]*\n[\s\S]*?```/g, '')
    .replace(/^#### Navigation Hierarchy\s*$/gm, '')
    .trim();
  const hasText = (section: string) => withoutDiagrams(section)
    .replace(/^#{2,6} .+$/gm, '')
    .replace(/\[Back to Top\]\(#table-of-contents\)/g, '')
    .trim().length > 0;
  const diagramTitles = new Set(sections.filter((section) => !hasText(section))
    .map((section) => section.match(/^## (.+)$/m)?.[1])
    .filter((title): title is string => Boolean(title)));
  const mainPreamble = preamble
    .split('\n')
    .filter((line) => {
      const match = /^- \[(.+)\]\((#[^)]+)\)/.exec(line);
      return !match || !diagramTitles.has(match[1]);
    })
    .join('\n');
  const mainSections = matches.map((match, index) => {
    const start = match.index ?? 0;
    const end = index + 1 < matches.length ? (matches[index + 1].index ?? markdown.length) : markdown.length;
    const section = markdown.slice(start, end).trim();
    if (diagramTitles.has(match[1].trim())) return '';
    if (match[1].trim() !== 'Table of Contents') return withoutDiagrams(section);
    return section.split('\n').filter((line) => {
      const link = /^- \[(.+)\]\((#[^)]+)\)/.exec(line);
      return !link || !diagramTitles.has(link[1]);
    }).join('\n');
  }).filter(Boolean);

  const header = [
    preamble.replace(/^# .+$/m, `# ${title}`),
    '',
    heading(2, 'Table of Contents'),
    '',
    ...sections.map((section) => {
      const sectionTitle = section.match(/^## (.+)$/m)?.[1] ?? 'Diagram';
      return `- [${sectionTitle}](${headingAnchor(sectionTitle)})`;
    }),
  ].join('\n');

  return {
    mainMarkdown: [mainPreamble, ...mainSections].filter(Boolean).join('\n\n'),
    diagramsMarkdown: appendBackToTopLinks(`${header}\n\n${sections.join('\n\n')}`),
  };
}

export function extractDiagramsDocument(markdown: string, title = 'Diagrams'): string {
  return splitDiagramsDocument(markdown, title).diagramsMarkdown;
}
