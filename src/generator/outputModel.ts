export interface OutputMarkdownBlock {
  kind: 'markdown';
  content: string;
}

export interface OutputTableBlock {
  kind: 'table';
  headers: string[];
  rows: string[][];
}

export interface OutputDiagramBlock {
  kind: 'diagram';
  language: string;
  content: string;
}

export interface OutputMetadataGridRow {
  entity: string;
  attribute?: string;
  relationship?: string;
  value: string;
}

export type OutputBlock = OutputMarkdownBlock | OutputTableBlock | OutputDiagramBlock;

export interface OutputSection {
  id: string;
  blocks: OutputBlock[];
}

export interface MarkdownRenderOptions {
  addBackToTopLinks?: boolean;
}

export function createOutputSection(id: string, markdown: string): OutputSection {
  return {
    id,
    blocks: [{ kind: 'markdown', content: markdown }],
  };
}

export function createTableSection(options: {
  id: string;
  title: string;
  introMarkdown?: string;
  headers: string[];
  rows: string[][];
}): OutputSection {
  const blocks: OutputBlock[] = [
    { kind: 'markdown', content: `## ${options.title}\n` },
  ];

  if (options.introMarkdown && options.introMarkdown.trim().length > 0) {
    blocks.push({ kind: 'markdown', content: `${options.introMarkdown}\n` });
  }

  blocks.push({
    kind: 'table',
    headers: options.headers,
    rows: options.rows,
  });

  return {
    id: options.id,
    blocks,
  };
}

export function createMetadataGridSection(options: {
  id: string;
  title: string;
  introMarkdown?: string;
  rows: OutputMetadataGridRow[];
}): OutputSection {
  return createTableSection({
    id: options.id,
    title: options.title,
    introMarkdown: options.introMarkdown,
    headers: ['Entity', 'Attribute', 'Relationship', 'Value'],
    rows: options.rows.map((row) => [
      row.entity,
      row.attribute ?? '–',
      row.relationship ?? '–',
      row.value,
    ]),
  });
}

export function createDiagramSection(options: {
  id: string;
  title: string;
  introMarkdown?: string;
  language: string;
  content: string;
}): OutputSection {
  const blocks: OutputBlock[] = [
    { kind: 'markdown', content: `## ${options.title}\n` },
  ];

  if (options.introMarkdown && options.introMarkdown.trim().length > 0) {
    blocks.push({ kind: 'markdown', content: `${options.introMarkdown}\n` });
  }

  blocks.push({
    kind: 'diagram',
    language: options.language,
    content: options.content,
  });

  return {
    id: options.id,
    blocks,
  };
}

export function renderMarkdownDocument(
  sections: OutputSection[],
  options: MarkdownRenderOptions = {},
): string {
  const markdown = sections
    .map(renderSection)
    .filter((sectionMarkdown) => sectionMarkdown.trim().length > 0)
    .join('\n');

  if (!options.addBackToTopLinks) return markdown;
  return appendBackToTopLinks(markdown);
}

function renderSection(section: OutputSection): string {
  return section.blocks
    .map((block) => {
      if (block.kind === 'markdown') return block.content;
      if (block.kind === 'table') return renderTable(block.headers, block.rows);
      return renderDiagram(block.language, block.content);
    })
    .filter((part) => part.trim().length > 0)
    .join('\n');
}

function renderTable(headers: string[], rows: string[][]): string {
  const lines: string[] = [];
  lines.push(`| ${headers.join(' | ')} |`);
  lines.push(`|${headers.map(() => '---').join('|')}|`);
  rows.forEach((row) => {
    lines.push(`| ${row.join(' | ')} |`);
  });
  lines.push('');
  return lines.join('\n');
}

function renderDiagram(language: string, content: string): string {
  return `\`\`\`${language}\n${content}\n\`\`\`\n`;
}

function appendBackToTopLinks(markdown: string): string {
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
