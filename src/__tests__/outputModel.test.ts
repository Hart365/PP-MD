import { describe, expect, it } from 'vitest';
import {
  createDiagramSection,
  createMetadataGridSection,
  createTableSection,
  renderMarkdownDocument,
} from '../generator/outputModel';

describe('outputModel', () => {
  it('renders a table section into markdown', () => {
    const section = createTableSection({
      id: 'component-inventory',
      title: 'Solution Component Inventory',
      headers: ['Component Type', 'Count', 'Covered in PP-MD'],
      rows: [['Entity', '12', '✅']],
    });

    const markdown = renderMarkdownDocument([section]);

    expect(markdown).toContain('## Solution Component Inventory');
    expect(markdown).toContain('| Component Type | Count | Covered in PP-MD |');
    expect(markdown).toContain('| Entity | 12 | ✅ |');
  });

  it('renders metadata-grid rows with expected columns', () => {
    const section = createMetadataGridSection({
      id: 'entity-grid',
      title: 'Metadata Grid',
      rows: [
        {
          entity: 'account',
          attribute: 'name',
          relationship: 'account_primary_contact',
          value: 'Primary account name',
        },
      ],
    });

    const markdown = renderMarkdownDocument([section]);

    expect(markdown).toContain('| Entity | Attribute | Relationship | Value |');
    expect(markdown).toContain('| account | name | account_primary_contact | Primary account name |');
  });

  it('renders diagram blocks using fenced code syntax', () => {
    const section = createDiagramSection({
      id: 'erd',
      title: 'Entity Relationship Diagram',
      introMarkdown: '> Diagram summary',
      language: 'mermaid',
      content: 'erDiagram\n  account ||--o{ contact : has',
    });

    const markdown = renderMarkdownDocument([section]);

    expect(markdown).toContain('## Entity Relationship Diagram');
    expect(markdown).toContain('```mermaid');
    expect(markdown).toContain('account ||--o{ contact : has');
  });
});
