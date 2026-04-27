import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, resolve, basename } from 'node:path';
import { parseSolutionZip } from '../src/parser/solutionParser';
import { generateMarkdown } from '../src/generator/markdownGenerator';

interface RelationshipEdge {
  from: string;
  to: string;
  type: string;
  label: string;
}

function collectRelationshipStats(parsed: Awaited<ReturnType<typeof parseSolutionZip>>) {
  const included = new Set(parsed.entities.map((entity) => entity.logicalName.toLowerCase()));
  const edges: RelationshipEdge[] = [];
  const dedupe = new Set<string>();
  const degree = new Map<string, number>();

  const bump = (name: string) => {
    degree.set(name, (degree.get(name) ?? 0) + 1);
  };

  parsed.entities.forEach((entity) => {
    entity.relationships.forEach((rel) => {
      const referenced = rel.referencedEntity.toLowerCase();
      const referencing = rel.referencingEntity.toLowerCase();
      if (!included.has(referenced) || !included.has(referencing)) return;

      const label = `${rel.name || `${referenced}_${referencing}`}${rel.referencingAttribute ? ` [${rel.referencingAttribute}]` : ''}`;

      if (rel.type === 'ManyToMany') {
        const [left, right] = [referenced, referencing].sort();
        const key = `${left}|many-to-many|${right}|${label}`;
        if (dedupe.has(key)) return;
        dedupe.add(key);
        edges.push({ from: left, to: right, type: rel.type, label });
        bump(left);
        bump(right);
        return;
      }

      const key = `${referenced}|one-to-many|${referencing}|${label}`;
      if (dedupe.has(key)) return;
      dedupe.add(key);
      edges.push({ from: referenced, to: referencing, type: rel.type, label });
      bump(referenced);
      bump(referencing);
    });
  });

  const degreeValues = Array.from(degree.values()).sort((a, b) => b - a);
  const maxDegree = degreeValues[0] ?? 0;
  const avgDegree = degreeValues.length > 0
    ? degreeValues.reduce((sum, value) => sum + value, 0) / degreeValues.length
    : 0;
  const relationshipDensity = parsed.entities.length > 0 ? edges.length / parsed.entities.length : 0;

  return {
    edgeCount: edges.length,
    entityCount: parsed.entities.length,
    maxDegree,
    avgDegree,
    relationshipDensity,
    topHubs: Array.from(degree.entries()).sort((a, b) => b[1] - a[1]).slice(0, 10),
  };
}

async function run() {
  const inputPath = process.argv[2];
  const outputPath = process.argv[3] || resolve(process.cwd(), 'samples', 'generated-sample.md');

  if (!inputPath) {
    throw new Error('Usage: tsx scripts/sample-doc-from-zip.ts "<zip-path>" [output-markdown-path]');
  }

  const zipBytes = readFileSync(inputPath);
  const parsed = await parseSolutionZip(zipBytes as unknown as Blob);
  const markdown = generateMarkdown(parsed, { erdMode: 'detailed-relationships' });

  mkdirSync(dirname(outputPath), { recursive: true });
  writeFileSync(outputPath, markdown, 'utf8');

  const stats = collectRelationshipStats(parsed);
  console.log('Sample markdown generated');
  console.log(`Input: ${inputPath}`);
  console.log(`Output: ${outputPath}`);
  console.log(`Entities: ${stats.entityCount}`);
  console.log(`Unique relationships: ${stats.edgeCount}`);
  console.log(`Relationship density (rels/entities): ${stats.relationshipDensity.toFixed(2)}`);
  console.log(`Max degree: ${stats.maxDegree}`);
  console.log(`Avg degree: ${stats.avgDegree.toFixed(2)}`);
  console.log('Top hubs:');
  stats.topHubs.forEach(([name, count], index) => {
    console.log(`  ${index + 1}. ${name} -> ${count}`);
  });

  const sampleName = basename(outputPath);
  console.log(`Saved markdown sample file: ${sampleName}`);
}

run().catch((err) => {
  console.error('Failed to generate sample markdown.');
  console.error(err);
  process.exit(1);
});
