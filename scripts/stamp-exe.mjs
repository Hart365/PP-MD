import { readdirSync, existsSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import rcedit from 'rcedit';

function toWindowsVersion(version) {
  const core = (version ?? '0.0.0').split('-')[0].split('+')[0];
  const parts = core.split('.').map((part) => Number.parseInt(part, 10));
  const numericParts = parts.map((part) => (Number.isFinite(part) ? Math.max(0, part) : 0));
  while (numericParts.length < 4) {
    numericParts.push(0);
  }
  return numericParts.slice(0, 4).join('.');
}

const releaseDir = resolve(process.cwd(), 'release');
if (!existsSync(releaseDir)) {
  console.error('release directory not found. Run desktop build first.');
  process.exit(1);
}

const exeCandidates = readdirSync(releaseDir)
  .filter((name) => {
    const lower = name.toLowerCase();
    return lower.endsWith('.exe') && lower.startsWith('pp-md-') && !lower.startsWith('__uninstaller');
  })
  .sort();

if (exeCandidates.length === 0) {
  console.error('No PP-MD executable found in release directory.');
  process.exit(1);
}

const packageJsonPath = resolve(process.cwd(), 'package.json');
const packageJson = JSON.parse(readFileSync(packageJsonPath, 'utf8'));
const version = toWindowsVersion(packageJson.version);

for (const exeName of exeCandidates) {
  const exePath = join(releaseDir, exeName);
  await rcedit(exePath, {
    'file-version': version,
    'product-version': version,
    'version-string': {
      CompanyName: 'Hart of the Midlands',
      FileDescription: 'Power Platform solution documentation generator for Windows desktop.',
      ProductName: 'PP-MD',
      LegalCopyright: 'Copyright (c) Mike Hartley, Hart of the Midlands',
      OriginalFilename: exeName,
      InternalName: 'PP-MD',
    },
  });

  console.log(`Stamped executable metadata: ${exePath}`);
}
