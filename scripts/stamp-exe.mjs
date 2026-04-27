import { readdirSync, existsSync } from 'node:fs';
import { join, resolve } from 'node:path';
import rcedit from 'rcedit';

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

const version = '1.0.0.0';

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
