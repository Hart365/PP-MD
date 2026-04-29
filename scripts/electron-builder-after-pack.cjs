const path = require('node:path');
const fs = require('node:fs');

function normalizeWindowsVersion(version) {
  const parts = String(version || '1.0.0')
    .split('.')
    .map((part) => Number.parseInt(part, 10))
    .filter((part) => Number.isFinite(part));

  while (parts.length < 4) {
    parts.push(0);
  }

  return parts.slice(0, 4).join('.');
}

module.exports = async function afterPack(context) {
  if (context.electronPlatformName !== 'win32') {
    return;
  }

  const appExeName = context.packager.appInfo.productFilename + '.exe';
  const appExePath = path.join(context.appOutDir, appExeName);
  const iconPath = path.resolve(context.packager.projectDir, 'build', 'icon.ico');

  if (!fs.existsSync(appExePath) || !fs.existsSync(iconPath)) {
    return;
  }

  const rceditModule = await import('rcedit');
  const rcedit = rceditModule.rcedit || rceditModule.default || rceditModule;
  const version = normalizeWindowsVersion(context.packager.appInfo.version);

  await rcedit(appExePath, {
    icon: iconPath,
    'file-version': version,
    'product-version': version,
    'version-string': {
      CompanyName: 'Hart of the Midlands',
      FileDescription: 'Power Platform solution documentation generator for Windows desktop.',
      ProductName: 'PP-MD',
      LegalCopyright: 'Copyright (c) Mike Hartley, Hart of the Midlands',
      OriginalFilename: appExeName,
      InternalName: context.packager.appInfo.productFilename,
    },
  });

  console.log(`[afterPack] Stamped executable icon/metadata: ${appExePath}`);
};
