import { test, expect } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

async function solutionFile(name: string) {
  const zip = new JSZip();
  zip.file('solution.xml', `<ImportExportXml><SolutionManifest><UniqueName>${name}</UniqueName><Version>1.0</Version><Publisher><UniqueName>contoso</UniqueName><CustomizationPrefix>new</CustomizationPrefix></Publisher><MissingDependencies><MissingDependency><Required type="1" schemaName="account"/><Dependent type="2" schemaName="new_customer" parentSchemaName="new_case"/></MissingDependency></MissingDependencies></SolutionManifest></ImportExportXml>`);
  zip.file('customizations.xml', '<ImportExportXml><Entities><Entity><Name>new_case</Name><EntityInfo><entity><Name>new_case</Name><IsCustomEntity>1</IsCustomEntity><IsAuditEnabled>1</IsAuditEnabled><attributes><attribute><Name>new_name</Name><Type>nvarchar</Type><IsCustomField>1</IsCustomField><RequiredLevel>None</RequiredLevel></attribute></attributes></entity></EntityInfo></Entity></Entities></ImportExportXml>');
  zip.file('Workflows/invalid.json', '{');
  return { name: `${name}.zip`, mimeType: 'application/zip', buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

test('home screen meets WCAG 2.2 AA baseline checks', async ({ page }) => {
  await page.goto('/');

  // Keyboard users should land on the skip link first.
  await page.keyboard.press('Tab');
  await expect(page.getByRole('link', { name: /skip to main content/i })).toBeFocused();

  const results = await new AxeBuilder({ page }).analyze();
  const blocking = results.violations.filter((violation) =>
    violation.impact === 'serious' || violation.impact === 'critical'
  );

  expect(
    blocking,
    blocking.map((violation) => `${violation.id}: ${violation.help}`).join('\n')
  ).toEqual([]);
});

test('settings controls support keyboard navigation and visible focus flow', async ({ page }) => {
  await page.goto('/');

  const configSelect = page.getByRole('combobox', { name: /select document configuration/i });
  await configSelect.focus();
  await expect(configSelect).toBeFocused();

  await page.keyboard.press('Tab');
  await expect(page.getByRole('textbox', { name: /configuration name/i })).toBeFocused();

  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: /save configuration/i })).toBeFocused();

  await page.keyboard.press('Tab');
  await expect(page.getByRole('textbox', { name: 'Client' })).toBeFocused();

  const detailSelect = page.getByRole('combobox', { name: /select documentation detail level/i });
  await detailSelect.focus();
  await expect(detailSelect).toBeFocused();

  const attributeModeSelect = page.getByRole('combobox', { name: /select attribute selection mode/i });
  await attributeModeSelect.focus();
  await expect(attributeModeSelect).toBeFocused();
});

test('local solution parity workflow exports real files and remains accessible', async ({ page }) => {
  await page.goto('/');
  await page.getByLabel('Include missing import dependencies').check();
  await page.locator('input[type="file"][accept*=".zip"]').setInputFiles([await solutionFile('Alpha'), await solutionFile('Beta')]);
  await page.getByRole('button', { name: 'Generate documentation for 2 files' }).click();
  await page.getByRole('button', { name: 'View documentation for All Selected Solutions', exact: true }).click();
  await expect(page.getByRole('region', { name: 'Documentation for All Selected Solutions', exact: true }).getByRole('heading', { name: 'Tables & Columns', exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: 'View documentation for All Selected Solutions - Diagrams', exact: true })).toHaveCount(0);
  await page.getByRole('button', { name: 'View documentation for Alpha', exact: true }).click();
  const viewer = page.getByRole('region', { name: 'Documentation for Alpha', exact: true });
  await expect(viewer.locator('figure svg').first()).toBeVisible();
  await expect(viewer.getByRole('heading', { name: 'Missing Dependencies', exact: true })).toBeAttached();

  const analysis = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa']).analyze();
  expect(analysis.violations.map((violation) => `${violation.id}: ${violation.nodes.map((node) => `${node.target.join(' ')} ${node.failureSummary}`).join('\n')}`)).toEqual([]);
  await page.screenshot({ path: 'test-results/parity-desktop.png', fullPage: true });

  for (const format of ['html', 'pdf', 'xlsx']) {
    await page.getByRole('button', { name: 'Export documentation', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('menuitem', { name: new RegExp(`Export \\.${format}$`, 'i') }).click();
    const download = await pending;
    expect(download.suggestedFilename()).toMatch(new RegExp(`\\.${format}$`));
    const downloadedPath = await download.path();
    expect(downloadedPath).not.toBeNull();
    const bytes = await readFile(downloadedPath!);
    expect(bytes.length).toBeGreaterThan(100);
    if (format === 'html') expect(bytes.toString()).toContain('<svg');
    if (format === 'pdf') {
      expect(bytes.toString('latin1')).toContain('%PDF-');
      expect(bytes.toString('latin1')).toContain('/Subtype /Image');
    }
    if (format === 'xlsx') {
      const workbook = await JSZip.loadAsync(bytes);
      expect(await workbook.file('xl/workbook.xml')?.async('string')).toContain('Tables');
    }
  }

  await page.getByRole('button', { name: 'Export All', exact: true }).click();
  const categoryPending = page.waitForEvent('download');
  await page.getByRole('button', { name: 'By Category (ZIP)', exact: true }).click();
  const categoryPath = await (await categoryPending).path();
  const categories = await JSZip.loadAsync(await readFile(categoryPath!));
  expect(Object.keys(categories.files).some((name) => name.startsWith('Consolidated/'))).toBe(true);
  expect(Object.keys(categories.files).some((name) => name.startsWith('Solutions/'))).toBe(true);

  for (const format of ['HTML', 'PDF', 'Excel']) {
    await page.getByRole('button', { name: 'Export All', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: `All ${format} (ZIP)`, exact: true }).click();
    const path = await (await pending).path();
    const zip = await JSZip.loadAsync(await readFile(path!));
    expect(Object.keys(zip.files).filter((name) => !zip.files[name].dir).length).toBeGreaterThan(3);
    if (format === 'Excel') {
      expect(Object.keys(zip.files).some((name) => name.endsWith('.xlsx'))).toBe(true);
      expect(Object.keys(zip.files).some((name) => name.endsWith('Diagnostics.md'))).toBe(true);
    }
  }

  await page.setViewportSize({ width: 320, height: 800 });
  await expect(page.getByRole('button', { name: 'Open toolbar menu' })).toBeVisible();
  await page.getByRole('button', { name: 'Open toolbar menu' }).click();
  await page.getByRole('switch', { name: /switch to dark/i }).click();
  await page.getByRole('button', { name: 'Close toolbar menu' }).click();
  const mobileAnalysis = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa']).analyze();
  expect(mobileAnalysis.violations.map((violation) => `${violation.id}: ${violation.nodes.map((node) => node.failureSummary).join('\n')}`)).toEqual([]);
  const controls = await page.getByRole('toolbar', { name: 'Documentation actions' }).getByRole('button').evaluateAll((buttons) => buttons.map((button) => {
    const bounds = button.getBoundingClientRect();
    return { right: bounds.right, left: bounds.left };
  }));
  expect(controls.every((bounds) => bounds.left >= 0 && bounds.right <= 320)).toBe(true);
  await page.screenshot({ path: 'test-results/parity-viewer-mobile-dark.png', fullPage: true });
});

test('diagram options control inline and companion output with a dedicated sidebar section', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: /diagram options/i }).click();
  await expect(page.getByLabel(/generate companion diagrams document/i)).not.toBeChecked();
  await page.locator('input[type="file"][accept*=".zip"]').setInputFiles([await solutionFile('Alpha'), await solutionFile('Beta')]);
  await page.getByRole('button', { name: 'Generate documentation for 2 files' }).click();
  const sidebar = page.getByRole('navigation', { name: 'Solution navigation' });
  const diagrams = sidebar.getByRole('region', { name: 'Diagrams', exact: true });
  await expect(sidebar.getByRole('button', { name: 'View documentation for All Selected Solutions', exact: true })).toBeVisible();
  await expect(diagrams).toHaveCount(0);

  async function markdownArchive() {
    await page.getByRole('button', { name: 'Export All', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'All Markdown (ZIP)', exact: true }).click();
    const path = await (await pending).path();
    return JSZip.loadAsync(await readFile(path!));
  }
  const inlineArchive = await markdownArchive();
  expect(Object.keys(inlineArchive.files).some((name) => / - Diagrams\.md$/.test(name))).toBe(false);
  expect(await inlineArchive.file('All Selected Solutions.md')?.async('string')).toContain('```mermaid');
  expect(await inlineArchive.file('Alpha.md')?.async('string')).toContain('```mermaid');

  const options = page.getByRole('button', { name: 'Options', exact: true });
  const optionsPanel = page.getByRole('group', { name: 'Options actions', exact: true });
  await options.click();
  await optionsPanel.getByRole('checkbox', { name: 'Separate diagrams document', exact: true }).check();
  await page.keyboard.press('Escape');
  await expect(diagrams).toBeVisible();
  await expect(diagrams.getByRole('button')).toHaveCount(3);
  await expect(sidebar.getByRole('region', { name: 'Solutions', exact: true }).getByRole('button', { name: /Diagrams/ })).toHaveCount(0);
  const splitArchive = await markdownArchive();
  expect(await splitArchive.file('All Selected Solutions.md')?.async('string')).not.toContain('```mermaid');
  expect(await splitArchive.file('Alpha.md')?.async('string')).not.toContain('```mermaid');
  expect(await splitArchive.file('Alpha - Diagrams.md')?.async('string')).toContain('```mermaid');
  expect(await splitArchive.file('All Selected Solutions - Diagrams.md')?.async('string')).toContain('```mermaid');
  await diagrams.getByRole('button', { name: 'View documentation for Alpha - Diagrams', exact: true }).click();
  await expect(page.getByRole('region', { name: 'Documentation for Alpha (Diagrams)', exact: true })).toBeVisible();
  await page.screenshot({ path: 'test-results/diagrams-sidebar.png', fullPage: true });
  await page.locator('input[type="file"][accept=".md,text/markdown"]').setInputFiles({
    name: 'Notes.md', mimeType: 'text/markdown', buffer: Buffer.from('# Notes\n\nImported document.'),
  });
  await expect(diagrams.getByRole('button')).toHaveCount(3);
  await diagrams.getByRole('button', { name: 'View documentation for Alpha - Diagrams', exact: true }).click();

  await options.click();
  await optionsPanel.getByRole('checkbox', { name: 'Separate diagrams document', exact: true }).uncheck();
  await page.keyboard.press('Escape');
  await expect(diagrams).toHaveCount(0);
  await expect(page.getByRole('region', { name: 'Documentation for Alpha', exact: true })).toBeVisible();
  await options.click();
  await optionsPanel.getByRole('checkbox', { name: 'Separate diagrams document', exact: true }).check();
  await optionsPanel.getByRole('checkbox', { name: 'Include diagrams', exact: true }).uncheck();
  await page.keyboard.press('Escape');
  await expect(diagrams).toHaveCount(0);
  const excludedArchive = await markdownArchive();
  const contents = await Promise.all(Object.values(excludedArchive.files).filter((entry) => !entry.dir).map((entry) => entry.async('string')));
  contents.forEach((markdown) => expect(markdown).not.toContain('```mermaid'));
  await options.click();
  await optionsPanel.getByRole('checkbox', { name: 'Create consolidated document', exact: true }).uncheck();
  await page.keyboard.press('Escape');
  await expect(sidebar.getByRole('region', { name: 'Consolidated', exact: true })).toHaveCount(0);
  await expect(page.getByRole('region', { name: 'Documentation for Alpha', exact: true })).toBeVisible();
});

for (const width of [1280, 320]) {
  test(`header dropdowns dismiss and work without blocking the app at ${width}px`, async ({ page }) => {
    await page.setViewportSize({ width, height: 900 });
    await page.goto('/');
    await page.getByLabel('Include diagrams', { exact: true }).uncheck();
    await page.locator('input[type="file"][accept*=".zip"]').setInputFiles([await solutionFile('Alpha'), await solutionFile('Beta')]);
    await page.getByRole('button', { name: 'Generate documentation for 2 files' }).click();
    const combined = page.getByRole('button', { name: 'View documentation for All Selected Solutions', exact: true });
    await expect(combined).toBeVisible();
    if (width === 320) await page.getByRole('button', { name: 'Open toolbar menu' }).click();

    const dimensions = await page.locator('header').locator('button:visible, a:visible').evaluateAll((controls) => controls.map((control) => ({
      height: control.getBoundingClientRect().height,
      radius: getComputedStyle(control).borderRadius,
    })));
    expect(new Set(dimensions.map((control) => control.height))).toEqual(new Set([40]));
    expect(new Set(dimensions.map((control) => control.radius))).toEqual(new Set(['6px']));

    const options = page.getByRole('button', { name: 'Options', exact: true });
    const exports = page.getByRole('button', { name: 'Export All', exact: true });
    const optionsPanel = page.getByRole('group', { name: 'Options actions', exact: true });
    const exportPanel = page.getByRole('group', { name: 'Export All actions', exact: true });
    await options.click();
    await expect(options).toHaveAttribute('aria-expanded', 'true');
    await expect(optionsPanel).toBeVisible();
    expect(await optionsPanel.evaluate((panel) => panel.matches(':popover-open'))).toBe(true);
    await optionsPanel.getByRole('checkbox', { name: 'Create consolidated document', exact: true }).uncheck();
    await expect(combined).toHaveCount(0);
    await optionsPanel.getByRole('checkbox', { name: 'Create consolidated document', exact: true }).check();
    await expect(combined).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(optionsPanel).toBeHidden();
    await expect(options).toBeFocused();
    await expect(options).toHaveAttribute('aria-expanded', 'false');

    await options.click();
    await exports.click();
    await expect(optionsPanel).toBeHidden();
    await expect(exportPanel).toBeVisible();
    const bounds = await exportPanel.boundingBox();
    expect(bounds).not.toBeNull();
    expect(bounds!.x).toBeGreaterThanOrEqual(0);
    expect(bounds!.x + bounds!.width).toBeLessThanOrEqual(width);
    await page.getByRole('heading', { name: /^PP-MD v/ }).click();
    await expect(exportPanel).toBeHidden();

    await exports.focus();
    await page.keyboard.press('Enter');
    await expect(exportPanel).toBeVisible();
    const pending = page.waitForEvent('download');
    await exportPanel.getByRole('button', { name: 'All Markdown (ZIP)', exact: true }).click();
    expect((await pending).suggestedFilename()).toMatch(/\.zip$/);
    await expect(exportPanel).toBeHidden();
    await expect(exports).toHaveAttribute('aria-expanded', 'false');
    await options.click();
    await expect(optionsPanel).toBeVisible();
    await page.screenshot({ path: `test-results/header-dropdowns-${width}.png`, fullPage: true });
    const analysis = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa']).analyze();
    expect(analysis.violations.map((violation) => `${violation.id}: ${violation.help}`)).toEqual([]);
  });
}

for (const theme of ['light', 'dark']) {
  test(`dropdown hover and keyboard focus are clearly highlighted in ${theme} mode`, async ({ page }) => {
    await page.goto('/');
    if (theme === 'dark') await page.getByRole('switch', { name: /switch to dark/i }).click();
    await page.getByLabel('Include diagrams', { exact: true }).uncheck();
    await page.locator('input[type="file"][accept*=".zip"]').setInputFiles(await solutionFile('Highlight'));
    await page.getByRole('button', { name: 'Generate documentation for 1 file' }).click();
    const background = theme === 'light' ? 'rgb(29, 78, 216)' : 'rgb(147, 197, 253)';
    const foreground = theme === 'light' ? 'rgb(255, 255, 255)' : 'rgb(26, 26, 46)';
    const exports = page.getByRole('button', { name: 'Export All', exact: true });
    await exports.click();
    const exportPanel = page.getByRole('group', { name: 'Export All actions', exact: true });
    const exportItem = exportPanel.getByRole('button', { name: 'All Markdown (ZIP)', exact: true });
    await exportItem.hover();
    await expect(exportItem).toHaveCSS('background-color', background);
    await expect(exportItem).toHaveCSS('color', foreground);
    await page.screenshot({ path: `test-results/export-hover-${theme}.png`, fullPage: true });
    await exports.focus();
    await page.keyboard.press('Tab');
    await expect(exportItem).toBeFocused();
    await expect(exportItem).toHaveCSS('outline-style', 'solid');
    await expect(exportItem).toHaveCSS('outline-width', '2px');

    const options = page.getByRole('button', { name: 'Options', exact: true });
    await options.click();
    const optionsPanel = page.getByRole('group', { name: 'Options actions', exact: true });
    const optionRow = optionsPanel.locator('label').filter({ hasText: 'Create consolidated document' });
    await optionRow.hover();
    await expect(optionRow).toHaveCSS('background-color', background);
    await expect(optionRow).toHaveCSS('color', foreground);
    await page.screenshot({ path: `test-results/options-hover-${theme}.png`, fullPage: true });
    await options.focus();
    await page.keyboard.press('Tab');
    await expect(optionRow.getByRole('checkbox')).toBeFocused();
    await expect(optionRow).toHaveCSS('outline-style', 'solid');
    await expect(optionRow).toHaveCSS('outline-width', '2px');
    await expect(optionRow).toHaveCSS('background-color', background);
    await expect(optionRow).toHaveCSS('color', foreground);
    const analysis = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa']).analyze();
    expect(analysis.violations.map((violation) => `${violation.id}: ${violation.help}`)).toEqual([]);
    await page.keyboard.press('Escape');

    const viewerExport = page.getByRole('button', { name: 'Export documentation', exact: true });
    await viewerExport.click();
    const viewerItem = page.getByRole('menuitem', { name: /Export \.html$/i });
    await viewerItem.hover();
    await expect(viewerItem).toHaveCSS('background-color', background);
    await expect(viewerItem).toHaveCSS('color', foreground);
    await viewerItem.focus();
    await page.keyboard.press('Tab');
    await page.keyboard.press('Shift+Tab');
    await expect(viewerItem).toBeFocused();
    await expect(viewerItem).toHaveCSS('outline-style', 'solid');
    await expect(viewerItem).toHaveCSS('outline-width', '2px');
    await page.keyboard.press('Escape');

    await page.emulateMedia({ forcedColors: 'active' });
    await options.click();
    await optionRow.hover();
    await expect(optionRow).toHaveCSS('outline-style', 'solid');
    await expect(optionRow).toHaveCSS('outline-width', '2px');
    await expect(optionRow).toHaveCSS('forced-color-adjust', 'none');
    await expect(optionRow.getByRole('checkbox')).toHaveCSS('forced-color-adjust', 'auto');
    await page.screenshot({ path: `test-results/options-hover-forced-colours-${theme}.png`, fullPage: true });
    await exports.click();
    await exportItem.hover();
    await expect(exportItem).toHaveCSS('outline-style', 'solid');
    await expect(exportItem).toHaveCSS('outline-width', '2px');
    await expect(exportItem).toHaveCSS('forced-color-adjust', 'none');
  });
}

test('narrow toolbar reflows and preserves keyboard access in both themes and forced colours', async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 800 });
  await page.goto('/');
  await expect(page.getByRole('button', { name: 'Open toolbar menu' })).toBeVisible();
  await page.getByRole('button', { name: 'Open toolbar menu' }).click();
  await expect(page.getByRole('button', { name: 'Open previously generated Markdown files' })).toBeVisible();
  await page.evaluate(() => {
    document.documentElement.style.setProperty('--transition-base', '2s');
    document.documentElement.style.setProperty('--transition-fast', '1s');
  });
  await page.getByRole('switch', { name: /switch to dark/i }).click();
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
  const analysis = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa']).analyze();
  expect(analysis.violations.map((violation) => `${violation.id}: ${violation.nodes.map((node) => `${node.target.join(' ')} ${node.failureSummary}`).join('\n')}`)).toEqual([]);
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
  expect(await page.evaluate(() => ({
    viewport: window.innerHeight,
    page: document.documentElement.scrollHeight,
    root: document.getElementById('root')?.getBoundingClientRect().height,
    main: document.querySelector('main')?.getBoundingClientRect().height,
    footer: document.querySelector('footer')?.getBoundingClientRect().bottom,
  }))).toMatchObject({ page: 800, root: 800 });
  await page.screenshot({ path: 'test-results/parity-mobile-dark.png', fullPage: true });
  await page.emulateMedia({ forcedColors: 'active', reducedMotion: 'reduce' });
  await page.getByRole('button', { name: 'Close toolbar menu' }).focus();
  await expect(page.getByRole('button', { name: 'Close toolbar menu' })).toBeFocused();
  await page.screenshot({ path: 'test-results/parity-mobile-forced-colours.png', fullPage: true });
  await page.emulateMedia({ forcedColors: 'none', reducedMotion: 'no-preference' });
  await page.getByRole('switch', { name: /switch to light/i }).click();
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
  const lightAnalysis = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa']).analyze();
  expect(lightAnalysis.violations.map((violation) => `${violation.id}: ${violation.nodes.map((node) => `${node.target.join(' ')} ${node.failureSummary}`).join('\n')}`)).toEqual([]);
});
