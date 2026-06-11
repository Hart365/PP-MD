import { test, expect } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';

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
