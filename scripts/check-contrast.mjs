import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const cssPath = resolve(process.cwd(), 'src/assets/global.css');
const css = readFileSync(cssPath, 'utf8');

const themeBlocks = {
  light: /:root\s*,\s*\[data-theme="light"\]\s*\{([\s\S]*?)\}/m,
  dark: /\[data-theme="dark"\]\s*\{([\s\S]*?)\}/m,
};

function extractVars(block) {
  const vars = new Map();
  const regex = /--([a-z0-9-]+)\s*:\s*([^;]+);/gi;
  let match = regex.exec(block);
  while (match) {
    vars.set(match[1], match[2].trim());
    match = regex.exec(block);
  }
  return vars;
}

function hexToRgb(value) {
  const hex = value.replace('#', '').trim();
  if (!/^[0-9a-f]{3}([0-9a-f]{3})?$/i.test(hex)) return undefined;
  if (hex.length === 3) {
    return {
      r: parseInt(hex[0] + hex[0], 16),
      g: parseInt(hex[1] + hex[1], 16),
      b: parseInt(hex[2] + hex[2], 16),
    };
  }
  return {
    r: parseInt(hex.slice(0, 2), 16),
    g: parseInt(hex.slice(2, 4), 16),
    b: parseInt(hex.slice(4, 6), 16),
  };
}

function toLuminance({ r, g, b }) {
  const channel = (n) => {
    const srgb = n / 255;
    return srgb <= 0.03928 ? srgb / 12.92 : ((srgb + 0.055) / 1.055) ** 2.4;
  };
  return (0.2126 * channel(r)) + (0.7152 * channel(g)) + (0.0722 * channel(b));
}

function contrastRatio(foreground, background) {
  const fg = hexToRgb(foreground);
  const bg = hexToRgb(background);
  if (!fg || !bg) return Number.NaN;

  const l1 = toLuminance(fg);
  const l2 = toLuminance(bg);
  const lighter = Math.max(l1, l2);
  const darker = Math.min(l1, l2);
  return (lighter + 0.05) / (darker + 0.05);
}

const textChecks = [
  ['color-text-primary', 'color-bg', 4.5],
  ['color-text-primary', 'color-surface', 4.5],
  ['color-text-secondary', 'color-bg', 4.5],
  ['color-text-secondary', 'color-surface', 4.5],
  ['color-text-muted', 'color-bg', 4.5],
  ['color-text-muted', 'color-surface', 4.5],
  ['color-text-inverse', 'color-accent', 4.5],
  ['color-text-inverse', 'color-accent-hover', 4.5],
  ['color-text-code', 'color-code-bg', 4.5],
  ['color-accent', 'color-bg', 4.5],
];

const nonTextChecks = [
  ['color-border', 'color-surface', 3.0],
  ['color-border', 'color-bg', 3.0],
  ['color-border-focus', 'color-bg', 3.0],
  ['color-dropzone-border', 'color-dropzone-bg', 3.0],
  ['color-progress-fill', 'color-progress-track', 3.0],
  ['color-sidebar-item-active', 'color-sidebar-bg', 3.0],
];

function evaluateTheme(name, vars) {
  const failures = [];
  const outputs = [];

  const runChecks = (checks, checkType) => {
    checks.forEach(([fgToken, bgToken, min]) => {
      const fg = vars.get(fgToken);
      const bg = vars.get(bgToken);

      if (!fg || !bg) {
        failures.push(`[${name}] Missing token(s): ${fgToken} or ${bgToken}`);
        return;
      }

      const ratio = contrastRatio(fg, bg);
      if (Number.isNaN(ratio)) {
        failures.push(`[${name}] Unsupported color format for ${fgToken}=${fg} or ${bgToken}=${bg}`);
        return;
      }

      const label = `${fgToken} on ${bgToken}`;
      outputs.push(`[${name}] ${checkType}: ${label} = ${ratio.toFixed(2)}:1 (min ${min}:1)`);

      if (ratio < min) {
        failures.push(`[${name}] FAIL ${label}: ${ratio.toFixed(2)}:1 < ${min}:1`);
      }
    });
  };

  runChecks(textChecks, 'text');
  runChecks(nonTextChecks, 'non-text');

  return { outputs, failures };
}

function parseTheme(name, regex) {
  const match = css.match(regex);
  if (!match) {
    throw new Error(`Could not locate ${name} theme block in ${cssPath}`);
  }
  return extractVars(match[1]);
}

const lightVars = parseTheme('light', themeBlocks.light);
const darkVars = parseTheme('dark', themeBlocks.dark);

const lightResult = evaluateTheme('light', lightVars);
const darkResult = evaluateTheme('dark', darkVars);

const allOutputs = [...lightResult.outputs, ...darkResult.outputs];
const allFailures = [...lightResult.failures, ...darkResult.failures];

console.log('WCAG contrast audit for theme tokens');
allOutputs.forEach((line) => console.log(line));

if (allFailures.length > 0) {
  console.error('\nContrast audit failed:');
  allFailures.forEach((line) => console.error(`- ${line}`));
  process.exit(1);
}

console.log('\nContrast audit passed.');
