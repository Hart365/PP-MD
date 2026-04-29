/**
 * @file versionUtils.ts
 * @description Utilities for version detection, platform identification, and update checking.
 */

/**
 * Platform type for determining which asset to download
 */
export type Platform = 'windows' | 'linux' | 'macos';

/**
 * Architecture type
 */
export type Architecture = 'x64' | 'arm64';

/**
 * Install type: portable vs installer
 */
export type InstallType = 'portable' | 'installer';

/**
 * Available release asset
 */
export interface ReleaseAsset {
  name: string;
  downloadUrl: string;
  size: number;
}

/**
 * GitHub release info
 */
export interface GitHubRelease {
  tag_name: string;
  name: string;
  assets: Array<{
    name: string;
    browser_download_url: string;
    size: number;
  }>;
  published_at: string;
}

/**
 * Detect the current platform (Windows, Linux, macOS)
 */
export function getPlatform(): Platform {
  const ua = navigator.userAgent.toLowerCase();

  if (ua.includes('win')) return 'windows';
  if (ua.includes('linux')) return 'linux';
  if (ua.includes('mac')) return 'macos';

  return 'windows'; // Default fallback
}

/**
 * Detect the current architecture (x64 or arm64)
 * Note: In Electron context, this will be called from the main process
 */
export function getArchitecture(): Architecture {
  // This will be populated by the Electron main process via IPC
  // For web context, we'll default to x64
  return 'x64';
}

/**
 * Compare semantic versions (returns: -1 if a < b, 0 if a == b, 1 if a > b)
 */
export function compareVersions(a: string, b: string): number {
  const parseVersion = (v: string) => {
    const clean = v.replace(/^v/, '').split('-')[0]; // Handle v prefix and pre-release
    const parts = clean.split('.').map(p => parseInt(p, 10));
    return [parts[0] || 0, parts[1] || 0, parts[2] || 0];
  };

  const [aMajor, aMinor, aPatch] = parseVersion(a);
  const [bMajor, bMinor, bPatch] = parseVersion(b);

  if (aMajor !== bMajor) return aMajor < bMajor ? -1 : 1;
  if (aMinor !== bMinor) return aMinor < bMinor ? -1 : 1;
  if (aPatch !== bPatch) return aPatch < bPatch ? -1 : 1;
  return 0;
}

/**
 * Find the best matching asset for the given platform, architecture, and install type
 */
export function findMatchingAsset(
  assets: ReleaseAsset[],
  platform: Platform,
  arch: Architecture,
  installType: InstallType,
): ReleaseAsset | null {
  // Build expected filename patterns
  let patterns: string[] = [];

  if (platform === 'windows') {
    const archSuffix = arch === 'arm64' ? 'arm64' : 'x64';
    const typeSuffix = installType === 'portable' ? 'portable' : 'installer';
    patterns = [`PP-MD-*-${archSuffix}-${typeSuffix}.exe`];
  } else if (platform === 'linux') {
    const archSuffix = arch === 'arm64' ? 'arm64' : 'x64';
    patterns = [`PP-MD-*-${archSuffix}.AppImage`, `PP-MD-*.AppImage`];
  } else if (platform === 'macos') {
    const archSuffix = arch === 'arm64' ? 'arm64' : 'x64';
    const ext = installType === 'portable' ? 'mac.zip' : 'dmg';
    patterns = [`PP-MD-*-${archSuffix}-${ext}`, `PP-MD-*-${ext}`];
  }

  // Try to find a matching asset
  for (const pattern of patterns) {
    const regex = new RegExp(`^${pattern.replace(/\*/g, '[^/]*')}$`);
    const match = assets.find(a => regex.test(a.name));
    if (match) return match;
  }

  return null;
}

/**
 * Format file size for display
 */
export function formatFileSize(bytes: number): string {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
}
