/**
 * @file useUpdateCheck.ts
 * @description React hook for managing update checking logic.
 */

import { useState, useCallback } from 'react';
import type {
  ReleaseAsset,
  Platform,
  Architecture,
  InstallType,
  GitHubRelease,
} from '../utils/versionUtils';
import { compareVersions, findMatchingAsset, getPlatform } from '../utils/versionUtils';

export interface UpdateCheckResult {
  hasUpdate: boolean;
  currentVersion: string;
  latestVersion: string;
  releaseUrl: string;
  downloadAsset: ReleaseAsset | null;
  releaseDate: string;
}

export interface UpdateCheckState {
  checking: boolean;
  result: UpdateCheckResult | null;
  error: string | null;
}

/**
 * Hook for checking for app updates via GitHub API
 */
export function useUpdateCheck() {
  const [state, setState] = useState<UpdateCheckState>({
    checking: false,
    result: null,
    error: null,
  });

  const getRuntimeWindow = (): Window & {
    electron?: {
      invoke: (channel: string, ...args: unknown[]) => Promise<any>;
    };
    __PPMD_VERSION__?: string;
  } => window as Window & {
    electron?: {
      invoke: (channel: string, ...args: unknown[]) => Promise<any>;
    };
    __PPMD_VERSION__?: string;
  };

  const fetchLatestRelease = useCallback(async (): Promise<GitHubRelease & { html_url: string }> => {
    const runtimeWindow = getRuntimeWindow();

    if (runtimeWindow.electron?.invoke) {
      return runtimeWindow.electron.invoke('check-for-updates');
    }

    const response = await fetch(
      'https://api.github.com/repos/Hart365/PP-MD/releases/latest',
      {
        headers: {
          Accept: 'application/vnd.github+json',
        },
      },
    );

    if (!response.ok) {
      throw new Error(`GitHub API error: ${response.status}`);
    }

    return response.json();
  }, []);

  /**
   * Check for updates by querying GitHub API
   */
  const checkForUpdates = useCallback(async () => {
    setState({ checking: true, result: null, error: null });

    try {
      // Get current app version from window (set by Electron)
      const runtimeWindow = getRuntimeWindow();
      const currentVersion = runtimeWindow.__PPMD_VERSION__ || '1.0.0';
      
      // Get current platform/arch info from Electron main process
      let platform: Platform = getPlatform();
      let arch: Architecture = 'x64';
      let installType: InstallType = 'portable';

      // Request info from Electron main process via IPC
      if (runtimeWindow.electron?.invoke) {
        try {
          const info = await runtimeWindow.electron.invoke('get-app-info');
          platform = info.platform;
          arch = info.architecture;
          installType = info.installType;
        } catch (e) {
          console.warn('Failed to get app info from Electron:', e);
        }
      }

      // Fetch latest release from GitHub (via main process in Electron builds)
      const release = await fetchLatestRelease();
      const latestVersion = release.tag_name;

      // Compare versions
      const versionComparison = compareVersions(currentVersion, latestVersion);

      if (versionComparison >= 0) {
        // Already on latest or newer version
        setState({
          checking: false,
          result: {
            hasUpdate: false,
            currentVersion,
            latestVersion,
            releaseUrl: release.html_url,
            downloadAsset: null,
            releaseDate: release.published_at,
          },
          error: null,
        });
        return;
      }

      // Parse release assets
      const assets: ReleaseAsset[] = (release.assets || []).map((asset: any) => ({
        name: asset.name,
        downloadUrl: asset.browser_download_url,
        size: asset.size,
      }));

      // Find matching asset for current platform/arch
      const downloadAsset = findMatchingAsset(assets, platform, arch, installType);

      setState({
        checking: false,
        result: {
          hasUpdate: true,
          currentVersion,
          latestVersion,
          releaseUrl: release.html_url,
          downloadAsset,
          releaseDate: release.published_at,
        },
        error: null,
      });
    } catch (err) {
      const errorMessage =
        err instanceof TypeError
          ? 'Network error while checking updates. Please verify your connection and try again.'
          : err instanceof Error
            ? err.message
            : 'Unknown error';
      setState({
        checking: false,
        result: null,
        error: errorMessage,
      });
    }
  }, [fetchLatestRelease]);

  /**
   * Clear the update check result
   */
  const dismiss = useCallback(() => {
    setState({ checking: false, result: null, error: null });
  }, []);

  return {
    ...state,
    checkForUpdates,
    dismiss,
  };
}
