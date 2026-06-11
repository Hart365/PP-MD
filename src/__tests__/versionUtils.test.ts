import { compareVersions, formatFileSize } from '../utils/versionUtils';

describe('versionUtils', () => {
  it('compares semantic versions correctly', () => {
    expect(compareVersions('v1.0.4', '1.0.3')).toBe(1);
    expect(compareVersions('1.0.4', '1.0.4')).toBe(0);
    expect(compareVersions('1.0.2', '1.0.4')).toBe(-1);
  });

  it('formats file sizes for display', () => {
    expect(formatFileSize(0)).toBe('0 B');
    expect(formatFileSize(1024)).toBe('1 KB');
  });
});
