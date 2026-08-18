import { describe, expect, it } from 'vitest';
import { isSafeExternalUrl } from '../utils/security';

describe('security utilities', () => {
  it('allows secure http and https destinations', () => {
    expect(isSafeExternalUrl('https://example.com/download')).toBe(true);
    expect(isSafeExternalUrl('http://example.com/download')).toBe(true);
  });

  it('rejects dangerous schemes and malformed targets', () => {
    expect(isSafeExternalUrl('javascript:alert(1)')).toBe(false);
    expect(isSafeExternalUrl('data:text/html;base64,PHNjcmlwdD4=')).toBe(false);
    expect(isSafeExternalUrl('ftp://example.com/file.zip')).toBe(false);
    expect(isSafeExternalUrl('')).toBe(false);
  });
});
