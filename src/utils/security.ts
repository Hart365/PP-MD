const SAFE_PROTOCOLS = new Set(['http:', 'https:']);
const BASE64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

function base64ToUint8Array(value: string): Uint8Array {
  const normalized = value.replace(/\s+/g, '');
  if (!normalized) {
    return new Uint8Array();
  }

  const padded = normalized.length % 4 === 0
    ? normalized
    : normalized + '='.repeat((4 - (normalized.length % 4)) % 4);

  const bytes: number[] = [];
  for (let index = 0; index < padded.length; index += 4) {
    const chunk = padded.slice(index, index + 4);
    const values = chunk.split('').map((char) => {
      if (char === '=') return 0;
      const code = BASE64_ALPHABET.indexOf(char);
      return code >= 0 ? code : 0;
    });

    const a = values[0] ?? 0;
    const b = values[1] ?? 0;
    const c = values[2] ?? 0;
    const d = values[3] ?? 0;

    bytes.push((a << 2) | (b >> 4));
    if (chunk[2] !== '=') {
      bytes.push(((b & 0x0f) << 4) | (c >> 2));
    }
    if (chunk[3] !== '=') {
      bytes.push(((c & 0x03) << 6) | d);
    }
  }

  return Uint8Array.from(bytes);
}

export function encodeBase64Text(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let output = '';

  for (let index = 0; index < bytes.length; index += 3) {
    const a = bytes[index] ?? 0;
    const b = bytes[index + 1] ?? 0;
    const c = bytes[index + 2] ?? 0;

    output += BASE64_ALPHABET[a >> 2];
    output += BASE64_ALPHABET[((a & 0x03) << 4) | (b >> 4)];
    output += bytes[index + 1] !== undefined ? BASE64_ALPHABET[((b & 0x0f) << 2) | (c >> 6)] : '=';
    output += bytes[index + 2] !== undefined ? BASE64_ALPHABET[c & 0x3f] : '=';
  }

  return output;
}

export function decodeBase64Text(value: string): string {
  const normalized = value.replace(/\s+/g, '');
  if (!normalized) {
    return '';
  }

  const bytes = base64ToUint8Array(normalized);
  return new TextDecoder('utf-8', { fatal: false }).decode(bytes);
}

export function isSafeExternalUrl(value: string | null | undefined): boolean {
  if (!value || typeof value !== 'string') {
    return false;
  }

  try {
    const url = new URL(value);
    return SAFE_PROTOCOLS.has(url.protocol);
  } catch {
    return false;
  }
}

export function openSafeExternalUrl(url: string): boolean {
  if (!isSafeExternalUrl(url)) {
    return false;
  }

  const opened = window.open(url, '_blank', 'noopener,noreferrer');
  return opened !== null;
}
