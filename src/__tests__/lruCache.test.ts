import { describe, expect, it } from 'vitest';
import { LruCache } from '../utils/lruCache';

describe('LruCache', () => {
  it('evicts the least recently used diagram', () => {
    const cache = new LruCache<string, number>(2);
    cache.set('first', 1);
    cache.set('second', 2);
    expect(cache.get('first')).toBe(1);
    cache.set('third', 3);
    expect(cache.get('second')).toBeUndefined();
    expect(cache.get('first')).toBe(1);
    expect(cache.get('third')).toBe(3);
    expect(cache.size).toBe(2);
  });

  it('refreshes keys, clears entries and rejects invalid limits', () => {
    const cache = new LruCache<string, number>(1);
    cache.set('first', 1);
    cache.set('first', 2);
    expect(cache.size).toBe(1);
    expect(cache.get('first')).toBe(2);
    cache.clear();
    expect(cache.size).toBe(0);
    expect(() => new LruCache(0)).toThrow(RangeError);
  });
});