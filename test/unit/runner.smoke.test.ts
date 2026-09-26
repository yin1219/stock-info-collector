import { describe, expect, it } from 'vitest';

describe('test infrastructure / unit runner', () => {
  it('runs TypeScript tests in the node project', () => {
    expect(process.release.name).toBe('node');
  });
});
