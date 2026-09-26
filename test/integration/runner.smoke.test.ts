import { describe, expect, it } from 'vitest';
import { createTestPaths } from '../support';

describe('test infrastructure / integration runner', () => {
  it('creates an isolated data directory for integration scenarios', async () => {
    const paths = await createTestPaths();
    try {
      expect(paths.userData).toContain(paths.root);
      expect(paths.database).toContain(paths.root);
    } finally {
      await paths.cleanup();
    }
  });
});
