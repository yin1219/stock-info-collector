import { createHash } from 'node:crypto';
import type { MaterialSourceRecord } from '../providers/mops-material';

export interface ExistingMaterialVersion {
  id: string;
  sourceKey: string;
  contentFingerprint: string;
}

export interface PreparedMaterialEvent {
  sourceKey: string;
  contentFingerprint: string;
  revisionOf: string | null;
  kind: 'new' | 'duplicate' | 'revision';
  shouldNotify: boolean;
}

export function materialContentFingerprint(content: string): string {
  const normalized = content.normalize('NFC').replace(/\s+/g, ' ').trim();
  return createHash('sha256').update(normalized).digest('hex');
}

export function prepareMaterialEvent(
  record: MaterialSourceRecord,
  existing: readonly ExistingMaterialVersion[] = [],
): PreparedMaterialEvent {
  const contentFingerprint = materialContentFingerprint(record.content);
  const sourceVersions = existing.filter(({ sourceKey }) => sourceKey === record.sourceKey);
  const exact = sourceVersions.find(({ contentFingerprint: fingerprint }) => fingerprint === contentFingerprint);
  if (exact) {
    return { sourceKey: exact.sourceKey, contentFingerprint, revisionOf: null, kind: 'duplicate', shouldNotify: false };
  }

  const base = sourceVersions[0];
  const explicitParent = record.revisionOf
    ? existing.find(({ id, sourceKey }) => id === record.revisionOf || sourceKey === record.revisionOf)
    : undefined;
  const parent = base ?? explicitParent;
  if (parent) {
    const revisionKey = `${record.sourceKey}:revision:${contentFingerprint.slice(0, 20)}`;
    const repeatedRevision = existing.find(({ sourceKey, contentFingerprint: fingerprint }) =>
      sourceKey === revisionKey && fingerprint === contentFingerprint,
    );
    if (repeatedRevision) {
      return { sourceKey: repeatedRevision.sourceKey, contentFingerprint, revisionOf: parent.id, kind: 'duplicate', shouldNotify: false };
    }
    return { sourceKey: revisionKey, contentFingerprint, revisionOf: parent.id, kind: 'revision', shouldNotify: true };
  }

  return { sourceKey: record.sourceKey, contentFingerprint, revisionOf: null, kind: 'new', shouldNotify: true };
}
