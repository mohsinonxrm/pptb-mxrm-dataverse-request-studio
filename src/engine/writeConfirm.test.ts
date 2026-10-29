import { describe, expect, it } from 'vitest';
import {
  confirmWordOf,
  describeWriteScope,
  isProductionEnvironment,
  isTypedConfirmation,
  needsWriteConfirmation,
  requiresTypedConfirmation,
  type WriteConfirmSpec,
} from './writeConfirm';

const update: WriteConfirmSpec = {
  operation: 'Update Record',
  table: 'account',
  scope: { kind: 'records', count: 1 },
  highImpact: false,
};
const merge: WriteConfirmSpec = {
  operation: 'Merge Records',
  table: 'account',
  scope: { kind: 'records', count: 2 },
  highImpact: true,
};
const bulk = (count: number): WriteConfirmSpec => ({
  operation: 'Bulk Write · CreateMultiple',
  table: 'account',
  scope: { kind: 'records', count },
  highImpact: true,
});

describe('isProductionEnvironment', () => {
  it('matches PPTB\'s "Production" type only', () => {
    expect(isProductionEnvironment('Production')).toBe(true);
    expect(isProductionEnvironment(' production ')).toBe(true);
    for (const e of ['Dev', 'Test', 'UAT', '', undefined])
      expect(isProductionEnvironment(e)).toBe(false);
  });
});

describe('needsWriteConfirmation', () => {
  it('never asks for a read', () => {
    expect(needsWriteConfirmation(null, 'Production', true)).toBe(false);
  });

  it('always asks for high-impact writes', () => {
    expect(needsWriteConfirmation(merge, 'Dev', true)).toBe(true);
    expect(needsWriteConfirmation(merge, 'Dev', false)).toBe(true);
  });

  it('asks for other writes only on Production with safe mode on', () => {
    expect(needsWriteConfirmation(update, 'Dev', true)).toBe(false);
    expect(needsWriteConfirmation(update, 'Production', true)).toBe(true);
    expect(needsWriteConfirmation(update, 'Production', false)).toBe(false);
  });
});

describe('requiresTypedConfirmation', () => {
  it('requires typing on Production with safe mode on', () => {
    expect(requiresTypedConfirmation(merge, 'Production', true)).toBe(true);
    expect(requiresTypedConfirmation(update, 'Production', true)).toBe(true);
  });

  it('requires typing for more than 100 records anywhere', () => {
    expect(requiresTypedConfirmation(bulk(100), 'Dev', true)).toBe(false);
    expect(requiresTypedConfirmation(bulk(101), 'Dev', true)).toBe(true);
    expect(requiresTypedConfirmation(bulk(101), 'Production', false)).toBe(true);
  });

  it('counts a worded scope when it carries a count (a batch with more than 100 writes)', () => {
    const batch = (count?: number) => ({
      operation: '$batch',
      scope: { kind: 'text' as const, text: 'n requests', count },
      highImpact: true,
    });
    expect(requiresTypedConfirmation(batch(100), 'Dev', true)).toBe(false);
    expect(requiresTypedConfirmation(batch(101), 'Dev', true)).toBe(true);
    expect(requiresTypedConfirmation(batch(), 'Dev', true)).toBe(false);
  });

  it("doesn't require typing for a small high-impact write outside Production", () => {
    expect(requiresTypedConfirmation(merge, 'Dev', true)).toBe(false);
    expect(requiresTypedConfirmation(merge, 'Production', false)).toBe(false);
  });
});

describe('wording', () => {
  it('describes the scope', () => {
    expect(describeWriteScope({ kind: 'records', count: 1 })).toBe('1 record');
    expect(describeWriteScope({ kind: 'records', count: 1500 })).toBe(
      `${(1500).toLocaleString()} records`,
    );
    expect(describeWriteScope({ kind: 'text', text: '3 requests, 2 of them writes' })).toBe(
      '3 requests, 2 of them writes',
    );
  });

  it('picks the confirmation word: given, else table, else operation', () => {
    expect(confirmWordOf({ ...merge, confirmWord: 'PublishAllXml' })).toBe('PublishAllXml');
    expect(confirmWordOf(merge)).toBe('account');
    expect(confirmWordOf({ ...merge, table: undefined })).toBe('Merge Records');
  });

  it('matches the typed word ignoring case and spaces, never an empty word', () => {
    expect(isTypedConfirmation(' Account ', 'account')).toBe(true);
    expect(isTypedConfirmation('accounts', 'account')).toBe(false);
    expect(isTypedConfirmation('', '')).toBe(false);
  });
});
