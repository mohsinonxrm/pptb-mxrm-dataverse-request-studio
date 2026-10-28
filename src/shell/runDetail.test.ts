import { describe, expect, it } from 'vitest';
import { runDetail } from './Sidebar';

// #309: Recent runs read "204 · 1 rows" for writes.
describe('runDetail', () => {
  it('counts rows for reads, with the right plural', () => {
    expect(runDetail({ method: 'GET', rowCount: 1 })).toBe(' · 1 row');
    expect(runDetail({ method: 'GET', rowCount: 0 })).toBe(' · 0 rows');
    expect(runDetail({ method: 'GET', rowCount: 1234 })).toBe(` · ${(1234).toLocaleString()} rows`);
    expect(runDetail({ method: 'GET' })).toBe(' · — rows');
  });

  it('shows nothing extra for a write without a summary', () => {
    expect(runDetail({ method: 'PATCH', rowCount: 1 })).toBe('');
    expect(runDetail({ method: 'DELETE', rowCount: 0 })).toBe('');
    expect(runDetail({ method: 'POST', rowCount: 1 })).toBe('');
  });

  it("prefers the mode's summary", () => {
    expect(runDetail({ method: 'GET', rowCount: 6, summary: '6 rows · all pages' })).toBe(
      ' · 6 rows · all pages',
    );
    expect(runDetail({ method: 'POST', summary: '3 requests · all succeeded' })).toBe(
      ' · 3 requests · all succeeded',
    );
  });
});
