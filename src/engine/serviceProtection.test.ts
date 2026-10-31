import { describe, expect, it } from 'vitest';
import { mapError } from './dataverseExecutor';
import { serviceProtectionLimit } from './serviceProtection';

// The host rejects with "<code>: <message>" behind Electron's IPC wrapper.
const hostError = (code: string, message: string) =>
  new Error(
    `Error invoking remote method 'dataverse.queryData': Error: Dataverse queryData failed: ${code}: ${message}`,
  );

describe('service protection limits (#269)', () => {
  it.each([
    [
      '0x80072322',
      'Number of requests exceeded the limit of 6000 over time window of 300 seconds.',
      'number of requests',
      /6,000 requests in 5 minutes/,
    ],
    [
      '0x80072321',
      'Combined execution time of incoming requests exceeded limit of 1,200,000 milliseconds over time window of 300 seconds.',
      'execution time',
      /20 minutes of server time/,
    ],
    [
      '0x80072326',
      'Number of concurrent requests exceeded the limit of 52.',
      'concurrent requests',
      /52 of your requests/,
    ],
  ])('%s maps to 429 and names its limit', (code, message, limit, advice) => {
    const r = mapError(hostError(code, message), 12);
    expect(r.status).toBe(429);
    expect(r.statusText).toBe('Too Many Requests');
    const l = serviceProtectionLimit(r);
    expect(l).toMatchObject({ code, limit });
    expect(l?.advice).toMatch(advice);
  });

  it('a 429 without a known code is still throttling', () => {
    const r = mapError(Object.assign(new Error('Too Many Requests'), { statusCode: 429 }), 1);
    expect(serviceProtectionLimit(r)?.limit).toBe('too many requests');
  });

  it('other failures are not throttling', () => {
    expect(
      serviceProtectionLimit(mapError(hostError('0x80040217', 'Does Not Exist'), 1)),
    ).toBeNull();
    expect(serviceProtectionLimit(null)).toBeNull();
  });
});
