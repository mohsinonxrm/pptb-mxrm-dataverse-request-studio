// The scope chip keeps these hooks mounted all session (#42): their lists
// must follow "Refresh metadata", never show another org's data, and not
// stay "loading" after a fetch is cut short.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const host = vi.hoisted(() => ({
  org: 'a.crm.dynamics.com',
  resetListeners: new Set<() => void>(),
  cache: new Map<string, unknown>(),
  pending: [] as Array<{
    org: string;
    resolve: (v: unknown) => void;
    reject: (e: unknown) => void;
  }>,
  fetches: 0,
}));
vi.mock('./activeOrg', () => ({ useActiveOrg: () => host.org }));
vi.mock('./metadataProvider', () => ({
  subscribeMetadataReset: (cb: () => void) => {
    host.resetListeners.add(cb);
    return () => host.resetListeners.delete(cb);
  },
}));
vi.mock('./cache', () => ({
  pinMetadataCache: () => {
    const org = host.org;
    return {
      getPublishersWithSolutions: () => host.cache.get(`p:${org}`),
      setPublishersWithSolutions: (v: unknown) => host.cache.set(`p:${org}`, v),
      getSolutions: () => host.cache.get(`s:${org}`),
      setSolutions: (_k: unknown, v: unknown) => host.cache.set(`s:${org}`, v),
      getFilteredEntities: () => undefined,
      setFilteredEntities: () => {},
    };
  },
}));
vi.mock('./pptbClient', () => ({
  getPublishersWithSolutions: () => {
    host.fetches++;
    const org = host.org;
    return new Promise((resolve, reject) => host.pending.push({ org, resolve, reject }));
  },
  getAllSolutionsWithEntities: () => {
    host.fetches++;
    const org = host.org;
    return new Promise((resolve, reject) => host.pending.push({ org, resolve, reject }));
  },
  getSolutionComponents: async () => [],
  filterCachedEntitiesByNames: () => [],
}));

import { usePublisherFilter } from './usePublisherFilter';
import { useSolutionFilter } from './useSolutionFilter';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const pub = (org: string) => [
  {
    publisher: {
      publisherid: `${org}-p`,
      friendlyname: `${org} publisher`,
      uniquename: 'p',
      customizationprefix: 'p',
    },
    solutions: [],
  },
];
const sol = (org: string) => [
  { solutionid: `${org}-s`, friendlyname: `${org} solution`, uniquename: 's' },
];

let root: Root | null = null;
let pubResult: ReturnType<typeof usePublisherFilter> | null = null;
let solResult: ReturnType<typeof useSolutionFilter> | null = null;
function PublisherProbe() {
  pubResult = usePublisherFilter(true);
  return null;
}
function SolutionProbe() {
  solResult = useSolutionFilter(true);
  return null;
}
const render = (which: 'pub' | 'sol') =>
  act(() => root!.render(which === 'pub' ? <PublisherProbe /> : <SolutionProbe />));
const settle = async (i: number, value: unknown) => {
  await act(async () => {
    host.pending[i].resolve(value);
    await Promise.resolve();
  });
};

beforeEach(() => {
  host.org = 'a.crm.dynamics.com';
  host.cache.clear();
  host.pending = [];
  host.fetches = 0;
  host.resetListeners.clear();
  root = createRoot(document.createElement('div'));
});
afterEach(() => {
  act(() => root?.unmount());
  root = null;
});

describe('usePublisherFilter / useSolutionFilter', () => {
  it('refetch after "Refresh metadata" clears the cache', async () => {
    render('pub');
    await settle(0, pub('a'));
    expect(pubResult!.publishers.map((p) => p.friendlyname)).toEqual(['a publisher']);

    host.cache.clear();
    act(() => host.resetListeners.forEach((l) => l()));
    expect(host.fetches).toBe(2);
    await settle(1, [...pub('a'), ...pub('new')]);
    expect(pubResult!.publishers).toHaveLength(2);
  });

  it("never show another org's list, and don't stay loading when switched to a cached org", async () => {
    host.cache.set('p:b.crm.dynamics.com', pub('b'));
    render('pub');
    expect(pubResult!.publishersLoading).toBe(true);

    // Switch to org b (cached) while a's fetch is still in flight.
    host.org = 'b.crm.dynamics.com';
    render('pub');
    expect(pubResult!.publishersLoading).toBe(false);
    expect(pubResult!.publishers.map((p) => p.friendlyname)).toEqual(['b publisher']);

    // a's late answer doesn't replace b's list.
    await settle(0, pub('a'));
    expect(pubResult!.publishers.map((p) => p.friendlyname)).toEqual(['b publisher']);
  });

  it("drop a failed org's list after switching, instead of offering it", async () => {
    render('sol');
    await settle(0, sol('a'));
    expect(solResult!.solutions.map((s) => s.friendlyname)).toEqual(['a solution']);

    host.org = 'c.crm.dynamics.com';
    render('sol');
    expect(solResult!.solutions).toEqual([]);
    await act(async () => {
      host.pending[1].reject(new Error('no access'));
      await Promise.resolve();
    });
    expect(solResult!.solutions).toEqual([]);
    expect(solResult!.solutionsError).toBe('no access');
    expect(solResult!.solutionsLoading).toBe(false);
  });
});
