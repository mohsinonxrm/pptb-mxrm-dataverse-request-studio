import { createContext, useContext, useEffect, useState, type ReactNode } from 'react';
import { subscribe, getSession, type HostSession } from './pptbBridge';

const HostContext = createContext<HostSession>(getSession());

export function HostProvider({ children }: { children: ReactNode }) {
  const [session, setSession] = useState<HostSession>(getSession());
  useEffect(() => subscribe((s) => setSession(s)), []);
  return <HostContext.Provider value={session}>{children}</HostContext.Provider>;
}

export function useHostSession(): HostSession {
  return useContext(HostContext);
}
