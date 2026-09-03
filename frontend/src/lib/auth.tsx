import { createContext, useCallback, useContext, useEffect, useState, type ReactNode } from "react";
import { toast } from "sonner";

import { setAccessToken, setSessionExpiredHandler } from "@/api/client";
import { getMe, login as apiLogin, logout as apiLogout } from "@/api/auth";
import { demoLogin, demoMe, demoRefresh, isDemoApiEnabled } from "@/demo/mockApi";

export const ROLES = ["Gestor", "Técnico", "Usuário"] as const;
export type Role = (typeof ROLES)[number];

export interface AuthUser {
  id: number;
  username: string;
  role: string;
}

interface AuthContextValue {
  user: AuthUser | null;
  isLoading: boolean;
  login: (username: string, password: string) => Promise<void>;
  logout: () => Promise<void>;
  hasRole: (...roles: string[]) => boolean;
}

const AuthContext = createContext<AuthContextValue | null>(null);

function demoUser(): AuthUser {
  const user = demoMe();
  if (!user) throw new Error("Modo Demo indisponível.");
  return { id: user.id, username: user.email, role: user.role };
}

export function AuthProvider({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<AuthUser | null>(null);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    let cancelled = false;
    async function initialize() {
      try {
        if (isDemoApiEnabled()) {
          const data = demoRefresh();
          if (data) {
            setAccessToken(data.access_token);
            if (!cancelled) setUser(demoUser());
          }
          return;
        }
        const res = await fetch("/api/auth/refresh", { method: "POST", credentials: "include" });
        if (res.ok) {
          const data = (await res.json()) as { access_token: string };
          setAccessToken(data.access_token);
          const me = await getMe();
          if (!cancelled) setUser(me);
        }
      } catch {
        // Não autenticado — fluxo normal.
      } finally {
        if (!cancelled) setIsLoading(false);
      }
    }
    void initialize();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    setSessionExpiredHandler(() => {
      setAccessToken(null);
      setUser(null);
      toast.error("Sessão expirada", { description: "Faça login novamente para continuar." });
    });
    return () => setSessionExpiredHandler(null);
  }, []);

  const login = useCallback(async (username: string, password: string) => {
    if (isDemoApiEnabled()) {
      const data = demoLogin(username, password);
      if (!data) throw new Error("Modo Demo indisponível.");
      setAccessToken(data.access_token);
      setUser({ id: data.user.id, username: data.user.email, role: data.user.role });
      return;
    }
    const data = await apiLogin(username, password);
    setAccessToken(data.access_token);
    const me = await getMe();
    setUser(me);
  }, []);

  const logout = useCallback(async () => {
    if (isDemoApiEnabled()) {
      setAccessToken(null);
      setUser(null);
      return;
    }
    try {
      await apiLogout();
    } finally {
      setAccessToken(null);
      setUser(null);
    }
  }, []);

  const hasRole = useCallback((...roles: string[]) => !!user && roles.includes(user.role), [user]);

  return (
    <AuthContext.Provider value={{ user, isLoading, login, logout, hasRole }}>
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error("useAuth deve ser usado dentro de AuthProvider");
  return ctx;
}
