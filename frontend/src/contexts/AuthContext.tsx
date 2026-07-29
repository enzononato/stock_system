import React, { createContext, useContext, useState, useCallback, useEffect } from 'react'
import { setAccessToken, setSessionExpiredHandler } from '@/api/client'
import { login as apiLogin, logout as apiLogout, getMe } from '@/api/auth'
import { toast } from '@/components/ui/toast'

interface AuthUser {
  id: number
  username: string
  role: string
}

interface AuthContextType {
  user: AuthUser | null
  isLoading: boolean
  login: (username: string, password: string) => Promise<void>
  logout: () => Promise<void>
  hasRole: (...roles: string[]) => boolean
}

const AuthContext = createContext<AuthContextType | null>(null)

export function AuthProvider({ children }: { children: React.ReactNode }) {
  const [user, setUser] = useState<AuthUser | null>(null)
  const [isLoading, setIsLoading] = useState(true)

  // Ao carregar, tenta renovar o access token via cookie
  useEffect(() => {
    async function tryRefresh() {
      try {
        const res = await fetch('/api/auth/refresh', { method: 'POST', credentials: 'include' })
        if (res.ok) {
          const data = await res.json()
          setAccessToken(data.access_token)
          const me = await getMe()
          setUser(me)
        }
      } catch {
        // Não estava autenticado - normal
      } finally {
        setIsLoading(false)
      }
    }
    tryRefresh()
  }, [])

  // Registra o callback chamado pelo interceptor de 401 do axios (api/client.ts)
  // quando o refresh do token falha de verdade (sessão expirada). Em vez de forçar
  // um hard reload com `window.location.href`, apenas limpamos o usuário aqui — o
  // AppLayout já redireciona para /login normalmente (via <Navigate>) assim que
  // `user` fica null, preservando o estado do React (SPA) durante a transição.
  useEffect(() => {
    setSessionExpiredHandler(() => {
      setAccessToken(null)
      setUser(null)
      toast('Sessão expirada. Faça login novamente.', 'error')
    })
    return () => setSessionExpiredHandler(null)
  }, [])

  const login = useCallback(async (username: string, password: string) => {
    const data = await apiLogin(username, password)
    setAccessToken(data.access_token)
    const me = await getMe()
    setUser(me)
  }, [])

  const logout = useCallback(async () => {
    await apiLogout()
    setAccessToken(null)
    setUser(null)
  }, [])

  const hasRole = useCallback(
    (...roles: string[]) => !!user && roles.includes(user.role),
    [user]
  )

  return (
    <AuthContext.Provider value={{ user, isLoading, login, logout, hasRole }}>
      {children}
    </AuthContext.Provider>
  )
}

export function useAuth() {
  const ctx = useContext(AuthContext)
  if (!ctx) throw new Error('useAuth deve ser usado dentro de AuthProvider')
  return ctx
}
