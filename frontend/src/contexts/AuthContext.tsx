import React, { createContext, useContext, useState, useCallback, useEffect } from 'react'
import { setAccessToken, setSessionExpiredHandler } from '@/api/client'
import { login as apiLogin, logout as apiLogout, getMe, type LoginResponse } from '@/api/auth'
import { toast } from '@/components/ui/toast'
import { MSG_CADASTRO_2FA_OBRIGATORIO } from '@/lib/auth-errors'

interface AuthUser {
  id: number
  username: string
  role: string
}

interface AuthContextType {
  user: AuthUser | null
  isLoading: boolean
  login: (username: string, password: string) => Promise<LoginResponse>
  concluirLogin: (accessToken: string) => Promise<void>
  /** Aviso para a tela de login (ex.: Gestor barrado por falta de cadastro do 2FA). */
  avisoLogin: string | null
  limparAviso: () => void
  logout: () => Promise<void>
  hasRole: (...roles: string[]) => boolean
}

/** Pergunta ao servidor por que o refresh falha e devolve o aviso que a tela de login deve mostrar, se houver. */
async function motivoDoRefresh(): Promise<string | null> {
  try {
    const res = await fetch('/api/auth/refresh', { method: 'POST', credentials: 'include' })
    if (res.ok) return null
    const data = (await res.json()) as { detail?: unknown }
    return data.detail === MSG_CADASTRO_2FA_OBRIGATORIO ? MSG_CADASTRO_2FA_OBRIGATORIO : null
  } catch {
    return null
  }
}

const AuthContext = createContext<AuthContextType | null>(null)

export function AuthProvider({ children }: { children: React.ReactNode }) {
  const [user, setUser] = useState<AuthUser | null>(null)
  const [isLoading, setIsLoading] = useState(true)
  const [avisoLogin, setAvisoLogin] = useState<string | null>(null)

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
        } else {
          const data = (await res.json().catch(() => ({}))) as { detail?: unknown }
          if (data.detail === MSG_CADASTRO_2FA_OBRIGATORIO) setAvisoLogin(MSG_CADASTRO_2FA_OBRIGATORIO)
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
      // Gestor sem 2FA cadastrado também recebe 401: em vez de "sessão expirada", a tela de
      // login explica o que falta (o motivo vem do detail do refresh).
      void motivoDoRefresh().then((aviso) => {
        if (aviso) setAvisoLogin(aviso)
        else toast('Sessão expirada. Faça login novamente.', 'error')
      })
    })
    return () => setSessionExpiredHandler(null)
  }, [])

  const concluirLogin = useCallback(async (accessToken: string) => {
    setAccessToken(accessToken)
    const me = await getMe()
    setUser(me)
  }, [])

  // Só a etapa `ok` abre a sessão; `mfa` e `cadastro_2fa` devolvem a resposta
  // para a tela de login continuar (nenhum token nem usuário é guardado ainda).
  const login = useCallback(
    async (username: string, password: string) => {
      const data = await apiLogin(username, password)
      if (data.etapa === 'ok' && data.access_token) {
        await concluirLogin(data.access_token)
      }
      return data
    },
    [concluirLogin]
  )

  const limparAviso = useCallback(() => setAvisoLogin(null), [])

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
    <AuthContext.Provider value={{ user, isLoading, login, concluirLogin, avisoLogin, limparAviso, logout, hasRole }}>
      {children}
    </AuthContext.Provider>
  )
}

export function useAuth() {
  const ctx = useContext(AuthContext)
  if (!ctx) throw new Error('useAuth deve ser usado dentro de AuthProvider')
  return ctx
}
