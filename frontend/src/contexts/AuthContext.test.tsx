import type { ReactNode } from 'react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { act, renderHook, waitFor } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import api, { setAccessToken } from '@/api/client'
import { toast } from '@/components/ui/toast'
import { MSG_CADASTRO_2FA_OBRIGATORIO } from '@/lib/auth-errors'
import { AuthProvider, useAuth } from '@/contexts/AuthContext'

vi.mock('@/components/ui/toast', () => ({ toast: vi.fn() }))

function wrapper({ children }: { children: ReactNode }) {
  return (
    <MemoryRouter>
      <AuthProvider>{children}</AuthProvider>
    </MemoryRouter>
  )
}

async function montar() {
  const hook = renderHook(() => useAuth(), { wrapper })
  await waitFor(() => expect(hook.result.current.isLoading).toBe(false))
  return hook
}

afterEach(() => {
  setAccessToken(null)
  vi.mocked(toast).mockClear()
})

describe('AuthContext', () => {
  it('login com etapa ok guarda o usuário', async () => {
    server.use(
      http.post('/api/auth/refresh', () => new HttpResponse(null, { status: 401 })),
      http.post('/api/auth/login', () =>
        HttpResponse.json({ etapa: 'ok', access_token: 'tok-1', token_type: 'bearer' })
      ),
      http.get('/api/auth/me', () => HttpResponse.json({ id: 1, username: 'gestor', role: 'Gestor' }))
    )
    const { result } = await montar()

    await act(async () => {
      await result.current.login('gestor', 'senha')
    })

    expect(result.current.user?.username).toBe('gestor')
  })

  it('login com etapa mfa não chama /auth/me nem guarda o usuário', async () => {
    let chamadasMe = 0
    server.use(
      http.post('/api/auth/refresh', () => new HttpResponse(null, { status: 401 })),
      http.post('/api/auth/login', () =>
        HttpResponse.json({ etapa: 'mfa', mfa_token: 'mfa-1', expira_em_segundos: 300 })
      ),
      http.get('/api/auth/me', () => {
        chamadasMe += 1
        return HttpResponse.json({ id: 1, username: 'gestor', role: 'Gestor' })
      })
    )
    const { result } = await montar()

    let resposta: Awaited<ReturnType<typeof result.current.login>> | undefined
    await act(async () => {
      resposta = await result.current.login('gestor', 'senha')
    })

    expect(resposta?.etapa).toBe('mfa')
    expect(resposta?.mfa_token).toBe('mfa-1')
    expect(chamadasMe).toBe(0)
    expect(result.current.user).toBeNull()
  })

  it('concluirLogin usa o token recebido', async () => {
    let autorizacao: string | null = null
    server.use(
      http.post('/api/auth/refresh', () => new HttpResponse(null, { status: 401 })),
      http.get('/api/auth/me', ({ request }) => {
        autorizacao = request.headers.get('Authorization')
        return HttpResponse.json({ id: 2, username: 'tecnico', role: 'Técnico' })
      })
    )
    const { result } = await montar()

    await act(async () => {
      await result.current.concluirLogin('tok-2')
    })

    expect(autorizacao).toBe('Bearer tok-2')
    expect(result.current.user?.username).toBe('tecnico')
  })
})

describe('AuthContext — Gestor sem 2FA cadastrado', () => {
  const recusaPorCadastro = http.post('/api/auth/refresh', () =>
    HttpResponse.json({ detail: MSG_CADASTRO_2FA_OBRIGATORIO }, { status: 401 })
  )

  it('refresh recusado por falta de cadastro vira aviso para a tela de login', async () => {
    server.use(recusaPorCadastro)
    const { result } = await montar()

    await waitFor(() => expect(result.current.avisoLogin).toBe(MSG_CADASTRO_2FA_OBRIGATORIO))
    expect(result.current.user).toBeNull()
  })

  it('sessão aberta barrada no meio do uso mostra o aviso, não "sessão expirada"', async () => {
    server.use(
      http.post('/api/auth/refresh', () => new HttpResponse(null, { status: 401 })),
      http.get('/api/auth/me', () => HttpResponse.json({ id: 1, username: 'gestor', role: 'Gestor' })),
      http.get('/api/users', () =>
        HttpResponse.json({ detail: MSG_CADASTRO_2FA_OBRIGATORIO }, { status: 401 })
      )
    )
    const { result } = await montar()
    await act(async () => {
      await result.current.concluirLogin('tok-1')
    })
    server.use(recusaPorCadastro)

    await act(async () => {
      await api.get('/users').catch(() => undefined)
    })

    await waitFor(() => expect(result.current.avisoLogin).toBe(MSG_CADASTRO_2FA_OBRIGATORIO))
    expect(result.current.user).toBeNull()
    expect(toast).not.toHaveBeenCalled()
  })

  it('sessão expirada comum continua avisando com o toast', async () => {
    server.use(
      http.post('/api/auth/refresh', () => new HttpResponse(null, { status: 401 })),
      http.get('/api/auth/me', () => HttpResponse.json({ id: 1, username: 'gestor', role: 'Gestor' })),
      http.get('/api/users', () => new HttpResponse(null, { status: 401 }))
    )
    const { result } = await montar()
    await act(async () => {
      await result.current.concluirLogin('tok-1')
    })

    await act(async () => {
      await api.get('/users').catch(() => undefined)
    })

    await waitFor(() => expect(toast).toHaveBeenCalledWith('Sessão expirada. Faça login novamente.', 'error'))
    expect(result.current.avisoLogin).toBeNull()
  })
})
