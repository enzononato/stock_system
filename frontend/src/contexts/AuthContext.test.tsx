import type { ReactNode } from 'react'
import { afterEach, describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { act, renderHook, waitFor } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { setAccessToken } from '@/api/client'
import { AuthProvider, useAuth } from '@/contexts/AuthContext'

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

afterEach(() => setAccessToken(null))

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
