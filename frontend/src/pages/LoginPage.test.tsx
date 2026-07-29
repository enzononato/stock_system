import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { render, screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { AuthProvider } from '@/contexts/AuthContext'
import LoginPage from '@/pages/LoginPage'

function renderLoginPage() {
  return render(
    <MemoryRouter initialEntries={['/login']}>
      <AuthProvider>
        <LoginPage />
      </AuthProvider>
    </MemoryRouter>
  )
}

describe('LoginPage', () => {
  it('credencial inválida mostra a mensagem de erro e não deixa a tela em estado de carregando', async () => {
    server.use(
      // AuthProvider tenta um refresh silencioso ao montar — sem sessão prévia, 401 é o esperado.
      http.post('/api/auth/refresh', () => new HttpResponse(null, { status: 401 })),
      http.post('/api/auth/login', () =>
        HttpResponse.json({ detail: 'Usuário ou senha inválidos.' }, { status: 401 })
      )
    )

    const user = userEvent.setup()
    renderLoginPage()

    await user.type(screen.getByLabelText('Usuário'), 'usuario.teste')
    await user.type(screen.getByLabelText('Senha'), 'senha-errada')

    const botao = screen.getByRole('button', { name: /Entrar/i })
    await user.click(botao)

    expect(await screen.findByText('Usuário ou senha inválidos.')).toBeInTheDocument()

    // Não fica "preso" em carregando: o botão volta ao texto normal e fica habilitado.
    await waitFor(() => expect(screen.getByRole('button', { name: 'Entrar' })).toBeInTheDocument())
    expect(screen.getByRole('button', { name: 'Entrar' })).not.toBeDisabled()
  })

  it('credencial válida navega para fora da tela de login (não mostra erro)', async () => {
    server.use(
      http.post('/api/auth/refresh', () => new HttpResponse(null, { status: 401 })),
      http.post('/api/auth/login', () => HttpResponse.json({ access_token: 'tok-abc' })),
      http.get('/api/auth/me', () => HttpResponse.json({ id: 1, username: 'gestor', role: 'Gestor' }))
    )

    const user = userEvent.setup()
    renderLoginPage()

    await user.type(screen.getByLabelText('Usuário'), 'gestor')
    await user.type(screen.getByLabelText('Senha'), 'senha-certa')
    await user.click(screen.getByRole('button', { name: /Entrar/i }))

    await waitFor(() => expect(screen.queryByText('Usuário ou senha inválidos.')).not.toBeInTheDocument())
  })
})
