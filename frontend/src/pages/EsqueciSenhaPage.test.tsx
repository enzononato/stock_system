import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import App from '@/App'
import { AuthProvider } from '@/contexts/AuthContext'
import EsqueciSenhaPage from '@/pages/EsqueciSenhaPage'

const MSG_FORGOT =
  'Se o usuário existir e tiver e-mail cadastrado, enviaremos um link para redefinir a senha.'

function renderPagina() {
  return renderWithClient(
    <MemoryRouter initialEntries={['/esqueci-senha']}>
      <EsqueciSenhaPage />
    </MemoryRouter>
  )
}

describe('EsqueciSenhaPage', () => {
  it('envia o identificador e mostra a resposta genérica do servidor', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/auth/forgot', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ detail: MSG_FORGOT })
      })
    )
    const user = userEvent.setup()
    renderPagina()

    expect(screen.getByRole('heading', { name: 'Esqueci minha senha' })).toBeInTheDocument()
    await user.type(screen.getByLabelText('Usuário ou e-mail'), 'joao.silva')
    await user.click(screen.getByRole('button', { name: 'Enviar link' }))

    expect(await screen.findByRole('status')).toHaveTextContent(MSG_FORGOT)
    expect(corpo).toEqual({ identificador: 'joao.silva' })
    expect(screen.queryByLabelText('Usuário ou e-mail')).not.toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'Voltar para o login' })).toHaveAttribute('href', '/login')
  })

  it('429 mostra a mensagem de muitas tentativas e mantém o formulário', async () => {
    server.use(
      http.post('/api/auth/forgot', () =>
        HttpResponse.json({ error: 'Rate limit exceeded: 5 per 1 minute' }, { status: 429 })
      )
    )
    const user = userEvent.setup()
    renderPagina()

    await user.type(screen.getByLabelText('Usuário ou e-mail'), 'joao.silva')
    await user.click(screen.getByRole('button', { name: 'Enviar link' }))

    expect(
      await screen.findByText('Muitas tentativas. Aguarde um minuto e tente de novo.')
    ).toBeInTheDocument()
    expect(screen.getByLabelText('Usuário ou e-mail')).toBeInTheDocument()
  })

  it('rota pública: abre sem sessão, sem redirecionar para o login', async () => {
    // O App monta o PosLoginRedirect (F6), que precisa do AuthProvider; o refresh devolve 401 (sem sessão).
    server.use(http.post('/api/auth/refresh', () => new HttpResponse(null, { status: 401 })))
    renderWithClient(
      <MemoryRouter initialEntries={['/esqueci-senha']}>
        <AuthProvider>
          <App />
        </AuthProvider>
      </MemoryRouter>
    )

    expect(await screen.findByRole('heading', { name: 'Esqueci minha senha' })).toBeInTheDocument()
  })
})
