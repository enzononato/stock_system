import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import RedefinirSenhaPage from '@/pages/RedefinirSenhaPage'

const MSG_LINK = 'Link inválido ou expirado. Peça um novo.'

function renderPagina(url = '/redefinir-senha?token=abc') {
  return render(
    <MemoryRouter initialEntries={[url]}>
      <RedefinirSenhaPage />
    </MemoryRouter>
  )
}

async function preencher(user: ReturnType<typeof userEvent.setup>, nova: string, confirma: string) {
  await user.type(screen.getByLabelText('Nova senha'), nova)
  await user.type(screen.getByLabelText('Confirme a nova senha'), confirma)
  await user.click(screen.getByRole('button', { name: 'Salvar nova senha' }))
}

describe('RedefinirSenhaPage', () => {
  it('envia o token da URL com a nova senha e oferece ir para o login', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/auth/reset', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ detail: 'Senha redefinida. Entre com a nova senha.' })
      })
    )
    const user = userEvent.setup()
    renderPagina()
    await preencher(user, 'Nova#123', 'Nova#123')

    expect(await screen.findByText('Senha redefinida. Entre com a nova senha.')).toBeInTheDocument()
    expect(corpo).toEqual({ token: 'abc', nova_senha: 'Nova#123' })
    expect(screen.getByRole('link', { name: 'Ir para o login' })).toHaveAttribute('href', '/login')
  })

  it('senhas diferentes não chamam a API', async () => {
    let chamadas = 0
    server.use(
      http.post('/api/auth/reset', () => {
        chamadas += 1
        return HttpResponse.json({ detail: 'ok' })
      })
    )
    const user = userEvent.setup()
    renderPagina()
    await preencher(user, 'Nova#123', 'Outra#123')

    expect(await screen.findByText('As senhas não conferem.')).toBeInTheDocument()
    expect(chamadas).toBe(0)
  })

  it('o botão fica desabilitado com a nova senha vazia', () => {
    renderPagina()
    expect(screen.getByRole('button', { name: 'Salvar nova senha' })).toBeDisabled()
  })

  it('sem token mostra o link inválido e leva a pedir um novo', () => {
    renderPagina('/redefinir-senha')

    expect(screen.getByText(MSG_LINK)).toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'Pedir novo link' })).toHaveAttribute('href', '/esqueci-senha')
    expect(screen.queryByLabelText('Nova senha')).not.toBeInTheDocument()
  })

  it('link recusado pelo servidor mostra a mensagem e o atalho para pedir outro', async () => {
    server.use(
      http.post('/api/auth/reset', () => HttpResponse.json({ detail: MSG_LINK }, { status: 400 }))
    )
    const user = userEvent.setup()
    renderPagina()
    await preencher(user, 'Nova#123', 'Nova#123')

    expect(await screen.findByText(MSG_LINK)).toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'Pedir novo link' })).toHaveAttribute('href', '/esqueci-senha')
  })
})
