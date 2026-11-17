import { beforeAll, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, within } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import UsersPage from '@/pages/UsersPage'
import type { User } from '@/api/users'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'gestor', role: 'Gestor' },
    isLoading: false,
    login: vi.fn(),
    concluirLogin: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Gestor'),
  }),
}))

// jsdom não implementa a Pointer Events API que o Select (Radix) usa.
beforeAll(() => {
  Element.prototype.hasPointerCapture = Element.prototype.hasPointerCapture ?? (() => false)
  Element.prototype.setPointerCapture = Element.prototype.setPointerCapture ?? (() => {})
  Element.prototype.releasePointerCapture = Element.prototype.releasePointerCapture ?? (() => {})
  Element.prototype.scrollIntoView = Element.prototype.scrollIntoView ?? (() => {})
})

const conta = { email_confirmado: true, senha_definida: true }
const gestor: User = { id: 1, username: 'gestor', role: 'Gestor', email: 'gestor@revalle.com.br', totp_ativo: true, ...conta }
const tecnico: User = { id: 2, username: 'tecnico', role: 'Técnico', email: null, totp_ativo: false, email_confirmado: false, senha_definida: true }
const outro: User = { id: 3, username: 'outro', role: 'Técnico', email: 'outro@revalle.com.br', totp_ativo: true, ...conta }
const gestorPendente: User = { id: 4, username: 'gestor.dois', role: 'Gestor', email: null, totp_ativo: false, email_confirmado: false, senha_definida: true }

function servirUsuarios(usuarios: User[]) {
  server.use(http.get('/api/users', () => HttpResponse.json(usuarios)))
}

function renderPagina() {
  return renderWithClient(
    <MemoryRouter>
      <UsersPage />
      <ToastContainer />
    </MemoryRouter>
  )
}

async function linhaDe(username: string) {
  const celula = await screen.findByText(username, { selector: 'td p, td span, td' })
  return celula.closest('tr') as HTMLElement
}

function muitos(n: number): User[] {
  return Array.from({ length: n }, (_, i) => {
    const num = String(i + 1).padStart(2, '0')
    return { id: 100 + i, username: `usuario${num}`, role: 'Técnico', email: `usuario${num}@revalle.com.br`, totp_ativo: false, ...conta }
  })
}

describe('UsersPage', () => {
  it('lista em tabela com e-mail e as pílulas de 2FA (Ativo, Pendente, Inativo)', async () => {
    servirUsuarios([gestor, gestorPendente, tecnico])
    renderPagina()

    const linhaGestor = await linhaDe('gestor')
    expect(within(linhaGestor).getByText('gestor@revalle.com.br')).toBeInTheDocument()
    expect(within(linhaGestor).getByText('Ativo')).toBeInTheDocument()

    expect(within(await linhaDe('gestor.dois')).getByText('Pendente')).toBeInTheDocument()

    const linhaTecnico = await linhaDe('tecnico')
    expect(within(linhaTecnico).getByText('Inativo')).toBeInTheDocument()
    expect(within(linhaTecnico).getByText('—')).toBeInTheDocument()
  })

  it('30 usuários: mostra 25 por página e a próxima traz o 26º', async () => {
    servirUsuarios(muitos(30))
    const user = userEvent.setup()
    renderPagina()

    expect(await screen.findByText('Mostrando 1–25 de 30')).toBeInTheDocument()
    expect(screen.queryByText('usuario26')).not.toBeInTheDocument()

    await user.click(screen.getByRole('button', { name: 'Próxima' }))

    expect(await screen.findByText('usuario26')).toBeInTheDocument()
    expect(screen.getByText('Mostrando 26–30 de 30')).toBeInTheDocument()
  })

  it('busca por e-mail filtra e volta à página 1', async () => {
    servirUsuarios(muitos(30))
    const user = userEvent.setup()
    renderPagina()

    await screen.findByText('Mostrando 1–25 de 30')
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await screen.findByText('Mostrando 26–30 de 30')

    await user.type(screen.getByLabelText('Buscar por usuário ou e-mail'), 'usuario07@')

    expect(await screen.findByText('Mostrando 1–1 de 1')).toBeInTheDocument()
    expect(screen.getByText('usuario07')).toBeInTheDocument()
    expect(screen.queryByText('usuario08')).not.toBeInTheDocument()
  })

  it('cria conta só com e-mail e papel', async () => {
    servirUsuarios([gestor])
    const corpos: unknown[] = []
    server.use(
      http.post('/api/users', async ({ request }) => {
        corpos.push(await request.json())
        return HttpResponse.json({
          detail: "Usuário 'novo' cadastrado com sucesso. Enviamos o convite para n***@revalle.com.br.",
          convite_enviado: true,
        })
      })
    )
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('gestor@revalle.com.br')

    expect(screen.queryByLabelText('Senha *')).not.toBeInTheDocument()
    expect(screen.queryByLabelText('Nome de Usuário *')).not.toBeInTheDocument()
    expect(screen.getByText('@revalle.com.br', { exact: false, selector: 'p' })).toBeInTheDocument()

    await user.type(screen.getByLabelText('E-mail *'), ' novo@revalle.com.br ')
    await user.click(screen.getByRole('combobox', { name: 'Papel *' }))
    await user.click(await screen.findByRole('option', { name: 'Técnico' }))
    await user.click(screen.getByRole('button', { name: 'Criar e enviar convite' }))

    await vi.waitFor(() => expect(corpos).toEqual([{ email: 'novo@revalle.com.br', role: 'Técnico' }]))
    expect(
      await screen.findByText("Usuário 'novo' cadastrado com sucesso. Enviamos o convite para n***@revalle.com.br.")
    ).toBeInTheDocument()
    await vi.waitFor(() => expect(screen.getByLabelText('E-mail *')).toHaveValue(''))
  })

  it('conta criada sem envio do convite avisa o Gestor', async () => {
    servirUsuarios([gestor])
    const detail = 'Conta criada, mas não foi possível enviar o e-mail. Use “Reenviar convite”.'
    server.use(http.post('/api/users', () => HttpResponse.json({ detail, convite_enviado: false })))
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('gestor@revalle.com.br')

    await user.type(screen.getByLabelText('E-mail *'), 'novo@revalle.com.br')
    await user.click(screen.getByRole('combobox', { name: 'Papel *' }))
    await user.click(await screen.findByRole('option', { name: 'Técnico' }))
    await user.click(screen.getByRole('button', { name: 'Criar e enviar convite' }))

    expect(await screen.findByText(detail)).toBeInTheDocument()
  })

  it('mostra Convite pendente, Aguardando confirmação e Confirmado', async () => {
    const convidado: User = {
      id: 5, username: 'novo', role: 'Técnico', email: 'novo@revalle.com.br', totp_ativo: false,
      email_confirmado: false, senha_definida: false,
    }
    const naoConfirmado: User = { ...outro, id: 6, username: 'antigo', email: 'antigo@revalle.com.br', email_confirmado: false }
    servirUsuarios([gestor, convidado, naoConfirmado])
    renderPagina()

    expect(within(await linhaDe('novo')).getByText('Convite pendente')).toBeInTheDocument()
    expect(within(await linhaDe('antigo')).getByText('Aguardando confirmação')).toBeInTheDocument()
    expect(within(await linhaDe('gestor')).getByText('Confirmado')).toBeInTheDocument()
    expect(screen.getByRole('columnheader', { name: 'Usuário (histórico)' })).toBeInTheDocument()
  })

  it('reenviar convite e enviar link chamam enviar-link-senha; não há troca de senha pelo Gestor', async () => {
    const convidado: User = {
      id: 5, username: 'novo', role: 'Técnico', email: 'novo@revalle.com.br', totp_ativo: false,
      email_confirmado: false, senha_definida: false,
    }
    servirUsuarios([gestor, convidado, outro])
    const chamados: string[] = []
    server.use(
      http.post('/api/users/:id/enviar-link-senha', ({ params }) => {
        chamados.push(String(params.id))
        return HttpResponse.json({ detail: `Enviamos o link (${params.id}).` })
      })
    )
    const user = userEvent.setup()
    renderPagina()

    const linhaConvidado = await linhaDe('novo')
    expect(within(linhaConvidado).queryByRole('button', { name: 'Enviar link para criar nova senha' })).not.toBeInTheDocument()
    await user.click(within(linhaConvidado).getByRole('button', { name: 'Reenviar convite' }))
    expect(await screen.findByText('Enviamos o link (5).')).toBeInTheDocument()

    const linhaOutro = await linhaDe('outro')
    expect(within(linhaOutro).queryByRole('button', { name: 'Reenviar convite' })).not.toBeInTheDocument()
    await user.click(within(linhaOutro).getByRole('button', { name: 'Enviar link para criar nova senha' }))
    expect(await screen.findByText('Enviamos o link (3).')).toBeInTheDocument()

    expect(chamados).toEqual(['5', '3'])
    expect(screen.queryByRole('button', { name: 'Alterar senha' })).not.toBeInTheDocument()
  })

  it('editar e-mail envia o PUT; vazio envia null; duplicado mostra a mensagem do servidor', async () => {
    servirUsuarios([gestor, tecnico])
    const corpos: unknown[] = []
    let duplicado = false
    server.use(
      http.put('/api/users/2', async ({ request }) => {
        corpos.push(await request.json())
        if (duplicado) {
          return HttpResponse.json({ detail: 'Este e-mail já está em uso por outro usuário.' }, { status: 400 })
        }
        return HttpResponse.json({ ...tecnico, email: 'tec@revalle.com.br' })
      })
    )
    const user = userEvent.setup()
    renderPagina()

    async function abrirEdicao() {
      const linha = await linhaDe('tecnico')
      await user.click(within(linha).getByRole('button', { name: 'Editar e-mail' }))
      const dialogo = await screen.findByRole('dialog')
      expect(within(dialogo).getByText('Editar e-mail — tecnico')).toBeInTheDocument()
      return dialogo
    }

    let dialogo = await abrirEdicao()
    await user.type(within(dialogo).getByLabelText('E-mail'), 'tec@revalle.com.br')
    await user.click(within(dialogo).getByRole('button', { name: 'Salvar' }))
    expect(await screen.findByText('E-mail atualizado.')).toBeInTheDocument()

    duplicado = true
    dialogo = await abrirEdicao()
    await user.clear(within(dialogo).getByLabelText('E-mail'))
    await user.type(within(dialogo).getByLabelText('E-mail'), 'gestor@revalle.com.br')
    await user.click(within(dialogo).getByRole('button', { name: 'Salvar' }))
    expect(
      await within(dialogo).findByText('Este e-mail já está em uso por outro usuário.')
    ).toBeInTheDocument()

    duplicado = false
    await user.clear(within(dialogo).getByLabelText('E-mail'))
    await user.click(within(dialogo).getByRole('button', { name: 'Salvar' }))
    await vi.waitFor(() => expect(corpos).toHaveLength(3))

    expect(corpos).toEqual([
      { email: 'tec@revalle.com.br' },
      { email: 'gestor@revalle.com.br' },
      { email: null },
    ])
  })

  it('"Zerar 2FA" só aparece para outro usuário com 2FA ativo; confirmar chama o DELETE', async () => {
    servirUsuarios([gestor, tecnico, outro])
    let chamado = false
    server.use(
      http.delete('/api/users/3/2fa', () => {
        chamado = true
        return HttpResponse.json({ detail: 'Verificação em duas etapas zerada.' })
      })
    )
    const user = userEvent.setup()
    renderPagina()

    const linhaGestor = await linhaDe('gestor')
    expect(within(linhaGestor).queryByRole('button', { name: 'Zerar 2FA' })).not.toBeInTheDocument()
    expect(within(await linhaDe('tecnico')).queryByRole('button', { name: 'Zerar 2FA' })).not.toBeInTheDocument()

    await user.click(within(await linhaDe('outro')).getByRole('button', { name: 'Zerar 2FA' }))
    const dialogo = await screen.findByRole('alertdialog')
    expect(
      within(dialogo).getByText('Zerar a verificação em duas etapas de "outro"?')
    ).toBeInTheDocument()
    await user.click(within(dialogo).getByRole('button', { name: 'Zerar 2FA' }))

    expect(await screen.findByText('Verificação em duas etapas zerada.')).toBeInTheDocument()
    expect(chamado).toBe(true)
  })
})
