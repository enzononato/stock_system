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

const gestor: User = { id: 1, username: 'gestor', role: 'Gestor', email: 'gestor@revalle.com.br', totp_ativo: true }
const tecnico: User = { id: 2, username: 'tecnico', role: 'Técnico', email: null, totp_ativo: false }
const outro: User = { id: 3, username: 'outro', role: 'Técnico', email: 'outro@revalle.com.br', totp_ativo: true }
const gestorPendente: User = { id: 4, username: 'gestor.dois', role: 'Gestor', email: null, totp_ativo: false }

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
    return { id: 100 + i, username: `usuario${num}`, role: 'Técnico', email: `usuario${num}@revalle.com.br`, totp_ativo: false }
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

  it('criar usuário envia o e-mail informado, ou null quando vazio', async () => {
    servirUsuarios([gestor])
    const corpos: unknown[] = []
    server.use(
      http.post('/api/users', async ({ request }) => {
        corpos.push(await request.json())
        return HttpResponse.json({ detail: 'ok' })
      })
    )
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('gestor@revalle.com.br')

    async function criar(username: string, email: string) {
      await user.type(screen.getByLabelText('Nome de Usuário *'), username)
      await user.type(screen.getByLabelText('Senha *'), 'Abc#12345')
      await user.click(screen.getByRole('combobox', { name: 'Função *' }))
      await user.click(await screen.findByRole('option', { name: 'Técnico' }))
      if (email) await user.type(screen.getByLabelText('E-mail'), email)
      const esperados = corpos.length + 1
      await user.click(screen.getByRole('button', { name: 'Criar Usuário' }))
      await vi.waitFor(() => expect(corpos.length).toBe(esperados))
      await vi.waitFor(() => expect(screen.getByLabelText('Nome de Usuário *')).toHaveValue(''))
    }

    await criar('novo.um', 'novo@revalle.com.br')
    await criar('novo.dois', '')

    expect(corpos).toEqual([
      { username: 'novo.um', password: 'Abc#12345', role: 'Técnico', email: 'novo@revalle.com.br' },
      { username: 'novo.dois', password: 'Abc#12345', role: 'Técnico', email: null },
    ])
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
