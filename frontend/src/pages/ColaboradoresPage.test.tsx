import { beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import ColaboradoresPage from '@/pages/ColaboradoresPage'
import type { Colaborador } from '@/api/colaboradores'

let role = 'Gestor'
vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'teste', role },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes(role),
  }),
}))

const maria: Colaborador = {
  id: 7,
  cpf: '12345678909',
  nome: 'MARIA SILVA',
  cargo: 'Analista',
  setor: 'TI',
  unidade_nome: 'Juazeiro',
  unidade_id: 1,
  situacao: 'ativo',
  email: 'maria@revalle.com',
  telefone: '74999990000',
  email_manual: false,
  telefone_manual: false,
  superior_nome: 'JOAO CHEFE',
  origem: 'rh',
}

const jose: Colaborador = {
  ...maria,
  id: 8,
  cpf: '98765432100',
  nome: 'JOSE LIMA',
  situacao: 'afastado',
  email: null,
  telefone: null,
  superior_nome: null,
  origem: 'manual',
}

const syncOk = {
  ultima: {
    iniciado_em: '2026-11-10T10:00:00',
    concluido_em: '2026-11-10T10:01:00',
    status: 'ok',
    total_base: 100,
    criados: 3,
    atualizados: 5,
    desligados: 1,
    erro: null,
  },
  configurado: true,
}

function mockBase(opts: { items?: Colaborador[]; total?: number; sync?: unknown } = {}) {
  const pedidos: URLSearchParams[] = []
  const items = opts.items ?? [maria, jose]
  server.use(
    http.get('/api/colaboradores', ({ request }) => {
      pedidos.push(new URL(request.url).searchParams)
      return HttpResponse.json({ items, total: opts.total ?? items.length })
    }),
    http.get('/api/colaboradores/sincronizacao', () => HttpResponse.json(opts.sync ?? syncOk)),
    http.get('/api/colaboradores/cargos', () => HttpResponse.json(['Analista', 'Gerente'])),
    http.get('/api/unidades', () =>
      HttpResponse.json([{ id: 1, nome: 'Juazeiro', razao_social: 'X', cnpj: '1', is_active: true }])
    )
  )
  return pedidos
}

function renderPage() {
  return renderWithClient(
    <>
      <ColaboradoresPage />
      <ToastContainer />
    </>
  )
}

beforeEach(() => {
  role = 'Gestor'
  localStorage.clear()
})

describe('ColaboradoresPage', () => {
  it('lista colaboradores com CPF formatado, unidade em maiúsculas e situação', async () => {
    mockBase()
    renderPage()
    expect(await screen.findByText('MARIA SILVA')).toBeInTheDocument()
    expect(screen.getByText('123.456.789-09')).toBeInTheDocument()
    expect(screen.getAllByText('JUAZEIRO').length).toBeGreaterThan(0)
    const tabela = within(screen.getByRole('table'))
    expect(tabela.getByText('Ativo')).toBeInTheDocument()
    expect(tabela.getByText('Afastado')).toBeInTheDocument()
    expect(screen.getAllByText('(74) 99999-0000').length).toBeGreaterThan(0)
    expect(screen.getByText('JOAO CHEFE')).toBeInTheDocument()
  })

  it('pagina no servidor com offset e limit conforme a página', async () => {
    const pedidos = mockBase({ items: [maria], total: 60 })
    renderPage()
    expect(await screen.findByText('Mostrando 1–25 de 60')).toBeInTheDocument()
    expect(pedidos[0].get('limit')).toBe('25')
    expect(pedidos[0].get('offset')).toBe('0')

    await userEvent.setup().click(screen.getByRole('button', { name: 'Próxima' }))
    await waitFor(() => expect(pedidos.some((p) => p.get('offset') === '25')).toBe(true))
  })

  it('busca e filtro de situação voltam para a primeira página e vão para a API', async () => {
    const pedidos = mockBase({ items: [maria], total: 60 })
    renderPage()
    await screen.findByText('MARIA SILVA')
    const user = userEvent.setup()
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await waitFor(() => expect(pedidos.some((p) => p.get('offset') === '25')).toBe(true))

    await user.selectOptions(screen.getByLabelText('Situação'), 'desligado')
    await waitFor(() =>
      expect(pedidos.some((p) => p.get('situacao') === 'desligado' && p.get('offset') === '0')).toBe(true)
    )
    await user.type(screen.getByPlaceholderText(/buscar colaborador/i), 'mar')
    await waitFor(() =>
      expect(pedidos.some((p) => p.get('busca') === 'mar' && p.get('offset') === '0')).toBe(true)
    )
  })

  it('Gestor vê "Sincronizar agora" e o resumo vai para o toast', async () => {
    mockBase()
    server.use(
      http.post('/api/colaboradores/sincronizar', () =>
        HttpResponse.json({ total_base: 100, criados: 3, atualizados: 5, desligados: 1 })
      )
    )
    renderPage()
    expect(await screen.findByText(/última sincronização/i)).toBeInTheDocument()
    await userEvent.setup().click(await screen.findByRole('button', { name: /sincronizar agora/i }))
    expect(await screen.findByText(/Sincronização concluída: 100 na base, 3 criados, 5 atualizados, 1 desligado/)).toBeInTheDocument()
  })

  it('mostra a mensagem da API quando a base de RH está indisponível', async () => {
    mockBase()
    server.use(
      http.post('/api/colaboradores/sincronizar', () =>
        HttpResponse.json({ detail: 'Base de colaboradores indisponível no momento.' }, { status: 503 })
      )
    )
    renderPage()
    await userEvent.setup().click(await screen.findByRole('button', { name: /sincronizar agora/i }))
    expect(await screen.findByText('Base de colaboradores indisponível no momento.')).toBeInTheDocument()
  })

  it('mostra o erro da última sincronização', async () => {
    mockBase({ sync: { ...syncOk, ultima: { ...syncOk.ultima, status: 'erro', erro: 'timeout na base' } } })
    renderPage()
    expect(await screen.findByText(/timeout na base/)).toBeInTheDocument()
  })

  it('Técnico não vê "Sincronizar agora" nem o status da sincronização', async () => {
    role = 'Técnico'
    mockBase()
    renderPage()
    await screen.findByText('MARIA SILVA')
    expect(screen.queryByRole('button', { name: /sincronizar agora/i })).not.toBeInTheDocument()
    expect(screen.queryByText(/última sincronização/i)).not.toBeInTheDocument()
  })

  it('editar com e-mail inválido mostra a mensagem da API', async () => {
    mockBase()
    server.use(
      http.put('/api/colaboradores/7', () =>
        HttpResponse.json({ detail: 'E-mail: E-mail inválido.' }, { status: 422 })
      )
    )
    renderPage()
    const user = userEvent.setup()
    await user.click(await screen.findByRole('button', { name: /editar maria silva/i }))

    expect(screen.getAllByText('Vem da base de RH').length).toBeGreaterThan(0)
    const email = screen.getByLabelText('E-mail')
    await user.clear(email)
    await user.type(email, 'invalido')
    await user.click(screen.getByRole('button', { name: /salvar/i }))

    expect(await screen.findByText('E-mail: E-mail inválido.')).toBeInTheDocument()
  })

  it('editar colaborador da base só envia e-mail, telefone e superior', async () => {
    mockBase()
    let corpo: Record<string, unknown> | null = null
    server.use(
      http.put('/api/colaboradores/7', async ({ request }) => {
        corpo = (await request.json()) as Record<string, unknown>
        return HttpResponse.json({ ...maria, email: 'novo@revalle.com' })
      })
    )
    renderPage()
    const user = userEvent.setup()
    await user.click(await screen.findByRole('button', { name: /editar maria silva/i }))
    const email = screen.getByLabelText('E-mail')
    await user.clear(email)
    await user.type(email, 'novo@revalle.com')
    await user.click(screen.getByRole('button', { name: /salvar/i }))

    await waitFor(() => expect(corpo).not.toBeNull())
    expect(Object.keys(corpo!).sort()).toEqual(['email', 'superior_id', 'telefone'])
    expect(corpo!.email).toBe('novo@revalle.com')
  })

  it('cadastra colaborador manual e mostra erro de CPF duplicado (409)', async () => {
    mockBase()
    server.use(
      http.post('/api/colaboradores', () =>
        HttpResponse.json({ detail: 'Já existe colaborador com este CPF.' }, { status: 409 })
      )
    )
    renderPage()
    const user = userEvent.setup()
    await user.click(await screen.findByRole('button', { name: /novo colaborador/i }))
    await user.type(screen.getByLabelText(/CPF/), '52998224725')
    await user.type(screen.getByLabelText(/^Nome/), 'Fulano de Tal')
    await user.click(screen.getByRole('button', { name: /salvar/i }))
    expect(await screen.findByText('Já existe colaborador com este CPF.')).toBeInTheDocument()
  })

  it('recusa CPF inválido no cadastro sem chamar a API', async () => {
    mockBase()
    renderPage()
    const user = userEvent.setup()
    await user.click(await screen.findByRole('button', { name: /novo colaborador/i }))
    await user.type(screen.getByLabelText(/CPF/), '11111111111')
    await user.type(screen.getByLabelText(/^Nome/), 'Fulano')
    await user.click(screen.getByRole('button', { name: /salvar/i }))
    expect(await screen.findByText(/CPF inválido/)).toBeInTheDocument()
  })
})
