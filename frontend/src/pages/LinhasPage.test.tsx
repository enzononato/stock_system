import { beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import LinhasPage from '@/pages/LinhasPage'
import { getVisibleNavItems } from '@/components/layout/navigation'
import { formatMoeda } from '@/lib/format'
import type { Linha } from '@/api/linhas'

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

const ativa: Linha = {
  id: 5,
  numero: '74999990001',
  operadora: 'Vivo',
  plano: 'Controle 20GB',
  custo_mensal: '49.90',
  status: 'ativa',
  unidade_id: 1,
  unidade_nome: 'Unidade A',
  item_id: 7,
  item_tipo: 'Celular',
  item_brand: 'Samsung',
  item_model: 'A15',
  item_identificador: 'CEL-A',
  item_ativo: true,
  colaborador_id: 3,
  colaborador_nome: 'Maria Souza',
  colaborador_situacao: 'ativo',
  iccid: null,
  observacoes: null,
  criada_em: '2026-11-12T10:00:00',
  atualizada_em: '2026-11-12T10:00:00',
}

const cancelada: Linha = {
  ...ativa,
  id: 6,
  numero: '74999990002',
  status: 'cancelada',
  item_id: null,
  item_tipo: null,
  item_brand: null,
  item_model: null,
  item_identificador: null,
  item_ativo: null,
}

function mockBase(opts: { items?: Linha[]; total?: number | ((offset: number) => number) } = {}) {
  const pedidos: URLSearchParams[] = []
  const items = opts.items ?? [ativa]
  server.use(
    http.get('/api/linhas', ({ request }) => {
      const params = new URL(request.url).searchParams
      pedidos.push(params)
      const t = typeof opts.total === 'function' ? opts.total(Number(params.get('offset') ?? 0)) : opts.total
      return HttpResponse.json({ items, total: t ?? items.length })
    }),
    http.get('/api/unidades', () =>
      HttpResponse.json([
        { id: 1, nome: 'Unidade A', razao_social: 'X', cnpj: '1', is_active: true },
        { id: 2, nome: 'Unidade B', razao_social: 'X', cnpj: '2', is_active: true },
      ])
    )
  )
  return pedidos
}

function renderPage(path = '/linhas') {
  return renderWithClient(
    <MemoryRouter initialEntries={[path]}>
      <LinhasPage />
    </MemoryRouter>
  )
}

beforeEach(() => {
  role = 'Gestor'
  localStorage.clear()
})

describe('LinhasPage', () => {
  it('lista número formatado, custo, status, unidade em maiúsculas e link do aparelho', async () => {
    mockBase()
    renderPage()

    const tabela = within(await screen.findByRole('table'))
    expect(await tabela.findByText('(74) 99999-0001')).toBeInTheDocument()
    const custo = formatMoeda('49.90').replace(/\s/g, ' ')
    expect(tabela.getByText(custo)).toBeInTheDocument()
    expect(tabela.getByText('Ativa')).toBeInTheDocument()
    expect(tabela.getByText('UNIDADE A')).toBeInTheDocument()
    expect(tabela.getByText('Controle 20GB')).toBeInTheDocument()
    expect(tabela.getByRole('link', { name: 'Celular Samsung A15 · CEL-A' })).toHaveAttribute('href', '/items/7')
    expect(tabela.getByRole('link', { name: 'MARIA SOUZA' })).toHaveAttribute('href', '/colaboradores/3')
  })

  it('aparelho baixado, colaborador desligado e linha sem vínculos continuam legíveis', async () => {
    mockBase({
      items: [
        { ...ativa, item_ativo: false, colaborador_situacao: 'desligado' },
        { ...cancelada, plano: null, colaborador_id: null, colaborador_nome: null, colaborador_situacao: null },
      ],
    })
    renderPage()

    const tabela = within(await screen.findByRole('table'))
    expect(await tabela.findByText(/\(baixado\)/)).toBeInTheDocument()
    // aparelho baixado não tem página de detalhe: aparece como texto, sem link
    expect(tabela.queryByRole('link', { name: /Celular Samsung A15/ })).not.toBeInTheDocument()
    expect(tabela.getByText(/Celular Samsung A15 · CEL-A \(baixado\)/)).toBeInTheDocument()
    expect(tabela.getByText(/\(desligado\)/)).toBeInTheDocument()
    // plano, aparelho e colaborador vazios viram traço
    expect(tabela.getAllByText('—').length).toBeGreaterThanOrEqual(3)
  })

  it('filtros de status, unidade e operadora vão para a API e voltam à página 1', async () => {
    const pedidos = mockBase({ total: 120 })
    renderPage()
    await screen.findByText('(74) 99999-0001')
    const user = userEvent.setup()

    await user.selectOptions(screen.getByLabelText('Status'), 'suspensa')
    await waitFor(() =>
      expect(pedidos.some((p) => p.get('status') === 'suspensa' && p.get('offset') === '0')).toBe(true)
    )
    await user.selectOptions(screen.getByLabelText('Unidade'), '2')
    await waitFor(() => expect(pedidos.some((p) => p.get('unidade_id') === '2')).toBe(true))

    // estando na página 3, trocar a operadora volta à página 1 na mesma requisição
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await waitFor(() => expect(pedidos.some((p) => p.get('offset') === '50')).toBe(true))
    const antes = pedidos.length
    await user.selectOptions(screen.getByLabelText('Operadora'), 'Claro')
    await waitFor(() =>
      expect(pedidos.some((p) => p.get('operadora') === 'Claro' && p.get('offset') === '0')).toBe(true)
    )
    expect(pedidos.slice(antes).some((p) => p.get('offset') === '50')).toBe(false)
  })

  it('trocar o tamanho da página envia limit=50 e offset=0', async () => {
    const pedidos = mockBase({ total: 120 })
    renderPage()
    await screen.findByText('(74) 99999-0001')
    const user = userEvent.setup()
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await waitFor(() => expect(pedidos.some((p) => p.get('offset') === '25')).toBe(true))

    await user.selectOptions(screen.getByLabelText('Itens por página'), '50')

    await waitFor(() => expect(pedidos.some((p) => p.get('limit') === '50' && p.get('offset') === '0')).toBe(true))
    expect(pedidos.some((p) => p.get('limit') === '50' && p.get('offset') !== '0')).toBe(false)
  })

  it('total que encolhe leva de volta à última página válida', async () => {
    // Na página 3 (offset 50) o servidor já responde com total 10: só existe a página 1.
    const pedidos = mockBase({ total: (offset) => (offset === 50 ? 10 : 120) })
    renderPage()
    await screen.findByText('(74) 99999-0001')
    const user = userEvent.setup()

    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await waitFor(() => expect(pedidos.some((p) => p.get('offset') === '50')).toBe(true))

    const indice50 = pedidos.findIndex((p) => p.get('offset') === '50')
    await waitFor(() => expect(pedidos.slice(indice50 + 1).some((p) => p.get('offset') === '0')).toBe(true))
  })

  it('a busca da URL vira o primeiro filtro enviado', async () => {
    const pedidos = mockBase()
    renderPage('/linhas?busca=99999')

    await screen.findByText('(74) 99999-0001')

    expect(pedidos[0].get('busca')).toBe('99999')
    expect(screen.getByLabelText('Buscar linha')).toHaveValue('99999')
  })

  it('Jovem Aprendiz só consulta: sem ações de escrita, com histórico e colaborador sem link', async () => {
    role = 'Jovem Aprendiz'
    mockBase()
    renderPage()

    await screen.findByText('(74) 99999-0001')

    expect(screen.queryByRole('button', { name: 'Nova linha' })).not.toBeInTheDocument()
    expect(screen.queryByRole('link', { name: 'Relatório de custo' })).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: /^Editar/ })).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: /^Suspender/ })).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: /^Cancelar/ })).not.toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Histórico (74) 99999-0001' })).toBeInTheDocument()
    expect(screen.getByText('MARIA SOUZA')).toBeInTheDocument()
    expect(screen.queryByRole('link', { name: 'MARIA SOUZA' })).not.toBeInTheDocument()
  })

  it('Gestor vê as ações por status e o link do relatório', async () => {
    mockBase({ items: [ativa, cancelada] })
    renderPage()
    await screen.findByText('(74) 99999-0001')

    expect(screen.getByRole('link', { name: 'Relatório de custo' })).toHaveAttribute('href', '/linhas/relatorio')
    expect(screen.getByRole('button', { name: 'Nova linha' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Editar (74) 99999-0001' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Suspender (74) 99999-0001' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Cancelar (74) 99999-0001' })).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Reativar (74) 99999-0001' })).not.toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Reativar (74) 99999-0002' })).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Editar (74) 99999-0002' })).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Suspender (74) 99999-0002' })).not.toBeInTheDocument()
  })

  it('Histórico abre o diálogo e consulta os eventos da linha', async () => {
    mockBase()
    let chamadas = 0
    server.use(
      http.get('/api/linhas/5/eventos', () => {
        chamadas += 1
        return HttpResponse.json([])
      })
    )
    renderPage()
    await screen.findByText('(74) 99999-0001')

    await userEvent.setup().click(screen.getByRole('button', { name: 'Histórico (74) 99999-0001' }))

    expect(await screen.findByText('Histórico da linha (74) 99999-0001')).toBeInTheDocument()
    await waitFor(() => expect(chamadas).toBe(1))
  })

  it('lista vazia mostra o estado vazio', async () => {
    mockBase({ items: [], total: 0 })
    renderPage()

    expect(await screen.findByText('Nenhuma linha encontrada')).toBeInTheDocument()
  })
})

describe('navegação', () => {
  it('o item Linhas aparece para todos os papéis', () => {
    expect(getVisibleNavItems('Jovem Aprendiz').map((i) => i.to)).toContain('/linhas')
    expect(getVisibleNavItems('Gestor').map((i) => i.to)).toContain('/linhas')
  })
})
