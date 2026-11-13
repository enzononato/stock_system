import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import ColaboradorDetailPage from '@/pages/ColaboradorDetailPage'
import tab from '@/features/colaborador-detail/tabs/linhas'
import type { Colaborador } from '@/api/colaboradores'
import type { Linha } from '@/api/linhas'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'tecnico.teste', role: 'Técnico' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Técnico'),
  }),
}))

const colaborador: Colaborador = {
  id: 3,
  cpf: '12345678909',
  nome: 'Maria Souza',
  situacao: 'ativo',
  email_manual: false,
  telefone_manual: false,
  origem: 'rh',
}

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
  plano: null,
  item_id: null,
  item_tipo: null,
  item_brand: null,
  item_model: null,
  item_identificador: null,
  item_ativo: null,
}

function mockLinhas(items: Linha[], total = items.length) {
  const pedidos: URLSearchParams[] = []
  server.use(
    http.get('/api/linhas', ({ request }) => {
      pedidos.push(new URL(request.url).searchParams)
      return HttpResponse.json({ items, total })
    })
  )
  return pedidos
}

function renderAba(c: Colaborador = colaborador) {
  const Component = tab.Component
  return renderWithClient(
    <MemoryRouter initialEntries={['/']}>
      <Routes>
        <Route path="/" element={<Component colaborador={c} />} />
      </Routes>
    </MemoryRouter>
  )
}

describe('aba Linhas do colaborador', () => {
  it('é a aba "linhas" de ordem 40', () => {
    expect(tab.id).toBe('linhas')
    expect(tab.label).toBe('Linhas')
    expect(tab.order).toBe(40)
  })

  it('lista as linhas do colaborador, inclusive canceladas, e filtra por colaborador_id', async () => {
    const pedidos = mockLinhas([ativa, cancelada])
    renderAba()

    expect(await screen.findByText('(74) 99999-0001')).toBeInTheDocument()
    expect(screen.getByText('(74) 99999-0002')).toBeInTheDocument()
    expect(screen.getByText('Ativa')).toBeInTheDocument()
    expect(screen.getByText('Cancelada')).toBeInTheDocument()
    expect(screen.getByText('Vivo · Controle 20GB')).toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'Celular Samsung A15 · CEL-A' })).toHaveAttribute('href', '/items/7')
    expect(pedidos[0].get('colaborador_id')).toBe('3')
    expect(pedidos[0].get('limit')).toBe('25')
    expect(pedidos[0].get('offset')).toBe('0')
  })

  it('o botão de histórico abre o diálogo e consulta os eventos da linha', async () => {
    mockLinhas([ativa])
    let chamadas = 0
    server.use(
      http.get('/api/linhas/5/eventos', () => {
        chamadas += 1
        return HttpResponse.json([])
      })
    )
    renderAba()

    await userEvent.setup().click(await screen.findByRole('button', { name: 'Histórico (74) 99999-0001' }))

    expect(await screen.findByText('Histórico da linha (74) 99999-0001')).toBeInTheDocument()
    await waitFor(() => expect(chamadas).toBe(1))
  })

  it('sem linhas mostra a mensagem de vazio', async () => {
    mockLinhas([], 0)
    renderAba()

    expect(await screen.findByText('Nenhuma linha com este colaborador.')).toBeInTheDocument()
  })

  it('trocar o tamanho da página envia limit=50 e offset=0', async () => {
    localStorage.clear()
    const pedidos = mockLinhas([ativa], 80)
    renderAba()
    await screen.findByText('(74) 99999-0001')
    const user = userEvent.setup()
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await waitFor(() => expect(pedidos.some((p) => p.get('offset') === '25')).toBe(true))

    await user.selectOptions(screen.getByLabelText('Itens por página'), '50')

    await waitFor(() => expect(pedidos.some((p) => p.get('limit') === '50' && p.get('offset') === '0')).toBe(true))
    expect(within(screen.getByRole('list')).getAllByRole('listitem')).toHaveLength(1)
  })

  it('aparelho baixado aparece como texto com (baixado), sem link', async () => {
    mockLinhas([{ ...ativa, item_ativo: false }])
    renderAba()

    expect(await screen.findByText(/Celular Samsung A15 · CEL-A \(baixado\)/)).toBeInTheDocument()
    expect(screen.queryByRole('link', { name: /Celular Samsung A15/ })).not.toBeInTheDocument()
  })
})

describe('ColaboradorDetailPage com a aba Linhas', () => {
  it('mostra a aba Linhas depois de Equipamentos e lista as linhas do colaborador', async () => {
    const pedidos = mockLinhas([ativa, cancelada])
    server.use(http.get('/api/colaboradores/3', () => HttpResponse.json(colaborador)))
    const { unmount } = renderWithClient(
      <MemoryRouter initialEntries={['/colaboradores/3?aba=linhas']}>
        <Routes>
          <Route path="/colaboradores/:id" element={<ColaboradorDetailPage />} />
        </Routes>
      </MemoryRouter>
    )

    const abas = within(await screen.findByRole('tablist')).getAllByRole('tab').map((t) => t.textContent)
    expect(abas).toContain('Linhas')
    expect(abas.indexOf('Equipamentos')).toBeLessThan(abas.indexOf('Linhas'))
    expect(screen.getByRole('tab', { name: 'Linhas' })).toHaveAttribute('aria-selected', 'true')
    expect(await screen.findByText('(74) 99999-0001')).toBeInTheDocument()
    expect(screen.getByText('(74) 99999-0002')).toBeInTheDocument()
    expect(pedidos[0].get('colaborador_id')).toBe('3')
    unmount()
  })
})
