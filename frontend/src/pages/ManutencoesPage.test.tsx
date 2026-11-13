import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import type { Manutencao } from '@/api/manutencoes'
import ManutencoesPage from '@/pages/ManutencoesPage'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'gestor', role: 'Gestor' },
    hasRole: () => true,
  }),
}))

function m(parcial: Partial<Manutencao> = {}): Manutencao {
  return {
    id: 1,
    item_id: 7,
    status: 'retornada',
    defeito: 'Tela quebrada',
    aberta_por: 'gestor',
    aberta_em: '2026-01-10T09:00:00',
    custo: '150.00',
    custo_acumulado_item: '239.90',
    item_tipo: 'Notebook',
    item_brand: 'Dell',
    item_model: 'Latitude',
    item_identificador: 'NB-0007',
    item_revenda: 'Revalle Juazeiro',
    item_status: 'Disponível',
    ...parcial,
  }
}

function mockServidor(opcoes: { total?: number; custoTotal?: string; linhas?: Manutencao[] } = {}) {
  const chamadas: URLSearchParams[] = []
  server.use(
    http.get('/api/constants', () =>
      HttpResponse.json({
        center_costs: [],
        revendas: ['Revalle Juazeiro', 'Revalle Petrolina'],
        setores: [],
        equipment_types: ['Notebook', 'Impressora'],
        peripheral_types: [],
        removal_reasons: [],
        removal_reasons_attachment: {},
      })
    ),
    http.get('/api/manutencoes', ({ request }) => {
      chamadas.push(new URL(request.url).searchParams)
      const linhas = opcoes.linhas ?? [m()]
      return HttpResponse.json({
        items: linhas,
        total: opcoes.total ?? linhas.length,
        custo_total: opcoes.custoTotal ?? '1239.90',
      })
    })
  )
  return chamadas
}

function renderPage(inicial = '/manutencoes') {
  return renderWithClient(
    <MemoryRouter initialEntries={[inicial]}>
      <ManutencoesPage />
      <ToastContainer />
    </MemoryRouter>
  )
}

describe('ManutencoesPage', () => {
  it('lista as manutenções e o custo total', async () => {
    mockServidor()
    renderPage()

    expect(await screen.findByText(/Custo total: R\$\s1\.239,90/)).toBeInTheDocument()
    expect(screen.getByText(/^1 manutenção ·/)).toBeInTheDocument()
    expect(screen.getByRole('link', { name: /Notebook Dell Latitude/ })).toHaveAttribute(
      'href',
      '/items/7?aba=manutencoes'
    )
    expect(within(screen.getByRole('table')).getByText('REVALLE JUAZEIRO')).toBeInTheDocument()
    expect(screen.getByText('Tela quebrada')).toBeInTheDocument()
    expect(within(screen.getByRole('table')).getByText('Retornada')).toBeInTheDocument()
  })

  it('primeira consulta usa 25 por página e sem filtros', async () => {
    const chamadas = mockServidor()
    renderPage()

    await screen.findByText('Tela quebrada')
    expect(Object.fromEntries(chamadas[0].entries())).toEqual({ limit: '25', offset: '0' })
  })

  it('filtro de status volta à página 1', async () => {
    const chamadas = mockServidor({ total: 60, linhas: [m()] })
    const user = userEvent.setup()
    renderPage()

    await screen.findByText('Tela quebrada')
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await waitFor(() => expect(chamadas.at(-1)?.get('offset')).toBe('25'))

    await user.selectOptions(screen.getByLabelText('Status'), 'Em andamento')

    await waitFor(() => {
      const ultima = chamadas.at(-1)!
      expect(ultima.get('status')).toBe('em_andamento')
      expect(ultima.get('offset')).toBe('0')
    })
  })

  it('filtros de unidade, tipo e período seguem para a consulta', async () => {
    const chamadas = mockServidor()
    const user = userEvent.setup()
    renderPage()

    await screen.findByText('Tela quebrada')
    await user.selectOptions(screen.getByLabelText('Unidade'), 'REVALLE PETROLINA')
    await user.selectOptions(screen.getByLabelText('Tipo'), 'Impressora')
    await user.type(screen.getByLabelText('Aberta de'), '2026-01-01')
    await user.type(screen.getByLabelText('até'), '2026-01-31')

    await waitFor(() => {
      const ultima = chamadas.at(-1)!
      expect(ultima.get('revenda')).toBe('Revalle Petrolina')
      expect(ultima.get('tipo')).toBe('Impressora')
      expect(ultima.get('de')).toBe('2026-01-01')
      expect(ultima.get('ate')).toBe('2026-01-31')
    })
  })

  it('a busca espera a pausa na digitação', async () => {
    const chamadas = mockServidor()
    const user = userEvent.setup()
    renderPage()

    await screen.findByText('Tela quebrada')
    await user.type(screen.getByLabelText('Buscar manutenções'), 'TecInfo')

    await waitFor(() => expect(chamadas.at(-1)?.get('busca')).toBe('TecInfo'))
    expect(chamadas.filter((c) => c.get('busca') !== null && c.get('busca') !== 'TecInfo')).toHaveLength(0)
  })

  it('?item_id=7 filtra por equipamento e mostra o chip', async () => {
    const chamadas = mockServidor()
    const user = userEvent.setup()
    renderPage('/manutencoes?item_id=7')

    expect(await screen.findByText('Equipamento #7')).toBeInTheDocument()
    expect(chamadas[0].get('item_id')).toBe('7')

    await user.click(screen.getByRole('button', { name: 'Limpar filtro de equipamento' }))
    await waitFor(() => expect(screen.queryByText('Equipamento #7')).not.toBeInTheDocument())
    await waitFor(() => expect(chamadas.at(-1)?.get('item_id')).toBeNull())
  })

  it('Nova manutenção abre o diálogo com busca de equipamento', async () => {
    mockServidor()
    const buscas: URLSearchParams[] = []
    server.use(
      http.get('/api/items', ({ request }) => {
        buscas.push(new URL(request.url).searchParams)
        return HttpResponse.json({
          items: [{ id: 7, tipo: 'Notebook', brand: 'Dell', model: 'Latitude', status: 'Disponível' }],
          total: 1,
        })
      })
    )
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: 'Nova manutenção' }))
    const dialogo = await screen.findByRole('dialog')
    expect(within(dialogo).getByRole('heading', { name: 'Abrir manutenção' })).toBeInTheDocument()
    await waitFor(() => expect(buscas.length).toBeGreaterThan(0))
    expect(buscas[0].get('status')).toBe('Disponível')
    expect(buscas[0].get('limit')).toBe('20')
  })

  it('sem manutenções mostra o estado vazio', async () => {
    mockServidor({ linhas: [], total: 0, custoTotal: '0.00' })
    renderPage()

    expect(await screen.findByText('Nenhuma manutenção encontrada')).toBeInTheDocument()
    expect(screen.getByText(/Custo total: R\$\s0,00/)).toBeInTheDocument()
  })
})
