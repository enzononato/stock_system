import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within, fireEvent } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import StockPage from '@/pages/StockPage'

const papel = vi.hoisted(() => ({ atual: 'Gestor' }))

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'gestor.teste', role: papel.atual },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes(papel.atual),
  }),
}))

// jsdom não implementa a Pointer Events API que o Select (Radix) usa.
beforeAll(() => {
  Element.prototype.hasPointerCapture = Element.prototype.hasPointerCapture ?? (() => false)
  Element.prototype.setPointerCapture = Element.prototype.setPointerCapture ?? (() => {})
  Element.prototype.releasePointerCapture = Element.prototype.releasePointerCapture ?? (() => {})
  Element.prototype.scrollIntoView = Element.prototype.scrollIntoView ?? (() => {})
})

const maria = {
  id: 7,
  cpf: '12345678909',
  nome: 'MARIA SILVA',
  cargo: 'Vendedora',
  setor: 'Comercial',
  situacao: 'ativo',
  email_manual: false,
  telefone_manual: false,
  origem: 'rh',
}

const itemComColaborador = {
  id: 10,
  tipo: 'Notebook',
  brand: 'Dell',
  model: 'Latitude',
  status: 'Indisponível',
  revenda: 'Revalle Juazeiro',
  colaborador_nome: 'MARIA SILVA',
  colaborador_setor: 'Comercial',
  colaborador_cargo: 'Vendedora',
}
const itemSemColaborador = {
  id: 11,
  tipo: 'Desktop',
  brand: 'HP',
  model: 'Prodesk',
  status: 'Disponível',
  revenda: 'Revalle Juazeiro',
  colaborador_nome: null,
  assigned_to: null,
}

let pedidos: URLSearchParams[] = []

function mockServidor(opts: { total?: () => number; items?: unknown[] } = {}) {
  pedidos = []
  server.use(
    http.get('/api/items', ({ request }) => {
      const params = new URL(request.url).searchParams
      pedidos.push(params)
      if (opts.total) {
        const total = opts.total()
        const limit = Number(params.get('limit'))
        const offset = Number(params.get('offset'))
        const items = Array.from({ length: Math.max(0, Math.min(limit, total - offset)) }, (_, i) => ({
          id: 1000 + offset + i,
          tipo: 'Notebook',
          status: 'Disponível',
        }))
        return HttpResponse.json({ items, total })
      }
      const items = opts.items ?? [itemComColaborador, itemSemColaborador]
      return HttpResponse.json({ items, total: items.length })
    }),
    http.get('/api/unidades', () =>
      HttpResponse.json([{ id: 1, nome: 'Revalle Juazeiro', razao_social: 'X', cnpj: '1', is_active: true }])
    ),
    http.get('/api/colaboradores/:id', () => HttpResponse.json(maria)),
    http.get('/api/colaboradores', ({ request }) => {
      const params = new URL(request.url).searchParams
      colaboradoresPedidos.push(params)
      const busca = params.get('busca')
      return HttpResponse.json({ items: busca === 'mar' ? [maria] : [], total: busca === 'mar' ? 1 : 0 })
    })
  )
}
let colaboradoresPedidos: URLSearchParams[] = []

function renderPage() {
  return renderWithClient(
    <MemoryRouter>
      <StockPage />
    </MemoryRouter>
  )
}

const ultimo = () => pedidos[pedidos.length - 1]

describe('StockPage', () => {
  beforeEach(() => {
    localStorage.clear()
    colaboradoresPedidos = []
    papel.atual = 'Gestor'
  })

  it('mostra Setor e Cargo do colaborador, unidade em CAIXA ALTA e "—" sem colaborador', async () => {
    mockServidor()
    renderPage()
    const linha10 = (await screen.findByRole('link', { name: /#10/ })).closest('tr') as HTMLElement
    expect(within(linha10).getByText('Comercial')).toBeInTheDocument()
    expect(within(linha10).getByText('Vendedora')).toBeInTheDocument()
    expect(within(linha10).getByText('REVALLE JUAZEIRO')).toBeInTheDocument()
    expect(within(linha10).getByText('Indisponível')).toHaveClass('font-bold')

    const linha11 = screen.getByRole('link', { name: /#11/ }).closest('tr') as HTMLElement
    const celulas = within(linha11).getAllByRole('cell')
    // Patrimônio, Tipo, Equipamento, Status, Colaborador, Setor, Cargo, Unidade, Cadastro, ações
    expect(celulas[4]).toHaveTextContent('—')
    expect(celulas[5]).toHaveTextContent('—')
    expect(celulas[6]).toHaveTextContent('—')
  })

  it('pede 25 por página e lembra a troca para 50', async () => {
    mockServidor()
    renderPage()
    await screen.findByRole('link', { name: /#10/ })
    expect(pedidos[0].get('limit')).toBe('25')
    expect(pedidos[0].get('offset')).toBe('0')
    expect(pedidos[0].get('sort')).toBe('id')
    expect(pedidos[0].get('dir')).toBe('desc')

    const user = userEvent.setup()
    await user.selectOptions(screen.getByLabelText('Itens por página'), '50')
    await waitFor(() => expect(ultimo().get('limit')).toBe('50'))
    expect(ultimo().get('offset')).toBe('0')
    expect(localStorage.getItem('revalle_page_size:estoque')).toBe('50')
  })

  it('ordenação por coluna vai para o servidor e volta à página 1', async () => {
    mockServidor({ total: () => 60 })
    renderPage()
    const user = userEvent.setup()
    await screen.findByText('Mostrando 1–25 de 60')
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await waitFor(() => expect(ultimo().get('offset')).toBe('25'))

    await user.click(screen.getByRole('button', { name: 'Tipo' }))
    await waitFor(() => expect(ultimo().get('sort')).toBe('tipo'))
    expect(ultimo().get('dir')).toBe('asc')
    expect(ultimo().get('offset')).toBe('0')

    await user.click(screen.getByRole('button', { name: 'Tipo' }))
    await waitFor(() => expect(ultimo().get('dir')).toBe('desc'))
    expect(screen.getByRole('columnheader', { name: /Tipo/ })).toHaveAttribute('aria-sort', 'descending')
  })

  it('filtro de unidade mostra CAIXA ALTA mas envia o nome original', async () => {
    mockServidor()
    renderPage()
    const user = userEvent.setup()
    await screen.findByRole('link', { name: /#10/ })
    await user.click(screen.getByRole('combobox', { name: 'Unidade' }))
    await user.click(await screen.findByRole('option', { name: 'REVALLE JUAZEIRO' }))
    await waitFor(() => expect(ultimo().get('revenda')).toBe('Revalle Juazeiro'))
    expect(ultimo().get('offset')).toBe('0')
  })

  it('filtro de colaborador envia colaborador_id', async () => {
    mockServidor({ items: [itemSemColaborador] })
    renderPage()
    const user = userEvent.setup()
    await screen.findByRole('link', { name: /#11/ })
    await user.click(screen.getByRole('button', { name: /Todos os colaboradores/ }))
    await user.type(screen.getByPlaceholderText(/buscar por nome/i), 'mar')
    await user.click(await screen.findByText('MARIA SILVA'))
    await waitFor(() => expect(ultimo().get('colaborador_id')).toBe('7'))
    expect(ultimo().get('offset')).toBe('0')
    expect(colaboradoresPedidos.length).toBeGreaterThan(0)
    expect(colaboradoresPedidos.every((p) => p.get('situacao') === null)).toBe(true)
  })

  it('página além da última depois de o total encolher volta para a última', async () => {
    let total = 60
    mockServidor({ total: () => total })
    renderPage()
    await screen.findByText('Mostrando 1–25 de 60')
    total = 30
    fireEvent.change(screen.getByLabelText('Ir para a página'), { target: { value: '3' } })
    fireEvent.keyDown(screen.getByLabelText('Ir para a página'), { key: 'Enter' })
    await waitFor(() => expect(pedidos.some((p) => p.get('offset') === '25')).toBe(true))
    expect(await screen.findByText('Mostrando 26–30 de 30')).toBeInTheDocument()
  })

  it('busca com atraso volta à página 1 sem pedir (novo termo, página antiga)', async () => {
    mockServidor({ total: () => 60 })
    renderPage()
    const user = userEvent.setup()
    await screen.findByText('Mostrando 1–25 de 60')
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await waitFor(() => expect(ultimo().get('offset')).toBe('25'))
    await user.type(screen.getByLabelText('Buscar equipamentos'), 'dell')
    await waitFor(() => expect(pedidos.some((p) => p.get('search') === 'dell')).toBe(true))
    expect(pedidos.filter((p) => p.get('search') === 'dell').every((p) => p.get('offset') === '0')).toBe(true)
  })

  it('limpar a busca não ressuscita a página antiga (pedido com offset 0 e sem termo)', async () => {
    mockServidor({ total: () => 60 })
    renderPage()
    const user = userEvent.setup()
    await screen.findByText('Mostrando 1–25 de 60')
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await waitFor(() => expect(ultimo().get('offset')).toBe('25'))
    const campo = screen.getByLabelText('Buscar equipamentos')
    await user.type(campo, 'dell')
    await waitFor(() => expect(pedidos.some((p) => p.get('search') === 'dell')).toBe(true))
    const antes = pedidos.length
    await user.clear(campo)
    await waitFor(() => expect(pedidos.slice(antes).some((p) => p.get('search') === null)).toBe(true))
    expect(pedidos.slice(antes).every((p) => p.get('offset') === '0')).toBe(true)
  })

  it('enquanto a página se corrige não mostra "Não há equipamentos" (total > 0)', async () => {
    let total = 60
    mockServidor({ total: () => total })
    renderPage()
    await screen.findByText('Mostrando 1–25 de 60')
    total = 30
    fireEvent.change(screen.getByLabelText('Ir para a página'), { target: { value: '3' } })
    fireEvent.keyDown(screen.getByLabelText('Ir para a página'), { key: 'Enter' })
    await screen.findByText('Mostrando 26–30 de 30')
    expect(screen.queryByText(/Não há equipamentos/)).not.toBeInTheDocument()
  })

  it('sem nenhum equipamento mostra "Não há equipamentos cadastrados"', async () => {
    mockServidor({ items: [] })
    renderPage()
    expect(await screen.findByText(/Não há equipamentos cadastrados/)).toBeInTheDocument()
  })

  it('#id leva à página do equipamento e a tabela rola dentro do painel', async () => {
    mockServidor()
    renderPage()
    const link = await screen.findByRole('link', { name: /#10/ })
    expect(link).toHaveAttribute('href', '/items/10')
    expect(screen.getByRole('table').parentElement).toHaveClass('overflow-x-auto')
  })

  describe('exportação (F7)', () => {
    function mockExportacao() {
      const urls: URL[] = []
      URL.createObjectURL = vi.fn(() => 'blob:teste')
      URL.revokeObjectURL = vi.fn()
      vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
      server.use(
        http.get('/api/exportacoes/estoque', ({ request }) => {
          urls.push(new URL(request.url))
          return new HttpResponse(new Blob(['x']))
        })
      )
      return urls
    }

    it('exporta com os filtros e a ordenação da tela', async () => {
      const urls = mockExportacao()
      mockServidor()
      renderPage()
      const user = userEvent.setup()
      await screen.findByRole('link', { name: /#10/ })
      await user.click(screen.getByRole('combobox', { name: 'Status' }))
      await user.click(await screen.findByRole('option', { name: 'Disponível' }))
      await waitFor(() => expect(ultimo().get('status')).toBe('Disponível'))
      await user.click(screen.getByRole('button', { name: 'Colaborador' }))
      await waitFor(() => expect(ultimo().get('sort')).toBe('colaborador'))

      await user.click(screen.getByRole('button', { name: 'Exportar em Excel' }))

      await waitFor(() => expect(urls).toHaveLength(1))
      const p = urls[0].searchParams
      expect(p.get('status')).toBe('Disponível')
      expect(p.get('sort')).toBe('colaborador')
      expect(p.get('dir')).toBe('asc')
      expect(p.get('formato')).toBe('xlsx')
      expect(p.has('revenda')).toBe(false)
    })

    it('Jovem Aprendiz não vê os botões de exportar', async () => {
      papel.atual = 'Jovem Aprendiz'
      mockServidor()
      renderPage()
      await screen.findByRole('link', { name: /#10/ })
      expect(screen.queryByRole('group', { name: 'Exportar' })).not.toBeInTheDocument()
    })
  })
})
