import { beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import App from '@/App'
import { ThemeProvider } from '@/lib/theme'

let role = 'Técnico'
vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'teste', role },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes(role),
  }),
}))

const unidades = [
  { id: 1, nome: 'Juazeiro', razao_social: 'X', cnpj: '1', is_active: true },
  { id: 2, nome: 'Petrolina', razao_social: 'Y', cnpj: '2', is_active: true },
]

const itens = [3, 5].map((id) => ({
  id,
  tipo: 'Notebook',
  brand: 'Dell',
  model: 'Latitude 5420',
  identificador: `SN-${id}`,
  revenda: 'Juazeiro',
  status: 'Disponível',
}))

function etiquetaDe(id: number) {
  return { ...itens[0], id, url: `${window.location.origin}/e/${id}` }
}

let baseUrl = ''
let corpoGerar: unknown = null
let consultasItens: string[] = []

beforeEach(() => {
  role = 'Técnico'
  baseUrl = window.location.origin
  corpoGerar = null
  consultasItens = []
  localStorage.clear()
  window.print = vi.fn()
  server.use(
    http.get('/api/unidades', () => HttpResponse.json(unidades)),
    http.get('/api/constants', () =>
      HttpResponse.json({
        center_costs: [],
        revendas: [],
        setores: [],
        equipment_types: ['Notebook', 'Monitor'],
        peripheral_types: [],
        removal_reasons: [],
        removal_reasons_attachment: {},
      })
    ),
    http.get('/api/items', ({ request }) => {
      consultasItens.push(new URL(request.url).search)
      return HttpResponse.json({ items: itens, total: 2 })
    }),
    http.post('/api/etiquetas', async ({ request }) => {
      corpoGerar = await request.json()
      const ids = (corpoGerar as { item_ids: number[] }).item_ids
      return HttpResponse.json({ base_url: baseUrl, etiquetas: ids.map(etiquetaDe), ignorados: [] })
    })
  )
})

function renderRota() {
  return renderWithClient(
    <ThemeProvider>
      <MemoryRouter initialEntries={['/etiquetas']}>
        <App />
      </MemoryRouter>
    </ThemeProvider>
  )
}

async function gerarDosDoisItens(user: ReturnType<typeof userEvent.setup>) {
  await user.click(await screen.findByLabelText('Selecionar #5'))
  await user.click(screen.getByLabelText('Selecionar #3'))
  await user.click(screen.getByRole('button', { name: 'Gerar etiquetas' }))
  await screen.findByText('Folha de etiquetas')
}

describe('EtiquetasPage (/etiquetas)', () => {
  it('filtra pela unidade e pagina no servidor', async () => {
    const user = userEvent.setup()
    renderRota()
    expect(await screen.findByRole('heading', { name: 'Etiquetas' })).toBeInTheDocument()
    expect(await screen.findByRole('option', { name: 'JUAZEIRO' })).toBeInTheDocument()

    await user.selectOptions(screen.getByLabelText('Unidade'), 'JUAZEIRO')

    await waitFor(() => expect(consultasItens).toContain('?revenda=Juazeiro&limit=25&offset=0'))
  })

  it('gera as etiquetas dos selecionados em ordem crescente', async () => {
    const user = userEvent.setup()
    renderRota()

    await gerarDosDoisItens(user)

    expect(corpoGerar).toEqual({ item_ids: [3, 5] })
    const previa = screen.getByLabelText('Prévia da folha')
    expect(within(previa).getByText('#3')).toBeInTheDocument()
    expect(screen.getByText('2 etiqueta(s) em 1 folha(s)')).toBeInTheDocument()
  })

  it('selecionar todos da unidade usa os filtros e soma à seleção', async () => {
    let query = ''
    server.use(
      http.get('/api/etiquetas/ids', ({ request }) => {
        query = new URL(request.url).search
        return HttpResponse.json({ ids: Array.from({ length: 40 }, (_, i) => i + 1), total: 40 })
      })
    )
    const user = userEvent.setup()
    renderRota()
    const botao = await screen.findByRole('button', { name: 'Selecionar todos da unidade' })
    expect(botao).toBeDisabled()

    await user.selectOptions(await screen.findByLabelText('Unidade'), 'JUAZEIRO')
    await user.click(screen.getByRole('button', { name: 'Selecionar todos da unidade' }))

    expect(await screen.findByText('40 selecionado(s)')).toBeInTheDocument()
    expect(query).toBe('?revenda=Juazeiro')
  })

  it('a seleção sobrevive à troca de filtro e "Limpar seleção" zera', async () => {
    const user = userEvent.setup()
    renderRota()
    await user.click(await screen.findByLabelText('Selecionar #5'))
    await user.selectOptions(screen.getByLabelText('Tipo'), 'Notebook')
    expect(screen.getByText('1 selecionado(s)')).toBeInTheDocument()

    await user.click(screen.getByRole('button', { name: 'Limpar seleção' }))

    expect(screen.getByText('0 selecionado(s)')).toBeInTheDocument()
  })

  it('layout inválido bloqueia a impressão', async () => {
    const user = userEvent.setup()
    renderRota()
    await gerarDosDoisItens(user)
    const colunas = screen.getByLabelText('Colunas')

    await user.clear(colunas)
    await user.type(colunas, '4')

    expect(screen.getByRole('alert')).toHaveTextContent('O layout não cabe numa folha A4 (210 × 297 mm).')
    expect(screen.getByRole('button', { name: 'Imprimir' })).toBeDisabled()

    await user.clear(colunas)
    await user.type(colunas, '3')

    expect(screen.queryByRole('alert')).not.toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: 'Imprimir' }))
    expect(window.print).toHaveBeenCalledTimes(1)
  })

  it('posição inicial fora da folha também bloqueia', async () => {
    const user = userEvent.setup()
    renderRota()
    await gerarDosDoisItens(user)
    const posicao = screen.getByLabelText('Posição inicial')

    await user.clear(posicao)
    await user.type(posicao, '25')

    expect(screen.getByRole('alert')).toHaveTextContent('A posição inicial deve estar entre 1 e 24.')
    expect(screen.getByRole('button', { name: 'Imprimir' })).toBeDisabled()
  })

  it('guarda o layout e não a posição inicial', async () => {
    const user = userEvent.setup()
    const primeira = renderRota()
    await gerarDosDoisItens(user)
    const colunas = screen.getByLabelText('Colunas')
    await user.clear(colunas)
    await user.type(colunas, '2')
    const posicao = screen.getByLabelText('Posição inicial')
    await user.clear(posicao)
    await user.type(posicao, '5')

    const salvo = JSON.parse(localStorage.getItem('revalle_etiquetas_layout')!)
    expect(salvo.colunas).toBe(2)
    expect(salvo).not.toHaveProperty('posicaoInicial')
    primeira.unmount()

    renderRota()
    await gerarDosDoisItens(user)

    expect(screen.getByLabelText('Colunas')).toHaveValue(2)
    expect(screen.getByLabelText('Posição inicial')).toHaveValue(1)
  })

  it('"Restaurar padrão" volta ao 3 × 8', async () => {
    const user = userEvent.setup()
    renderRota()
    await gerarDosDoisItens(user)
    const colunas = screen.getByLabelText('Colunas')
    await user.clear(colunas)
    await user.type(colunas, '2')

    await user.click(screen.getByRole('button', { name: 'Restaurar padrão' }))

    expect(screen.getByLabelText('Colunas')).toHaveValue(3)
    expect(screen.getByLabelText('Linhas')).toHaveValue(8)
  })

  it('avisa quando o endereço do QR difere', async () => {
    baseUrl = 'https://outro.exemplo'
    const user = userEvent.setup()
    renderRota()

    await gerarDosDoisItens(user)

    expect(screen.getByText(/Confira APP_PUBLIC_URL no servidor/)).toBeInTheDocument()
  })

  it('não avisa quando o endereço do QR é o da página', async () => {
    const user = userEvent.setup()
    renderRota()
    await gerarDosDoisItens(user)
    expect(screen.queryByText(/Confira APP_PUBLIC_URL no servidor/)).not.toBeInTheDocument()
  })

  it('sair da tela limpa a impressão', async () => {
    const user = userEvent.setup()
    const { unmount } = renderRota()
    await gerarDosDoisItens(user)
    expect(document.body.classList.contains('imprimindo-etiquetas')).toBe(true)

    unmount()

    expect(document.body.classList.contains('imprimindo-etiquetas')).toBe(false)
    expect(document.querySelector('.etiquetas-impressao')).toBeNull()
  })

  it('Jovem Aprendiz não acessa', async () => {
    role = 'Jovem Aprendiz'
    renderRota()
    expect(await screen.findByText('Acesso restrito')).toBeInTheDocument()
    expect(screen.queryByLabelText('Unidade')).not.toBeInTheDocument()
  })
})
