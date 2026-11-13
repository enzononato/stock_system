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

const resumoBase = {
  esperados: 3,
  lidos: 1,
  encontrados: 1,
  ok: 1,
  local_errado: 0,
  usuario_diferente: 0,
  outra_unidade: 0,
  nao_cadastrado: 0,
  nao_encontrados: 2,
}

function linha(extra: Record<string, unknown>) {
  return {
    categoria: 'ok',
    leitura_id: null,
    item_id: null,
    tipo: 'Notebook',
    brand: 'Dell',
    model: 'Latitude',
    identificador: 'SN-1',
    unidade_item: 'Juazeiro',
    ...extra,
  }
}

const linhas = [
  linha({ categoria: 'nao_encontrado', item_id: 7 }),
  linha({ categoria: 'nao_encontrado', item_id: 8 }),
  linha({ categoria: 'ok', item_id: 5, leitura_id: 1, lido_por: 'ana', lido_em: '2026-11-10T10:00:00' }),
]

function relatorio(status: string) {
  return {
    inventario: {
      id: 12,
      unidade_id: 1,
      unidade_nome: 'Juazeiro',
      status,
      iniciado_por: 'tecnico.teste',
      iniciado_em: '2026-11-10T09:00:00',
      encerrado_por: status === 'aberto' ? null : 'tecnico.teste',
      encerrado_em: status === 'aberto' ? null : '2026-11-10T18:00:00',
      observacao: null,
      motivo_cancelamento: null,
      total_leituras: 1,
      resumo: resumoBase,
    },
    linhas,
  }
}

let status = 'encerrado'

beforeEach(() => {
  role = 'Técnico'
  status = 'encerrado'
  localStorage.clear()
  server.use(http.get('/api/inventarios/12/relatorio', () => HttpResponse.json(relatorio(status))))
})

function renderRota(caminho = '/inventarios/12') {
  return renderWithClient(
    <ThemeProvider>
      <MemoryRouter initialEntries={[caminho]}>
        <App />
      </MemoryRouter>
    </ThemeProvider>
  )
}

describe('InventarioRelatorioPage', () => {
  it('mostra o resumo e as abas com contagem', async () => {
    const user = userEvent.setup()
    renderRota()

    expect(await screen.findByRole('heading', { name: 'Inventário #12' })).toBeInTheDocument()
    const abas = screen.getByRole('tablist', { name: 'Resultados do inventário' })
    expect(within(abas).getByRole('tab', { name: 'Não encontrados (2)' })).toBeInTheDocument()
    expect(within(abas).getByRole('tab', { name: 'Conferidos (1)' })).toBeInTheDocument()
    // Abre na primeira categoria com itens.
    expect(within(abas).getByRole('tab', { name: 'Não encontrados (2)' })).toHaveAttribute('aria-selected', 'true')
    expect(screen.getByRole('link', { name: '#7' })).toHaveAttribute('href', '/items/7')

    await user.click(within(abas).getByRole('tab', { name: 'Conferidos (1)' }))
    expect(await screen.findByRole('link', { name: '#5' })).toBeInTheDocument()
    expect(screen.queryByRole('link', { name: '#7' })).not.toBeInTheDocument()

    await user.click(within(abas).getByRole('tab', { name: 'Não encontrados (2)' }))
    expect(await screen.findByRole('link', { name: '#7' })).toBeInTheDocument()
  })

  it('a aba escolhida fica na URL (?aba=)', async () => {
    renderRota('/inventarios/12?aba=ok')
    expect(await screen.findByRole('link', { name: '#5' })).toBeInTheDocument()
    expect(screen.getByRole('tab', { name: 'Conferidos (1)' })).toHaveAttribute('aria-selected', 'true')
  })

  it('inventário aberto chama de "Ainda não lidos" e oferece continuar', async () => {
    status = 'aberto'
    renderRota()

    expect(await screen.findByRole('tab', { name: 'Ainda não lidos (2)' })).toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'Continuar leitura' })).toHaveAttribute('href', '/inventarios/12/leitura')
    expect(screen.getByRole('button', { name: 'Encerrar inventário' })).toBeInTheDocument()
  })

  it('inventário encerrado não oferece continuar nem encerrar', async () => {
    renderRota()
    await screen.findByRole('heading', { name: 'Inventário #12' })
    expect(screen.queryByRole('link', { name: 'Continuar leitura' })).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Encerrar inventário' })).not.toBeInTheDocument()
  })

  it('encerrar pede confirmação e chama a API', async () => {
    status = 'aberto'
    let chamou = false
    server.use(
      http.post('/api/inventarios/12/encerrar', () => {
        chamou = true
        return HttpResponse.json(relatorio('encerrado').inventario)
      })
    )
    const user = userEvent.setup()
    renderRota()

    await user.click(await screen.findByRole('button', { name: 'Encerrar inventário' }))
    const dialogo = await screen.findByRole('alertdialog')
    expect(within(dialogo).getByText(/2 equipamento\(s\) ainda não lido\(s\)/)).toBeInTheDocument()
    expect(chamou).toBe(false)
    await user.click(within(dialogo).getByRole('button', { name: 'Encerrar' }))

    await waitFor(() => expect(chamou).toBe(true))
    expect(await screen.findByText('Inventário encerrado.')).toBeInTheDocument()
  })

  it('cancelar só aparece para o Gestor e exige o motivo', async () => {
    status = 'aberto'
    let corpo: unknown = null
    server.use(
      http.post('/api/inventarios/12/cancelar', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json(relatorio('cancelado').inventario)
      })
    )
    const tecnico = renderRota()
    await screen.findByRole('button', { name: 'Encerrar inventário' })
    expect(screen.queryByRole('button', { name: 'Cancelar inventário' })).not.toBeInTheDocument()
    tecnico.unmount()

    role = 'Gestor'
    const user = userEvent.setup()
    renderRota()
    await user.click(await screen.findByRole('button', { name: 'Cancelar inventário' }))
    const dialogo = await screen.findByRole('dialog')
    const confirmar = within(dialogo).getByRole('button', { name: 'Cancelar inventário' })
    expect(confirmar).toBeDisabled()
    await user.type(within(dialogo).getByLabelText('Motivo do cancelamento'), 'Contagem refeita')
    await user.click(confirmar)

    await waitFor(() => expect(corpo).toEqual({ motivo: 'Contagem refeita' }))
  })

  it('exportar Excel baixa o arquivo', async () => {
    let query = ''
    server.use(
      http.get('/api/inventarios/12/relatorio/exportar', ({ request }) => {
        query = new URL(request.url).search
        return new HttpResponse(new Blob(['x']), { headers: { 'Content-Type': 'application/octet-stream' } })
      })
    )
    const criar = vi.fn(() => 'blob:x')
    Object.defineProperty(URL, 'createObjectURL', { value: criar, configurable: true })
    Object.defineProperty(URL, 'revokeObjectURL', { value: vi.fn(), configurable: true })
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
    const user = userEvent.setup()
    renderRota()

    await user.click(await screen.findByRole('button', { name: 'Exportar Excel' }))

    await waitFor(() => expect(query).toBe('?formato=xlsx'))
    await waitFor(() => expect(criar).toHaveBeenCalled())
  })
})
