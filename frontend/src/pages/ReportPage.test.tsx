import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import ReportPage from '@/pages/ReportPage'

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

const linhas = Array.from({ length: 30 }, (_, i) => ({
  history_id: i + 1,
  item_id: i + 1,
  operador: 'gestor.teste',
  operation_type: 'Cadastro',
  tipo: 'Notebook',
  revenda: i < 20 ? 'Revalle Juazeiro' : 'Revalle Petrolina',
  data_emprestimo: '2026-01-10T09:00:00',
}))

function renderPage() {
  server.use(
    http.get('/api/unidades', () => HttpResponse.json([])),
    http.get('/api/reports/monthly', () => HttpResponse.json(linhas))
  )
  return renderWithClient(<ReportPage />)
}

describe('ReportPage', () => {
  beforeEach(() => {
    localStorage.clear()
    papel.atual = 'Gestor'
  })

  it('pagina no cliente com 25 por padrão', async () => {
    renderPage()
    expect(await screen.findByText('Mostrando 1–25 de 30')).toBeInTheDocument()
  })

  it('revenda em CAIXA ALTA e filtro pelo nome original', async () => {
    renderPage()
    const user = userEvent.setup()
    expect((await screen.findAllByText('REVALLE JUAZEIRO')).length).toBeGreaterThan(0)
    await user.click(screen.getByRole('combobox', { name: 'Revenda' }))
    await user.click(await screen.findByRole('option', { name: 'REVALLE PETROLINA' }))
    expect(await screen.findByText('Mostrando 1–10 de 10')).toBeInTheDocument()
  })

  it('indicadores em caixas azuis', async () => {
    renderPage()
    const rotulo = await screen.findByText('Total de Registros')
    expect(rotulo.closest('div')).toHaveClass('bg-brand-soft')
  })

  it('opções de revenda sem duplicata por caixa e filtro sem diferenciar caixa', async () => {
    server.use(
      http.get('/api/unidades', () =>
        HttpResponse.json([{ id: 1, nome: 'Revalle Juazeiro', razao_social: 'X', cnpj: '1', is_active: true }])
      ),
      http.get('/api/reports/monthly', () =>
        HttpResponse.json([
          { history_id: 1, item_id: 1, operation_type: 'Cadastro', revenda: 'REVALLE JUAZEIRO' },
          { history_id: 2, item_id: 2, operation_type: 'Cadastro', revenda: 'Revalle Juazeiro' },
        ])
      )
    )
    renderWithClient(<ReportPage />)
    const user = userEvent.setup()
    await screen.findByText('Mostrando 1–2 de 2')
    await user.click(screen.getByRole('combobox', { name: 'Revenda' }))
    expect(await screen.findAllByRole('option', { name: 'REVALLE JUAZEIRO' })).toHaveLength(1)
    await user.click(screen.getByRole('option', { name: 'REVALLE JUAZEIRO' }))
    expect(await screen.findByText('Mostrando 1–2 de 2')).toBeInTheDocument()
  })

  it('o Gestor exporta o mês gerado em PDF e o CSV continua lá', async () => {
    const urls: URL[] = []
    URL.createObjectURL = vi.fn(() => 'blob:teste')
    URL.revokeObjectURL = vi.fn()
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
    server.use(
      http.get('/api/exportacoes/relatorio-mensal', ({ request }) => {
        urls.push(new URL(request.url))
        return new HttpResponse(new Blob(['x']))
      })
    )
    renderPage()
    const user = userEvent.setup()
    await screen.findByText('Mostrando 1–25 de 30')
    expect(screen.getByRole('button', { name: 'CSV' })).toBeInTheDocument()

    await user.click(screen.getByRole('button', { name: 'Exportar em PDF' }))

    await vi.waitFor(() => expect(urls).toHaveLength(1))
    const agora = new Date()
    expect(urls[0].searchParams.get('year')).toBe(String(agora.getFullYear()))
    expect(urls[0].searchParams.get('month')).toBe(String(agora.getMonth() + 1))
    expect(urls[0].searchParams.get('formato')).toBe('pdf')
  })

  it('o Técnico não vê nenhuma exportação do relatório (a API é só do Gestor)', async () => {
    papel.atual = 'Técnico'
    renderPage()
    await screen.findByText('Mostrando 1–25 de 30')
    expect(screen.queryByRole('group', { name: 'Exportar' })).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'CSV' })).not.toBeInTheDocument()
    expect(screen.queryByText(/O CSV exportado/)).not.toBeInTheDocument()
  })
})
