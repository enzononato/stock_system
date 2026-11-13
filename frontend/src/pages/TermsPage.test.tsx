import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import TermsPage from '@/pages/TermsPage'

const item5 = {
  id: 5,
  tipo: 'Notebook',
  brand: 'Dell',
  model: 'Latitude',
  identificador: 'PAT-005',
  status: 'Pendente',
  assigned_to: 'Maria Souza',
  cpf: '52998224725',
  revenda: 'Revalle Juazeiro',
  setor: 'Financeiro',
  date_issued: '2026-11-10T09:00:00',
}

function mockItens(captura?: (status: string, p: URLSearchParams) => void) {
  server.use(
    http.get('/api/items', ({ request }) => {
      const params = new URL(request.url).searchParams
      const status = params.get('status') ?? ''
      captura?.(status, params)
      const lista = status === 'Pendente' ? [item5] : []
      return HttpResponse.json({ items: lista, total: lista.length })
    })
  )
}

function renderPage() {
  return renderWithClient(
    <>
      <TermsPage />
      <ToastContainer />
    </>
  )
}

describe('TermsPage', () => {
  it('mostra a assinatura eletrônica do termo selecionado', async () => {
    mockItens()
    server.use(
      http.get('/api/items/5/termos', () =>
        HttpResponse.json({
          revsign_ativo: false,
          vigente: {
            id: 31,
            item_id: 5,
            tipo: 'entrega',
            status: 'manual',
            modelo_versao: 'termo.docx@1a2b3c4d',
            solicitado_por: 'gestor.teste',
            solicitado_em: '2026-11-10T09:00:00',
            motivo: 'Integração com o RevSign desligada: baixe o termo e anexe o PDF assinado.',
            tem_documento: true,
            tem_assinado: false,
            legado: false,
          },
          historico: { items: [], total: 0 },
        })
      )
    )
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByText('Maria Souza'))

    expect(await screen.findByText('Assinatura eletrônica')).toBeInTheDocument()
    expect(await screen.findByText('Assinatura manual')).toBeInTheDocument()
    expect(
      screen.getByText('Integração com o RevSign desligada: baixe o termo e anexe o PDF assinado.')
    ).toBeInTheDocument()
    expect(screen.getByText('REVALLE JUAZEIRO', { selector: 'span.font-medium' })).toBeInTheDocument()
  })

  it('paginação padrão de 25', async () => {
    const limites: string[] = []
    mockItens((status, p) => {
      if (status === 'Pendente') limites.push(p.get('limit') ?? '')
    })
    renderPage()

    expect(await screen.findByText('Mostrando 1–1 de 1')).toBeInTheDocument()
    await waitFor(() => expect(limites).toContain('25'))
    expect(screen.getByText('REVALLE JUAZEIRO')).toBeInTheDocument()
  })

  it('avisa quando a assinatura eletrônica não carrega', async () => {
    mockItens()
    server.use(http.get('/api/items/5/termos', () => HttpResponse.json({ detail: 'x' }, { status: 500 })))
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByText('Maria Souza'))

    expect(await screen.findByText('Não foi possível carregar a assinatura.')).toBeInTheDocument()
  })

  it('exporta os termos pendentes em Excel', async () => {
    mockItens()
    const urls: URL[] = []
    URL.createObjectURL = vi.fn(() => 'blob:teste')
    URL.revokeObjectURL = vi.fn()
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
    server.use(
      http.get('/api/exportacoes/termos-pendentes', ({ request }) => {
        urls.push(new URL(request.url))
        return new HttpResponse(new Blob(['x']))
      })
    )
    const user = userEvent.setup()
    renderPage()
    await screen.findByText('Maria Souza')

    await user.click(screen.getByRole('button', { name: 'Exportar termos pendentes em Excel' }))

    await waitFor(() => expect(urls).toHaveLength(1))
    expect(urls[0].pathname).toBe('/api/exportacoes/termos-pendentes')
    expect(urls[0].search).toBe('?formato=xlsx')
  })
})
