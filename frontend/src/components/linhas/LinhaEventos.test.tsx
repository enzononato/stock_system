import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { render, screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { formatMoeda } from '@/lib/format'
import { ToastContainer } from '@/components/ui/toast'
import { LinhaEventosList } from '@/components/linhas/LinhaEventosList'
import { LinhaHistoricoDialog } from '@/components/linhas/LinhaHistoricoDialog'
import { LinhaStatusDialog } from '@/components/linhas/LinhaStatusDialog'
import type { Linha, LinhaEvento } from '@/api/linhas'

const base: LinhaEvento = {
  id: 1,
  linha_id: 5,
  linha_numero: '74999990001',
  tipo: 'troca_aparelho',
  campo: null,
  de_valor: 'Celular Samsung A15 · CEL-A',
  para_valor: 'Celular Motorola G54 · CEL-B',
  item_id: 8,
  item_id_anterior: 7,
  ocorrido_em: '2026-11-12T10:00:00',
  por_usuario: 'gestor.teste',
  observacao: 'Tela quebrada',
}

const eventoCusto: LinhaEvento = {
  ...base,
  id: 2,
  tipo: 'edicao',
  campo: 'custo_mensal',
  de_valor: '49.90',
  para_valor: '59.90',
  observacao: null,
}

const linha = {
  id: 5,
  numero: '74999990001',
  operadora: 'Vivo',
  status: 'ativa',
} as Linha

describe('LinhaEventosList', () => {
  it('mostra rótulo, descrição, autor e observação do evento', () => {
    render(<LinhaEventosList eventos={[base, eventoCusto]} />)

    expect(screen.getByText('Troca de aparelho')).toBeInTheDocument()
    expect(screen.getByText('Celular Samsung A15 · CEL-A → Celular Motorola G54 · CEL-B')).toBeInTheDocument()
    expect(screen.getByText('Tela quebrada')).toBeInTheDocument()
    // O Intl separa "R$" do valor com espaço não separável; o getByText normaliza o texto do nó para espaço comum.
    const custo = `Custo mensal: ${formatMoeda('49.90')} → ${formatMoeda('59.90')}`.replace(/\s/g, ' ')
    expect(screen.getByText(custo)).toBeInTheDocument()
    expect(screen.getAllByText(/gestor\.teste/).length).toBe(2)
    expect(screen.queryByText('(74) 99999-0001')).not.toBeInTheDocument()
  })

  it('com mostrarNumero exibe o número da linha', () => {
    render(<LinhaEventosList eventos={[base]} mostrarNumero />)

    expect(screen.getByText('(74) 99999-0001')).toBeInTheDocument()
  })

  it('lista vazia mostra a mensagem', () => {
    render(<LinhaEventosList eventos={[]} />)

    expect(screen.getByText('Nenhum evento registrado.')).toBeInTheDocument()
  })
})

describe('LinhaHistoricoDialog', () => {
  it('aberto, busca os eventos da linha e mostra o título com o número', async () => {
    let chamadas = 0
    server.use(
      http.get('/api/linhas/5/eventos', () => {
        chamadas += 1
        return HttpResponse.json([base])
      })
    )
    renderWithClient(<LinhaHistoricoDialog linha={linha} open onOpenChange={vi.fn()} />)

    expect(await screen.findByText('Histórico da linha (74) 99999-0001')).toBeInTheDocument()
    expect(await screen.findByText('Troca de aparelho')).toBeInTheDocument()
    expect(chamadas).toBe(1)
  })

  it('fechado, não faz requisição', async () => {
    // Sem handler: o MSW (onUnhandledRequest: 'error') falharia se saísse uma requisição.
    renderWithClient(<LinhaHistoricoDialog linha={linha} open={false} onOpenChange={vi.fn()} />)

    await new Promise((r) => setTimeout(r, 50))
    expect(screen.queryByText(/Histórico da linha/)).not.toBeInTheDocument()
  })

  it('mostra o erro quando o histórico não carrega', async () => {
    server.use(http.get('/api/linhas/5/eventos', () => HttpResponse.json({ detail: 'x' }, { status: 500 })))
    renderWithClient(<LinhaHistoricoDialog linha={linha} open onOpenChange={vi.fn()} />)

    expect(await screen.findByText('Não foi possível carregar o histórico.')).toBeInTheDocument()
  })
})

describe('LinhaStatusDialog', () => {
  it('cancelar avisa que libera o aparelho, envia a observação e fecha', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/linhas/5/status', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...linha, status: 'cancelada' })
      })
    )
    const onOpenChange = vi.fn()
    renderWithClient(
      <>
        <LinhaStatusDialog linha={linha} acao="cancelar" open onOpenChange={onOpenChange} />
        <ToastContainer />
      </>
    )
    const user = userEvent.setup()

    expect(
      screen.getByText('A linha deixa de contar no custo mensal e o aparelho vinculado é liberado.')
    ).toBeInTheDocument()
    await user.type(screen.getByLabelText('Observação (opcional)'), 'Chip perdido')
    await user.click(screen.getByRole('button', { name: 'Cancelar linha' }))

    await waitFor(() => expect(corpo).toEqual({ status: 'cancelada', observacao: 'Chip perdido' }))
    await waitFor(() => expect(onOpenChange).toHaveBeenCalledWith(false))
    expect(await screen.findByText('Linha cancelada.')).toBeInTheDocument()
  })

  it('suspender mostra o aviso do custo e envia sem observação quando vazia', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/linhas/5/status', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...linha, status: 'suspensa' })
      })
    )
    renderWithClient(<LinhaStatusDialog linha={linha} acao="suspender" open onOpenChange={vi.fn()} />)
    const user = userEvent.setup()

    expect(screen.getByText('Suspender linha')).toBeInTheDocument()
    expect(screen.getByText('A linha deixa de contar no custo mensal até ser reativada.')).toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: 'Suspender' }))

    await waitFor(() => expect(corpo).toEqual({ status: 'suspensa' }))
  })

  it('reativar usa status ativa', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/linhas/5/status', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...linha, status: 'ativa' })
      })
    )
    renderWithClient(<LinhaStatusDialog linha={linha} acao="reativar" open onOpenChange={vi.fn()} />)
    const user = userEvent.setup()

    await user.click(screen.getByRole('button', { name: 'Reativar' }))

    await waitFor(() => expect(corpo).toEqual({ status: 'ativa' }))
  })

  it('erro da API aparece no diálogo e o diálogo continua aberto', async () => {
    server.use(
      http.post('/api/linhas/5/status', () =>
        HttpResponse.json({ detail: 'A linha já está cancelada.' }, { status: 409 })
      )
    )
    const onOpenChange = vi.fn()
    renderWithClient(<LinhaStatusDialog linha={linha} acao="cancelar" open onOpenChange={onOpenChange} />)
    const user = userEvent.setup()

    await user.click(screen.getByRole('button', { name: 'Cancelar linha' }))

    expect(await screen.findByText('A linha já está cancelada.')).toBeInTheDocument()
    expect(onOpenChange).not.toHaveBeenCalled()
  })

  it('Voltar fecha sem requisição', async () => {
    const onOpenChange = vi.fn()
    renderWithClient(<LinhaStatusDialog linha={linha} acao="cancelar" open onOpenChange={onOpenChange} />)

    await userEvent.setup().click(screen.getByRole('button', { name: 'Voltar' }))

    expect(onOpenChange).toHaveBeenCalledWith(false)
  })
})
