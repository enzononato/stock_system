import { describe, expect, it } from 'vitest'
import { screen } from '@testing-library/react'
import { http, HttpResponse } from 'msw'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import type { Colaborador } from '@/api/colaboradores'
import type { TermoSolicitacao } from '@/api/termos'
import tab from './termos'

const Termos = tab.Component

const colaborador = { id: 8, nome: 'Maria Souza' } as Colaborador

function termo(extra: Partial<TermoSolicitacao> = {}): TermoSolicitacao {
  return {
    id: 30,
    item_id: 12,
    item_descricao: 'Notebook Dell Latitude 5420',
    tipo: 'entrega',
    status: 'assinado',
    solicitado_por: 'tecnico.teste',
    solicitado_em: '2026-11-10T09:00:00',
    assinado_em: '2026-11-12T10:00:00',
    tem_documento: true,
    tem_assinado: true,
    legado: false,
    ...extra,
  }
}

function renderTab() {
  return renderWithClient(
    <MemoryRouter>
      <Termos colaborador={colaborador} />
    </MemoryRouter>
  )
}

describe('aba Termos do colaborador', () => {
  it('é uma aba só de Gestor e Técnico, na ordem 30', () => {
    expect(tab.id).toBe('termos')
    expect(tab.order).toBe(30)
    expect(tab.roles).toEqual(['Gestor', 'Técnico'])
  })

  it('lista os termos com o link do equipamento e as pílulas de status', async () => {
    let query = ''
    const itens = [
      termo(),
      termo({ id: 29, item_id: 15, item_descricao: 'Celular Samsung A54', tipo: 'devolucao', status: 'enviado', assinado_em: null }),
    ]
    server.use(
      http.get('/api/colaboradores/8/termos', ({ request }) => {
        query = new URL(request.url).search
        return HttpResponse.json({ items: itens, total: 2 })
      })
    )
    renderTab()

    const link = await screen.findByRole('link', { name: 'Notebook Dell Latitude 5420' })
    expect(link).toHaveAttribute('href', '/items/12?aba=termos')
    expect(screen.getByRole('link', { name: 'Celular Samsung A54' })).toHaveAttribute('href', '/items/15?aba=termos')
    expect(screen.getByText('Assinado')).toBeInTheDocument()
    expect(screen.getByText('Enviado')).toBeInTheDocument()
    expect(screen.getByText('Termo de entrega')).toBeInTheDocument()
    expect(screen.getByText('Termo de devolução')).toBeInTheDocument()
    expect(screen.getByText('Mostrando 1–2 de 2')).toBeInTheDocument()
    expect(query).toContain('limit=25')
    expect(query).toContain('offset=0')
  })

  it('mostra a mensagem de vazio', async () => {
    server.use(http.get('/api/colaboradores/8/termos', () => HttpResponse.json({ items: [], total: 0 })))
    renderTab()

    expect(await screen.findByText('Nenhum termo registrado para este colaborador.')).toBeInTheDocument()
  })

  it('avisa quando a lista não carrega', async () => {
    server.use(http.get('/api/colaboradores/8/termos', () => HttpResponse.json({ detail: 'x' }, { status: 500 })))
    renderTab()

    expect(await screen.findByText('Não foi possível carregar os termos.')).toBeInTheDocument()
  })
})
