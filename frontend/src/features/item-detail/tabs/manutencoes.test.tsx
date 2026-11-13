import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import type { Item } from '@/api/items'
import type { Manutencao } from '@/api/manutencoes'
import tab from './manutencoes'

const auth = vi.hoisted(() => ({ role: 'Gestor' }))

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'teste', role: auth.role },
    hasRole: (...roles: string[]) => roles.includes(auth.role),
  }),
}))

function m(parcial: Partial<Manutencao>): Manutencao {
  return {
    id: 1,
    item_id: 7,
    status: 'retornada',
    defeito: 'Tela quebrada',
    aberta_por: 'gestor',
    aberta_em: '2026-01-10T09:00:00',
    custo: '150.00',
    custo_acumulado_item: '239.90',
    ...parcial,
  }
}

const item: Item = { id: 7, tipo: 'Notebook', brand: 'Dell', model: 'Latitude', status: 'Disponível' }

/** Responde `GET /api/manutencoes` respeitando o filtro `status=em_andamento` da consulta auxiliar. */
function mockLista(linhas: Manutencao[], custoTotal = '239.90') {
  const chamadas: URLSearchParams[] = []
  server.use(
    http.get('/api/manutencoes', ({ request }) => {
      const params = new URL(request.url).searchParams
      chamadas.push(params)
      const emAndamento = params.get('status') === 'em_andamento'
      const itens = emAndamento
        ? linhas.filter((l) => ['aberta', 'enviada', 'em_reparo'].includes(l.status))
        : linhas
      return HttpResponse.json({
        items: itens,
        total: itens.length,
        custo_total: emAndamento ? '0.00' : custoTotal,
      })
    })
  )
  return chamadas
}

function renderTab(overrides: Partial<Item> = {}, role = 'Gestor') {
  auth.role = role
  const Aba = tab.Component
  return renderWithClient(
    <MemoryRouter>
      <Aba item={{ ...item, ...overrides }} />
      <ToastContainer />
    </MemoryRouter>
  )
}

describe('aba Manutenções', () => {
  it('está registrada como aba 40 para Gestor e Técnico', () => {
    expect(tab.id).toBe('manutencoes')
    expect(tab.label).toBe('Manutenções')
    expect(tab.order).toBe(40)
    expect(tab.roles).toEqual(['Gestor', 'Técnico'])
  })

  it('lista as manutenções do equipamento e o custo acumulado', async () => {
    const chamadas = mockLista([m({ id: 2 }), m({ id: 1, custo: '89.90', defeito: 'Bateria' })])
    renderTab()

    expect(await screen.findByText('Custo acumulado: R$ 239,90')).toBeInTheDocument()
    expect(screen.getByText('Tela quebrada')).toBeInTheDocument()
    expect(screen.getByText('Bateria')).toBeInTheDocument()
    expect(chamadas.some((p) => p.get('item_id') === '7' && p.get('limit') === '25' && p.get('offset') === '0')).toBe(true)
  })

  it('item com colaborador desabilita Abrir manutenção e explica o motivo', async () => {
    mockLista([])
    renderTab({ colaborador_id: 3, status: 'Indisponível' })

    expect(await screen.findByRole('button', { name: 'Abrir manutenção' })).toBeDisabled()
    expect(
      screen.getByText('Equipamento com colaborador: faça a devolução antes de abrir a manutenção.')
    ).toBeInTheDocument()
  })

  it('manutenção em andamento desabilita Abrir manutenção', async () => {
    mockLista([m({ id: 3, status: 'em_reparo' })])
    renderTab({ status: 'Em Manutenção' })

    await waitFor(() => expect(screen.getByRole('button', { name: 'Abrir manutenção' })).toBeDisabled())
    expect(await screen.findByText('Já existe uma manutenção em andamento.')).toBeInTheDocument()
  })

  it('item Disponível abre a manutenção e refaz a consulta', async () => {
    const chamadas = mockLista([])
    let corpo: unknown
    server.use(
      http.post('/api/manutencoes', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ detail: 'Manutenção aberta.', id: 9 })
      })
    )
    const user = userEvent.setup()
    renderTab()

    await user.click(await screen.findByRole('button', { name: 'Abrir manutenção' }))
    const dialogo = await screen.findByRole('dialog')
    const antes = chamadas.length
    await user.type(within(dialogo).getByLabelText('Defeito'), 'Tela quebrada')
    await user.click(within(dialogo).getByRole('button', { name: 'Abrir manutenção' }))

    await waitFor(() => expect(corpo).toEqual({ item_id: 7, defeito: 'Tela quebrada' }))
    expect(await screen.findByText('Manutenção aberta.')).toBeInTheDocument()
    await waitFor(() => expect(chamadas.length).toBeGreaterThan(antes))
  })

  it('sem conserto mostra Dar baixa para o Gestor', async () => {
    mockLista([m({ id: 4, status: 'sem_conserto' })])
    renderTab({ status: 'Em Manutenção' })

    expect(await screen.findByText('Equipamento sem conserto. Avalie a baixa.')).toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'Dar baixa' })).toHaveAttribute('href', '/remove?item=7')
    // Sem ordem em andamento, uma nova manutenção pode ser aberta.
    await waitFor(() => expect(screen.getByRole('button', { name: 'Abrir manutenção' })).toBeEnabled())
  })

  it('sem conserto não mostra Dar baixa para o Técnico', async () => {
    mockLista([m({ id: 4, status: 'sem_conserto' })])
    renderTab({ status: 'Em Manutenção' }, 'Técnico')

    expect(await screen.findByText('Equipamento sem conserto. Avalie a baixa.')).toBeInTheDocument()
    expect(screen.queryByRole('link', { name: 'Dar baixa' })).not.toBeInTheDocument()
  })

  it('sem manutenções mostra o texto vazio', async () => {
    mockLista([], '0.00')
    renderTab()

    expect(await screen.findByText('Nenhuma manutenção registrada para este equipamento.')).toBeInTheDocument()
    expect(screen.getByText('Custo acumulado: R$ 0,00')).toBeInTheDocument()
  })
})
