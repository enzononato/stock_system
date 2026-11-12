import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import AuditPage from '@/pages/AuditPage'
import type { AuditEntry } from '@/api/audit'

const statusEntry: AuditEntry = {
  id: 1,
  ocorrido_em: '2026-11-10T15:30:00',
  username: 'maria',
  papel: 'Gestor',
  ip: '10.0.0.1',
  entidade: 'item',
  entidade_id: '42',
  acao: 'atualizar',
  antes: { status: 'Disponível' },
  depois: { status: 'Indisponível' },
}

const criacao: AuditEntry = {
  id: 2,
  ocorrido_em: '2026-11-10T16:00:00',
  username: null,
  entidade: 'colaborador',
  entidade_id: null,
  acao: 'criar',
  antes: null,
  depois: { cpf: '***.456.789-**', ativo: true },
}

function mockAudit(items: AuditEntry[] = [statusEntry, criacao], total = items.length) {
  const pedidos: URLSearchParams[] = []
  server.use(
    http.get('/api/audit', ({ request }) => {
      pedidos.push(new URL(request.url).searchParams)
      return HttpResponse.json({ items, total })
    })
  )
  return pedidos
}

describe('AuditPage', () => {
  it('mostra o resumo "campo: antes → depois" a partir de antes/depois', async () => {
    mockAudit()
    renderWithClient(<AuditPage />)

    expect(await screen.findByText('status: Disponível → Indisponível')).toBeInTheDocument()
    expect(screen.getByText('maria')).toBeInTheDocument()
    expect(screen.getByText('item #42')).toBeInTheDocument()
    expect(screen.getByText('atualizar')).toBeInTheDocument()
  })

  it('trata criação (sem antes), usuário e id ausentes', async () => {
    mockAudit()
    renderWithClient(<AuditPage />)

    expect(await screen.findByText('cpf: — → ***.456.789-**')).toBeInTheDocument()
    expect(screen.getByText('ativo: — → sim')).toBeInTheDocument()
    expect(screen.getByText('colaborador')).toBeInTheDocument()
  })

  it('interpreta ocorrido_em como UTC e exibe no horário local', async () => {
    // Fuso fixado para o teste não passar "por acaso" numa máquina em UTC:
    // sem o `Z`, 15:00 seria lido como local e sairia 15:00 em vez de 12:00.
    const tzOriginal = process.env.TZ
    process.env.TZ = 'America/Sao_Paulo'
    try {
      mockAudit([{ ...statusEntry, ocorrido_em: '2026-11-12T15:00:00' }])
      renderWithClient(<AuditPage />)
      expect(await screen.findByText('12/11/2026, 12:00:00')).toBeInTheDocument()
    } finally {
      if (tzOriginal === undefined) delete process.env.TZ
      else process.env.TZ = tzOriginal
    }
  })

  it('não consulta o servidor e avisa quando a data inicial é maior que a final', async () => {
    const pedidos = mockAudit()
    const user = userEvent.setup()
    renderWithClient(<AuditPage />)
    await screen.findByText('status: Disponível → Indisponível')
    const antes = pedidos.length

    await user.type(screen.getByLabelText('Até'), '2026-11-01')
    await user.type(screen.getByLabelText('De'), '2026-11-12')

    expect(await screen.findByText('Período inválido: a data inicial é maior que a final.')).toBeInTheDocument()
    expect(pedidos.slice(antes).some((p) => p.get('de') === '2026-11-12' && p.get('ate') === '2026-11-01')).toBe(false)
  })

  it('envia filtros e paginação ao servidor e volta à página 1 ao filtrar', async () => {
    const pedidos = mockAudit([statusEntry], 80)
    const user = userEvent.setup()
    renderWithClient(<AuditPage />)
    await screen.findByText('status: Disponível → Indisponível')
    expect(pedidos[0].get('limit')).toBe('25')
    expect(pedidos[0].get('offset')).toBe('0')

    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await waitFor(() => expect(pedidos.at(-1)?.get('offset')).toBe('25'))

    await user.type(screen.getByLabelText('Entidade'), 'item')
    await waitFor(() => {
      const ultimo = pedidos.at(-1)!
      expect(ultimo.get('entidade')).toBe('item')
      expect(ultimo.get('offset')).toBe('0')
    })

    await user.type(screen.getByLabelText('Usuário'), 'maria')
    await waitFor(() => expect(pedidos.at(-1)?.get('username')).toBe('maria'))

    await user.type(screen.getByLabelText('De'), '2026-11-01')
    await user.type(screen.getByLabelText('Até'), '2026-11-12')
    await waitFor(() => {
      const ultimo = pedidos.at(-1)!
      expect(ultimo.get('de')).toBe('2026-11-01')
      expect(ultimo.get('ate')).toBe('2026-11-12')
      expect(ultimo.get('offset')).toBe('0')
    })
  })

  it('mostra estado vazio', async () => {
    mockAudit([], 0)
    renderWithClient(<AuditPage />)
    expect(await screen.findByText('Nenhum registro encontrado')).toBeInTheDocument()
  })
})
