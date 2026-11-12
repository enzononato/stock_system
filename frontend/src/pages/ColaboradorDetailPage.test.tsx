import { beforeEach, describe, expect, it, vi } from 'vitest'
import { act, screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { http, HttpResponse } from 'msw'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import ColaboradorDetailPage from '@/pages/ColaboradorDetailPage'
import { buildColaboradorDetailTabs } from '@/features/colaborador-detail/registry'
import type { ColaboradorDetailTab } from '@/features/colaborador-detail/types'
import type { Colaborador } from '@/api/colaboradores'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'tecnico.teste', role: 'Técnico' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Técnico'),
  }),
}))

const maria: Colaborador = {
  id: 7,
  cpf: '12345678909',
  nome: 'Maria Silva',
  matricula: 'M-0042',
  cargo: 'Analista',
  setor: 'Financeiro',
  unidade_nome: 'Juazeiro',
  unidade_id: 1,
  situacao: 'ativo',
  email: 'maria@revalle.com',
  telefone: '74999990000',
  email_manual: false,
  telefone_manual: false,
  superior_id: 3,
  superior_nome: 'João Chefe',
  origem: 'rh',
  sincronizado_em: '2026-11-10T10:01:00',
}

const equipamento = (id: number) => ({
  id,
  tipo: 'Notebook',
  brand: 'Dell',
  model: `Latitude ${id}`,
  identificador: `NB-${String(id).padStart(4, '0')}`,
  status: 'Em uso',
})

function renderPage(path: string) {
  return renderWithClient(
    <MemoryRouter initialEntries={[path]}>
      <Routes>
        <Route path="/colaboradores/:id" element={<ColaboradorDetailPage />} />
        <Route path="/colaboradores" element={<p>Lista de colaboradores</p>} />
        <Route path="/items/:id" element={<p>Detalhe do equipamento</p>} />
      </Routes>
    </MemoryRouter>
  )
}

beforeEach(() => {
  localStorage.clear()
})

describe('ColaboradorDetailPage', () => {
  it('mostra o cabeçalho (nome, CPF, cargo, setor, unidade em maiúsculas e situação) e as abas Dados e Equipamentos', async () => {
    server.use(http.get('/api/colaboradores/7', () => HttpResponse.json(maria)))
    renderPage('/colaboradores/7')

    expect(await screen.findByRole('heading', { level: 1, name: 'MARIA SILVA' })).toBeInTheDocument()
    expect(screen.getByText('123.456.789-09')).toBeInTheDocument()
    expect(screen.getAllByText('Analista').length).toBeGreaterThan(0)
    expect(screen.getAllByText('Financeiro').length).toBeGreaterThan(0)
    expect(screen.getByText('JUAZEIRO')).toBeInTheDocument()
    expect(screen.getByText('Ativo')).toBeInTheDocument()

    const tabs = within(screen.getByRole('tablist')).getAllByRole('tab')
    const nomesAbas = tabs.map((t) => t.textContent)
    expect(nomesAbas).toContain('Dados')
    expect(nomesAbas).toContain('Equipamentos')
    expect(nomesAbas.indexOf('Dados')).toBeLessThan(nomesAbas.indexOf('Equipamentos'))
    expect(screen.getByRole('tab', { name: 'Dados' })).toHaveAttribute('aria-selected', 'true')
    expect(screen.getByRole('tabpanel')).toBeInTheDocument()
  })

  it('aba Dados mostra matrícula, e-mail, telefone, origem e última sincronização', async () => {
    server.use(http.get('/api/colaboradores/7', () => HttpResponse.json(maria)))
    renderPage('/colaboradores/7')

    const painel = within(await screen.findByRole('tabpanel'))
    expect(painel.getByText('M-0042')).toBeInTheDocument()
    expect(painel.getByText('maria@revalle.com')).toBeInTheDocument()
    expect(painel.getByText('(74) 99999-0000')).toBeInTheDocument()
    expect(painel.getByText('RH')).toBeInTheDocument()
    expect(painel.getByText(/10\/11\/2026/)).toBeInTheDocument()
  })

  it('origem manual aparece como "Manual" e sem sincronização mostra "—"', async () => {
    server.use(
      http.get('/api/colaboradores/7', () => HttpResponse.json({ ...maria, origem: 'manual', sincronizado_em: null }))
    )
    renderPage('/colaboradores/7')
    const painel = within(await screen.findByRole('tabpanel'))
    expect(painel.getByText('Manual')).toBeInTheDocument()
    const dt = painel.getByText('Última sincronização')
    expect(dt.parentElement).toHaveTextContent('—')
  })

  it('o superior imediato é um link para o detalhe do superior', async () => {
    server.use(http.get('/api/colaboradores/7', () => HttpResponse.json(maria)))
    renderPage('/colaboradores/7')

    const link = await screen.findByRole('link', { name: 'JOÃO CHEFE' })
    expect(link).toHaveAttribute('href', '/colaboradores/3')
  })

  it('sem superior mostra "Não informado"', async () => {
    server.use(
      http.get('/api/colaboradores/7', () => HttpResponse.json({ ...maria, superior_id: null, superior_nome: null }))
    )
    renderPage('/colaboradores/7')
    const painel = within(await screen.findByRole('tabpanel'))
    expect(painel.getByText('Não informado')).toBeInTheDocument()
    expect(painel.queryByRole('link')).not.toBeInTheDocument()
  })

  it('?aba=equipamentos consulta /api/items com colaborador_id e lista com link para o equipamento', async () => {
    server.use(http.get('/api/colaboradores/7', () => HttpResponse.json(maria)))
    const pedidos: URLSearchParams[] = []
    server.use(
      http.get('/api/items', ({ request }) => {
        pedidos.push(new URL(request.url).searchParams)
        return HttpResponse.json({ items: [equipamento(11)], total: 1 })
      })
    )
    renderPage('/colaboradores/7?aba=equipamentos')

    const link = await screen.findByRole('link', { name: /Dell Latitude 11/ })
    expect(link).toHaveAttribute('href', '/items/11')
    expect(screen.getByRole('tab', { name: 'Equipamentos' })).toHaveAttribute('aria-selected', 'true')
    expect(pedidos[0].get('colaborador_id')).toBe('7')
    expect(pedidos[0].get('limit')).toBe('25')
    expect(pedidos[0].get('offset')).toBe('0')
  })

  it('aba Equipamentos sem itens mostra mensagem vazia', async () => {
    server.use(
      http.get('/api/colaboradores/7', () => HttpResponse.json(maria)),
      http.get('/api/items', () => HttpResponse.json({ items: [], total: 0 }))
    )
    renderPage('/colaboradores/7?aba=equipamentos')
    expect(await screen.findByText('Nenhum equipamento com este colaborador.')).toBeInTheDocument()
  })

  it('aba Equipamentos com erro mostra mensagem de falha', async () => {
    server.use(
      http.get('/api/colaboradores/7', () => HttpResponse.json(maria)),
      http.get('/api/items', () => HttpResponse.json({ detail: 'x' }, { status: 500 }))
    )
    renderPage('/colaboradores/7?aba=equipamentos')
    expect(await screen.findByText('Não foi possível carregar os equipamentos.')).toBeInTheDocument()
  })

  it('aba desconhecida em ?aba= cai na primeira aba', async () => {
    server.use(http.get('/api/colaboradores/7', () => HttpResponse.json(maria)))
    renderPage('/colaboradores/7?aba=nao-existe')
    await screen.findByRole('tablist')
    expect(screen.getByRole('tab', { name: 'Dados' })).toHaveAttribute('aria-selected', 'true')
  })

  it('trocar de aba atualiza a seleção e as setas navegam entre abas (roving tabindex)', async () => {
    server.use(
      http.get('/api/colaboradores/7', () => HttpResponse.json(maria)),
      http.get('/api/items', () => HttpResponse.json({ items: [], total: 0 }))
    )
    renderPage('/colaboradores/7')
    await screen.findByRole('tablist')
    const dados = screen.getByRole('tab', { name: 'Dados' })
    const equip = screen.getByRole('tab', { name: 'Equipamentos' })
    expect(dados).toHaveAttribute('tabindex', '0')
    expect(equip).toHaveAttribute('tabindex', '-1')

    dados.focus()
    await userEvent.keyboard('{ArrowRight}')
    await waitFor(() => expect(equip).toHaveAttribute('aria-selected', 'true'))
    expect(equip).toHaveFocus()
    expect(equip).toHaveAttribute('tabindex', '0')

    await userEvent.click(dados)
    await waitFor(() => expect(dados).toHaveAttribute('aria-selected', 'true'))
  })

  it('colaborador inexistente mostra "Colaborador não encontrado" sem tentar de novo', async () => {
    let chamadas = 0
    server.use(
      http.get('/api/colaboradores/999', () => {
        chamadas += 1
        return HttpResponse.json({ detail: 'Colaborador não encontrado' }, { status: 404 })
      })
    )
    renderPage('/colaboradores/999')
    expect(await screen.findByText('Colaborador não encontrado')).toBeInTheDocument()
    expect(screen.getByRole('link', { name: /Voltar aos colaboradores/ })).toHaveAttribute('href', '/colaboradores')
    expect(chamadas).toBe(1)
  })

  it('id não numérico mostra "Colaborador não encontrado" sem chamar a API', async () => {
    renderPage('/colaboradores/abc')
    expect(await screen.findByText('Colaborador não encontrado')).toBeInTheDocument()
  })
})

describe('ColaboradorDetailPage — paginação dos equipamentos', () => {
  function mockItens(total: () => number) {
    const calls: Array<{ offset: string | null; limit: string | null }> = []
    server.use(
      http.get('/api/colaboradores/7', () => HttpResponse.json(maria)),
      http.get('/api/items', ({ request }) => {
        const sp = new URL(request.url).searchParams
        const offset = sp.get('offset')
        calls.push({ offset, limit: sp.get('limit') })
        const t = total()
        const start = Number(offset)
        return HttpResponse.json({ items: start < t ? [equipamento(start + 1)] : [], total: t })
      })
    )
    return calls
  }

  it('envia limit/offset, avança de página e volta ao offset 0 ao trocar o tamanho', async () => {
    const calls = mockItens(() => 120)
    renderPage('/colaboradores/7?aba=equipamentos')

    await screen.findByRole('link', { name: /Latitude 1$/ })
    expect(calls[0]).toEqual({ offset: '0', limit: '25' })

    await userEvent.click(screen.getByRole('button', { name: 'Próxima' }))
    await screen.findByRole('link', { name: /Latitude 26$/ })
    expect(calls.at(-1)).toEqual({ offset: '25', limit: '25' })

    const antes = calls.length
    await userEvent.selectOptions(screen.getByLabelText('Itens por página'), '50')
    await screen.findByRole('link', { name: /Latitude 1$/ })
    expect(calls.slice(antes)).toEqual([{ offset: '0', limit: '50' }])
  })

  it('quando o total encolhe e a página passa da última, consulta com o offset ajustado', async () => {
    let total = 60
    const calls = mockItens(() => total)
    const { client } = renderPage('/colaboradores/7?aba=equipamentos')

    await screen.findByRole('link', { name: /Latitude 1$/ })
    await userEvent.click(screen.getByRole('button', { name: 'Próxima' }))
    await userEvent.click(screen.getByRole('button', { name: 'Próxima' }))
    await screen.findByRole('link', { name: /Latitude 51$/ })
    expect(calls.at(-1)).toEqual({ offset: '50', limit: '25' })

    total = 30
    await act(() => client.invalidateQueries({ queryKey: ['colaborador-equipamentos'] }))
    await screen.findByRole('link', { name: /Latitude 26$/ })
    expect(calls.at(-1)).toEqual({ offset: '25', limit: '25' })
    expect(screen.getByText('2 / 2')).toBeInTheDocument()
  })
})

describe('registry do colaborador', () => {
  const Vazio = () => null
  const tab = (over: Partial<ColaboradorDetailTab>): ColaboradorDetailTab => ({
    id: 'x',
    label: 'X',
    order: 10,
    Component: Vazio,
    ...over,
  })

  it('filtra por papel e ordena por order', () => {
    const mods = {
      './tabs/a.tsx': { default: tab({ id: 'a', order: 30 }) },
      './tabs/b.tsx': { default: tab({ id: 'b', order: 20, roles: ['Gestor'] }) },
      './tabs/c.tsx': { default: tab({ id: 'c', order: 10 }) },
    }
    expect(buildColaboradorDetailTabs(mods, 'Técnico').map((t) => t.id)).toEqual(['c', 'a'])
    expect(buildColaboradorDetailTabs(mods, 'Gestor').map((t) => t.id)).toEqual(['c', 'b', 'a'])
    expect(buildColaboradorDetailTabs(mods, undefined).map((t) => t.id)).toEqual(['c', 'a'])
  })

  it('ignora módulos sem default válido', () => {
    const mods = {
      './tabs/a.tsx': { default: tab({ id: 'a', order: 20 }) },
      './tabs/helper.tsx': {},
      './tabs/b.tsx': { default: { foo: 1 } },
      './tabs/c.tsx': { default: tab({ id: 'c', order: 10 }) },
    }
    expect(buildColaboradorDetailTabs(mods, 'Gestor').map((t) => t.id)).toEqual(['c', 'a'])
  })

  it('avisa e ignora ids duplicados', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const mods = {
      './tabs/a.tsx': { default: tab({ id: 'x', order: 10 }) },
      './tabs/b.tsx': { default: tab({ id: 'x', order: 20 }) },
    }
    expect(buildColaboradorDetailTabs(mods, 'Gestor')).toHaveLength(1)
    expect(warn).toHaveBeenCalledWith(expect.stringContaining('duplicado'))
    warn.mockRestore()
  })
})
