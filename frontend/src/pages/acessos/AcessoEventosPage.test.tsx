import { beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import { getVisibleNavItems } from '@/components/layout/navigation'
import { AVISO_SEM_SENHA } from '@/components/acessos/AvisoSemSenha'
import type { Colaborador } from '@/api/colaboradores'
import AcessoEventosPage from './AcessoEventosPage'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 2, username: 'tecnico', role: 'Técnico' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Técnico'),
  }),
}))

const maria: Colaborador = {
  id: 7,
  cpf: '11144477735',
  nome: 'Maria Silva',
  cargo: 'Supervisor',
  situacao: 'ativo',
  email_manual: false,
  telefone_manual: false,
  origem: 'rh',
}

const evento = {
  id: 5,
  colaborador_id: 7,
  colaborador_nome: 'Maria Silva',
  tipo: 'admissao',
  cargo_anterior: null,
  cargo_novo: 'Vendedor',
  status: 'aberto',
  observacao: null,
  aberto_por: 'tecnico',
  aberto_em: '2026-11-10T09:30:00',
  concluido_por: null,
  concluido_em: null,
  cancelado_por: null,
  cancelado_em: null,
  motivo_cancelamento: null,
  total_tarefas: 3,
  tarefas_pendentes: 2,
}

let consultas: URLSearchParams[] = []
let total = 1

function renderPagina() {
  return renderWithClient(
    <MemoryRouter initialEntries={['/acessos']}>
      <Routes>
        <Route path="/acessos" element={<AcessoEventosPage />} />
        <Route path="/acessos/eventos/:id" element={<p>Detalhe do evento aberto</p>} />
      </Routes>
      <ToastContainer />
    </MemoryRouter>
  )
}

beforeEach(() => {
  consultas = []
  total = 1
  server.use(
    http.get('/api/acessos/eventos', ({ request }) => {
      consultas.push(new URL(request.url).searchParams)
      return HttpResponse.json({ items: [evento], total })
    }),
    http.get('/api/colaboradores/cargos', () => HttpResponse.json(['Supervisor', 'Vendedor'])),
    http.get('/api/colaboradores', () => HttpResponse.json({ items: [maria], total: 1 }))
  )
})

describe('AcessoEventosPage', () => {
  it('lista os eventos abertos por padrão', async () => {
    renderPagina()
    const linha = (await screen.findByText('MARIA SILVA')).closest('tr') as HTMLElement
    expect(within(linha).getByText('Admissão')).toBeInTheDocument()
    expect(within(linha).getByText('Aberto')).toBeInTheDocument()
    expect(within(linha).getByText('1/3 resolvidas')).toBeInTheDocument()
    expect(within(linha).getByText('Vendedor')).toBeInTheDocument()
    expect(within(linha).getByRole('link', { name: 'Abrir evento #5' })).toHaveAttribute(
      'href',
      '/acessos/eventos/5'
    )
    expect(consultas[0].get('status')).toBe('aberto')
    expect(consultas[0].get('limit')).toBe('25')
    expect(consultas[0].get('offset')).toBe('0')
  })

  it('"Todos" refaz a consulta sem status e voltando ao início', async () => {
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('MARIA SILVA')
    await user.selectOptions(screen.getByLabelText('Situação'), 'todos')
    await waitFor(() => expect(consultas.some((c) => !c.has('status'))).toBe(true))
    const ultima = consultas[consultas.length - 1]
    expect(ultima.has('status')).toBe(false)
    expect(ultima.get('offset')).toBe('0')
  })

  it('filtra por tipo e pela busca (com debounce)', async () => {
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('MARIA SILVA')
    await user.selectOptions(screen.getByLabelText('Tipo'), 'desligamento')
    await waitFor(() => expect(consultas.some((c) => c.get('tipo') === 'desligamento')).toBe(true))
    await user.type(screen.getByLabelText('Buscar por nome ou CPF'), 'maria')
    await waitFor(() => expect(consultas.some((c) => c.get('busca') === 'maria')).toBe(true))
    expect(consultas.some((c) => c.get('busca') === 'mar')).toBe(false)
  })

  it('pagina pelo servidor', async () => {
    total = 60
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('MARIA SILVA')
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    await waitFor(() => expect(consultas.some((c) => c.get('offset') === '25')).toBe(true))
  })

  it('sem eventos mostra a mensagem de vazio', async () => {
    server.use(http.get('/api/acessos/eventos', () => HttpResponse.json({ items: [], total: 0 })))
    renderPagina()
    expect(await screen.findByText('Nenhum evento encontrado.')).toBeInTheDocument()
  })

  it('registra uma mudança de função e abre o evento criado', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/acessos/eventos', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...evento, id: 12, tipo: 'mudanca_funcao' }, { status: 201 })
      })
    )
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('MARIA SILVA')
    await user.click(screen.getByRole('button', { name: 'Registrar evento' }))
    const dialogo = await screen.findByRole('dialog')
    expect(within(dialogo).getByText(AVISO_SEM_SENHA)).toBeInTheDocument()

    await user.selectOptions(within(dialogo).getByLabelText('Tipo de evento'), 'mudanca_funcao')
    await user.click(within(dialogo).getByRole('button', { name: /selecione um colaborador/i }))
    await user.click(await within(dialogo).findByText('MARIA SILVA'))

    const registrar = within(dialogo).getByRole('button', { name: 'Registrar evento' })
    expect(registrar).toBeDisabled() // falta o cargo anterior
    await user.selectOptions(within(dialogo).getByLabelText(/^Cargo anterior/), 'Vendedor')
    expect(within(dialogo).getByLabelText('Cargo novo')).toHaveValue('Supervisor')
    await user.click(registrar)

    await waitFor(() =>
      expect(corpo).toEqual({
        colaborador_id: 7,
        tipo: 'mudanca_funcao',
        cargo_anterior: 'Vendedor',
        cargo_novo: 'Supervisor',
        observacao: null,
      })
    )
    expect(await screen.findByText('Detalhe do evento aberto')).toBeInTheDocument()
  })

  it('mostra o erro da API ao registrar um evento duplicado', async () => {
    server.use(
      http.post('/api/acessos/eventos', () =>
        HttpResponse.json(
          { detail: 'Já existe um evento de acesso aberto para este colaborador.' },
          { status: 409 }
        )
      )
    )
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('MARIA SILVA')
    await user.click(screen.getByRole('button', { name: 'Registrar evento' }))
    const dialogo = await screen.findByRole('dialog')
    await user.click(within(dialogo).getByRole('button', { name: /selecione um colaborador/i }))
    await user.click(await within(dialogo).findByText('MARIA SILVA'))
    await user.click(within(dialogo).getByRole('button', { name: 'Registrar evento' }))
    expect(
      await screen.findByText('Já existe um evento de acesso aberto para este colaborador.')
    ).toBeInTheDocument()
  })

  it('desligamento manda os dois cargos nulos', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/acessos/eventos', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...evento, id: 13, tipo: 'desligamento' }, { status: 201 })
      })
    )
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('MARIA SILVA')
    await user.click(screen.getByRole('button', { name: 'Registrar evento' }))
    const dialogo = await screen.findByRole('dialog')
    await user.selectOptions(within(dialogo).getByLabelText('Tipo de evento'), 'desligamento')
    expect(within(dialogo).getByText(/Serão geradas tarefas para desativar todos os acessos/)).toBeInTheDocument()
    await user.click(within(dialogo).getByRole('button', { name: /selecione um colaborador/i }))
    await user.click(await within(dialogo).findByText('MARIA SILVA'))
    await user.type(within(dialogo).getByLabelText('Observação'), 'Pedido de demissão')
    await user.click(within(dialogo).getByRole('button', { name: 'Registrar evento' }))
    await waitFor(() =>
      expect(corpo).toEqual({
        colaborador_id: 7,
        tipo: 'desligamento',
        cargo_anterior: null,
        cargo_novo: null,
        observacao: 'Pedido de demissão',
      })
    )
  })

  it('o menu mostra Eventos para Técnico e não para Jovem Aprendiz', () => {
    expect(getVisibleNavItems('Técnico').map((i) => i.to)).toContain('/acessos')
    expect(getVisibleNavItems('Jovem Aprendiz').map((i) => i.to)).not.toContain('/acessos')
  })
})
