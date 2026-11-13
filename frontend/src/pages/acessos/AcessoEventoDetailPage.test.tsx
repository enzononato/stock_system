import { beforeEach, describe, expect, it, vi } from 'vitest'
import { delay, http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import { AVISO_SEM_SENHA } from '@/components/acessos/AvisoSemSenha'
import AcessoEventoDetailPage from './AcessoEventoDetailPage'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 2, username: 'tecnico', role: 'Técnico' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Técnico'),
  }),
}))

type Obj = Record<string, unknown>

const base: Obj = {
  id: 5,
  colaborador_id: 7,
  colaborador_nome: 'Maria Silva',
  colaborador_situacao: 'ativo',
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
  total_tarefas: 2,
  tarefas_pendentes: 2,
  equipamentos: [],
  pode_concluir: false,
  pendencias: ['2 tarefas pendentes'],
}

function tarefa(id: number, sistema: string, acao: string, extra: Obj = {}): Obj {
  return {
    id, evento_id: 5, sistema_id: id, sistema_nome: sistema, acao, status: 'pendente',
    login: null, acesso_id: null, feito_por: null, feito_em: null, observacao: null, ...extra,
  }
}

const admissao: Obj = { ...base, tarefas: [tarefa(1, 'AD', 'criar'), tarefa(2, 'E-mail', 'criar')] }

const equipamento = {
  id: 11, tipo: 'Notebook', brand: 'Dell', model: 'X', identificador: 'NB-1',
  status: 'Indisponível', revenda: 'Revalle Juazeiro',
}

let atual: Obj = admissao

function usarEvento(evento: Obj) {
  atual = evento
  server.use(http.get('/api/acessos/eventos/5', () => HttpResponse.json(atual)))
}

function renderPagina(id = 5) {
  return renderWithClient(
    <MemoryRouter initialEntries={[`/acessos/eventos/${id}`]}>
      <Routes>
        <Route path="/acessos/eventos/:id" element={<AcessoEventoDetailPage />} />
        <Route path="/acessos" element={<p>Lista de eventos</p>} />
      </Routes>
      <ToastContainer />
    </MemoryRouter>
  )
}

beforeEach(() => {
  usarEvento(admissao)
  server.use(
    http.get('/api/acessos/colaboradores/7/equipamentos', () => HttpResponse.json([equipamento]))
  )
})

describe('AcessoEventoDetailPage', () => {
  it('mostra a admissão aberta com as tarefas pendentes', async () => {
    renderPagina()
    expect(await screen.findByRole('heading', { level: 1, name: 'Admissão' })).toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'MARIA SILVA' })).toHaveAttribute('href', '/colaboradores/7')
    expect(screen.getByRole('link', { name: 'Voltar aos eventos' })).toHaveAttribute('href', '/acessos')
    expect(screen.getByText('Vendedor')).toBeInTheDocument()
    const tarefas = within(screen.getByRole('list', { name: 'Tarefas' }))
    expect(tarefas.getAllByText('Pendente')).toHaveLength(2)
    expect(tarefas.getByText('AD')).toBeInTheDocument()
    expect(tarefas.getByText('E-mail')).toBeInTheDocument()
    expect(screen.getByText(/Aberto por tecnico em/)).toBeInTheDocument()
    expect(screen.getByRole('status')).toHaveTextContent('2 tarefas pendentes')
    expect(screen.getByRole('button', { name: 'Concluir evento' })).toBeDisabled()
  })

  it('marca a criação como feita com login e equipamento, sem campo de senha', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/acessos/tarefas/1/resolver', async ({ request }) => {
        corpo = await request.json()
        atual = {
          ...admissao,
          tarefas_pendentes: 1,
          pendencias: ['1 tarefa pendente'],
          tarefas: [
            tarefa(1, 'AD', 'criar', { status: 'feito', login: 'maria.silva', feito_por: 'tecnico', feito_em: '2026-11-10T10:00:00' }),
            tarefa(2, 'E-mail', 'criar'),
          ],
        }
        return HttpResponse.json(atual)
      })
    )
    const user = userEvent.setup()
    renderPagina()
    const linhaAd = (await screen.findByText('AD')).closest('li') as HTMLElement
    await user.click(within(linhaAd).getByRole('button', { name: 'Marcar como feito' }))

    const dialogo = await screen.findByRole('dialog')
    expect(within(dialogo).getByText(AVISO_SEM_SENHA)).toBeInTheDocument()
    expect(within(dialogo).queryByLabelText(/senha/i)).toBeNull()

    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar' }))
    expect(await within(dialogo).findByText('Informe o login.')).toBeInTheDocument()
    expect(corpo).toBeNull()

    await user.type(within(dialogo).getByLabelText('Login'), 'maria.silva')
    await within(dialogo).findByRole('option', { name: 'Notebook Dell X · NB-1' })
    await user.selectOptions(within(dialogo).getByLabelText('Equipamento (opcional)'), '11')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar' }))

    await waitFor(() => expect(corpo).toEqual({ status: 'feito', login: 'maria.silva', item_id: 11 }))
    expect(await screen.findByText('Feito')).toBeInTheDocument()
    expect(await screen.findByText('Tarefa atualizada.')).toBeInTheDocument()
    expect(screen.getByText('maria.silva')).toBeInTheDocument()
  })

  it('dois cliques rápidos em Confirmar geram uma só requisição', async () => {
    let chamadas = 0
    server.use(
      http.post('/api/acessos/tarefas/1/resolver', async () => {
        chamadas += 1
        await delay(80)
        return HttpResponse.json(admissao)
      })
    )
    const user = userEvent.setup()
    renderPagina()
    const linhaAd = (await screen.findByText('AD')).closest('li') as HTMLElement
    await user.click(within(linhaAd).getByRole('button', { name: 'Marcar como feito' }))
    const dialogo = await screen.findByRole('dialog')
    await user.type(within(dialogo).getByLabelText('Login'), 'maria.silva')
    await user.dblClick(within(dialogo).getByRole('button', { name: 'Confirmar' }))
    await waitFor(() => expect(chamadas).toBeGreaterThanOrEqual(1))
    await delay(150)
    expect(chamadas).toBe(1)
  })

  it('"Não se aplica" exige observação e a envia', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/acessos/tarefas/2/resolver', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json(admissao)
      })
    )
    const user = userEvent.setup()
    renderPagina()
    const linha = (await screen.findByText('E-mail')).closest('li') as HTMLElement
    await user.click(within(linha).getByRole('button', { name: 'Não se aplica' }))
    const dialogo = await screen.findByRole('dialog')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar' }))
    expect(await within(dialogo).findByText('Informe a observação.')).toBeInTheDocument()
    await user.type(within(dialogo).getByLabelText('Observação'), 'Sem licença disponível')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar' }))
    await waitFor(() =>
      expect(corpo).toEqual({ status: 'nao_se_aplica', observacao: 'Sem licença disponível' })
    )
  })

  it('mudança de função: desativar oferece "Manter acesso" no lugar de "Não se aplica"', async () => {
    usarEvento({
      ...base,
      tipo: 'mudanca_funcao',
      cargo_anterior: 'Vendedor',
      cargo_novo: 'Supervisor',
      total_tarefas: 2,
      tarefas: [
        tarefa(3, 'VPN', 'criar'),
        tarefa(4, 'Promax', 'desativar', { login: 'MSILVA', acesso_id: 9 }),
      ],
    })
    let corpo: unknown = null
    server.use(
      http.post('/api/acessos/tarefas/4/resolver', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json(atual)
      })
    )
    const user = userEvent.setup()
    renderPagina()
    expect(await screen.findByText('Vendedor → Supervisor')).toBeInTheDocument()
    const linhaPromax = screen.getByText('Promax').closest('li') as HTMLElement
    expect(within(linhaPromax).queryByRole('button', { name: 'Não se aplica' })).toBeNull()
    expect(within(linhaPromax).getByRole('button', { name: 'Marcar como desativado' })).toBeInTheDocument()
    await user.click(within(linhaPromax).getByRole('button', { name: 'Manter acesso' }))
    const dialogo = await screen.findByRole('dialog')
    await user.type(within(dialogo).getByLabelText('Motivo'), 'continua na rota')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar' }))
    await waitFor(() => expect(corpo).toEqual({ status: 'manter', observacao: 'continua na rota' }))
  })

  it('desligamento: confirma a desativação e lista os equipamentos com atalho para a devolução', async () => {
    usarEvento({
      ...base,
      tipo: 'desligamento',
      cargo_anterior: 'Vendedor',
      cargo_novo: null,
      total_tarefas: 1,
      tarefas_pendentes: 1,
      tarefas: [tarefa(6, 'AD', 'desativar', { login: 'maria.silva', acesso_id: 3 })],
      equipamentos: [equipamento],
      pendencias: ['1 tarefa pendente', '1 equipamento ainda com o colaborador'],
    })
    let corpo: unknown = null
    server.use(
      http.post('/api/acessos/tarefas/6/resolver', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json(atual)
      })
    )
    const user = userEvent.setup()
    renderPagina()
    expect(await screen.findByRole('heading', { level: 1, name: 'Desligamento' })).toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'Notebook Dell X · NB-1' })).toHaveAttribute('href', '/items/11')
    expect(screen.getByText('REVALLE JUAZEIRO')).toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'Ir para Devolução' })).toHaveAttribute('href', '/return')
    expect(screen.getByRole('status')).toHaveTextContent('1 equipamento ainda com o colaborador')
    expect(screen.getByRole('button', { name: 'Concluir evento' })).toBeDisabled()

    await user.click(screen.getByRole('button', { name: 'Marcar como desativado' }))
    const dialogo = await screen.findByRole('dialog')
    expect(
      within(dialogo).getByText('Confirme que o acesso "maria.silva" no sistema AD foi desativado.')
    ).toBeInTheDocument()
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar desativação' }))
    await waitFor(() => expect(corpo).toEqual({ status: 'feito' }))
  })

  it('desligamento sem equipamentos mostra a frase de vazio', async () => {
    usarEvento({
      ...base, tipo: 'desligamento', total_tarefas: 0, tarefas_pendentes: 0, tarefas: [],
      pode_concluir: true, pendencias: [],
    })
    renderPagina()
    expect(await screen.findByText('Nenhum equipamento com o colaborador.')).toBeInTheDocument()
    expect(screen.queryByRole('link', { name: 'Ir para Devolução' })).toBeNull()
  })

  it('conclui o evento quando pode_concluir', async () => {
    usarEvento({ ...admissao, tarefas_pendentes: 0, pode_concluir: true, pendencias: [] })
    let chamado = false
    server.use(
      http.post('/api/acessos/eventos/5/concluir', () => {
        chamado = true
        return HttpResponse.json({ ...atual, status: 'concluido', concluido_por: 'tecnico', concluido_em: '2026-11-10T11:00:00' })
      })
    )
    const user = userEvent.setup()
    renderPagina()
    const botao = await screen.findByRole('button', { name: 'Concluir evento' })
    expect(botao).toBeEnabled()
    await user.click(botao)
    await waitFor(() => expect(chamado).toBe(true))
    expect(await screen.findByText('Evento concluído.')).toBeInTheDocument()
  })

  it('cancela o evento com o motivo', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/acessos/eventos/5/cancelar', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...atual, status: 'cancelado' })
      })
    )
    const user = userEvent.setup()
    renderPagina()
    await user.click(await screen.findByRole('button', { name: 'Cancelar evento' }))
    const dialogo = await screen.findByRole('dialog')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar cancelamento' }))
    expect(await within(dialogo).findByText('Informe o motivo do cancelamento.')).toBeInTheDocument()
    await user.type(within(dialogo).getByLabelText('Motivo do cancelamento'), 'Aberto por engano')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar cancelamento' }))
    await waitFor(() => expect(corpo).toEqual({ motivo_cancelamento: 'Aberto por engano' }))
    expect(await screen.findByText('Evento cancelado.')).toBeInTheDocument()
  })

  it('evento concluído não mostra botões de ação', async () => {
    usarEvento({
      ...admissao,
      status: 'concluido',
      concluido_por: 'tecnico',
      concluido_em: '2026-11-10T11:00:00',
      tarefas_pendentes: 0,
      pendencias: [],
      tarefas: [
        tarefa(1, 'AD', 'criar', { status: 'feito', login: 'maria.silva' }),
        tarefa(2, 'E-mail', 'criar', { status: 'nao_se_aplica', observacao: 'Sem licença' }),
      ],
    })
    renderPagina()
    expect(await screen.findByText('Sem licença')).toBeInTheDocument()
    expect(screen.getByText('Concluído')).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Concluir evento' })).toBeNull()
    expect(screen.queryByRole('button', { name: 'Cancelar evento' })).toBeNull()
    expect(screen.queryByRole('button', { name: 'Marcar como feito' })).toBeNull()
  })

  it('evento inexistente mostra "Evento não encontrado"', async () => {
    server.use(
      http.get('/api/acessos/eventos/5', () =>
        HttpResponse.json({ detail: 'Evento não encontrado.' }, { status: 404 })
      )
    )
    renderPagina()
    expect(await screen.findByText('Evento não encontrado')).toBeInTheDocument()
  })
})
