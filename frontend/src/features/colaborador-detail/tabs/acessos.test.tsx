import { beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import ColaboradorDetailPage from '@/pages/ColaboradorDetailPage'
import { getColaboradorDetailTabs } from '@/features/colaborador-detail/registry'
import { AVISO_SEM_SENHA } from '@/components/acessos/AvisoSemSenha'
import type { Colaborador } from '@/api/colaboradores'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 2, username: 'tecnico.teste', role: 'Técnico' },
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

const acessoAd = {
  id: 1,
  colaborador_id: 7,
  colaborador_nome: 'Maria Silva',
  sistema_id: 1,
  sistema_nome: 'AD',
  login: 'maria.silva',
  status: 'ativo',
  item_id: 11,
  item_descricao: 'Notebook Dell X · NB-1',
  criado_por: 'tecnico.teste',
  criado_em: '2026-11-10T09:30:00',
  desativado_por: null,
  desativado_em: null,
}

const acessoAntigo = {
  ...acessoAd,
  id: 2,
  sistema_id: 2,
  sistema_nome: 'E-mail',
  login: 'maria@revalle',
  status: 'desativado',
  item_id: null,
  item_descricao: null,
  desativado_por: 'gestor.teste',
  desativado_em: '2026-11-11T08:00:00',
}

const evento = {
  id: 5,
  colaborador_id: 7,
  colaborador_nome: 'Maria Silva',
  tipo: 'admissao',
  cargo_anterior: null,
  cargo_novo: 'Supervisor',
  status: 'aberto',
  observacao: null,
  aberto_por: 'tecnico.teste',
  aberto_em: '2026-11-10T09:00:00',
  concluido_por: null,
  concluido_em: null,
  cancelado_por: null,
  cancelado_em: null,
  motivo_cancelamento: null,
  total_tarefas: 2,
  tarefas_pendentes: 1,
}

const sistemas = [
  { id: 1, nome: 'AD', descricao: null, ativo: true },
  { id: 3, nome: 'VPN', descricao: null, ativo: true },
]

const equipamento = {
  id: 11, tipo: 'Notebook', brand: 'Dell', model: 'X', identificador: 'NB-1',
  status: 'Indisponível', revenda: null,
}

function renderPage(path = '/colaboradores/7?aba=acessos') {
  return renderWithClient(
    <MemoryRouter initialEntries={[path]}>
      <Routes>
        <Route path="/colaboradores/:id" element={<ColaboradorDetailPage />} />
        <Route path="/items/:id" element={<p>Detalhe do equipamento</p>} />
        <Route path="/acessos/eventos/:id" element={<p>Detalhe do evento aberto</p>} />
      </Routes>
      <ToastContainer />
    </MemoryRouter>
  )
}

beforeEach(() => {
  localStorage.clear()
  server.use(
    http.get('/api/colaboradores/7', () => HttpResponse.json(maria)),
    http.get('/api/colaboradores', () => HttpResponse.json({ items: [maria], total: 1 })),
    http.get('/api/acessos/ficha', () => HttpResponse.json([acessoAd, acessoAntigo])),
    http.get('/api/acessos/sistemas', () => HttpResponse.json(sistemas)),
    http.get('/api/acessos/colaboradores/7/equipamentos', () => HttpResponse.json([equipamento])),
    http.get('/api/acessos/eventos', () => HttpResponse.json({ items: [evento], total: 1 })),
    http.get('/api/colaboradores/cargos', () => HttpResponse.json(['Supervisor', 'Vendedor']))
  )
})

describe('aba Acessos do colaborador', () => {
  it('aparece depois de Equipamentos e mostra a ficha de acessos', async () => {
    renderPage()
    const painel = within(await screen.findByRole('tabpanel'))
    const abas = within(screen.getByRole('tablist')).getAllByRole('tab').map((t) => t.textContent)
    expect(abas.indexOf('Acessos')).toBeGreaterThan(abas.indexOf('Equipamentos'))

    const linha = (await painel.findByText('maria.silva')).closest('li') as HTMLElement
    expect(within(linha).getByText('AD')).toBeInTheDocument()
    expect(within(linha).getByText('Ativo')).toBeInTheDocument()
    expect(within(linha).getByRole('link', { name: 'Notebook Dell X · NB-1' })).toHaveAttribute(
      'href',
      '/items/11'
    )
    expect(within(linha).getByText(/Criado por tecnico.teste em/)).toBeInTheDocument()

    const antigo = painel.getByText('maria@revalle').closest('li') as HTMLElement
    expect(within(antigo).getByText('Desativado')).toBeInTheDocument()
    expect(within(antigo).getByText(/Desativado por gestor.teste em/)).toBeInTheDocument()
    expect(within(antigo).queryByRole('button', { name: /^Desativar/ })).toBeNull()
  })

  it('sem acessos mostra a frase de vazio', async () => {
    server.use(http.get('/api/acessos/ficha', () => HttpResponse.json([])))
    renderPage()
    expect(await screen.findByText('Nenhum acesso registrado para este colaborador.')).toBeInTheDocument()
  })

  it('consulta a ficha do colaborador', async () => {
    let query = ''
    server.use(
      http.get('/api/acessos/ficha', ({ request }) => {
        query = new URL(request.url).search
        return HttpResponse.json([])
      })
    )
    renderPage()
    await screen.findByText('Nenhum acesso registrado para este colaborador.')
    expect(query).toBe('?colaborador_id=7')
  })

  it('adiciona um acesso (só login, nunca senha)', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/acessos/ficha', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...acessoAd, id: 3 }, { status: 201 })
      })
    )
    const user = userEvent.setup()
    renderPage()
    await user.click(await screen.findByRole('button', { name: 'Adicionar acesso' }))
    const dialogo = await screen.findByRole('dialog')
    expect(within(dialogo).getByText(AVISO_SEM_SENHA)).toBeInTheDocument()
    expect(within(dialogo).queryByLabelText(/senha/i)).toBeNull()

    await within(dialogo).findByRole('option', { name: 'VPN' })
    await user.selectOptions(within(dialogo).getByLabelText('Sistema'), 'VPN')
    await user.type(within(dialogo).getByLabelText('Login'), 'msilva')
    await user.click(within(dialogo).getByRole('button', { name: 'Salvar' }))
    await waitFor(() =>
      expect(corpo).toEqual({ colaborador_id: 7, sistema_id: 3, login: 'msilva', item_id: null })
    )
    expect(await screen.findByText('Acesso registrado.')).toBeInTheDocument()
  })

  it('edita o login e o equipamento de um acesso', async () => {
    let corpo: unknown = null
    server.use(
      http.get('/api/acessos/ficha', () => HttpResponse.json([{ ...acessoAd, item_id: null, item_descricao: null }])),
      http.put('/api/acessos/ficha/1', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json(acessoAd)
      })
    )
    const user = userEvent.setup()
    renderPage()
    await user.click(await screen.findByRole('button', { name: /^Editar acesso AD/ }))
    const dialogo = await screen.findByRole('dialog')
    const login = within(dialogo).getByLabelText('Login')
    await user.clear(login)
    await user.type(login, 'm.silva')
    await within(dialogo).findByRole('option', { name: 'Notebook Dell X · NB-1' })
    await user.selectOptions(within(dialogo).getByLabelText('Equipamento (opcional)'), '11')
    await user.click(within(dialogo).getByRole('button', { name: 'Salvar' }))
    await waitFor(() => expect(corpo).toEqual({ login: 'm.silva', item_id: 11 }))
    expect(await screen.findByText('Acesso atualizado.')).toBeInTheDocument()
  })

  it('desativa um acesso depois de confirmar', async () => {
    let chamado = false
    server.use(
      http.post('/api/acessos/ficha/1/desativar', () => {
        chamado = true
        return HttpResponse.json({ ...acessoAd, status: 'desativado' })
      })
    )
    const user = userEvent.setup()
    renderPage()
    await user.click(await screen.findByRole('button', { name: /^Desativar acesso AD/ }))
    const confirmacao = await screen.findByRole('alertdialog')
    expect(
      within(confirmacao).getByText('Desativar o acesso "maria.silva" no sistema AD?')
    ).toBeInTheDocument()
    expect(chamado).toBe(false)
    await user.click(within(confirmacao).getByRole('button', { name: 'Desativar' }))
    await waitFor(() => expect(chamado).toBe(true))
    expect(await screen.findByText('Acesso desativado.')).toBeInTheDocument()
  })

  it('lista os eventos de acesso do colaborador e registra um novo', async () => {
    let query = ''
    let corpo: unknown = null
    server.use(
      http.get('/api/acessos/eventos', ({ request }) => {
        query = new URL(request.url).search
        return HttpResponse.json({ items: [evento], total: 1 })
      }),
      http.post('/api/acessos/eventos', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...evento, id: 9 }, { status: 201 })
      })
    )
    const user = userEvent.setup()
    renderPage()
    const painel = within(await screen.findByRole('tabpanel'))
    const itemEvento = (await painel.findByText('Admissão')).closest('li') as HTMLElement
    expect(within(itemEvento).getByText('Aberto')).toBeInTheDocument()
    expect(within(itemEvento).getByRole('link', { name: /Abrir evento #5/ })).toHaveAttribute(
      'href',
      '/acessos/eventos/5'
    )
    expect(query).toBe('?colaborador_id=7&limit=20&offset=0')

    await user.click(painel.getByRole('button', { name: 'Registrar evento' }))
    const dialogo = await screen.findByRole('dialog')
    // O colaborador já vem preenchido.
    expect(within(dialogo).getByText('MARIA SILVA')).toBeInTheDocument()
    await user.click(within(dialogo).getByRole('button', { name: 'Registrar evento' }))
    await waitFor(() =>
      expect(corpo).toEqual({
        colaborador_id: 7,
        tipo: 'admissao',
        cargo_anterior: null,
        cargo_novo: 'Supervisor',
        observacao: null,
      })
    )
    expect(await screen.findByText('Detalhe do evento aberto')).toBeInTheDocument()
  })

  it('Jovem Aprendiz não vê a aba', () => {
    expect(getColaboradorDetailTabs('Jovem Aprendiz').map((t) => t.id)).not.toContain('acessos')
    expect(getColaboradorDetailTabs('Técnico').map((t) => t.id)).toContain('acessos')
  })
})
