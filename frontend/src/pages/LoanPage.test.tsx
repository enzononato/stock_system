import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import LoanPage from '@/pages/LoanPage'
import type { Colaborador } from '@/api/colaboradores'

// Assistente de 4 passos (T6): Equipamento → Colaborador → Condições →
// Confirmação. Estes testes cobrem a lógica do assistente (validação por
// etapa, navegação, preenchimento automático de revenda, colaborador do
// cadastro com e-mail obrigatório e o padrão do toggle "É pessoa jurídica").

// Títulos das etapas seguem o texto visual da FT_STC (antes: "1. Equipamento",
// "2. Colaborador", "3. Condições") — só o texto mudou, o comportamento testado é o mesmo.
const TITULO_ETAPA_1 = '1. Seleção do Equipamento'
const TITULO_ETAPA_2 = '2. Dados do Colaborador Responsável'
const TITULO_ETAPA_3 = '3. Condições Operacionais e Alocação'

const itemDisponivel = {
  id: 42,
  tipo: 'Notebook',
  brand: 'Dell',
  model: 'Latitude',
  identificador: 'PAT-001',
  revenda: 'Matriz',
  status: 'Disponível',
}

function colaborador(extra: Partial<Colaborador> = {}): Colaborador {
  return {
    id: 7,
    cpf: '52998224725',
    nome: 'Maria Souza',
    cargo: 'Analista',
    setor: 'Financeiro',
    unidade_nome: 'Outra Unidade',
    situacao: 'ativo',
    email: 'maria@gmail.com',
    email_manual: true,
    telefone_manual: false,
    origem: 'rh',
    ...extra,
  }
}

/** Mocka `/api/items`: devolve o único item disponível para qualquer busca que não seja por "Pendente" (usada pela lista de pendentes de confirmação). */
function mockListItems(pendentes: unknown[] = [], paramsPendentes?: (p: URLSearchParams) => void) {
  server.use(
    http.get('/api/items', ({ request }) => {
      const url = new URL(request.url)
      if (url.searchParams.get('status') === 'Pendente') {
        paramsPendentes?.(url.searchParams)
        return HttpResponse.json({ items: pendentes, total: pendentes.length })
      }
      return HttpResponse.json({ items: [itemDisponivel], total: 1 })
    })
  )
}

function mockColaboradores(lista: Colaborador[]) {
  server.use(http.get('/api/colaboradores', () => HttpResponse.json({ items: lista, total: lista.length })))
}

function renderPage() {
  return renderWithClient(
    <>
      <LoanPage />
      <ToastContainer />
    </>
  )
}

/** Etapa 1: abre o SearchableSelect de equipamento e escolhe o único item disponível. */
async function selecionarEquipamento(user: ReturnType<typeof userEvent.setup>) {
  // Placeholder do seletor segue o texto da FT_STC ("Buscar por ID, modelo, tipo ou serial…").
  await user.click(await screen.findByRole('button', { name: /Buscar por ID, modelo, tipo ou serial/i }))
  await user.click(await screen.findByRole('button', { name: /Notebook Dell Latitude/i }))
}

/** Avança da etapa 1 (Equipamento) para a etapa 2 (Colaborador), já com o equipamento escolhido. */
async function avancarParaColaborador(user: ReturnType<typeof userEvent.setup>) {
  await selecionarEquipamento(user)
  await user.click(screen.getByRole('button', { name: /Próximo/i }))
  await screen.findByRole('heading', { name: TITULO_ETAPA_2 })
}

/** Etapa 2: abre o seletor de colaborador, busca e escolhe pelo nome. */
async function escolherColaborador(user: ReturnType<typeof userEvent.setup>, nome: string) {
  await user.click(await screen.findByRole('button', { name: /Selecione um colaborador/i }))
  await user.type(screen.getByPlaceholderText(/Buscar por nome, CPF ou matrícula/i), nome.slice(0, 3))
  await user.click(await screen.findByText(nome.toLocaleUpperCase('pt-BR')))
}

/** Avança da etapa 2 (Colaborador) para a etapa 3 (Condições), com o colaborador do cadastro. */
async function avancarParaCondicoes(user: ReturnType<typeof userEvent.setup>, c = colaborador()) {
  mockColaboradores([c])
  await avancarParaColaborador(user)
  await escolherColaborador(user, c.nome)
  await user.click(screen.getByRole('button', { name: /Próximo/i }))
  await screen.findByRole('heading', { name: TITULO_ETAPA_3 })
}

describe('LoanPage — assistente de 4 passos', () => {
  it('não avança da etapa Equipamento sem selecionar um item disponível', async () => {
    mockListItems()
    const user = userEvent.setup()
    renderPage()

    await screen.findByRole('heading', { name: TITULO_ETAPA_1 })
    await user.click(screen.getByRole('button', { name: /Próximo/i }))

    expect(await screen.findByText('Selecione um equipamento disponível.')).toBeInTheDocument()
    expect(screen.getByRole('heading', { name: TITULO_ETAPA_1 })).toBeInTheDocument()
  })

  // MUDANÇA INTENCIONAL F3: substitui o teste "CPF inválido barra o avanço no
  // passo 2". O CPF deixou de ser digitado: vem do cadastro de colaboradores, que
  // já o valida (F0). O que barra o avanço agora é não ter escolhido ninguém.
  it('não avança da etapa Colaborador sem escolher alguém do cadastro', async () => {
    mockListItems()
    mockColaboradores([colaborador()])
    const user = userEvent.setup()
    renderPage()

    await avancarParaColaborador(user)
    await user.click(screen.getByRole('button', { name: /Próximo/i }))

    expect(await screen.findByText('Selecione o colaborador.')).toBeInTheDocument()
    expect(screen.getByRole('heading', { name: TITULO_ETAPA_2 })).toBeInTheDocument()
  })

  it('exige e-mail válido e grava o e-mail no cadastro antes de seguir', async () => {
    mockListItems()
    const semEmail = colaborador({ id: 8, nome: 'João Silva', email: null, email_manual: false })
    mockColaboradores([semEmail])
    let corpoPut: unknown = null
    server.use(
      http.put('/api/colaboradores/8', async ({ request }) => {
        corpoPut = await request.json()
        return HttpResponse.json({ ...semEmail, email: 'joao.silva@gmail.com', email_manual: true })
      })
    )
    const user = userEvent.setup()
    renderPage()

    await avancarParaColaborador(user)
    await escolherColaborador(user, 'João Silva')
    const campo = screen.getByLabelText(/E-mail do colaborador/i)
    await user.type(campo, 'joao@')
    await user.click(screen.getByRole('button', { name: /Próximo/i }))
    expect(await screen.findByText('E-mail inválido.')).toBeInTheDocument()
    expect(corpoPut).toBeNull()

    await user.clear(campo)
    await user.type(campo, 'joao.silva@gmail.com')
    await user.click(screen.getByRole('button', { name: /Próximo/i }))

    expect(await screen.findByRole('heading', { name: TITULO_ETAPA_3 })).toBeInTheDocument()
    expect(corpoPut).toEqual({ email: 'joao.silva@gmail.com' })
  })

  it('exige informar o e-mail quando o cadastro não tem', async () => {
    mockListItems()
    mockColaboradores([colaborador({ id: 8, nome: 'João Silva', email: null })])
    const user = userEvent.setup()
    renderPage()

    await avancarParaColaborador(user)
    await escolherColaborador(user, 'João Silva')
    await user.click(screen.getByRole('button', { name: /Próximo/i }))

    expect(await screen.findByText('Informe o e-mail do colaborador.')).toBeInTheDocument()
    expect(screen.getByRole('heading', { name: TITULO_ETAPA_2 })).toBeInTheDocument()
  })

  it('preenche a revenda a partir do item e barra sem setor quando o cadastro não tem setor', async () => {
    mockListItems()
    const user = userEvent.setup()
    renderPage()

    await avancarParaCondicoes(user, colaborador({ setor: null }))
    await user.click(screen.getByRole('button', { name: /Próximo/i }))

    // A revenda já veio preenchida do item selecionado (Matriz) — só o setor
    // barra o avanço, confirmando que o preenchimento automático ocorreu.
    expect(await screen.findByText('Selecione o setor.')).toBeInTheDocument()
    expect(screen.queryByText('Selecione a revenda.')).not.toBeInTheDocument()
    expect(screen.getByRole('heading', { name: TITULO_ETAPA_3 })).toBeInTheDocument()
  })

  it('usa setor e cargo do cadastro e envia colaborador_id', async () => {
    mockListItems()
    let corpo: Record<string, unknown> | null = null
    server.use(
      http.post('/api/loans', async ({ request }) => {
        corpo = (await request.json()) as Record<string, unknown>
        return HttpResponse.json({
          detail: 'ok',
          termo: { id: 5, item_id: 42, tipo: 'entrega', status: 'enviado', destinatario_email: 'maria@gmail.com' },
        })
      })
    )
    const user = userEvent.setup()
    renderPage()

    await avancarParaCondicoes(user)
    expect(screen.getAllByText('Do cadastro do colaborador').length).toBe(2)
    await user.click(screen.getByRole('button', { name: /Próximo/i }))
    await screen.findByText('E-mail para assinatura')
    expect(screen.getByText('maria@gmail.com')).toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: /Confirmar Empréstimo/i }))

    await waitFor(() => expect(corpo).not.toBeNull())
    expect(corpo).toMatchObject({
      item_id: 42,
      colaborador_id: 7,
      setor: 'Financeiro',
      cargo: 'Analista',
      revenda: 'Matriz',
    })
    expect(corpo).not.toHaveProperty('usuario')
    expect(corpo).not.toHaveProperty('cpf')
    expect(
      await screen.findByText('Empréstimo iniciado. Termo enviado para assinatura em maria@gmail.com.')
    ).toBeInTheDocument()
    expect(screen.queryByRole('heading', { name: /Confirmar Empréstimo — Item/ })).not.toBeInTheDocument()
  })

  it('com assinatura manual abre o painel de confirmação', async () => {
    mockListItems()
    server.use(
      http.post('/api/loans', () =>
        HttpResponse.json({
          detail: 'ok',
          termo: { id: 5, item_id: 42, tipo: 'entrega', status: 'manual', motivo: 'Integração desligada.' },
        })
      )
    )
    const user = userEvent.setup()
    renderPage()

    await avancarParaCondicoes(user)
    await user.click(screen.getByRole('button', { name: /Próximo/i }))
    await user.click(await screen.findByRole('button', { name: /Confirmar Empréstimo/i }))

    const painel = await screen.findByRole('heading', { name: /Confirmar Empréstimo — Item/ })
    expect(painel).toHaveTextContent('Confirmar Empréstimo — Item #42')
  })

  it('pendentes de termo usam a paginação padrão', async () => {
    let params: URLSearchParams | null = null
    mockListItems(
      [{ id: 9, tipo: 'Celular', brand: 'Samsung', assigned_to: 'Ana Lima', revenda: 'Revalle Juazeiro', status: 'Pendente' }],
      (p) => {
        params = p
      }
    )
    renderPage()

    expect(await screen.findByText('Mostrando 1–1 de 1')).toBeInTheDocument()
    expect(params).not.toBeNull()
    expect((params as unknown as URLSearchParams).get('limit')).toBe('25')
    expect(screen.getByText('REVALLE JUAZEIRO')).toBeInTheDocument()
  })

  it('permite voltar a uma etapa já concluída, mas não pular para uma etapa futura', async () => {
    mockListItems()
    const user = userEvent.setup()
    renderPage()

    await avancarParaCondicoes(user)

    // Volta livre: "Equipamento" (etapa concluída) está clicável no Stepper.
    await user.click(screen.getByRole('button', { name: /Equipamento/i }))
    expect(await screen.findByRole('heading', { name: TITULO_ETAPA_1 })).toBeInTheDocument()

    // Sem avanço: de volta à etapa 1, "Condições" e "Confirmação" (etapas
    // futuras, ainda não concluídas) ficam desabilitadas no Stepper.
    const botaoCondicoes = screen.getByRole('button', { name: /Condições/i })
    const botaoConfirmacao = screen.getByRole('button', { name: /Confirmação/i })
    expect(botaoCondicoes).toBeDisabled()
    expect(botaoConfirmacao).toBeDisabled()

    await user.click(botaoCondicoes)
    expect(screen.getByRole('heading', { name: TITULO_ETAPA_1 })).toBeInTheDocument()
  })

  it('o toggle "É pessoa jurídica" continua desmarcado por padrão', async () => {
    mockListItems()
    mockColaboradores([colaborador()])
    const user = userEvent.setup()
    renderPage()

    await avancarParaColaborador(user)

    expect(screen.getByRole('checkbox')).not.toBeChecked()
  })

  it('o rótulo "Colaborador *" nomeia o seletor da etapa 2', async () => {
    mockListItems()
    mockColaboradores([colaborador()])
    const user = userEvent.setup()
    renderPage()

    await avancarParaColaborador(user)

    expect(screen.getByRole('group', { name: 'Colaborador *' })).toBeInTheDocument()
  })
})
