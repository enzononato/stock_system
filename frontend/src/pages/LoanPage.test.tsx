import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import LoanPage from '@/pages/LoanPage'

// Assistente de 4 passos (T6): Equipamento → Colaborador → Condições →
// Confirmação. Estes testes cobrem só a lógica do assistente (validação por
// etapa, navegação, preenchimento automático de revenda e o padrão do
// toggle "É pessoa jurídica") — a chamada de API, a validação de CPF em si e
// o painel de confirmação pós-empréstimo já têm cobertura própria em outros
// pontos da suíte (isValidCpf em utils, ConfirmacaoTermo embutido aqui).

const itemDisponivel = {
  id: 42,
  tipo: 'Notebook',
  brand: 'Dell',
  model: 'Latitude',
  identificador: 'PAT-001',
  revenda: 'Matriz',
  status: 'Disponível',
}

/** Mocka `/api/items`: devolve o único item disponível para qualquer busca que não seja por "Pendente" (usada pela lista de pendentes de confirmação). */
function mockListItems() {
  server.use(
    http.get('/api/items', ({ request }) => {
      const url = new URL(request.url)
      if (url.searchParams.get('status') === 'Pendente') {
        return HttpResponse.json({ items: [], total: 0 })
      }
      return HttpResponse.json({ items: [itemDisponivel], total: 1 })
    })
  )
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
  await user.click(await screen.findByRole('button', { name: /Selecione um equipamento disponível/i }))
  await user.click(await screen.findByRole('button', { name: /Notebook Dell Latitude/i }))
}

/** Avança da etapa 1 (Equipamento) para a etapa 2 (Colaborador), já com o equipamento escolhido. */
async function avancarParaColaborador(user: ReturnType<typeof userEvent.setup>) {
  await selecionarEquipamento(user)
  await user.click(screen.getByRole('button', { name: /Próximo/i }))
  await screen.findByRole('heading', { name: '2. Colaborador' })
}

/** Avança da etapa 2 (Colaborador) para a etapa 3 (Condições), preenchendo nome e CPF válidos. */
async function avancarParaCondicoes(user: ReturnType<typeof userEvent.setup>) {
  await avancarParaColaborador(user)
  const [funcionarioInput, cpfInput] = screen.getAllByRole('textbox')
  await user.type(funcionarioInput, 'João da Silva')
  await user.type(cpfInput, '52998224725') // CPF válido (dígitos verificadores corretos)
  await user.click(screen.getByRole('button', { name: /Próximo/i }))
  await screen.findByRole('heading', { name: '3. Condições' })
}

describe('LoanPage — assistente de 4 passos', () => {
  it('não avança da etapa Equipamento sem selecionar um item disponível', async () => {
    mockListItems()
    const user = userEvent.setup()
    renderPage()

    await screen.findByRole('heading', { name: '1. Equipamento' })
    await user.click(screen.getByRole('button', { name: /Próximo/i }))

    expect(await screen.findByText('Selecione um equipamento disponível.')).toBeInTheDocument()
    expect(screen.getByRole('heading', { name: '1. Equipamento' })).toBeInTheDocument()
  })

  it('CPF inválido barra o avanço no passo 2 (Colaborador)', async () => {
    mockListItems()
    const user = userEvent.setup()
    renderPage()

    await avancarParaColaborador(user)
    const [funcionarioInput, cpfInput] = screen.getAllByRole('textbox')
    await user.type(funcionarioInput, 'João da Silva')
    await user.type(cpfInput, '123.456.789-00') // sequência com dígitos verificadores errados

    await user.click(screen.getByRole('button', { name: /Próximo/i }))

    expect(await screen.findByText('CPF inválido.')).toBeInTheDocument()
    expect(screen.getByRole('heading', { name: '2. Colaborador' })).toBeInTheDocument()
  })

  it('preenche a revenda automaticamente a partir do item, mas ainda barra sem o setor no passo 3 (Condições)', async () => {
    mockListItems()
    const user = userEvent.setup()
    renderPage()

    await avancarParaCondicoes(user)
    await user.click(screen.getByRole('button', { name: /Próximo/i }))

    // A revenda já veio preenchida do item selecionado (Matriz) — só o setor
    // barra o avanço, confirmando que o preenchimento automático ocorreu.
    expect(await screen.findByText('Selecione o setor.')).toBeInTheDocument()
    expect(screen.queryByText('Selecione a revenda.')).not.toBeInTheDocument()
    expect(screen.getByRole('heading', { name: '3. Condições' })).toBeInTheDocument()
  })

  it('permite voltar a uma etapa já concluída, mas não pular para uma etapa futura', async () => {
    mockListItems()
    const user = userEvent.setup()
    renderPage()

    await avancarParaCondicoes(user)

    // Volta livre: "Equipamento" (etapa concluída) está clicável no Stepper.
    await user.click(screen.getByRole('button', { name: /Equipamento/i }))
    expect(await screen.findByRole('heading', { name: '1. Equipamento' })).toBeInTheDocument()

    // Sem avanço: de volta à etapa 1, "Condições" e "Confirmação" (etapas
    // futuras, ainda não concluídas) ficam desabilitadas no Stepper.
    const botaoCondicoes = screen.getByRole('button', { name: /Condições/i })
    const botaoConfirmacao = screen.getByRole('button', { name: /Confirmação/i })
    expect(botaoCondicoes).toBeDisabled()
    expect(botaoConfirmacao).toBeDisabled()

    await user.click(botaoCondicoes)
    expect(screen.getByRole('heading', { name: '1. Equipamento' })).toBeInTheDocument()
  })

  it('o toggle "É pessoa jurídica" continua desmarcado por padrão', async () => {
    mockListItems()
    const user = userEvent.setup()
    renderPage()

    await avancarParaColaborador(user)

    expect(screen.getByRole('checkbox')).not.toBeChecked()
  })
})
