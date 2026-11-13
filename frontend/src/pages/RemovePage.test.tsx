import { beforeAll, describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import RemovePage from '@/pages/RemovePage'
import { hojeBR } from '@/features/ciclo-vida/format'

// jsdom não implementa a Pointer Events API que o Select (Radix) usa — sem
// estes stubs, abrir o dropdown de motivo lançaria "hasPointerCapture is not
// a function". Escopo local a este arquivo, só para o teste conseguir
// exercitar o fluxo completo até o checkbox de ciência.
beforeAll(() => {
  Element.prototype.hasPointerCapture = Element.prototype.hasPointerCapture ?? (() => false)
  Element.prototype.setPointerCapture = Element.prototype.setPointerCapture ?? (() => {})
  Element.prototype.releasePointerCapture = Element.prototype.releasePointerCapture ?? (() => {})
  Element.prototype.scrollIntoView = Element.prototype.scrollIntoView ?? (() => {})
})

// RemovePage é o portão da única ação IRREVERSÍVEL do sistema (baixa
// definitiva de um ativo). O defeito herdado da FT_STC — o quadrado visível
// do checkbox de ciência não respondia ao clique quando o componente não
// recebia `id` (ver Task 1, checkbox.tsx) — é justamente o que trancava esse
// portão sem que ninguém percebesse: o operador clicava, nada acontecia, e o
// botão de confirmar continuava destravado ou travado por engano. Este teste
// cobre a integração real (não só o componente isolado): clica no QUADRADO
// (nunca no texto ao lado, que nem é mais clicável nesta tela) e confirma que
// o botão "Confirmar Baixa Definitiva" destrava só depois disso.

const itemDisponivel = {
  id: 77,
  tipo: 'Notebook',
  brand: 'Dell',
  model: 'Latitude',
  identificador: 'PAT-077',
  revenda: 'Matriz',
  nota_fiscal: '123456789',
  date_registered: '2026-01-10',
  status: 'Disponível',
}

function mockServidor() {
  server.use(
    http.get('/api/constants', () =>
      HttpResponse.json({
        center_costs: [],
        revendas: [],
        setores: [],
        equipment_types: [],
        peripheral_types: [],
        removal_reasons: ['Furto/Roubo'],
        removal_reasons_attachment: { 'Furto/Roubo': false },
      })
    ),
    http.get('/api/items', () => HttpResponse.json({ items: [itemDisponivel], total: 1 })),
    http.get('/api/items/:id', () => HttpResponse.json(itemDisponivel))
  )
}

// MemoryRouter: o layout da FT_STC trouxe os botões "Cancelar e Voltar" e
// "Cancelar", que usam `useNavigate` — sem um roteador em volta a página nem
// monta. Não altera nada do que o teste verifica.
function renderPage(inicial = '/remove') {
  return renderWithClient(
    <MemoryRouter initialEntries={[inicial]}>
      <RemovePage />
      <ToastContainer />
    </MemoryRouter>
  )
}

/** Seleciona o único equipamento disponível e o motivo — os dois pré-requisitos que, junto com a ciência, destravam o botão. */
async function selecionarEquipamentoEMotivo(user: ReturnType<typeof userEvent.setup>) {
  // Texto do gatilho passou a seguir o layout da FT_STC ("Pesquisar por marca,
  // modelo ou identificador…"); o significado do passo é o mesmo.
  await user.click(await screen.findByRole('button', { name: /Pesquisar por marca, modelo ou identificador/i }))
  await user.click(await screen.findByRole('button', { name: /Notebook Dell Latitude/i }))
  // Ficha do ativo aparece assim que a seleção é processada.
  await screen.findByText('Ficha do Ativo Selecionado')

  await user.click(screen.getByRole('combobox'))
  await user.click(await screen.findByRole('option', { name: 'Furto/Roubo' }))
}

describe('RemovePage — checkbox de ciência de irreversibilidade', () => {
  it('mantém o botão de confirmar travado até equipamento, motivo e ciência estarem completos', async () => {
    mockServidor()
    const user = userEvent.setup()
    renderPage()

    // Rótulo do botão segue a FT_STC ("Confirmar Baixa Definitiva").
    const botaoConfirmar = await screen.findByRole('button', { name: /Confirmar Baixa Definitiva/i })
    expect(botaoConfirmar).toBeDisabled()

    await selecionarEquipamentoEMotivo(user)
    // Equipamento e motivo escolhidos, mas a ciência ainda não foi marcada.
    expect(botaoConfirmar).toBeDisabled()
  })

  it('destrava o botão de confirmar ao clicar no QUADRADO do checkbox (não no texto ao lado)', async () => {
    mockServidor()
    const user = userEvent.setup()
    renderPage()

    await selecionarEquipamentoEMotivo(user)

    // Rótulo do botão segue a FT_STC ("Confirmar Baixa Definitiva").
    const botaoConfirmar = screen.getByRole('button', { name: /Confirmar Baixa Definitiva/i })
    expect(botaoConfirmar).toBeDisabled()

    // Clicar no texto ao lado não deve fazer nada: não há mais um <label>
    // externo envolvendo texto + checkbox (a própria correção de design desta
    // task) — só o quadrado em si é clicável.
    await user.click(screen.getByText(/Estou ciente de que a remoção do equipamento/))
    expect(botaoConfirmar).toBeDisabled()

    // O quadrado é o <label> interno do Checkbox — irmão do <input
    // type="checkbox"> oculto (sr-only), nunca o próprio input nem o texto.
    const inputOculto = screen.getByRole('checkbox')
    const quadrado = inputOculto.nextElementSibling
    expect(quadrado).toBeTruthy()
    expect(quadrado?.tagName).toBe('LABEL')

    await user.click(quadrado as HTMLLabelElement)

    expect(inputOculto).toBeChecked()
    expect(botaoConfirmar).not.toBeDisabled()
  })
})

// ── F4: apagamento de dados (LGPD) na baixa ──────────────────────────────────

const metodosApagamento = {
  formatacao_segura: 'Formatação segura',
  reset_fabrica: 'Reset de fábrica',
  destruicao_fisica: 'Destruição física',
}

function mockServidorApagamento(item: Record<string, unknown> = itemDisponivel) {
  server.use(
    http.get('/api/constants', () =>
      HttpResponse.json({
        center_costs: [],
        revendas: [],
        setores: [],
        equipment_types: [],
        peripheral_types: [],
        removal_reasons: ['Doação', 'Roubo'],
        removal_reasons_attachment: { Doação: false, Roubo: false },
        tipos_exigem_apagamento: ['Notebook'],
        motivos_sem_posse: ['Roubo'],
        metodos_apagamento: metodosApagamento,
      })
    ),
    http.get('/api/items', () => HttpResponse.json({ items: [item], total: 1 })),
    http.get('/api/items/:id', () => HttpResponse.json(item))
  )
}

async function selecionarItem(user: ReturnType<typeof userEvent.setup>, rotulo: RegExp) {
  await user.click(await screen.findByRole('button', { name: /Pesquisar por marca, modelo ou identificador/i }))
  await user.click(await screen.findByRole('button', { name: rotulo }))
  await screen.findByText('Ficha do Ativo Selecionado')
}

async function escolherMotivo(user: ReturnType<typeof userEvent.setup>, motivo: string) {
  await user.click(screen.getByRole('combobox'))
  await user.click(await screen.findByRole('option', { name: motivo }))
}

async function marcarCiencia(user: ReturnType<typeof userEvent.setup>) {
  const quadrado = screen.getByRole('checkbox').nextElementSibling as HTMLLabelElement
  await user.click(quadrado)
}

describe('RemovePage — apagamento de dados', () => {
  it('Notebook com Doação exige o apagamento antes de destravar', async () => {
    mockServidorApagamento()
    let campos: Record<string, string> = {}
    server.use(
      http.delete('/api/items/77', async ({ request }) => {
        const form = await request.formData()
        campos = Object.fromEntries([...form.entries()].map(([k, v]) => [k, String(v)]))
        return HttpResponse.json({ detail: 'Aparelho 77 removido do estoque.' })
      })
    )
    const user = userEvent.setup()
    renderPage()

    await selecionarItem(user, /Notebook Dell Latitude/i)
    await escolherMotivo(user, 'Doação')
    await marcarCiencia(user)

    const botao = screen.getByRole('button', { name: /Confirmar Baixa Definitiva/i })
    expect(botao).toBeDisabled()
    expect(
      screen.getByRole('group', { name: 'Apagamento de dados (LGPD) — obrigatório para Notebook' })
    ).toBeInTheDocument()

    await user.click(screen.getByRole('radio', { name: 'Sim, os dados foram apagados' }))
    expect(botao).toBeDisabled()
    await user.selectOptions(screen.getByLabelText('Método de apagamento'), 'Formatação segura')
    expect(botao).toBeDisabled()
    await user.type(screen.getByLabelText('Quem apagou os dados'), 'Ana TI')
    expect(botao).toBeEnabled()

    await user.click(botao)

    await waitFor(() => expect(campos.reason).toBe('Doação'))
    expect(campos.dados_apagados).toBe('sim')
    expect(campos.metodo_apagamento).toBe('formatacao_segura')
    expect(campos.apagado_por).toBe('Ana TI')
    expect(campos.apagado_em).toBe(hojeBR())
  })

  it('"Não se aplica" só aparece para Roubo', async () => {
    mockServidorApagamento()
    const user = userEvent.setup()
    renderPage()

    await selecionarItem(user, /Notebook Dell Latitude/i)
    await escolherMotivo(user, 'Doação')
    expect(screen.queryByRole('radio', { name: /Não se aplica/ })).not.toBeInTheDocument()

    await escolherMotivo(user, 'Roubo')
    const naoAplica = await screen.findByRole('radio', {
      name: 'Não se aplica — equipamento fora de posse (roubo ou perda)',
    })
    await marcarCiencia(user)
    const botao = screen.getByRole('button', { name: /Confirmar Baixa Definitiva/i })
    expect(botao).toBeDisabled()
    await user.click(naoAplica)
    expect(botao).toBeEnabled()

    // Voltar para um motivo sem posse zera a escolha "não se aplica".
    await escolherMotivo(user, 'Doação')
    expect(screen.queryByRole('radio', { name: /Não se aplica/ })).not.toBeInTheDocument()
    expect(screen.getByRole('radio', { name: 'Sim, os dados foram apagados' })).not.toBeChecked()
    expect(screen.getByRole('button', { name: /Confirmar Baixa Definitiva/i })).toBeDisabled()
  })

  it('Impressora não pede apagamento', async () => {
    mockServidorApagamento({ ...itemDisponivel, tipo: 'Impressora', brand: 'HP', model: 'LaserJet' })
    let campos: string[] = []
    server.use(
      http.delete('/api/items/77', async ({ request }) => {
        campos = [...(await request.formData()).keys()]
        return HttpResponse.json({ detail: 'ok' })
      })
    )
    const user = userEvent.setup()
    renderPage()

    await selecionarItem(user, /Impressora HP LaserJet/i)
    await escolherMotivo(user, 'Doação')
    await marcarCiencia(user)

    expect(
      screen.getByText(
        'Este tipo de equipamento não armazena dados pessoais: o apagamento será registrado como "não se aplica".'
      )
    ).toBeInTheDocument()
    expect(screen.queryByRole('group', { name: /Apagamento de dados/ })).not.toBeInTheDocument()
    const botao = screen.getByRole('button', { name: /Confirmar Baixa Definitiva/i })
    expect(botao).toBeEnabled()
    await user.click(botao)

    await waitFor(() => expect(campos).toContain('reason'))
    expect(campos).not.toContain('dados_apagados')
  })

  it('?item=77 abre com o equipamento selecionado', async () => {
    mockServidorApagamento()
    renderPage('/remove?item=77')

    expect(await screen.findByText('Ficha do Ativo Selecionado')).toBeInTheDocument()
    expect(screen.getByText('Disponível')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /#77 — Notebook Dell Latitude/ })).toBeInTheDocument()
  })

  it('?item=77 de equipamento sem conserto (fora da busca de Disponível) também abre', async () => {
    const semConserto = { ...itemDisponivel, status: 'Em Manutenção' }
    server.use(
      http.get('/api/constants', () =>
        HttpResponse.json({
          center_costs: [],
          revendas: [],
          setores: [],
          equipment_types: [],
          peripheral_types: [],
          removal_reasons: ['Doação'],
          removal_reasons_attachment: {},
        })
      ),
      http.get('/api/items', () => HttpResponse.json({ items: [], total: 0 })),
      http.get('/api/items/:id', () => HttpResponse.json(semConserto))
    )
    renderPage('/remove?item=77')

    expect(await screen.findByText('Ficha do Ativo Selecionado')).toBeInTheDocument()
    expect(screen.getByText('Em Manutenção')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /#77 — Notebook Dell Latitude/ })).toBeInTheDocument()
  })
})
