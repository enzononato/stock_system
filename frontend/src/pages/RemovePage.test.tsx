import { beforeAll, describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import RemovePage from '@/pages/RemovePage'

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
// o botão "Confirmar Remoção" destrava só depois disso.

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

function renderPage() {
  return renderWithClient(
    <>
      <RemovePage />
      <ToastContainer />
    </>
  )
}

/** Seleciona o único equipamento disponível e o motivo — os dois pré-requisitos que, junto com a ciência, destravam o botão. */
async function selecionarEquipamentoEMotivo(user: ReturnType<typeof userEvent.setup>) {
  await user.click(await screen.findByRole('button', { name: /Selecione um equipamento disponível/i }))
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

    const botaoConfirmar = await screen.findByRole('button', { name: /Confirmar Remoção/i })
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

    const botaoConfirmar = screen.getByRole('button', { name: /Confirmar Remoção/i })
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
