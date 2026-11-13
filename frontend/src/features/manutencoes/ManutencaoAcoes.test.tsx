import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import type { Manutencao } from '@/api/manutencoes'
import { ManutencaoAcoes } from './ManutencaoAcoes'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({ user: { id: 1, username: 'gestor', role: 'Gestor' }, hasRole: () => true }),
}))

function manutencao(parcial: Partial<Manutencao> = {}): Manutencao {
  return {
    id: 5,
    item_id: 7,
    status: 'aberta',
    defeito: 'Tela quebrada',
    aberta_por: 'gestor',
    aberta_em: '2026-01-10T09:00:00',
    custo_acumulado_item: '0.00',
    ...parcial,
  }
}

function renderAcoes(m: Manutencao) {
  return renderWithClient(
    <>
      <ManutencaoAcoes manutencao={m} />
      <ToastContainer />
    </>
  )
}

const nomes = () => screen.queryAllByRole('button').map((b) => b.textContent)

describe('ManutencaoAcoes — botões por status', () => {
  it('aberta mostra os quatro botões', () => {
    renderAcoes(manutencao({ status: 'aberta' }))
    expect(nomes()).toEqual(['Registrar envio', 'Marcar em reparo', 'Registrar retorno', 'Cancelar manutenção'])
  })

  it('enviada mostra os três últimos', () => {
    renderAcoes(manutencao({ status: 'enviada' }))
    expect(nomes()).toEqual(['Marcar em reparo', 'Registrar retorno', 'Cancelar manutenção'])
  })

  it('em reparo mostra retorno e cancelamento', () => {
    renderAcoes(manutencao({ status: 'em_reparo' }))
    expect(nomes()).toEqual(['Registrar retorno', 'Cancelar manutenção'])
  })

  it('retornada não mostra nenhuma ação, só o anexo quando existe', () => {
    const { unmount } = renderAcoes(manutencao({ status: 'retornada' }))
    expect(nomes()).toEqual([])
    unmount()
    renderAcoes(manutencao({ status: 'retornada', anexo_key: 'manutencoes/manutencao_5_x.pdf' }))
    expect(nomes()).toEqual(['Baixar anexo'])
  })
})

describe('ManutencaoAcoes — diálogos', () => {
  it('retorno com custo 150,00 envia resultado e custo com ponto', async () => {
    let campos: Record<string, string> = {}
    server.use(
      http.post('/api/manutencoes/5/retorno', async ({ request }) => {
        const form = await request.formData()
        campos = Object.fromEntries([...form.entries()].map(([k, v]) => [k, String(v)]))
        return HttpResponse.json({ detail: 'Retorno registrado. O equipamento voltou a Disponível.', sugerir_baixa: false })
      })
    )
    const user = userEvent.setup()
    renderAcoes(manutencao())

    await user.click(screen.getByRole('button', { name: 'Registrar retorno' }))
    const dialogo = await screen.findByRole('dialog')
    expect(within(dialogo).getByRole('radio', { name: 'Consertado — volta para Disponível' })).toBeChecked()
    await user.type(within(dialogo).getByLabelText('Custo (R$)'), '150,00')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar retorno' }))

    await waitFor(() => expect(campos.resultado).toBe('retornada'))
    expect(campos.custo).toBe('150.00')
    expect(await screen.findByText('Retorno registrado. O equipamento voltou a Disponível.')).toBeInTheDocument()
  })

  it('sem conserto envia o resultado correspondente', async () => {
    let campos: Record<string, string> = {}
    server.use(
      http.post('/api/manutencoes/5/retorno', async ({ request }) => {
        const form = await request.formData()
        campos = Object.fromEntries([...form.entries()].map(([k, v]) => [k, String(v)]))
        return HttpResponse.json({ detail: 'Manutenção encerrada sem conserto. Avalie a baixa do equipamento.', sugerir_baixa: true })
      })
    )
    const user = userEvent.setup()
    renderAcoes(manutencao({ status: 'em_reparo' }))

    await user.click(screen.getByRole('button', { name: 'Registrar retorno' }))
    const dialogo = await screen.findByRole('dialog')
    await user.click(within(dialogo).getByRole('radio', { name: 'Sem conserto' }))
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar retorno' }))

    await waitFor(() => expect(campos.resultado).toBe('sem_conserto'))
    expect(campos).not.toHaveProperty('custo')
  })

  it('custo inválido não chama a API', async () => {
    let chamadas = 0
    server.use(
      http.post('/api/manutencoes/5/retorno', () => {
        chamadas += 1
        return HttpResponse.json({ detail: 'x', sugerir_baixa: false })
      })
    )
    const user = userEvent.setup()
    renderAcoes(manutencao())

    await user.click(screen.getByRole('button', { name: 'Registrar retorno' }))
    const dialogo = await screen.findByRole('dialog')
    await user.type(within(dialogo).getByLabelText('Custo (R$)'), '1,234')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar retorno' }))

    expect(await screen.findByText('Custo inválido. Use números com até duas casas decimais.')).toBeInTheDocument()
    expect(chamadas).toBe(0)
  })

  it('cancelar sem motivo não chama a API', async () => {
    let chamadas = 0
    server.use(
      http.post('/api/manutencoes/5/cancelar', () => {
        chamadas += 1
        return HttpResponse.json({ detail: 'ok' })
      })
    )
    const user = userEvent.setup()
    renderAcoes(manutencao())

    await user.click(screen.getByRole('button', { name: 'Cancelar manutenção' }))
    const dialogo = await screen.findByRole('dialog')
    const confirmar = within(dialogo).getByRole('button', { name: 'Confirmar cancelamento' })
    expect(confirmar).toBeDisabled()
    await user.click(confirmar)
    expect(chamadas).toBe(0)
  })

  it('cancelar com motivo envia o motivo', async () => {
    let corpo: unknown
    server.use(
      http.post('/api/manutencoes/5/cancelar', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ detail: 'Manutenção cancelada. O equipamento voltou a Disponível.' })
      })
    )
    const user = userEvent.setup()
    renderAcoes(manutencao())

    await user.click(screen.getByRole('button', { name: 'Cancelar manutenção' }))
    const dialogo = await screen.findByRole('dialog')
    await user.type(within(dialogo).getByLabelText('Motivo do cancelamento'), 'Aberta por engano')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar cancelamento' }))

    await waitFor(() => expect(corpo).toEqual({ motivo: 'Aberta por engano' }))
  })

  it('envio repassa assistência e previsão de retorno', async () => {
    let corpo: unknown
    server.use(
      http.post('/api/manutencoes/5/enviar', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ detail: 'Envio à assistência registrado.' })
      })
    )
    const user = userEvent.setup()
    renderAcoes(manutencao({ assistencia: 'TecInfo' }))

    await user.click(screen.getByRole('button', { name: 'Registrar envio' }))
    const dialogo = await screen.findByRole('dialog')
    expect(within(dialogo).getByLabelText('Assistência técnica')).toHaveValue('TecInfo')
    await user.type(within(dialogo).getByLabelText('Previsão de retorno'), '20122026')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar envio' }))

    await waitFor(() => expect(corpo).toEqual({ assistencia: 'TecInfo', previsao_retorno: '20/12/2026' }))
  })

  it('envio exige a assistência', async () => {
    const user = userEvent.setup()
    renderAcoes(manutencao())

    await user.click(screen.getByRole('button', { name: 'Registrar envio' }))
    const dialogo = await screen.findByRole('dialog')
    expect(within(dialogo).getByRole('button', { name: 'Confirmar envio' })).toBeDisabled()
  })

  it('marcar em reparo é ação direta', async () => {
    let corpo: unknown
    server.use(
      http.post('/api/manutencoes/5/em-reparo', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ detail: 'Manutenção marcada como em reparo.' })
      })
    )
    const user = userEvent.setup()
    renderAcoes(manutencao())

    await user.click(screen.getByRole('button', { name: 'Marcar em reparo' }))

    expect(await screen.findByText('Manutenção marcada como em reparo.')).toBeInTheDocument()
    expect(corpo).toEqual({})
  })
})
