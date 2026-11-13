import { describe, expect, it } from 'vitest'
import { fireEvent, screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { http, HttpResponse } from 'msw'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import type { Item } from '@/api/items'
import type { TermoSolicitacao } from '@/api/termos'
import tab from './termos'

const Termos = tab.Component

function termo(extra: Partial<TermoSolicitacao> = {}): TermoSolicitacao {
  return {
    id: 30,
    item_id: 12,
    tipo: 'uso',
    modelo_versao: 'termo.docx@1a2b3c4d',
    status: 'assinado',
    solicitado_por: 'tecnico.teste',
    solicitado_em: '2026-11-10T09:00:00',
    assinado_em: '2026-11-12T10:00:00',
    tem_documento: true,
    tem_assinado: true,
    legado: false,
    ...extra,
  }
}

function resposta(vigente: TermoSolicitacao | null, items: TermoSolicitacao[], revsign = false) {
  return { revsign_ativo: revsign, vigente, historico: { items, total: items.length } }
}

const item = { id: 12, status: 'Indisponível' } as Item

function renderTab(i: Item = item) {
  return renderWithClient(
    <>
      <Termos item={i} />
      <ToastContainer />
    </>
  )
}

describe('aba Termos do equipamento', () => {
  it('é uma aba só de Gestor e Técnico, na ordem 30', () => {
    expect(tab.id).toBe('termos')
    expect(tab.order).toBe(30)
    expect(tab.roles).toEqual(['Gestor', 'Técnico'])
  })

  it('mostra o termo vigente com tipo, versão, status, solicitante e assinatura', async () => {
    let query = ''
    server.use(
      http.get('/api/items/12/termos', ({ request }) => {
        query = new URL(request.url).search
        return HttpResponse.json(resposta(termo(), [termo()]))
      })
    )
    renderTab()

    expect(await screen.findByText('termo.docx@1a2b3c4d')).toBeInTheDocument()
    expect(screen.getAllByText('Termo de uso').length).toBeGreaterThan(0)
    expect(screen.getAllByText('Assinado').length).toBeGreaterThan(0)
    expect(screen.getByText(/Solicitado por tecnico\.teste/)).toBeInTheDocument()
    expect(screen.getAllByText(/12\/11\/2026/).length).toBeGreaterThan(0)
    expect(query).toContain('limit=25')
    expect(query).toContain('offset=0')
  })

  it('mostra a mensagem de vazio quando não há solicitações', async () => {
    server.use(http.get('/api/items/12/termos', () => HttpResponse.json(resposta(null, []))))
    renderTab()

    expect(await screen.findByText('Nenhum termo vigente para este equipamento.')).toBeInTheDocument()
    expect(screen.getByText('Nenhuma solicitação de assinatura registrada.')).toBeInTheDocument()
  })

  it('solicitar com agendamento envia agendado_para', async () => {
    const user = userEvent.setup()
    let corpo: unknown = null
    server.use(
      http.get('/api/items/12/termos', () => HttpResponse.json(resposta(null, [], true))),
      http.post('/api/items/12/termos', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json(termo({ id: 40, status: 'agendado', agendado_para: '2026-11-20T09:30:00' }), {
          status: 201,
        })
      })
    )
    renderTab()

    await user.click(await screen.findByRole('button', { name: 'Solicitar assinatura' }))
    await user.click(screen.getByRole('radio', { name: 'Agendar envio' }))
    fireEvent.change(screen.getByLabelText('Data e hora do envio'), { target: { value: '2026-11-20T09:30' } })
    await user.click(screen.getByRole('button', { name: 'Solicitar' }))

    await waitFor(() => expect(corpo).toEqual({ agendado_para: '2026-11-20T09:30' }))
    expect(await screen.findByText(/Envio agendado para 20\/11\/2026/)).toBeInTheDocument()
  })

  it('solicitar agora envia corpo vazio e avisa o e-mail do destinatário', async () => {
    const user = userEvent.setup()
    let corpo: unknown = null
    server.use(
      http.get('/api/items/12/termos', () => HttpResponse.json(resposta(null, [], true))),
      http.post('/api/items/12/termos', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json(termo({ id: 41, status: 'enviado', destinatario_email: 'maria@gmail.com' }), {
          status: 201,
        })
      })
    )
    renderTab()

    await user.click(await screen.findByRole('button', { name: 'Solicitar assinatura' }))
    await user.click(screen.getByRole('button', { name: 'Solicitar' }))

    await waitFor(() => expect(corpo).toEqual({}))
    expect(await screen.findByText('Termo enviado para assinatura em maria@gmail.com.')).toBeInTheDocument()
  })

  it('sem RevSign o agendamento fica indisponível', async () => {
    const user = userEvent.setup()
    server.use(http.get('/api/items/12/termos', () => HttpResponse.json(resposta(null, [], false))))
    renderTab()

    await user.click(await screen.findByRole('button', { name: 'Solicitar assinatura' }))
    expect(screen.getByRole('radio', { name: 'Agendar envio' })).toBeDisabled()
    expect(screen.getByText('Agendamento exige a integração com o RevSign ligada.')).toBeInTheDocument()
  })

  it('avisa que a solicitação pendente será cancelada', async () => {
    const user = userEvent.setup()
    const pendente = termo({ id: 31, status: 'enviado', tem_assinado: false, assinado_em: null })
    server.use(http.get('/api/items/12/termos', () => HttpResponse.json(resposta(pendente, [pendente], true))))
    renderTab()

    await user.click(await screen.findByRole('button', { name: 'Solicitar assinatura' }))
    expect(screen.getByText('A solicitação pendente deste equipamento será cancelada.')).toBeInTheDocument()
  })

  it('cancelar solicitação pendente', async () => {
    const user = userEvent.setup()
    const pendente = termo({ id: 31, status: 'enviado', tem_assinado: false, assinado_em: null })
    let gets = 0
    let cancelou = false
    server.use(
      http.get('/api/items/12/termos', () => {
        gets += 1
        return HttpResponse.json(resposta(cancelou ? null : pendente, [pendente], true))
      }),
      http.post('/api/termos/31/cancelar', () => {
        cancelou = true
        return HttpResponse.json({ ...pendente, status: 'cancelado' })
      })
    )
    renderTab()

    await user.click(await screen.findByRole('button', { name: 'Cancelar' }))

    await waitFor(() => expect(cancelou).toBe(true))
    await waitFor(() => expect(gets).toBeGreaterThanOrEqual(2))
    expect(await screen.findByText('Solicitação cancelada.')).toBeInTheDocument()
  })

  it('reenvia a solicitação enviada', async () => {
    const user = userEvent.setup()
    const pendente = termo({ id: 31, status: 'enviado', tem_assinado: false, assinado_em: null })
    let reenviou = false
    server.use(
      http.get('/api/items/12/termos', () => HttpResponse.json(resposta(pendente, [pendente], true))),
      http.post('/api/termos/31/reenviar', () => {
        reenviou = true
        return HttpResponse.json(pendente)
      })
    )
    renderTab()

    await user.click(await screen.findByRole('button', { name: 'Reenviar' }))
    await waitFor(() => expect(reenviou).toBe(true))
  })

  it('equipamento disponível não permite solicitar', async () => {
    server.use(http.get('/api/items/12/termos', () => HttpResponse.json(resposta(null, []))))
    renderTab({ id: 12, status: 'Disponível' } as Item)

    const botao = await screen.findByRole('button', { name: 'Solicitar assinatura' })
    expect(botao).toBeDisabled()
    expect(screen.getByText('Disponível só para equipamento com colaborador.')).toBeInTheDocument()
  })

  it('histórico paginado', async () => {
    const enviado = termo({ id: 32, status: 'enviado', tem_assinado: false, assinado_em: null })
    const cancelado = termo({ id: 31, status: 'cancelado', tem_assinado: false, assinado_em: null })
    server.use(
      http.get('/api/items/12/termos', () => HttpResponse.json(resposta(enviado, [enviado, cancelado])))
    )
    renderTab()

    expect(await screen.findByText('Mostrando 1–2 de 2')).toBeInTheDocument()
    expect(screen.getAllByText('Enviado').length).toBeGreaterThan(0)
    expect(screen.getByText('Cancelado')).toBeInTheDocument()
  })

  it('termo legado só oferece ver o assinado', async () => {
    const legado = termo({ id: null, legado: true, modelo_versao: null })
    server.use(http.get('/api/items/12/termos', () => HttpResponse.json(resposta(legado, []))))
    renderTab()

    expect(await screen.findByRole('button', { name: 'Ver assinado' })).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Cancelar' })).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Baixar termo' })).not.toBeInTheDocument()
  })
})
