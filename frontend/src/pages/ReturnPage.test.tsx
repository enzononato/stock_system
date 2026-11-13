import { beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import ReturnPage from '@/pages/ReturnPage'
import type { Colaborador } from '@/api/colaboradores'

const maria: Colaborador = {
  id: 7,
  cpf: '52998224725',
  nome: 'Maria Souza',
  cargo: 'Analista',
  setor: 'Financeiro',
  situacao: 'ativo',
  email: 'maria@gmail.com',
  email_manual: true,
  telefone_manual: false,
  origem: 'rh',
}

function itemEmCampo(extra: Record<string, unknown> = {}) {
  return {
    id: 9,
    tipo: 'Notebook',
    brand: 'Dell',
    model: 'Latitude',
    status: 'Indisponível',
    assigned_to: 'Maria Souza',
    cpf: '52998224725',
    revenda: 'Revalle Juazeiro',
    date_issued: '2026-10-01T09:00:00',
    ...extra,
  }
}

function mockItens(emCampo: unknown[], pendentes: unknown[] = [], captura?: (status: string, p: URLSearchParams) => void) {
  server.use(
    http.get('/api/items', ({ request }) => {
      const params = new URL(request.url).searchParams
      const status = params.get('status') ?? ''
      captura?.(status, params)
      const lista = status === 'Indisponível' ? emCampo : status === 'Pendente Devolução' ? pendentes : []
      return HttpResponse.json({ items: lista, total: lista.length })
    })
  )
}

function renderPage() {
  return renderWithClient(
    <>
      <ReturnPage />
      <ToastContainer />
    </>
  )
}

beforeEach(() => {
  URL.createObjectURL = vi.fn(() => 'blob:x')
  URL.revokeObjectURL = vi.fn()
  vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
})

describe('ReturnPage — devolução com colaborador do cadastro', () => {
  it('item sem vínculo exige escolher do cadastro e mostra desligados', async () => {
    const buscas: URLSearchParams[] = []
    mockItens([itemEmCampo({ cpf: null })])
    server.use(
      http.get('/api/colaboradores', ({ request }) => {
        buscas.push(new URL(request.url).searchParams)
        return HttpResponse.json({ items: [maria], total: 1 })
      })
    )
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: /Iniciar Devolução/i }))
    expect(await screen.findByRole('heading', { name: 'Iniciar devolução' })).toBeInTheDocument()
    expect(
      screen.getByText(/Equipamento sem colaborador do cadastro: escolha quem está devolvendo \(responsável registrado: Maria Souza\)\./)
    ).toBeInTheDocument()

    await user.click(screen.getByRole('button', { name: 'Gerar termo e iniciar' }))
    expect(await screen.findByText('Selecione o colaborador.')).toBeInTheDocument()

    // Colaborador desligado também devolve: a busca não filtra por situação.
    await waitFor(() => expect(buscas.length).toBeGreaterThan(0))
    expect(buscas.every((p) => !p.has('situacao'))).toBe(true)
  })

  it('item vinculado envia colaborador_id e, com RevSign, não baixa o termo', async () => {
    let corpo: unknown = null
    let baixou = false
    mockItens([itemEmCampo({ colaborador_id: 7 })])
    server.use(
      http.get('/api/colaboradores', () => HttpResponse.json({ items: [maria], total: 1 })),
      http.get('/api/colaboradores/7', () => HttpResponse.json(maria)),
      http.post('/api/loans/9/return/initiate', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({
          detail: 'Termo de devolução gerado.',
          download_url: '/api/documents/files/termos_devolucao/termo.docx',
          filename: 'termo.docx',
          termo: { id: 3, item_id: 9, tipo: 'devolucao', status: 'enviado', destinatario_email: 'maria@gmail.com' },
        })
      }),
      http.get('/api/documents/files/*', () => {
        baixou = true
        return new HttpResponse('x')
      })
    )
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: /Iniciar Devolução/i }))
    await waitFor(() => expect(screen.getByLabelText(/E-mail do colaborador/i)).toHaveValue('maria@gmail.com'))
    await user.click(screen.getByRole('button', { name: 'Gerar termo e iniciar' }))

    expect(
      await screen.findByText('Termo de devolução enviado para assinatura em maria@gmail.com.')
    ).toBeInTheDocument()
    expect(corpo).toEqual({ colaborador_id: 7 })
    expect(baixou).toBe(false)
    expect(screen.queryByRole('heading', { name: /Conferência Física/ })).not.toBeInTheDocument()
  })

  it('assinatura manual baixa o termo e abre a conferência física', async () => {
    let baixou = false
    mockItens([itemEmCampo({ colaborador_id: 7 })])
    server.use(
      http.get('/api/colaboradores', () => HttpResponse.json({ items: [maria], total: 1 })),
      http.get('/api/colaboradores/7', () => HttpResponse.json(maria)),
      http.get('/api/items/9', () => HttpResponse.json(itemEmCampo({ colaborador_id: 7 }))),
      http.post('/api/loans/9/return/initiate', () =>
        HttpResponse.json({
          detail: 'Termo de devolução gerado.',
          download_url: '/api/documents/files/termos_devolucao/termo.docx',
          filename: 'termo.docx',
          termo: { id: 3, item_id: 9, tipo: 'devolucao', status: 'manual', motivo: 'Integração desligada.' },
        })
      ),
      http.get('/api/documents/files/termos_devolucao/termo.docx', () => {
        baixou = true
        return new HttpResponse(new Uint8Array([80, 75]))
      })
    )
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: /Iniciar Devolução/i }))
    await waitFor(() => expect(screen.getByLabelText(/E-mail do colaborador/i)).toHaveValue('maria@gmail.com'))
    await user.click(screen.getByRole('button', { name: 'Gerar termo e iniciar' }))

    expect(await screen.findByRole('heading', { name: 'Conferência Física & Termo de Devolução' })).toBeInTheDocument()
    await waitFor(() => expect(baixou).toBe(true))
  })

  it('grava o e-mail no cadastro quando o colaborador ainda não tem', async () => {
    let corpoPut: unknown = null
    const semEmail = { ...maria, email: null, email_manual: false }
    mockItens([itemEmCampo({ colaborador_id: 7 })])
    server.use(
      http.get('/api/colaboradores', () => HttpResponse.json({ items: [semEmail], total: 1 })),
      http.get('/api/colaboradores/7', () => HttpResponse.json(semEmail)),
      http.put('/api/colaboradores/7', async ({ request }) => {
        corpoPut = await request.json()
        return HttpResponse.json({ ...maria, email: 'maria.souza@gmail.com' })
      }),
      http.post('/api/loans/9/return/initiate', () =>
        HttpResponse.json({
          detail: 'Termo de devolução gerado.',
          download_url: '/api/documents/files/termos_devolucao/termo.docx',
          filename: 'termo.docx',
          termo: { id: 3, item_id: 9, tipo: 'devolucao', status: 'enviado', destinatario_email: 'maria.souza@gmail.com' },
        })
      )
    )
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: /Iniciar Devolução/i }))
    const campo = await screen.findByLabelText(/E-mail do colaborador/i)
    await user.click(screen.getByRole('button', { name: 'Gerar termo e iniciar' }))
    expect(await screen.findByText('Informe o e-mail do colaborador.')).toBeInTheDocument()

    await user.type(campo, 'maria.souza@gmail.com')
    await user.click(screen.getByRole('button', { name: 'Gerar termo e iniciar' }))
    expect(
      await screen.findByText('Termo de devolução enviado para assinatura em maria.souza@gmail.com.')
    ).toBeInTheDocument()
    expect(corpoPut).toEqual({ email: 'maria.souza@gmail.com' })
  })

  it('tabelas usam a paginação padrão', async () => {
    const limites: Record<string, string[]> = {}
    mockItens([itemEmCampo()], [itemEmCampo({ id: 10, status: 'Pendente Devolução' })], (status, p) => {
      ;(limites[status] ??= []).push(p.get('limit') ?? '')
    })
    renderPage()

    await waitFor(() => expect(screen.getAllByText('Mostrando 1–1 de 1')).toHaveLength(2))
    expect(limites['Indisponível']).toContain('25')
    expect(limites['Pendente Devolução']).toContain('25')
    expect(screen.getAllByText('REVALLE JUAZEIRO').length).toBeGreaterThan(0)
  })

  it('escolher um colaborador sem vínculo preenche o e-mail e não grava nada no cadastro', async () => {
    let put = false
    let corpo: unknown = null
    mockItens([itemEmCampo({ cpf: null })])
    server.use(
      http.get('/api/colaboradores', () => HttpResponse.json({ items: [maria], total: 1 })),
      http.put('/api/colaboradores/7', () => {
        put = true
        return HttpResponse.json(maria)
      }),
      http.post('/api/loans/9/return/initiate', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({
          detail: 'ok',
          download_url: '/api/documents/files/termos_devolucao/termo.docx',
          filename: 'termo.docx',
          termo: { id: 3, item_id: 9, tipo: 'devolucao', status: 'enviado', destinatario_email: 'maria@gmail.com' },
        })
      })
    )
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: /Iniciar Devolução/i }))
    await user.click(await screen.findByRole('button', { name: /Selecione um colaborador/i }))
    await user.type(screen.getByPlaceholderText(/Buscar por nome, CPF ou matrícula/i), 'Mar')
    await user.click(await screen.findByText('MARIA SOUZA'))

    expect(screen.getByLabelText(/E-mail do colaborador/i)).toHaveValue('maria@gmail.com')
    await user.click(screen.getByRole('button', { name: 'Gerar termo e iniciar' }))

    expect(
      await screen.findByText('Termo de devolução enviado para assinatura em maria@gmail.com.')
    ).toBeInTheDocument()
    expect(corpo).toEqual({ colaborador_id: 7 })
    expect(put).toBe(false)
  })

  it('pré-seleciona o colaborador cujo CPF é o do responsável pelo equipamento', async () => {
    let corpo: unknown = null
    mockItens([itemEmCampo()]) // cpf 52998224725 = o de Maria
    server.use(
      http.get('/api/colaboradores', () => HttpResponse.json({ items: [maria], total: 1 })),
      http.get('/api/colaboradores/cpf/52998224725', () => HttpResponse.json(maria)),
      http.post('/api/loans/9/return/initiate', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({
          detail: 'ok',
          download_url: '/x',
          filename: 'termo.docx',
          termo: { id: 3, item_id: 9, tipo: 'devolucao', status: 'enviado', destinatario_email: 'maria@gmail.com' },
        })
      })
    )
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: /Iniciar Devolução/i }))
    await waitFor(() => expect(screen.getByLabelText(/E-mail do colaborador/i)).toHaveValue('maria@gmail.com'))
    expect(screen.queryByText(/não confere/)).not.toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: 'Gerar termo e iniciar' }))

    await waitFor(() => expect(corpo).toEqual({ colaborador_id: 7 }))
  })

  it('avisa e bloqueia quando o CPF do colaborador escolhido não é o do responsável', async () => {
    mockItens([itemEmCampo({ cpf: '11144477735' })])
    server.use(
      http.get('/api/colaboradores', () => HttpResponse.json({ items: [maria], total: 1 })),
      http.get('/api/colaboradores/cpf/11144477735', () => HttpResponse.json({ detail: 'x' }, { status: 404 }))
    )
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: /Iniciar Devolução/i }))
    await user.click(await screen.findByRole('button', { name: /Selecione um colaborador/i }))
    await user.type(screen.getByPlaceholderText(/Buscar por nome, CPF ou matrícula/i), 'Mar')
    await user.click(await screen.findByText('MARIA SOUZA'))

    expect(
      await screen.findByText('O CPF do colaborador escolhido não confere com o do responsável pelo equipamento.')
    ).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Gerar termo e iniciar' })).toBeDisabled()
  })

  it('o rótulo "Colaborador *" nomeia o seletor', async () => {
    mockItens([itemEmCampo({ cpf: null })])
    server.use(http.get('/api/colaboradores', () => HttpResponse.json({ items: [maria], total: 1 })))
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: /Iniciar Devolução/i }))
    expect(await screen.findByRole('group', { name: 'Colaborador *' })).toBeInTheDocument()
  })

  it('responsável que só existe na base de RH não é pré-selecionado e bloqueia a devolução', async () => {
    mockItens([itemEmCampo()]) // cpf 52998224725
    server.use(
      http.get('/api/colaboradores', () => HttpResponse.json({ items: [], total: 0 })),
      http.get('/api/colaboradores/cpf/52998224725', () =>
        HttpResponse.json({ ...maria, id: null, local: false })
      )
    )
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: /Iniciar Devolução/i }))

    expect(
      await screen.findByText(
        'O responsável pelo equipamento (CPF 529.982.247-25) não está no cadastro de colaboradores. Sincronize ou cadastre-o antes da devolução.'
      )
    ).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Selecione um colaborador/i })).toBeInTheDocument()
    expect(screen.queryByLabelText(/E-mail do colaborador/i)).not.toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Gerar termo e iniciar' })).toBeDisabled()
  })
})
