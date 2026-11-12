import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import { LocaisDialog } from '@/components/locais/LocaisDialog'
import type { Unidade } from '@/api/unidades'
import type { Local } from '@/api/locais'

const unidade: Unidade = {
  id: 1,
  nome: 'Juazeiro',
  razao_social: 'Revalle Juazeiro Ltda',
  cnpj: '11.222.333/0001-81',
  is_active: true,
}

const almox: Local = { id: 10, unidade_id: 1, unidade_nome: 'Juazeiro', nome: 'Almoxarifado', descricao: 'Térreo', ativo: true }
const sala: Local = { id: 11, unidade_id: 1, unidade_nome: 'Juazeiro', nome: 'Sala TI', descricao: null, ativo: false }

function mockLista(locais: Local[] = [almox, sala]) {
  const pedidos: URLSearchParams[] = []
  server.use(
    http.get('/api/locais', ({ request }) => {
      pedidos.push(new URL(request.url).searchParams)
      return HttpResponse.json(locais)
    })
  )
  return pedidos
}

function renderDialog() {
  return renderWithClient(
    <>
      <LocaisDialog unidade={unidade} open onOpenChange={() => {}} />
      <ToastContainer />
    </>
  )
}

describe('LocaisDialog', () => {
  it('lista os locais da unidade (inclusive inativos) com o nome da unidade em caixa alta', async () => {
    const pedidos = mockLista()
    renderDialog()

    expect(await screen.findByText('Almoxarifado')).toBeInTheDocument()
    expect(screen.getByText('Sala TI')).toBeInTheDocument()
    expect(screen.getByText('Ativo')).toBeInTheDocument()
    expect(screen.getByText('Inativo')).toBeInTheDocument()
    expect(screen.getByText(/JUAZEIRO/)).toBeInTheDocument()
    expect(pedidos[0].get('unidade_id')).toBe('1')
    expect(pedidos[0].get('include_inactive')).toBe('true')
  })

  it('cria um local na unidade', async () => {
    mockLista()
    let corpo: unknown = null
    server.use(
      http.post('/api/locais', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ id: 12 }, { status: 201 })
      })
    )
    const user = userEvent.setup()
    renderDialog()
    await screen.findByText('Almoxarifado')

    await user.type(screen.getByLabelText('Nome do novo local'), '  Recepção ')
    await user.type(screen.getByLabelText('Descrição do novo local'), 'Entrada')
    await user.click(screen.getByRole('button', { name: 'Adicionar local' }))

    await waitFor(() => expect(corpo).toEqual({ unidade_id: 1, nome: 'Recepção', descricao: 'Entrada' }))
    expect(await screen.findByText('Local criado com sucesso!')).toBeInTheDocument()
    expect(screen.getByLabelText('Nome do novo local')).toHaveValue('')
  })

  it('mostra o erro de nome duplicado vindo do backend', async () => {
    mockLista()
    server.use(
      http.post('/api/locais', () =>
        HttpResponse.json({ detail: 'Já existe um local com esse nome nesta unidade.' }, { status: 400 })
      )
    )
    const user = userEvent.setup()
    renderDialog()
    await screen.findByText('Almoxarifado')

    await user.type(screen.getByLabelText('Nome do novo local'), 'Almoxarifado')
    await user.click(screen.getByRole('button', { name: 'Adicionar local' }))

    expect(await screen.findByText('Já existe um local com esse nome nesta unidade.')).toBeInTheDocument()
    expect(screen.getByLabelText('Nome do novo local')).toHaveValue('Almoxarifado')
  })

  it('não envia nome vazio', async () => {
    mockLista()
    let chamadas = 0
    server.use(
      http.post('/api/locais', () => {
        chamadas += 1
        return HttpResponse.json({ id: 1 })
      })
    )
    const user = userEvent.setup()
    renderDialog()
    await screen.findByText('Almoxarifado')

    await user.click(screen.getByRole('button', { name: 'Adicionar local' }))

    expect(await screen.findByText('Informe o nome do local.')).toBeInTheDocument()
    expect(chamadas).toBe(0)
  })

  it('renomeia um local enviando só o nome alterado', async () => {
    mockLista()
    let corpo: unknown = null
    server.use(
      http.put('/api/locais/10', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ detail: 'ok' })
      })
    )
    const user = userEvent.setup()
    renderDialog()
    await screen.findByText('Almoxarifado')

    await user.click(screen.getByLabelText('Renomear local Almoxarifado'))
    const campo = screen.getByLabelText('Novo nome do local Almoxarifado')
    await user.clear(campo)
    await user.type(campo, 'Almox Central')
    await user.click(screen.getByRole('button', { name: 'Salvar nome' }))

    await waitFor(() => expect(corpo).toEqual({ nome: 'Almox Central' }))
    expect(await screen.findByText('Local atualizado com sucesso!')).toBeInTheDocument()
  })

  it('desativa e reativa locais', async () => {
    mockLista()
    const corpos: Record<number, unknown> = {}
    server.use(
      http.put('/api/locais/:id', async ({ request, params }) => {
        corpos[Number(params.id)] = await request.json()
        return HttpResponse.json({ detail: 'ok' })
      })
    )
    const user = userEvent.setup()
    renderDialog()
    await screen.findByText('Almoxarifado')

    await user.click(screen.getByLabelText('Desativar local Almoxarifado'))
    await waitFor(() => expect(corpos[10]).toEqual({ ativo: false }))

    await user.click(screen.getByLabelText('Reativar local Sala TI'))
    await waitFor(() => expect(corpos[11]).toEqual({ ativo: true }))
  })

  it('mostra estado vazio quando a unidade não tem locais', async () => {
    mockLista([])
    renderDialog()
    expect(await screen.findByText('Nenhum local cadastrado')).toBeInTheDocument()
    const dialog = screen.getByRole('dialog')
    expect(within(dialog).getByRole('button', { name: 'Adicionar local' })).toBeInTheDocument()
  })
})
