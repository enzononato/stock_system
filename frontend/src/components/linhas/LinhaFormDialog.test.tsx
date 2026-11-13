import { beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import { LinhaFormDialog } from '@/components/linhas/LinhaFormDialog'
import type { Linha } from '@/api/linhas'

const linha: Linha = {
  id: 5,
  numero: '74999990001',
  operadora: 'Vivo',
  plano: 'Controle 20GB',
  custo_mensal: '49.90',
  status: 'ativa',
  unidade_id: 1,
  unidade_nome: 'Juazeiro',
  item_id: null,
  item_tipo: null,
  item_brand: null,
  item_model: null,
  item_identificador: null,
  item_ativo: null,
  colaborador_id: 3,
  colaborador_nome: 'MARIA SILVA',
  colaborador_situacao: 'ativo',
  iccid: null,
  observacoes: null,
  criada_em: '2026-11-12T10:00:00',
  atualizada_em: '2026-11-12T10:00:00',
}

const maria = {
  id: 3,
  cpf: '12345678909',
  nome: 'MARIA SILVA',
  situacao: 'ativo',
  email_manual: false,
  telefone_manual: false,
  origem: 'rh',
}

function mockBase() {
  server.use(
    http.get('/api/unidades', () =>
      HttpResponse.json([
        { id: 1, nome: 'Juazeiro', razao_social: 'X', cnpj: '1', is_active: true },
        { id: 2, nome: 'Petrolina', razao_social: 'X', cnpj: '2', is_active: true },
      ])
    ),
    http.get('/api/linhas/aparelhos', () => HttpResponse.json([])),
    http.get('/api/colaboradores', () => HttpResponse.json({ items: [maria], total: 1 }))
  )
}

function renderDialog(l: Linha | null, onOpenChange = vi.fn()) {
  const utils = renderWithClient(
    <>
      <LinhaFormDialog open onOpenChange={onOpenChange} linha={l} />
      <ToastContainer />
    </>
  )
  return { onOpenChange, ...utils }
}

beforeEach(() => {
  mockBase()
})

describe('LinhaFormDialog', () => {
  it('cria a linha enviando o custo convertido e os vínculos vazios como null', async () => {
    let corpo: Record<string, unknown> | null = null
    server.use(
      http.post('/api/linhas', async ({ request }) => {
        corpo = (await request.json()) as Record<string, unknown>
        return HttpResponse.json({ ...linha, id: 9 }, { status: 201 })
      })
    )
    const { onOpenChange } = renderDialog(null)
    const user = userEvent.setup()

    await user.type(screen.getByLabelText(/^Número/), '(74) 99999-0001')
    await user.selectOptions(screen.getByLabelText(/^Operadora/), 'Vivo')
    await user.type(screen.getByLabelText(/^Custo mensal/), '49,90')
    await screen.findByRole('option', { name: 'JUAZEIRO' })
    await user.selectOptions(screen.getByLabelText(/^Unidade/), '1')
    await user.click(screen.getByRole('button', { name: 'Salvar' }))

    await waitFor(() => expect(corpo).not.toBeNull())
    expect(corpo).toMatchObject({
      numero: '(74) 99999-0001',
      operadora: 'Vivo',
      custo_mensal: '49.90',
      unidade_id: 1,
      item_id: null,
      colaborador_id: null,
      plano: null,
      iccid: null,
      observacoes: null,
    })
    await waitFor(() => expect(onOpenChange).toHaveBeenCalledWith(false))
    expect(await screen.findByText('Linha cadastrada.')).toBeInTheDocument()
  })

  it('mostra no diálogo o erro devolvido pela API e mantém o diálogo aberto', async () => {
    server.use(
      http.post('/api/linhas', () =>
        HttpResponse.json({ detail: 'Já existe uma linha com este número.' }, { status: 409 })
      )
    )
    const { onOpenChange } = renderDialog(null)
    const user = userEvent.setup()

    await user.type(screen.getByLabelText(/^Número/), '74999990001')
    await user.selectOptions(screen.getByLabelText(/^Operadora/), 'Claro')
    await user.type(screen.getByLabelText(/^Custo mensal/), '10')
    await screen.findByRole('option', { name: 'JUAZEIRO' })
    await user.selectOptions(screen.getByLabelText(/^Unidade/), '1')
    await user.click(screen.getByRole('button', { name: 'Salvar' }))

    expect(await screen.findByText('Já existe uma linha com este número.')).toBeInTheDocument()
    expect(onOpenChange).not.toHaveBeenCalled()
  })

  it('recusa custo inválido sem fazer requisição', async () => {
    // Sem handler de POST: o MSW (onUnhandledRequest: 'error') falharia se saísse uma requisição.
    renderDialog(null)
    const user = userEvent.setup()

    await user.type(screen.getByLabelText(/^Número/), '74999990001')
    await user.selectOptions(screen.getByLabelText(/^Operadora/), 'Vivo')
    await user.type(screen.getByLabelText(/^Custo mensal/), '12,345')
    await screen.findByRole('option', { name: 'JUAZEIRO' })
    await user.selectOptions(screen.getByLabelText(/^Unidade/), '1')
    await user.click(screen.getByRole('button', { name: 'Salvar' }))

    expect(await screen.findByText('Informe o custo no formato 0,00.')).toBeInTheDocument()
  })

  it('na edição preenche os campos formatados e envia só o que mudou, com o motivo', async () => {
    let corpo: unknown = null
    server.use(
      http.put('/api/linhas/5', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...linha, plano: 'Controle 30GB' })
      })
    )
    const { onOpenChange } = renderDialog(linha)
    const user = userEvent.setup()

    expect(screen.getByLabelText(/^Número/)).toHaveValue('(74) 99999-0001')
    expect(screen.getByLabelText(/^Custo mensal/)).toHaveValue('49,90')

    const plano = screen.getByLabelText(/^Plano/)
    await user.clear(plano)
    await user.type(plano, 'Controle 30GB')
    await user.type(screen.getByLabelText(/^Motivo da alteração/), 'Upgrade')
    await user.click(screen.getByRole('button', { name: 'Salvar' }))

    await waitFor(() => expect(corpo).toEqual({ plano: 'Controle 30GB', motivo: 'Upgrade' }))
    await waitFor(() => expect(onOpenChange).toHaveBeenCalledWith(false))
    expect(await screen.findByText('Linha atualizada.')).toBeInTheDocument()
  })

  it('limpar o colaborador envia colaborador_id null', async () => {
    let corpo: unknown = null
    server.use(
      http.put('/api/linhas/5', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...linha, colaborador_id: null })
      })
    )
    renderDialog(linha)
    const user = userEvent.setup()

    await user.click(await screen.findByRole('button', { name: 'Limpar colaborador' }))
    await user.click(screen.getByRole('button', { name: 'Salvar' }))

    await waitFor(() => expect(corpo).toEqual({ colaborador_id: null }))
  })

  it('salvar sem mudar nada avisa e não faz requisição', async () => {
    const { onOpenChange } = renderDialog(linha)
    const user = userEvent.setup()

    await user.click(screen.getByRole('button', { name: 'Salvar' }))

    expect(await screen.findByText('Nenhuma alteração para salvar.')).toBeInTheDocument()
    expect(onOpenChange).not.toHaveBeenCalled()
  })

  it('mudar o número só conta quando os dígitos mudam', async () => {
    let corpo: unknown = null
    server.use(
      http.put('/api/linhas/5', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json(linha)
      })
    )
    renderDialog(linha)
    const user = userEvent.setup()

    const numero = screen.getByLabelText(/^Número/)
    await user.clear(numero)
    await user.type(numero, '74 99999 0002')
    await user.click(screen.getByRole('button', { name: 'Salvar' }))

    await waitFor(() => expect(corpo).toEqual({ numero: '74 99999 0002' }))
  })
})
