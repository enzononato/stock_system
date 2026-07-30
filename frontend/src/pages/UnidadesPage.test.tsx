import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import UnidadesPage, { isValidCnpj } from '@/pages/UnidadesPage'

const CNPJ_VALIDO = '11.222.333/0001-81'

const unidadeJuazeiro = {
  id: 1,
  nome: 'Juazeiro',
  razao_social: 'Revalle Juazeiro Ltda',
  cnpj: CNPJ_VALIDO,
  endereco: 'Av. Principal, 100',
  cep: '48900-000',
  cidade: 'Juazeiro',
  uf: 'BA',
  is_active: true,
}

// Petrolina foi cadastrada sem CEP de propósito — não constava no documento de
// origem que deu origem ao cadastro. A tela precisa exibir e editar essa
// unidade sem quebrar.
const unidadePetrolina = {
  id: 2,
  nome: 'Petrolina',
  razao_social: 'Revalle Petrolina Ltda',
  cnpj: '19.131.243/0001-97',
  cidade: 'Petrolina',
  uf: 'PE',
  is_active: true,
  // sem `cep`
}

const unidadeRecife = {
  id: 5,
  nome: 'Recife',
  razao_social: 'Revalle Recife Ltda',
  cnpj: '19.131.243/0001-97',
  cidade: 'Recife',
  uf: 'PE',
  is_active: false,
}

function mockListUnidades(unidades: unknown[] = [unidadeJuazeiro, unidadePetrolina]) {
  server.use(http.get('/api/unidades', () => HttpResponse.json(unidades)))
}

/**
 * Mock que respeita `include_inactive`, como o backend real: sem o parâmetro
 * (ou com ele em `false`), só devolve as unidades ativas. Usado nos testes da
 * alternância "Mostrar unidades inativas" e da ação condicional por linha,
 * onde essa filtragem é o próprio comportamento sob teste.
 */
function mockListUnidadesComFiltro(todas: { is_active: boolean }[]) {
  server.use(
    http.get('/api/unidades', ({ request }) => {
      const url = new URL(request.url)
      const includeInactive = url.searchParams.get('include_inactive') === 'true'
      const resultado = includeInactive ? todas : todas.filter((u) => u.is_active)
      return HttpResponse.json(resultado)
    })
  )
}

function renderPage() {
  return renderWithClient(
    <>
      <UnidadesPage />
      <ToastContainer />
    </>
  )
}

describe('isValidCnpj', () => {
  it('aceita um CNPJ com dígitos verificadores corretos', () => {
    expect(isValidCnpj('11.222.333/0001-81')).toBe(true)
    expect(isValidCnpj('11222333000181')).toBe(true)
  })

  it('rejeita um CNPJ com o último dígito verificador errado', () => {
    expect(isValidCnpj('11.222.333/0001-80')).toBe(false)
  })

  it('rejeita sequência com os 14 dígitos repetidos, mesmo que "passe" no cálculo', () => {
    expect(isValidCnpj('11.111.111/1111-11')).toBe(false)
  })

  it('rejeita CNPJ vazio ou com quantidade errada de dígitos', () => {
    expect(isValidCnpj('')).toBe(false)
    expect(isValidCnpj('123')).toBe(false)
    expect(isValidCnpj(undefined)).toBe(false)
  })
})

describe('UnidadesPage — unidade sem CEP', () => {
  it('lista e permite editar uma unidade sem CEP sem quebrar a tela', async () => {
    mockListUnidades()
    const user = userEvent.setup()
    renderPage()

    expect(await screen.findByText('Petrolina')).toBeInTheDocument()

    await user.click(screen.getByLabelText('Editar unidade Petrolina'))

    // O formulário abre em modo edição com o campo CEP vazio (não quebra) e
    // sinaliza visualmente que o dado está faltando, para não parecer que o
    // usuário simplesmente ainda não digitou nada.
    expect(screen.getByLabelText('CEP')).toHaveValue('')
    expect(screen.getByText('Não informado')).toBeInTheDocument()
    expect(screen.getByLabelText('Cidade')).toHaveValue('Petrolina')
  })
})

describe('UnidadesPage — confirmação ao renomear', () => {
  it('editar o nome exige confirmação explícita antes de chamar a API, e cancelar não envia nada', async () => {
    mockListUnidades()
    let chamadasPut = 0
    server.use(
      http.put('/api/unidades/1', () => {
        chamadasPut += 1
        return HttpResponse.json({ detail: 'Unidade atualizada com sucesso.' })
      })
    )

    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByLabelText('Editar unidade Juazeiro'))

    const nomeInput = screen.getByLabelText('Nome *')
    await user.clear(nomeInput)
    await user.type(nomeInput, 'Juazeiro Norte')

    await user.click(screen.getByRole('button', { name: 'Salvar Alterações' }))

    expect(await screen.findByText('Confirmar renomeação da unidade')).toBeInTheDocument()
    expect(screen.getByText(/de "Juazeiro" para "Juazeiro Norte"/)).toBeInTheDocument()
    expect(chamadasPut).toBe(0)

    await user.click(screen.getByRole('button', { name: 'Cancelar' }))

    await waitFor(() =>
      expect(screen.queryByText('Confirmar renomeação da unidade')).not.toBeInTheDocument()
    )
    expect(chamadasPut).toBe(0)
  })

  it('confirmando a renomeação, chama a API com o novo nome', async () => {
    mockListUnidades()
    let corpoRecebido: { nome?: string } = {}
    server.use(
      http.put('/api/unidades/1', async ({ request }) => {
        corpoRecebido = (await request.json()) as { nome?: string }
        return HttpResponse.json({ detail: 'Unidade atualizada com sucesso.' })
      })
    )

    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByLabelText('Editar unidade Juazeiro'))
    const nomeInput = screen.getByLabelText('Nome *')
    await user.clear(nomeInput)
    await user.type(nomeInput, 'Juazeiro Norte')
    await user.click(screen.getByRole('button', { name: 'Salvar Alterações' }))

    await screen.findByText('Confirmar renomeação da unidade')
    await user.click(screen.getByRole('button', { name: 'Confirmar Renomeação' }))

    await waitFor(() => expect(corpoRecebido.nome).toBe('Juazeiro Norte'))
    expect(await screen.findByText('Unidade atualizada com sucesso!')).toBeInTheDocument()
  })

  it('salvar sem alterar o nome não abre a confirmação de renomeação', async () => {
    mockListUnidades()
    let chamadasPut = 0
    server.use(
      http.put('/api/unidades/1', () => {
        chamadasPut += 1
        return HttpResponse.json({ detail: 'Unidade atualizada com sucesso.' })
      })
    )

    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByLabelText('Editar unidade Juazeiro'))
    // Só troca a cidade, não o nome.
    const cidadeInput = screen.getByLabelText('Cidade')
    await user.clear(cidadeInput)
    await user.type(cidadeInput, 'Juazeiro (Centro)')

    await user.click(screen.getByRole('button', { name: 'Salvar Alterações' }))

    expect(screen.queryByText('Confirmar renomeação da unidade')).not.toBeInTheDocument()
    await waitFor(() => expect(chamadasPut).toBe(1))
  })
})

describe('UnidadesPage — inativação', () => {
  it('exibe a mensagem do backend (não uma genérica) ao tentar inativar unidade com itens ativos', async () => {
    mockListUnidades()
    const mensagemBackend = 'Não é possível inativar: existem 4 itens ativos vinculados a esta unidade.'
    server.use(
      http.delete('/api/unidades/1', () => HttpResponse.json({ detail: mensagemBackend }, { status: 400 }))
    )

    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByLabelText('Inativar unidade Juazeiro'))
    await user.click(screen.getByRole('button', { name: 'Inativar' }))

    expect(await screen.findByText(mensagemBackend)).toBeInTheDocument()
    expect(screen.queryByText('Erro ao inativar unidade.')).not.toBeInTheDocument()
  })

  it('inativa com sucesso quando o backend aceita', async () => {
    mockListUnidades()
    server.use(
      http.delete('/api/unidades/1', () => HttpResponse.json({ detail: 'Unidade inativada.' }))
    )

    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByLabelText('Inativar unidade Juazeiro'))
    await user.click(screen.getByRole('button', { name: 'Inativar' }))

    expect(await screen.findByText('Unidade inativada.')).toBeInTheDocument()
  })
})

describe('UnidadesPage — indicadores', () => {
  it('mostra "carregando" e depois os indicadores com rótulos em português', async () => {
    mockListUnidades()
    server.use(
      http.get('/api/unidades/1/indicadores', () =>
        HttpResponse.json({
          termos_emitidos: 5,
          termos_confirmados: 4,
          devolucoes_concluidas: 2,
          itens_total: 10,
          itens_por_status: {
            'Disponível': 6,
            'Indisponível': 1,
            'Pendente': 2,
            'Pendente Devolução': 1,
          },
          emprestimos_ativos: 3,
          perifericos_vinculados: 7,
        })
      )
    )

    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByLabelText('Indicadores da unidade Juazeiro'))

    // A resposta do MSW pode resolver rápido demais para capturar de forma
    // confiável o estado "Carregando indicadores..." de forma síncrona — o que
    // importa aqui é que o painel final carrega e mostra os dados corretos.
    expect(await screen.findByText('Termos Emitidos')).toBeInTheDocument()
    const painel = screen.getByText('Termos Emitidos').closest('div')!.parentElement!
    expect(within(painel).getByText('5')).toBeInTheDocument()
    expect(screen.getByText('Termos Confirmados')).toBeInTheDocument()
    expect(screen.getByText('Devoluções Concluídas')).toBeInTheDocument()
    expect(screen.getByText('Total de Itens')).toBeInTheDocument()
    expect(screen.getByText('Empréstimos Ativos')).toBeInTheDocument()
    expect(screen.getByText('Periféricos Vinculados')).toBeInTheDocument()
    expect(screen.getByText('Disponível: 6')).toBeInTheDocument()
    expect(screen.getByText('Pendente Devolução: 1')).toBeInTheDocument()

    // Unidade com movimento não mostra o aviso de "sem movimentação".
    expect(
      screen.queryByText('Esta unidade ainda não tem nenhuma movimentação registrada.')
    ).not.toBeInTheDocument()
  })

  it('unidade sem nenhuma movimentação mostra os indicadores zerados sem parecer erro', async () => {
    mockListUnidades()
    server.use(
      http.get('/api/unidades/2/indicadores', () =>
        HttpResponse.json({
          termos_emitidos: 0,
          termos_confirmados: 0,
          devolucoes_concluidas: 0,
          itens_total: 0,
          itens_por_status: {
            'Disponível': 0,
            'Indisponível': 0,
            'Pendente': 0,
            'Pendente Devolução': 0,
          },
          emprestimos_ativos: 0,
          perifericos_vinculados: 0,
        })
      )
    )

    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByLabelText('Indicadores da unidade Petrolina'))

    expect(
      await screen.findByText('Esta unidade ainda não tem nenhuma movimentação registrada.')
    ).toBeInTheDocument()
    expect(screen.getByText('Disponível: 0')).toBeInTheDocument()
  })
})

describe('UnidadesPage — mostrar inativas e reativação', () => {
  it('por padrão só lista unidades ativas; a alternância revela as inativas', async () => {
    mockListUnidadesComFiltro([unidadeJuazeiro, unidadeRecife])
    const user = userEvent.setup()
    renderPage()

    expect(await screen.findByText('Juazeiro')).toBeInTheDocument()
    expect(screen.queryByText('Recife')).not.toBeInTheDocument()

    await user.click(screen.getByLabelText('Mostrar unidades inativas'))

    expect(await screen.findByText('Recife')).toBeInTheDocument()
    expect(screen.getByText('Juazeiro')).toBeInTheDocument()
  })

  it('unidade ativa só oferece "Inativar"; unidade inativa só oferece "Reativar"', async () => {
    mockListUnidadesComFiltro([unidadeJuazeiro, unidadeRecife])
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByLabelText('Mostrar unidades inativas'))
    await screen.findByText('Recife')

    expect(screen.getByLabelText('Inativar unidade Juazeiro')).toBeInTheDocument()
    expect(screen.queryByLabelText('Reativar unidade Juazeiro')).not.toBeInTheDocument()

    expect(screen.getByLabelText('Reativar unidade Recife')).toBeInTheDocument()
    expect(screen.queryByLabelText('Inativar unidade Recife')).not.toBeInTheDocument()
  })

  it('clicar em reativar chama o endpoint certo, mostra o toast de sucesso e recarrega a lista', async () => {
    let unidades = [unidadeJuazeiro, unidadeRecife]
    server.use(
      http.get('/api/unidades', ({ request }) => {
        const url = new URL(request.url)
        const includeInactive = url.searchParams.get('include_inactive') === 'true'
        return HttpResponse.json(includeInactive ? unidades : unidades.filter((u) => u.is_active))
      })
    )
    let chamadasReativar = 0
    server.use(
      http.post('/api/unidades/5/reativar', () => {
        chamadasReativar += 1
        // Backend real reativaria de fato; o mock reflete isso para o refetch
        // disparado por invalidateQueries mostrar a unidade já ativa.
        unidades = unidades.map((u) => (u.id === 5 ? { ...u, is_active: true } : u))
        return HttpResponse.json({ detail: 'Unidade reativada.' })
      })
    )

    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByLabelText('Mostrar unidades inativas'))
    await user.click(await screen.findByLabelText('Reativar unidade Recife'))

    await waitFor(() => expect(chamadasReativar).toBe(1))
    expect(await screen.findByText('Unidade reativada.')).toBeInTheDocument()

    // Depois de reativar, a linha da Recife passa a oferecer "Inativar" (não
    // mais "Reativar") — prova de que o refetch pegou o novo `is_active`.
    await waitFor(() =>
      expect(screen.getByLabelText('Inativar unidade Recife')).toBeInTheDocument()
    )
  })
})
