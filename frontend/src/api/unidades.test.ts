import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { server } from '@/test/server'
import {
  listUnidades,
  getUnidade,
  getIndicadoresUnidade,
  createUnidade,
  updateUnidade,
  deactivateUnidade,
  reactivateUnidade,
} from '@/api/unidades'

const unidadePetrolina = {
  id: 3,
  nome: 'Petrolina',
  razao_social: 'Revalle Petrolina Ltda',
  cnpj: '11.222.333/0001-81',
  cidade: 'Petrolina',
  uf: 'PE',
  is_active: true,
  // Sem `cep` de propósito — não constava no documento de origem.
}

describe('listUnidades', () => {
  it('sem parâmetros, não envia include_inactive', async () => {
    let url: URL | null = null
    server.use(
      http.get('/api/unidades', ({ request }) => {
        url = new URL(request.url)
        return HttpResponse.json([unidadePetrolina])
      })
    )

    const resultado = await listUnidades()

    expect(url!.search).toBe('')
    expect(resultado).toEqual([unidadePetrolina])
  })

  it('repassa include_inactive como query param quando fornecido', async () => {
    let queryRecebida: Record<string, string> = {}
    server.use(
      http.get('/api/unidades', ({ request }) => {
        const url = new URL(request.url)
        queryRecebida = Object.fromEntries(url.searchParams.entries())
        return HttpResponse.json([unidadePetrolina])
      })
    )

    await listUnidades({ include_inactive: true })

    expect(queryRecebida).toEqual({ include_inactive: 'true' })
  })
})

describe('getUnidade', () => {
  it('busca a unidade pelo id', async () => {
    server.use(
      http.get('/api/unidades/3', () => HttpResponse.json(unidadePetrolina))
    )

    const resultado = await getUnidade(3)

    expect(resultado).toEqual(unidadePetrolina)
  })
})

describe('getIndicadoresUnidade', () => {
  it('busca os indicadores da unidade pelo id', async () => {
    const indicadores = {
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
    }
    server.use(
      http.get('/api/unidades/3/indicadores', () => HttpResponse.json(indicadores))
    )

    const resultado = await getIndicadoresUnidade(3)

    expect(resultado).toEqual(indicadores)
  })
})

describe('createUnidade', () => {
  it('envia os dados da unidade via POST', async () => {
    let corpoRecebido: unknown = null
    server.use(
      http.post('/api/unidades', async ({ request }) => {
        corpoRecebido = await request.json()
        return HttpResponse.json({ detail: 'Unidade criada com sucesso.' })
      })
    )

    const resultado = await createUnidade({
      nome: 'Juazeiro',
      razao_social: 'Revalle Juazeiro Ltda',
      cnpj: '11.222.333/0001-81',
    })

    expect(corpoRecebido).toEqual({
      nome: 'Juazeiro',
      razao_social: 'Revalle Juazeiro Ltda',
      cnpj: '11.222.333/0001-81',
    })
    expect(resultado).toEqual({ detail: 'Unidade criada com sucesso.' })
  })
})

describe('updateUnidade', () => {
  it('envia os dados atualizados via PUT para o id correto', async () => {
    let corpoRecebido: unknown = null
    server.use(
      http.put('/api/unidades/3', async ({ request }) => {
        corpoRecebido = await request.json()
        return HttpResponse.json({ detail: 'Unidade atualizada com sucesso.' })
      })
    )

    const resultado = await updateUnidade(3, {
      nome: 'Petrolina',
      razao_social: 'Revalle Petrolina Ltda',
      cnpj: '11.222.333/0001-81',
      cidade: 'Petrolina',
      uf: 'PE',
    })

    expect(corpoRecebido).toEqual({
      nome: 'Petrolina',
      razao_social: 'Revalle Petrolina Ltda',
      cnpj: '11.222.333/0001-81',
      cidade: 'Petrolina',
      uf: 'PE',
    })
    expect(resultado).toEqual({ detail: 'Unidade atualizada com sucesso.' })
  })
})

describe('deactivateUnidade', () => {
  it('chama DELETE no id da unidade e devolve o detail de sucesso', async () => {
    server.use(
      http.delete('/api/unidades/3', () => HttpResponse.json({ detail: 'Unidade inativada.' }))
    )

    const resultado = await deactivateUnidade(3)

    expect(resultado).toEqual({ detail: 'Unidade inativada.' })
  })

  it('propaga o erro (com o detail do backend) quando a unidade tem itens ativos', async () => {
    server.use(
      http.delete('/api/unidades/3', () =>
        HttpResponse.json(
          { detail: 'Não é possível inativar: existem 4 itens ativos vinculados a esta unidade.' },
          { status: 400 }
        )
      )
    )

    await expect(deactivateUnidade(3)).rejects.toMatchObject({
      response: { status: 400, data: { detail: 'Não é possível inativar: existem 4 itens ativos vinculados a esta unidade.' } },
    })
  })
})

describe('reactivateUnidade', () => {
  it('chama POST /api/unidades/{id}/reativar e devolve o detail de sucesso', async () => {
    let chamadas = 0
    server.use(
      http.post('/api/unidades/3/reativar', () => {
        chamadas += 1
        return HttpResponse.json({ detail: 'Unidade reativada.' })
      })
    )

    const resultado = await reactivateUnidade(3)

    expect(chamadas).toBe(1)
    expect(resultado).toEqual({ detail: 'Unidade reativada.' })
  })
})
