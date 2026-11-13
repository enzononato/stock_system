import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import tab from './linha-do-tempo'
import type { Item } from '@/api/items'

const item: Item = {
  id: 7,
  tipo: 'Notebook',
  brand: 'Dell',
  model: 'Latitude',
  status: 'Disponível',
  revenda: 'Revalle Juazeiro',
}

function renderTab(overrides: Partial<Item> = {}) {
  const Aba = tab.Component
  return renderWithClient(<Aba item={{ ...item, ...overrides }} />)
}

function mockLinha(eventos: unknown[], fontesComErro: string[] = []) {
  server.use(
    http.get('/api/items/7/linha-do-tempo', () =>
      HttpResponse.json({ item_id: 7, eventos, fontes_com_erro: fontesComErro })
    )
  )
}

describe('aba Linha do tempo', () => {
  it('está registrada como aba 20 para Gestor e Técnico', () => {
    expect(tab.id).toBe('linha-do-tempo')
    expect(tab.label).toBe('Linha do tempo')
    expect(tab.order).toBe(20)
    expect(tab.roles).toEqual(['Gestor', 'Técnico'])
  })

  it('mostra os eventos na ordem da API', async () => {
    mockLinha([
      {
        ocorrido_em: '2026-01-01T10:00:00',
        origem: 'historico',
        tipo: 'cadastro',
        titulo: 'Cadastro do equipamento',
        usuario: 'gestor',
        referencia_id: 1,
      },
      {
        ocorrido_em: '2026-01-02T10:00:00',
        origem: 'transferencia',
        tipo: 'transferencia',
        titulo: 'Transferência Revalle Juazeiro → Revalle Petrolina',
        descricao: 'Mudança de filial',
        usuario: 'gestor',
        referencia_id: 2,
      },
      {
        ocorrido_em: '2026-01-03T10:00:00',
        origem: 'baixa',
        tipo: 'baixa',
        titulo: 'Baixa do equipamento — Doação',
        descricao: 'Apagamento de dados: não se aplica',
        usuario: 'gestor',
        referencia_id: 3,
      },
    ])
    renderTab()

    const lista = await screen.findByRole('list', { name: 'Linha do tempo' })
    const itens = within(lista).getAllByRole('listitem')
    expect(itens).toHaveLength(3)
    expect(itens[0]).toHaveTextContent('Cadastro do equipamento')
    expect(itens[1]).toHaveTextContent('Transferência Revalle Juazeiro → Revalle Petrolina')
    expect(itens[1]).toHaveTextContent('Mudança de filial')
    expect(itens[1]).toHaveTextContent('por gestor')
    expect(itens[2]).toHaveTextContent('Baixa do equipamento — Doação')
    expect(itens[2]).toHaveTextContent('Apagamento de dados: não se aplica')
  })

  it('avisa quando uma fonte falhou', async () => {
    mockLinha(
      [{ ocorrido_em: '2026-01-01T10:00:00', origem: 'historico', tipo: 'cadastro', titulo: 'Cadastro do equipamento' }],
      ['fonte_termos']
    )
    renderTab()

    expect(
      await screen.findByText('Parte da linha do tempo não pôde ser carregada. Tente novamente mais tarde.')
    ).toBeInTheDocument()
    expect(screen.getByText('Cadastro do equipamento')).toBeInTheDocument()
  })

  it('evento de origem desconhecida aparece', async () => {
    mockLinha([
      { ocorrido_em: '2026-01-01T10:00:00', origem: 'termo', tipo: 'termo_assinado', titulo: 'Termo assinado' },
    ])
    renderTab()

    expect(await screen.findByText('Termo assinado')).toBeInTheDocument()
  })

  it('sem eventos mostra o texto vazio', async () => {
    mockLinha([])
    renderTab()

    expect(await screen.findByText('Nenhum evento registrado para este equipamento.')).toBeInTheDocument()
  })

  it('erro ao carregar mostra aviso', async () => {
    server.use(http.get('/api/items/7/linha-do-tempo', () => HttpResponse.json({ detail: 'x' }, { status: 500 })))
    renderTab()

    expect(await screen.findByText('Não foi possível carregar a linha do tempo.')).toBeInTheDocument()
  })

  it('mostra o resumo de compra com data sem deslocar o dia', async () => {
    mockLinha([])
    renderTab({ valor_compra: '4599.90', data_compra: '2026-01-10', garantia_ate: '2027-01-10' })

    expect(await screen.findByText('10/01/2027')).toBeInTheDocument()
    expect(screen.getByText('10/01/2026')).toBeInTheDocument()
    expect(screen.getByText('R$ 4.599,90')).toBeInTheDocument()
    expect(screen.getByText('Garantia até')).toBeInTheDocument()
  })

  it('sem dados de compra não mostra o resumo', async () => {
    mockLinha([])
    renderTab()

    await screen.findByText('Nenhum evento registrado para este equipamento.')
    expect(screen.queryByText('Valor de compra')).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Baixar nota fiscal' })).not.toBeInTheDocument()
  })

  it('com NF anexada mostra o botão de baixar', async () => {
    mockLinha([])
    renderTab({ nf_anexo_key: 'notas_fiscais/nf_7_x.pdf' })

    expect(await screen.findByRole('button', { name: 'Baixar nota fiscal' })).toBeInTheDocument()
  })

  it('Transferir unidade abre o diálogo', async () => {
    mockLinha([])
    server.use(
      http.get('/api/unidades', () =>
        HttpResponse.json([
          { id: 1, nome: 'Revalle Juazeiro', razao_social: 'R1', cnpj: '1', is_active: true },
          { id: 2, nome: 'Revalle Petrolina', razao_social: 'R2', cnpj: '2', is_active: true },
        ])
      )
    )
    const user = userEvent.setup()
    renderTab()

    await user.click(await screen.findByRole('button', { name: 'Transferir unidade' }))

    expect(await screen.findByRole('heading', { name: 'Transferir equipamento' })).toBeInTheDocument()
  })
})
