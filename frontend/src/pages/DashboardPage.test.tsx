import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { screen, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { http, HttpResponse } from 'msw'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ThemeProvider } from '@/lib/theme'
import DashboardPage from '@/pages/DashboardPage'
import type { Alerta, AlertasPainel, ResumoPainel } from '@/api/painel'

const papel = vi.hoisted(() => ({ atual: 'Gestor' }))

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'teste', role: papel.atual },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes(papel.atual),
  }),
}))

const STATUS_ZERADOS = {
  Disponível: 0,
  Indisponível: 0,
  Pendente: 0,
  'Pendente Devolução': 0,
  'Em Manutenção': 0,
}

function resumo(sobrescrever: Partial<ResumoPainel> = {}): ResumoPainel {
  return {
    gerado_em: '2026-11-12T10:30:00',
    equipamentos: {
      total: 812,
      por_status: {
        Disponível: 400,
        Indisponível: 300,
        Pendente: 10,
        'Pendente Devolução': 2,
        'Em Manutenção': 100,
      },
      por_unidade: [
        {
          unidade: 'Revalle Juazeiro',
          total: 700,
          por_status: { ...STATUS_ZERADOS, Disponível: 400, Indisponível: 300 },
        },
        { unidade: null, total: 112, por_status: { ...STATUS_ZERADOS, 'Em Manutenção': 100, Pendente: 12 } },
      ],
    },
    manutencao: { equipamentos_em_manutencao: 100, ordens_em_andamento: 98 },
    termos: {
      pendentes: 7,
      por_status: { agendado: 1, enviado: 2, manual: 3, erro: 1 },
      sem_solicitacao: 3,
    },
    linhas: { ativas: 40, custo_mensal: '1996.00', por_unidade: [] },
    ...sobrescrever,
  }
}

function alerta(parcial: Partial<Alerta>): Alerta {
  return {
    chave: 'termo_pendente:1',
    tipo: 'termo_pendente',
    severidade: 'media',
    titulo: 'Termo de entrega pendente há 4 dias',
    descricao: '#1 Notebook Dell · Maria · Assinatura manual',
    link: '/items/1?aba=termos',
    item_id: 1,
    colaborador_id: null,
    desde: '2026-11-08T10:00:00',
    dias: 4,
    ...parcial,
  }
}

function alertas(lista: Alerta[]): AlertasPainel {
  return {
    gerado_em: '2026-11-12T10:30:00',
    limites: { garantia_dias: 30, termo_pendente_dias: 3, devolucao_pendente_dias: 7, manutencao_dias: 15 },
    total: lista.length,
    por_tipo: {},
    alertas: lista,
  }
}

const chamadas = { alertas: 0, resumo: 0 }

function montar(
  opcoes: { resumo?: ResumoPainel; alertas?: AlertasPainel; respostaResumo?: () => Response } = {}
) {
  server.use(
    http.get('/api/painel/resumo', () => {
      chamadas.resumo += 1
      return opcoes.respostaResumo?.() ?? HttpResponse.json(opcoes.resumo ?? resumo())
    }),
    http.get('/api/painel/alertas', () => {
      chamadas.alertas += 1
      return HttpResponse.json(opcoes.alertas ?? alertas([]))
    }),
    http.get('/api/reports/charts/loans', () =>
      HttpResponse.json({ days: [1, 2, 3], values: [1, 0, 2], values2: [0, 1, 0] })
    )
  )
  return renderWithClient(
    <ThemeProvider>
      <MemoryRouter>
        <DashboardPage />
      </MemoryRouter>
    </ThemeProvider>
  )
}

beforeAll(() => {
  // O ResponsiveContainer do recharts exige ResizeObserver, que o jsdom não traz.
  vi.stubGlobal('ResizeObserver', class { observe() {} unobserve() {} disconnect() {} })
})

beforeEach(() => {
  papel.atual = 'Gestor'
  chamadas.alertas = 0
  chamadas.resumo = 0
})

describe('DashboardPage', () => {
  it('mostra os indicadores do resumo (sem o limite de 500 do cliente)', async () => {
    montar()
    const indicadores = within(await screen.findByRole('region', { name: 'Indicadores' }))
    expect(await indicadores.findByText('812')).toBeInTheDocument()
    expect(indicadores.getByText('49% do total')).toBeInTheDocument()
    expect(indicadores.getByText('300')).toBeInTheDocument()
    expect(indicadores.getByText('12')).toBeInTheDocument()
    expect(indicadores.getByText('10 empréstimo(s) · 2 devolução(ões)')).toBeInTheDocument()
    expect(indicadores.getByText('98 ordem(ns) em andamento')).toBeInTheDocument()
    expect(indicadores.getByText('+3 sem solicitação')).toBeInTheDocument()
    expect(indicadores.getByText(/^R\$\s1\.996,00\/mês$/)).toBeInTheDocument()
    expect(screen.getByRole('heading', { name: 'Painel' })).toBeInTheDocument()
    expect(screen.getByText(/Atualizado em/)).toBeInTheDocument()
    for (const rotulo of [
      'Equipamentos',
      'Disponíveis',
      'Em uso',
      'Pendentes',
      'Em manutenção',
      'Termos pendentes',
      'Linhas ativas',
    ]) {
      expect(indicadores.getByText(rotulo)).toBeInTheDocument()
    }
  })

  it('sem solicitações legadas o termo mostra "aguardando assinatura"', async () => {
    const base = resumo()
    montar({ resumo: resumo({ termos: { ...base.termos, sem_solicitacao: 0 } }) })
    expect(await screen.findByText('aguardando assinatura')).toBeInTheDocument()
  })

  it('lista os alertas agrupados com link para o registro', async () => {
    montar({
      alertas: alertas([
        alerta({
          chave: 'garantia_vencendo:5',
          tipo: 'garantia_vencendo',
          severidade: 'baixa',
          titulo: 'Garantia vence em 10 dia(s)',
          link: '/items/5',
        }),
        alerta({
          chave: 'colaborador_desligado:9',
          tipo: 'colaborador_desligado',
          severidade: 'alta',
          titulo: 'Maria Souza (desligado) ainda está com equipamento',
          link: '/items/9',
          item_id: 9,
        }),
      ]),
    })
    const desligado = await screen.findByText('Colaborador desligado com equipamento')
    const garantia = screen.getByText('Garantia vencendo')
    expect(desligado.compareDocumentPosition(garantia) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy()
    const link = screen.getByRole('link', { name: /Maria Souza \(desligado\)/ })
    expect(link).toHaveAttribute('href', '/items/9')
    expect(within(link).getByText('Alta')).toBeInTheDocument()
    expect(screen.getByText(/Limites: garantia em até 30 dias, termo há mais de 3 dias/)).toBeInTheDocument()
  })

  it('mostra até 5 por grupo e expande', async () => {
    const lista = Array.from({ length: 7 }, (_, i) =>
      alerta({ chave: `termo_pendente:${i}`, titulo: `Termo pendente ${i}`, link: `/items/${i}?aba=termos` })
    )
    montar({ alertas: alertas(lista) })
    expect(await screen.findByText('Termo pendente 0')).toBeInTheDocument()
    expect(screen.getAllByRole('link', { name: /Termo pendente \d/ })).toHaveLength(5)
    await userEvent.click(screen.getByRole('button', { name: 'Mostrar todos (7)' }))
    expect(screen.getAllByRole('link', { name: /Termo pendente \d/ })).toHaveLength(7)
    await userEvent.click(screen.getByRole('button', { name: 'Mostrar menos' }))
    expect(screen.getAllByRole('link', { name: /Termo pendente \d/ })).toHaveLength(5)
  })

  it('sem alertas mostra a mensagem', async () => {
    montar()
    expect(await screen.findByText('Nenhum alerta no momento.')).toBeInTheDocument()
  })

  it('Jovem Aprendiz vê só os indicadores e não chama /api/painel/alertas', async () => {
    papel.atual = 'Jovem Aprendiz'
    montar()
    expect(await screen.findByText('812')).toBeInTheDocument()
    expect(screen.queryByText('Alertas')).not.toBeInTheDocument()
    expect(screen.queryByText('Nenhum alerta no momento.')).not.toBeInTheDocument()
    expect(chamadas.alertas).toBe(0)
    expect(screen.queryByRole('link', { name: /Manutenções/ })).not.toBeInTheDocument()
  })

  it('unidade sem nome aparece como SEM UNIDADE e os nomes em caixa alta', async () => {
    montar()
    expect(await screen.findByText('REVALLE JUAZEIRO')).toBeInTheDocument()
    expect(screen.getByText('SEM UNIDADE')).toBeInTheDocument()
  })

  it('unidades vazias mostram a mensagem', async () => {
    const base = resumo()
    montar({ resumo: resumo({ equipamentos: { ...base.equipamentos, total: 0, por_unidade: [] } }) })
    expect(await screen.findByText('Nenhum equipamento cadastrado.')).toBeInTheDocument()
  })

  it('não consulta a lista de itens nem o histórico', async () => {
    let chamouItens = false
    server.use(
      http.get('/api/items', () => {
        chamouItens = true
        return HttpResponse.json({ items: [], total: 0 })
      })
    )
    montar()
    await screen.findByText('812')
    expect(chamouItens).toBe(false)
    expect(screen.queryByText(/Últimas Movimentações/)).not.toBeInTheDocument()
  })

  it('erro no resumo oferece tentar novamente', async () => {
    let tentativas = 0
    montar({
      respostaResumo: () => {
        tentativas += 1
        return tentativas === 1
          ? HttpResponse.json({ detail: 'Falha' }, { status: 500 })
          : HttpResponse.json(resumo())
      },
    })
    await userEvent.click(await screen.findByRole('button', { name: 'Tentar novamente' }))
    expect(await screen.findByText('812')).toBeInTheDocument()
  })

  it('o botão Atualizar refaz as consultas do painel', async () => {
    montar()
    await screen.findByText('812')
    const antes = chamadas.resumo
    await userEvent.click(screen.getByRole('button', { name: /Atualizar/ }))
    await vi.waitFor(() => expect(chamadas.resumo).toBe(antes + 1))
  })

  it('o gestor vê o atalho para Manutenções nos fluxos rápidos', async () => {
    montar()
    await screen.findByText('812')
    expect(screen.getByRole('link', { name: /Manutenções/ })).toHaveAttribute('href', '/manutencoes')
  })
})
