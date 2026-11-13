import type { ResumoPainel } from '@/api/painel'
import { StatBlock } from '@/components/ui/StatBlock'
import { formatMoeda } from '@/lib/format'

/** Faixa de indicadores do painel; todos os números vêm do servidor (`GET /api/painel/resumo`). */
export function IndicadoresPainel({ resumo }: { resumo: ResumoPainel }) {
  const { equipamentos, manutencao, termos, linhas } = resumo
  const porStatus = equipamentos.por_status
  const total = equipamentos.total
  const disponiveis = porStatus['Disponível'] ?? 0
  const pendentes = porStatus['Pendente'] ?? 0
  const pendentesDevolucao = porStatus['Pendente Devolução'] ?? 0
  const pct = total > 0 ? Math.round((disponiveis / total) * 100) : 0

  return (
    <section
      aria-label="Indicadores"
      className="grid grid-cols-2 gap-4 sm:grid-cols-4 xl:grid-cols-7 border-b border-border pb-6"
    >
      <StatBlock label="Equipamentos" value={total} hint="ativos cadastrados" />
      <StatBlock label="Disponíveis" value={disponiveis} hint={`${pct}% do total`} />
      <StatBlock label="Em uso" value={porStatus['Indisponível'] ?? 0} />
      <StatBlock
        label="Pendentes"
        value={pendentes + pendentesDevolucao}
        hint={`${pendentes} empréstimo(s) · ${pendentesDevolucao} devolução(ões)`}
      />
      <StatBlock
        label="Em manutenção"
        value={manutencao.equipamentos_em_manutencao}
        hint={`${manutencao.ordens_em_andamento} ordem(ns) em andamento`}
      />
      <StatBlock
        label="Termos pendentes"
        value={termos.pendentes}
        hint={termos.sem_solicitacao > 0 ? `+${termos.sem_solicitacao} sem solicitação` : 'aguardando assinatura'}
      />
      <StatBlock label="Linhas ativas" value={linhas.ativas} hint={`${formatMoeda(linhas.custo_mensal)}/mês`} />
    </section>
  )
}
