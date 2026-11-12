import { useMemo } from 'react'
import { Link, useParams, useSearchParams } from 'react-router-dom'
import { useQuery } from '@tanstack/react-query'
import { isAxiosError } from 'axios'
import { ArrowLeft } from 'lucide-react'
import { getColaborador, type Colaborador } from '@/api/colaboradores'
import { DetailTabs } from '@/components/ui/DetailTabs'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { StatusPill } from '@/components/ui/StatusPill'
import { useAuth } from '@/contexts/AuthContext'
import { getColaboradorDetailTabs } from '@/features/colaborador-detail/registry'
import { upper } from '@/lib/format'
import type { StatusTone } from '@/lib/status'
import { formatCpf } from '@/lib/utils'

const SITUACOES: Record<Colaborador['situacao'], { label: string; tone: StatusTone }> = {
  ativo: { label: 'Ativo', tone: 'available' },
  afastado: { label: 'Afastado', tone: 'pending' },
  desligado: { label: 'Desligado', tone: 'retired' },
}

export default function ColaboradorDetailPage() {
  const { id } = useParams<{ id: string }>()
  const colaboradorId = Number(id)
  const { user } = useAuth()
  const [searchParams, setSearchParams] = useSearchParams()

  const tabs = useMemo(() => getColaboradorDetailTabs(user?.role), [user?.role])
  const activeTab = tabs.find((t) => t.id === searchParams.get('aba')) ?? tabs[0]

  const { data: colaborador, isLoading, error } = useQuery({
    queryKey: ['colaborador', colaboradorId],
    queryFn: () => getColaborador(colaboradorId),
    enabled: Number.isFinite(colaboradorId),
    retry: (count, err) => !(isAxiosError(err) && err.response?.status === 404) && count < 1,
  })

  const voltar = (
    <Link
      to="/colaboradores"
      className="inline-flex items-center gap-1.5 rounded-sm text-sm font-medium text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2"
    >
      <ArrowLeft className="size-4" aria-hidden />
      Voltar aos colaboradores
    </Link>
  )

  if (isLoading) return <LoadingState label="Carregando colaborador..." />

  const notFound = !Number.isFinite(colaboradorId) || (isAxiosError(error) && error.response?.status === 404)
  if (notFound) {
    return (
      <div className="mx-auto max-w-5xl space-y-4">
        {voltar}
        <EmptyState
          title="Colaborador não encontrado"
          description="O colaborador pode ter sido removido ou o endereço está incorreto."
        />
      </div>
    )
  }
  if (error || !colaborador) return <ErrorState error={error} />

  function selectTab(tabId: string) {
    const next = new URLSearchParams(searchParams)
    next.set('aba', tabId)
    setSearchParams(next, { replace: true })
  }

  const Active = activeTab?.Component
  const situacao = SITUACOES[colaborador.situacao]
  const unidade = upper(colaborador.unidade_nome ?? colaborador.unidade_rh)

  return (
    <div className="mx-auto max-w-5xl space-y-6">
      {voltar}

      <header className="space-y-1 border-b border-border pb-4">
        <div className="flex flex-wrap items-center gap-2">
          <StatusPill status={situacao?.label ?? colaborador.situacao} tone={situacao?.tone ?? 'neutral'} />
          {unidade && <span className="text-caption font-medium text-muted-foreground">{unidade}</span>}
        </div>
        <h1 className="text-heading-lg font-semibold tracking-tight text-foreground">{upper(colaborador.nome)}</h1>
        <p className="flex flex-wrap gap-x-3 text-sm text-muted-foreground">
          <span className="tabular-nums">{formatCpf(colaborador.cpf)}</span>
          {colaborador.cargo && <span>{colaborador.cargo}</span>}
          {colaborador.setor && <span>{colaborador.setor}</span>}
        </p>
      </header>

      <DetailTabs tabs={tabs} activeId={activeTab?.id} onSelect={selectTab} label="Seções do colaborador">
        {Active && <Active colaborador={colaborador} />}
      </DetailTabs>
    </div>
  )
}
