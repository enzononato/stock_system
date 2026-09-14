import type { ReactNode } from 'react'
import { cn } from '@/lib/utils'

interface StatBlockProps {
  /** Rótulo curto em caixa alta (ex.: "Total em Estoque"). */
  label: ReactNode
  /** Valor numérico (ou já formatado) em destaque, tipografia mono. */
  value: ReactNode
  /** Texto curto opcional ao lado do valor (ex.: unidade, variação). */
  hint?: ReactNode
  /** Caractere mono opcional (ex.: "#", "↑", "↓", "◯") no lugar de um ícone. */
  symbol?: ReactNode
  className?: string
}

/**
 * Bloco de métrica "Enterprise Editorial": legenda + valor mono + símbolo
 * opcional, sem ícone em tile — substitui os cartões com tile de 48×48 e
 * ícone lucide, que eram o elemento mais "SaaS genérico" da tela. Ver
 * `components/data/stat-block.tsx` e o bloco de estatísticas de
 * `features/reports/ReportPage.tsx` na referência (origin/redesign-frontend).
 */
export function StatBlock({ label, value, hint, symbol, className }: StatBlockProps) {
  return (
    <div className={cn('flex flex-col gap-0.5', className)}>
      <span className="text-caption text-muted-foreground">{label}</span>
      <div className="flex items-baseline gap-2">
        <span className="text-heading-lg font-semibold tracking-tight text-foreground num font-mono">{value}</span>
        {hint && <span className="text-body-sm text-muted-foreground">{hint}</span>}
        {symbol && <span className="text-body-lg text-muted-foreground font-mono">{symbol}</span>}
      </div>
    </div>
  )
}
