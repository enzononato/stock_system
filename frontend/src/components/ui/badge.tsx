import { cn } from '@/lib/utils'

interface BadgeProps {
  children: React.ReactNode
  className?: string
  variant?: 'default' | 'success' | 'warning' | 'danger' | 'info' | 'purple'
  showDot?: boolean
}

/*
 * A interface é monocromática, mas o status não: num inventário a cor do
 * status é sinal funcional, não enfeite. As variantes abaixo apontam para os
 * tokens --status-* (dessaturados) em vez da paleta cheia do Tailwind.
 *
 * Os nomes das variantes são preservados (success/warning/danger/info/purple)
 * porque as páginas já os usam — trocar os nomes obrigaria a mexer em todas
 * as chamadas sem ganho visual nenhum.
 */
const variantClasses = {
  default: 'badge-status-neutral',
  success: 'badge-status-available',
  warning: 'badge-status-pending',
  danger: 'badge-status-danger',
  info: 'badge-status-loaned',
  purple: 'badge-status-return',
}

const dotClasses = {
  default: 'bg-[var(--status-neutral)]',
  success: 'bg-[var(--status-available)]',
  warning: 'bg-[var(--status-pending)]',
  danger: 'bg-[var(--status-danger)]',
  info: 'bg-[var(--status-loaned)]',
  purple: 'bg-[var(--status-return)]',
}

export function Badge({ children, className, variant = 'default', showDot = false }: BadgeProps) {
  return (
    <span
      className={cn(
        'inline-flex items-center gap-1.5 rounded-sm border px-2 py-0.5 text-xs font-semibold tracking-wide select-none transition-colors duration-micro',
        variantClasses[variant],
        className
      )}
    >
      {showDot && <span className={cn('h-1.5 w-1.5 rounded-full', dotClasses[variant])} />}
      {children}
    </span>
  )
}

/** Mapeia os status do backend (pt-BR) para a cor semântica correspondente. */
export function StatusBadge({ status }: { status?: string }) {
  if (!status) return <Badge>-</Badge>
  if (status === 'Disponível') return <Badge variant="success" showDot>{status}</Badge>
  if (status === 'Indisponível') return <Badge variant="info" showDot>{status}</Badge>
  if (status === 'Pendente Devolução') return <Badge variant="purple" showDot>{status}</Badge>
  if (status.startsWith('Pendente')) return <Badge variant="warning" showDot>{status}</Badge>
  return <Badge variant="default">{status}</Badge>
}
