import { cn } from '@/lib/utils'

interface BadgeProps {
  children: React.ReactNode
  className?: string
  variant?: 'default' | 'success' | 'warning' | 'danger' | 'info' | 'purple'
  showDot?: boolean
}

/*
 * Meio-termo escolhido pelo dono do produto: o selo em si virou monocromático
 * (a interface já era), mas o status continua identificável à distância pelo
 * ponto de 6px (`dotClasses` abaixo) — só ele carrega cor. Isso evita cinco
 * tons brigando com uma base achromática, sem perder a varredura visual da
 * coluna de status numa tabela cheia.
 *
 * Os nomes das variantes são preservados (success/warning/danger/info/purple)
 * porque as páginas já os usam — trocar os nomes obrigaria a mexer em todas
 * as chamadas sem ganho visual nenhum. O mapa continua com um valor por
 * variante (em vez de uma string única) pelo mesmo motivo: nenhum call site
 * precisa mudar, mesmo que hoje os seis valores sejam idênticos.
 */
const variantClasses: Record<NonNullable<BadgeProps['variant']>, string> = {
  default: 'border-border bg-transparent text-foreground',
  success: 'border-border bg-transparent text-foreground',
  warning: 'border-border bg-transparent text-foreground',
  danger: 'border-border bg-transparent text-foreground',
  info: 'border-border bg-transparent text-foreground',
  purple: 'border-border bg-transparent text-foreground',
}

const dotClasses: Record<NonNullable<BadgeProps['variant']>, string> = {
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
        'inline-flex items-center gap-1.5 rounded-sm border px-2 py-0.5 font-mono text-xs font-medium uppercase tracking-wider select-none transition-colors duration-micro',
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
