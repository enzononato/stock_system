import type { ReactNode } from 'react'
import { cn } from '@/lib/utils'

interface PageHeaderProps {
  /** Rótulo curto de domínio, à esquerda do separador "•" (ex.: "Auditoria e Rastreabilidade"). */
  eyebrow?: ReactNode
  /** Rótulo mais específico, à direita do separador — só aparece junto com `eyebrow`. */
  eyebrowDetail?: ReactNode
  /** Título principal da página, renderizado como <h1>. */
  title: ReactNode
  /** Frase curta descrevendo o que a tela faz. */
  description?: ReactNode
  /** Slot opcional para botões de ação (ex.: "Novo Equipamento", "Exportar"). */
  actions?: ReactNode
  className?: string
}

/**
 * Cabeçalho de página no padrão "Enterprise Editorial": uma linha de contexto
 * em caixa alta (eyebrow) acima do título, seguida de uma descrição curta.
 * Reproduz o padrão escrito à mão em cada página de referência (não o
 * componente `PageHeader` compartilhado de lá, que depende de classes
 * inexistentes no nosso CSS) — ver anotações da tarefa.
 */
export function PageHeader({ eyebrow, eyebrowDetail, title, description, actions, className }: PageHeaderProps) {
  return (
    <div
      className={cn(
        'flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4',
        className
      )}
    >
      <div>
        {eyebrow && (
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              {eyebrow}
            </span>
            {eyebrowDetail && (
              <>
                <span className="text-muted-foreground">•</span>
                <span className="text-caption text-foreground font-medium">{eyebrowDetail}</span>
              </>
            )}
          </div>
        )}
        <h1 className="text-heading font-semibold tracking-tight text-foreground mt-0.5">{title}</h1>
        {description && <p className="text-body-sm text-muted-foreground mt-1">{description}</p>}
      </div>
      {actions && <div className="flex items-center gap-3">{actions}</div>}
    </div>
  )
}

interface PanelHeaderProps {
  /** Rótulo da faixa de legenda do painel (ex.: "Filtros de Consulta"). */
  title: ReactNode
  description?: ReactNode
  /** Slot opcional para ações alinhadas à direita da faixa (ex.: um botão de atualizar). */
  actions?: ReactNode
  className?: string
}

/**
 * Faixa de legenda usada no topo de painéis (filtros, formulários, listas)
 * para dar a cada bloco uma identidade nomeada, no mesmo padrão tipográfico
 * do `PageHeader`, em escala menor.
 */
export function PanelHeader({ title, description, actions, className }: PanelHeaderProps) {
  return (
    <div className={cn('border-b border-border pb-2.5 flex items-center justify-between gap-3', className)}>
      <div>
        <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">{title}</span>
        {description && <p className="text-body-sm text-muted-foreground mt-1">{description}</p>}
      </div>
      {actions && <div className="flex items-center gap-2 shrink-0">{actions}</div>}
    </div>
  )
}
