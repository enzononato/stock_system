import { Link } from 'react-router-dom'
import { ShieldAlert } from 'lucide-react'
import { Button } from '@/components/ui/button'

interface RestrictedAccessProps {
  /** Papéis que podem abrir a rota. */
  roles: string[]
  /** Papel do usuário logado (pode ser undefined se, por algum motivo, não houver usuário). */
  currentRole?: string
}

/**
 * Tela mostrada no lugar de uma rota protegida quando o papel do usuário
 * logado não está entre os permitidos. Substitui o antigo
 * `<Navigate to="/" replace />` silencioso de `RequireRole` (App.tsx): antes
 * o usuário clicava em algo, era jogado de volta para o Dashboard e não
 * entendia por quê. Aqui a explicação e o caminho de volta ficam explícitos
 * (mesma ideia de `origin/redesign-frontend:components/app/RequireRole.tsx`,
 * escrita em português e nos componentes deste projeto).
 */
export function RestrictedAccess({ roles, currentRole }: RestrictedAccessProps) {
  return (
    <div className="flex flex-col items-center justify-center gap-4 surface-panel py-16 px-6 text-center">
      <div className="h-12 w-12 rounded-full bg-surface-alt border border-border flex items-center justify-center text-muted-foreground">
        <ShieldAlert size={22} aria-hidden="true" />
      </div>
      <div className="space-y-1.5 max-w-md">
        <h2 className="text-heading-sm font-semibold text-foreground">Acesso restrito</h2>
        <p className="text-body-sm text-muted-foreground">
          Esta área é exclusiva para {roles.join(' ou ')}. Seu perfil atual é{' '}
          <span className="font-medium text-foreground">{currentRole ?? 'desconhecido'}</span>, por isso
          você foi impedido de continuar.
        </p>
      </div>
      <Button asChild variant="outline" size="sm">
        <Link to="/">Voltar ao Dashboard</Link>
      </Button>
    </div>
  )
}
