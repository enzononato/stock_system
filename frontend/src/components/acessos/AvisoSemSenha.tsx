import { ShieldAlert } from 'lucide-react'
import { cn } from '@/lib/utils'

export const AVISO_SEM_SENHA =
  'O Patrimônio nunca pede nem guarda senhas: registre só o login. Não digite senhas em nenhum campo.'

/** Aviso fixo dos formulários do módulo de Acessos: o sistema só registra o login. */
export function AvisoSemSenha({ className }: { className?: string }) {
  return (
    <p
      role="note"
      className={cn(
        'flex items-start gap-2 rounded-md bg-brand-soft p-3 text-sm text-brand-soft-foreground',
        className
      )}
    >
      <ShieldAlert className="mt-0.5 h-4 w-4 shrink-0" aria-hidden="true" />
      <span>{AVISO_SEM_SENHA}</span>
    </p>
  )
}
