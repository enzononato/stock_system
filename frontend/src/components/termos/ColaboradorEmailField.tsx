import type { Colaborador } from '@/api/colaboradores'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'

interface ColaboradorEmailFieldProps {
  colaborador: Colaborador | null
  value: string
  onChange: (value: string) => void
  disabled?: boolean
}

/** E-mail do colaborador que recebe o link de assinatura do termo. Não aparece sem colaborador. */
export function ColaboradorEmailField({ colaborador, value, onChange, disabled }: ColaboradorEmailFieldProps) {
  if (!colaborador) return null
  return (
    <div className="space-y-1.5">
      <Label htmlFor="colaborador-email" className="text-caption font-medium text-foreground">
        E-mail do colaborador *
      </Label>
      <Input
        id="colaborador-email"
        type="email"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        disabled={disabled}
        placeholder="nome@exemplo.com"
        className="h-9"
      />
      <p className="text-xs text-muted-foreground">
        O RevSign envia o link de assinatura para este e-mail (pode ser pessoal).
      </p>
    </div>
  )
}
