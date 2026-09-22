import * as React from 'react'
import { Check } from 'lucide-react'
import { cn } from '@/lib/utils'

export interface CheckboxProps
  extends Omit<React.InputHTMLAttributes<HTMLInputElement>, 'type' | 'onChange'> {
  checked?: boolean
  onCheckedChange?: (checked: boolean) => void
}

export const Checkbox = React.forwardRef<HTMLInputElement, CheckboxProps>(
  ({ className, checked = false, onCheckedChange, disabled, id, ...props }, ref) => {
    // Corrige defeito herdado da FT_STC: sem `id` explícito, o `htmlFor` do
    // label interno (o quadrado visível) ficava `undefined` e o clique nele
    // não ativava o input — só o clique no texto ao lado funcionava. Gera um
    // id interno estável com `useId()` quando a prop não vem preenchida.
    const generatedId = React.useId()
    const inputId = id ?? generatedId
    return (
      <div className="relative inline-flex items-center">
        <input
          type="checkbox"
          id={inputId}
          ref={ref}
          checked={checked}
          disabled={disabled}
          onChange={(e) => onCheckedChange?.(e.target.checked)}
          className="sr-only peer"
          {...props}
        />
        <label
          htmlFor={inputId}
          className={cn(
            'size-4 rounded-sm border border-border flex items-center justify-center cursor-pointer transition-colors select-none',
            'peer-focus-visible:ring-1 peer-focus-visible:ring-ring',
            'peer-disabled:cursor-not-allowed peer-disabled:opacity-50',
            checked ? 'bg-foreground border-foreground text-background' : 'bg-surface hover:border-border-strong',
            className
          )}
        >
          {checked && <Check className="size-3 stroke-[3]" />}
        </label>
      </div>
    )
  }
)
Checkbox.displayName = 'Checkbox'
