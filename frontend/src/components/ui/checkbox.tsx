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
    return (
      <div className="relative inline-flex items-center">
        <input
          type="checkbox"
          id={id}
          ref={ref}
          checked={checked}
          disabled={disabled}
          onChange={(e) => onCheckedChange?.(e.target.checked)}
          className="sr-only peer"
          {...props}
        />
        <label
          htmlFor={id}
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
