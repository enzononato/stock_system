import { Check } from 'lucide-react'
import { cn } from '@/lib/utils'

export interface Step {
  id: number | string
  label: string
  description?: string
}

interface StepperProps {
  steps: Step[]
  currentStep: number
  onStepClick?: (stepIndex: number) => void
  className?: string
}

export function Stepper({ steps, currentStep, onStepClick, className }: StepperProps) {
  return (
    <div className={cn('w-full py-2', className)}>
      <nav aria-label="Progresso da operação">
        <ol className="flex items-center justify-between gap-2">
          {steps.map((step, idx) => {
            const stepNumber = idx + 1
            const isCompleted = currentStep > idx
            const isCurrent = currentStep === idx
            const isClickable = Boolean(onStepClick && isCompleted)

            return (
              <li key={step.id ?? idx} className="flex-1 flex items-center gap-3">
                <button
                  type="button"
                  disabled={!isClickable}
                  onClick={() => isClickable && onStepClick?.(idx)}
                  className={cn(
                    'flex items-center gap-2.5 text-left transition-colors',
                    isClickable ? 'cursor-pointer' : 'cursor-default'
                  )}
                >
                  {/* Indicador monocromático */}
                  <div
                    className={cn(
                      'size-7 shrink-0 rounded flex items-center justify-center text-xs font-mono font-semibold border transition-colors',
                      isCompleted && 'bg-foreground text-background border-foreground font-semibold',
                      isCurrent && 'border-foreground bg-surface text-foreground font-bold shadow-sm',
                      !isCompleted && !isCurrent && 'border-border text-muted-foreground bg-transparent'
                    )}
                  >
                    {isCompleted ? <Check className="size-3.5 stroke-[2.5]" /> : stepNumber}
                  </div>

                  <div className="hidden sm:flex flex-col min-w-0">
                    <span
                      className={cn(
                        'text-caption font-semibold uppercase tracking-wider',
                        isCurrent && 'text-foreground',
                        isCompleted && 'text-foreground',
                        !isCompleted && !isCurrent && 'text-muted-foreground'
                      )}
                    >
                      {step.label}
                    </span>
                    {step.description && (
                      <span className="text-[11px] text-muted-foreground truncate">
                        {step.description}
                      </span>
                    )}
                  </div>
                </button>

                {idx < steps.length - 1 && (
                  <div
                    className={cn(
                      'flex-1 h-[1px] transition-colors',
                      isCompleted ? 'bg-border-strong' : 'bg-border'
                    )}
                  />
                )}
              </li>
            )
          })}
        </ol>
      </nav>
    </div>
  )
}
