import { useState, useEffect } from 'react'
import { X, CheckCircle, AlertCircle, Info } from 'lucide-react'
import { cn } from '@/lib/utils'

interface ToastAction {
  label: string
  onClick: () => void
}

interface Toast {
  id: number
  message: string
  type: 'success' | 'error' | 'info'
  action?: ToastAction
  duration: number
}

let listeners: ((t: Toast) => void)[] = []
let nextId = 1

interface ToastOptions {
  action?: ToastAction
  duration?: number
}

/** toast(message, type?, options?) — options.action renders a button
 *  (e.g. "Desfazer") inside the toast; options.duration overrides the
 *  default auto-dismiss time (longer by default when there's an action,
 *  so there's time to click it). */
export function toast(message: string, type: 'success' | 'error' | 'info' = 'success', options: ToastOptions = {}) {
  const duration = options.duration ?? (options.action ? 6000 : 4000)
  const t: Toast = { id: nextId++, message, type, action: options.action, duration }
  listeners.forEach((l) => l(t))
}

const ICONS = { success: CheckCircle, error: AlertCircle, info: Info }
const TONE = {
  success: 'bg-success/12 border-success/30 text-success [&_svg]:text-success',
  error: 'bg-destructive/12 border-destructive/30 text-destructive [&_svg]:text-destructive',
  info: 'bg-info/12 border-info/30 text-info [&_svg]:text-info',
}

export function ToastContainer() {
  const [toasts, setToasts] = useState<Toast[]>([])

  useEffect(() => {
    const handler = (t: Toast) => {
      setToasts((prev) => [...prev, t])
      setTimeout(() => setToasts((prev) => prev.filter((x) => x.id !== t.id)), t.duration)
    }
    listeners.push(handler)
    return () => { listeners = listeners.filter((l) => l !== handler) }
  }, [])

  return (
    <div className="fixed bottom-4 right-4 flex flex-col gap-2 z-[60]">
      {toasts.map((t) => {
        const Icon = ICONS[t.type]
        return (
          <div
            key={t.id}
            className={cn(
              'flex items-center gap-3 rounded-lg border px-4 py-3 shadow-lg text-xs font-medium min-w-[280px] max-w-sm bg-card animate-in slide-in-from-bottom-2 fade-in duration-200',
              TONE[t.type]
            )}
          >
            <Icon size={16} className="shrink-0" />
            <span className="flex-1 text-foreground">{t.message}</span>
            {t.action && (
              <button
                onClick={() => {
                  t.action!.onClick()
                  setToasts((p) => p.filter((x) => x.id !== t.id))
                }}
                className="shrink-0 text-primary font-semibold hover:underline"
              >
                {t.action.label}
              </button>
            )}
            <button onClick={() => setToasts((p) => p.filter((x) => x.id !== t.id))} className="shrink-0 text-muted-foreground hover:text-foreground">
              <X size={14} />
            </button>
          </div>
        )
      })}
    </div>
  )
}
