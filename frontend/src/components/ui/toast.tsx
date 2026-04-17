import { useState, useCallback } from 'react'
import { X, CheckCircle, AlertCircle } from 'lucide-react'
import { cn } from '@/lib/utils'

interface Toast {
  id: number
  message: string
  type: 'success' | 'error'
}

let listeners: ((t: Toast) => void)[] = []
let nextId = 1

export function toast(message: string, type: 'success' | 'error' = 'success') {
  const t: Toast = { id: nextId++, message, type }
  listeners.forEach((l) => l(t))
}

export function ToastContainer() {
  const [toasts, setToasts] = useState<Toast[]>([])

  useCallback(() => {
    const handler = (t: Toast) => {
      setToasts((prev) => [...prev, t])
      setTimeout(() => setToasts((prev) => prev.filter((x) => x.id !== t.id)), 4000)
    }
    listeners.push(handler)
    return () => { listeners = listeners.filter((l) => l !== handler) }
  }, [])()

  return (
    <div className="fixed bottom-4 right-4 flex flex-col gap-2 z-50">
      {toasts.map((t) => (
        <div
          key={t.id}
          className={cn(
            'flex items-center gap-3 rounded-lg px-4 py-3 shadow-lg text-sm min-w-[280px] max-w-sm',
            t.type === 'success' ? 'bg-green-600 text-white' : 'bg-red-600 text-white'
          )}
        >
          {t.type === 'success' ? <CheckCircle size={16} /> : <AlertCircle size={16} />}
          <span className="flex-1">{t.message}</span>
          <button onClick={() => setToasts((p) => p.filter((x) => x.id !== t.id))}>
            <X size={14} />
          </button>
        </div>
      ))}
    </div>
  )
}
