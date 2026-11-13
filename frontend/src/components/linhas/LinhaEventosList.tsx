import type { LinhaEvento } from '@/api/linhas'
import { descreverEvento, ROTULO_EVENTO } from '@/features/linhas/eventos'
import { formatTelefone } from '@/lib/format'
import { formatDateTime } from '@/lib/utils'

interface LinhaEventosListProps {
  eventos: LinhaEvento[]
  /** Mostra o número da linha antes do tipo (aba do aparelho, onde passam várias linhas). */
  mostrarNumero?: boolean
}

export function LinhaEventosList({ eventos, mostrarNumero = false }: LinhaEventosListProps) {
  if (eventos.length === 0) {
    return <p className="text-sm text-muted-foreground">Nenhum evento registrado.</p>
  }
  return (
    <ul className="divide-y divide-border rounded-lg border border-border">
      {eventos.map((e) => (
        <li key={e.id} className="space-y-0.5 px-4 py-3">
          <p className="flex flex-wrap items-center gap-x-2 text-sm font-semibold text-foreground">
            {mostrarNumero && <span className="tabular-nums">{formatTelefone(e.linha_numero)}</span>}
            <span>{ROTULO_EVENTO[e.tipo] ?? e.tipo}</span>
          </p>
          <p className="text-sm text-foreground">{descreverEvento(e)}</p>
          <p className="text-xs text-muted-foreground">
            {[e.por_usuario, formatDateTime(e.ocorrido_em)].filter(Boolean).join(' · ')}
          </p>
          {e.observacao && <p className="text-xs text-muted-foreground">{e.observacao}</p>}
        </li>
      ))}
    </ul>
  )
}
