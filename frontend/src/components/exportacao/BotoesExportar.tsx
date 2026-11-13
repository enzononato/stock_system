import { useState } from 'react'
import { Download } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import type { FormatoExportacao } from '@/features/exportacoes/api'

interface BotoesExportarProps {
  onExportar: (formato: FormatoExportacao) => Promise<void>
  disabled?: boolean
  rotulo?: string
}

const FORMATOS: { formato: FormatoExportacao; texto: string }[] = [
  { formato: 'xlsx', texto: 'Excel' },
  { formato: 'pdf', texto: 'PDF' },
]

/** Par de botões Excel/PDF; só um formato exporta por vez e o erro vira toast. */
export function BotoesExportar({ onExportar, disabled = false, rotulo = 'Exportar' }: BotoesExportarProps) {
  const [exportando, setExportando] = useState<FormatoExportacao | null>(null)

  async function exportar(formato: FormatoExportacao) {
    setExportando(formato)
    try {
      await onExportar(formato)
    } catch (erro) {
      toast(getErrorMessage(erro, 'Não foi possível exportar.'), 'error')
    } finally {
      setExportando(null)
    }
  }

  return (
    <div role="group" aria-label={rotulo} className="inline-flex items-center gap-2">
      {FORMATOS.map(({ formato, texto }) => (
        <Button
          key={formato}
          type="button"
          size="sm"
          variant="outline"
          aria-label={`${rotulo} em ${texto}`}
          disabled={disabled || (exportando !== null && exportando !== formato)}
          loading={exportando === formato}
          onClick={() => exportar(formato)}
        >
          <Download />
          {texto}
        </Button>
      ))}
    </div>
  )
}
