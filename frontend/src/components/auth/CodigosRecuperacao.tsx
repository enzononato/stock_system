import { useState } from 'react'
import { Download, Copy } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Checkbox } from '@/components/ui/checkbox'
import { Label } from '@/components/ui/label'

interface CodigosRecuperacaoProps {
  codigos: string[]
  onContinuar: () => void
  rotuloContinuar?: string
}

/** Os códigos de recuperação do 2FA, exibidos uma única vez: copiar, baixar e confirmar que foram guardados. */
export function CodigosRecuperacao({
  codigos,
  onContinuar,
  rotuloContinuar = 'Continuar',
}: CodigosRecuperacaoProps) {
  const [guardou, setGuardou] = useState(false)
  const [copiado, setCopiado] = useState(false)

  async function copiar() {
    try {
      await navigator.clipboard.writeText(codigos.join('\n'))
      setCopiado(true)
    } catch {
      setCopiado(false)
    }
  }

  function baixar() {
    const blob = new Blob([codigos.join('\n') + '\n'], { type: 'text/plain;charset=utf-8' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'codigos-recuperacao-patrimonio.txt'
    a.click()
    URL.revokeObjectURL(url)
  }

  return (
    <div className="space-y-4">
      <p className="text-sm text-muted-foreground leading-relaxed">
        Guarde estes 10 códigos em local seguro. Cada um serve uma única vez para entrar se você perder o
        celular. Eles não serão mostrados de novo.
      </p>

      <ul className="grid grid-cols-2 gap-x-4 gap-y-1.5 rounded-md border border-border bg-surface-alt p-3 font-mono text-sm">
        {codigos.map((c) => (
          <li key={c}>{c}</li>
        ))}
      </ul>

      <div className="flex flex-wrap gap-2">
        <Button type="button" variant="outline" size="sm" onClick={() => void copiar()}>
          <Copy /> Copiar códigos
        </Button>
        <Button type="button" variant="outline" size="sm" onClick={baixar}>
          <Download /> Baixar .txt
        </Button>
        {copiado && (
          <span role="status" className="self-center text-xs text-muted-foreground">
            Códigos copiados.
          </span>
        )}
      </div>

      <div className="flex items-center gap-2">
        <Checkbox id="guardou-codigos" checked={guardou} onCheckedChange={setGuardou} />
        <Label htmlFor="guardou-codigos" className="text-sm font-normal">
          Guardei os códigos em local seguro
        </Label>
      </div>

      <Button type="button" className="w-full" disabled={!guardou} onClick={onContinuar}>
        {rotuloContinuar}
      </Button>
    </div>
  )
}
