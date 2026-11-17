import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import type { LayoutEtiquetas } from '@/features/etiquetas/layout'

interface LayoutFormProps {
  layout: LayoutEtiquetas
  posicaoInicial: number
  onLayoutChange: (layout: LayoutEtiquetas) => void
  onPosicaoChange: (posicao: number) => void
  onRestaurar: () => void
}

const CAMPOS: { chave: keyof LayoutEtiquetas; rotulo: string; mm: boolean }[] = [
  { chave: 'colunas', rotulo: 'Colunas', mm: false },
  { chave: 'linhas', rotulo: 'Linhas', mm: false },
  { chave: 'larguraMm', rotulo: 'Largura (mm)', mm: true },
  { chave: 'alturaMm', rotulo: 'Altura (mm)', mm: true },
  { chave: 'margemSuperiorMm', rotulo: 'Margem superior (mm)', mm: true },
  { chave: 'margemEsquerdaMm', rotulo: 'Margem esquerda (mm)', mm: true },
  { chave: 'espacoHorizontalMm', rotulo: 'Espaço entre colunas (mm)', mm: true },
  { chave: 'espacoVerticalMm', rotulo: 'Espaço entre linhas (mm)', mm: true },
]

/** Campo vazio vira NaN (o layout fica inválido e o aviso aparece) em vez de virar 0 sozinho. */
function numero(texto: string): number {
  return texto.trim() === '' ? Number.NaN : Number(texto)
}

function exibir(n: number): string {
  return Number.isNaN(n) ? '' : String(n)
}

/** Colunas, linhas, tamanho, margens e posição inicial da folha de etiquetas. */
export function LayoutForm({ layout, posicaoInicial, onLayoutChange, onPosicaoChange, onRestaurar }: LayoutFormProps) {
  return (
    <div className="space-y-3">
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-5">
        {CAMPOS.map((campo) => (
          <label key={campo.chave} className="flex flex-col gap-1 text-caption text-muted-foreground">
            {campo.rotulo}
            <Input
              type="number"
              inputMode="decimal"
              step={campo.mm ? '0.5' : '1'}
              value={exibir(layout[campo.chave])}
              onChange={(e) => onLayoutChange({ ...layout, [campo.chave]: numero(e.target.value) })}
            />
          </label>
        ))}
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Posição inicial
          <Input
            type="number"
            inputMode="numeric"
            step="1"
            min={1}
            value={exibir(posicaoInicial)}
            onChange={(e) => onPosicaoChange(numero(e.target.value))}
          />
        </label>
      </div>
      <Button type="button" variant="outline" size="sm" onClick={onRestaurar}>
        Restaurar padrão
      </Button>
    </div>
  )
}
