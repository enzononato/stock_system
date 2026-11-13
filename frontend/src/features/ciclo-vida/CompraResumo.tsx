import { Download } from 'lucide-react'
import { baixarNotaFiscal, type Item } from '@/api/items'
import { Button } from '@/components/ui/button'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { formatMoeda } from '@/lib/format'
import { isoParaBR } from './format'

/**
 * Resumo da compra (valor, data, garantia, NF). Não mostra nada quando o equipamento
 * não tem nenhum desses dados. Datas só-dia vêm como `AAAA-MM-DD` e são convertidas por
 * texto (`new Date` mostraria um dia a menos no fuso de Brasília).
 */
export function CompraResumo({ item }: { item: Item }) {
  const campos = [
    { rotulo: 'Valor de compra', valor: formatMoeda(item.valor_compra) },
    { rotulo: 'Data da compra', valor: isoParaBR(item.data_compra) },
    { rotulo: 'Garantia até', valor: isoParaBR(item.garantia_ate) },
  ].filter((c) => c.valor !== '')
  const temNf = !!item.nf_anexo_key
  if (campos.length === 0 && !temNf) return null

  return (
    <section aria-label="Compra e garantia" className="space-y-3 rounded border border-border bg-surface-alt p-4">
      <h3 className="text-heading-sm text-foreground">Compra e garantia</h3>
      <dl className="grid grid-cols-1 gap-x-6 gap-y-3 sm:grid-cols-3">
        {campos.map((c) => (
          <div key={c.rotulo}>
            <dt className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">{c.rotulo}</dt>
            <dd className="mt-0.5 text-sm font-medium text-foreground">{c.valor}</dd>
          </div>
        ))}
      </dl>
      {temNf && (
        <Button
          type="button"
          variant="outline"
          size="sm"
          onClick={() =>
            void baixarNotaFiscal(item).catch((err: unknown) =>
              toast(getErrorMessage(err, 'Não foi possível baixar a nota fiscal.'), 'error')
            )
          }
        >
          <Download className="mr-1.5 size-3.5" />
          Baixar nota fiscal
        </Button>
      )}
    </section>
  )
}
