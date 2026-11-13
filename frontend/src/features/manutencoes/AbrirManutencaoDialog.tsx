import { useState } from 'react'
import { keepPreviousData, useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { listItemsPaginated, type Item } from '@/api/items'
import { abrirManutencao } from '@/api/manutencoes'
import { Button } from '@/components/ui/button'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { SearchableSelect } from '@/components/ui/SearchableSelect'
import { toast } from '@/components/ui/toast'
import { useDebouncedValue } from '@/hooks/useDebouncedValue'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'

type ItemResumo = Pick<Item, 'id' | 'tipo' | 'brand' | 'model' | 'identificador'>

export interface AbrirManutencaoDialogProps {
  /** Equipamento fixo (aba do detalhe). Sem ele, o diálogo oferece a busca de equipamentos Disponíveis. */
  item?: ItemResumo | null
  open: boolean
  onOpenChange: (open: boolean) => void
  onAberta?: (id: number) => void
}

const ITEM_SEARCH_LIMIT = 20

function descricao(i: ItemResumo): string {
  return [i.tipo, i.brand, i.model].filter(Boolean).join(' ')
}

/** Abre uma ordem de manutenção: defeito (obrigatório), assistência e observações (opcionais). */
export function AbrirManutencaoDialog({ item, open, onOpenChange, onAberta }: AbrirManutencaoDialogProps) {
  const queryClient = useQueryClient()
  const [itemId, setItemId] = useState('')
  const [busca, setBusca] = useState('')
  const [defeito, setDefeito] = useState('')
  const [assistencia, setAssistencia] = useState('')
  const [observacoes, setObservacoes] = useState('')

  // Busca no servidor (debounce de 300 ms): a lista nunca é carregada inteira.
  const buscaAplicada = useDebouncedValue(busca.trim(), 300)
  const { data } = useQuery({
    queryKey: ['items', 'select', 'Disponível', buscaAplicada],
    queryFn: () =>
      listItemsPaginated({ status: 'Disponível', search: buscaAplicada || undefined, limit: ITEM_SEARCH_LIMIT }),
    placeholderData: keepPreviousData,
    enabled: open && !item,
  })
  const disponiveis = data?.items ?? []

  const alvoId = item ? item.id : itemId ? Number(itemId) : null

  function reiniciar() {
    setItemId('')
    setBusca('')
    setDefeito('')
    setAssistencia('')
    setObservacoes('')
  }

  const mutation = useMutation({
    mutationFn: () =>
      abrirManutencao({
        item_id: alvoId as number,
        defeito: defeito.trim(),
        ...(assistencia.trim() ? { assistencia: assistencia.trim() } : {}),
        ...(observacoes.trim() ? { observacoes: observacoes.trim() } : {}),
      }),
    onSuccess: (r) => {
      toast(r.detail)
      void queryClient.invalidateQueries({ queryKey: ['manutencoes'] })
      void queryClient.invalidateQueries({ queryKey: ['item'] })
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      void queryClient.invalidateQueries({ queryKey: ['item-linha-do-tempo'] })
      onAberta?.(r.id)
      reiniciar()
      onOpenChange(false)
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Não foi possível abrir a manutenção.'), 'error')
    },
  })

  const podeAbrir = alvoId !== null && defeito.trim() !== ''

  return (
    <Dialog
      open={open}
      onOpenChange={(aberto) => {
        if (!aberto) reiniciar()
        onOpenChange(aberto)
      }}
    >
      <DialogContent className="max-w-lg">
        <DialogHeader>
          <DialogTitle>Abrir manutenção</DialogTitle>
          <DialogDescription>
            {item
              ? `${descricao(item)}${item.identificador ? ` (${item.identificador})` : ''}. O equipamento passa a Em Manutenção.`
              : 'O equipamento passa a Em Manutenção. Só equipamentos disponíveis (sem colaborador) aparecem na busca.'}
          </DialogDescription>
        </DialogHeader>

        <form
          className="space-y-4"
          onSubmit={(e) => {
            e.preventDefault()
            if (podeAbrir) mutation.mutate()
          }}
        >
          {!item && (
            <div className="flex flex-col gap-1.5">
              <Label className="text-caption font-medium">Equipamento</Label>
              <SearchableSelect
                options={disponiveis.map((i) => ({
                  value: String(i.id),
                  label: `#${i.id} — ${descricao(i)}`,
                  subtitle: [upper(i.revenda), i.identificador].filter(Boolean).join(' • '),
                }))}
                value={itemId}
                onValueChange={setItemId}
                placeholder="Pesquisar por marca, modelo ou identificador…"
                searchPlaceholder="Digite para filtrar equipamentos disponíveis…"
                search={busca}
                onSearchChange={setBusca}
              />
            </div>
          )}

          <div className="flex flex-col gap-1.5">
            <Label htmlFor="manutencao-defeito" className="text-caption font-medium">
              Defeito
            </Label>
            <Input
              id="manutencao-defeito"
              value={defeito}
              maxLength={500}
              onChange={(e) => setDefeito(e.target.value)}
              placeholder="Ex.: Tela quebrada"
              className="h-9"
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="manutencao-assistencia" className="text-caption font-medium">
              Assistência técnica
            </Label>
            <Input
              id="manutencao-assistencia"
              value={assistencia}
              maxLength={150}
              onChange={(e) => setAssistencia(e.target.value)}
              className="h-9"
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="manutencao-observacoes" className="text-caption font-medium">
              Observações
            </Label>
            <Input
              id="manutencao-observacoes"
              value={observacoes}
              maxLength={500}
              onChange={(e) => setObservacoes(e.target.value)}
              className="h-9"
            />
          </div>

          <DialogFooter className="gap-2 sm:gap-0">
            <Button type="button" variant="outline" size="sm" onClick={() => onOpenChange(false)}>
              Cancelar
            </Button>
            <Button type="submit" size="sm" loading={mutation.isPending} disabled={!podeAbrir}>
              Abrir manutenção
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  )
}
