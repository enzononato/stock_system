import { useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { transferirItem, type Item } from '@/api/items'
import { listUnidades } from '@/api/unidades'
import { LocalSelect } from '@/components/locais/LocalSelect'
import { Button } from '@/components/ui/button'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { Label } from '@/components/ui/label'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'

export interface TransferirItemDialogProps {
  item: Pick<Item, 'id' | 'revenda' | 'local_id'>
  open: boolean
  onOpenChange: (open: boolean) => void
  onTransferido?: (r: { para_unidade: string; para_local_id: number | null }) => void
}

const campoClass =
  'flex w-full rounded border border-input bg-surface-alt px-3 py-2 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring disabled:cursor-not-allowed disabled:opacity-40'

/**
 * Transferência do equipamento para outra unidade (e, opcionalmente, um local dela).
 * A unidade só muda por aqui: a edição do cadastro bloqueia o campo e o backend
 * recusa a troca direta. O motivo é obrigatório e fica na linha do tempo.
 */
export function TransferirItemDialog({ item, open, onOpenChange, onTransferido }: TransferirItemDialogProps) {
  const queryClient = useQueryClient()
  const [destinoId, setDestinoId] = useState<number | null>(null)
  const [localId, setLocalId] = useState<number | null>(null)
  const [motivo, setMotivo] = useState('')

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades', 'ativas'],
    queryFn: () => listUnidades(),
    enabled: open,
  })
  const atual = (item.revenda ?? '').trim().toLocaleLowerCase('pt-BR')
  const destinos = unidades.filter((u) => u.is_active && u.nome.trim().toLocaleLowerCase('pt-BR') !== atual)
  const destino = destinos.find((u) => u.id === destinoId)

  function reiniciar() {
    setDestinoId(null)
    setLocalId(null)
    setMotivo('')
  }

  const mutation = useMutation({
    mutationFn: () =>
      transferirItem(item.id, {
        para_unidade_id: destinoId as number,
        para_local_id: localId,
        motivo: motivo.trim(),
      }),
    onSuccess: (r) => {
      toast(r.detail)
      void queryClient.invalidateQueries({ queryKey: ['item'] })
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      void queryClient.invalidateQueries({ queryKey: ['item-linha-do-tempo'] })
      onTransferido?.({ para_unidade: destino?.nome ?? '', para_local_id: localId })
      reiniciar()
      onOpenChange(false)
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Não foi possível transferir o equipamento.'), 'error')
    },
  })

  const podeConfirmar = destinoId !== null && motivo.trim() !== ''

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
          <DialogTitle>Transferir equipamento</DialogTitle>
          <DialogDescription>
            {item.revenda ? `Unidade atual: ${upper(item.revenda)}. ` : ''}A transferência fica registrada na
            linha do tempo e na auditoria.
          </DialogDescription>
        </DialogHeader>

        <form
          className="space-y-4"
          onSubmit={(e) => {
            e.preventDefault()
            if (podeConfirmar) mutation.mutate()
          }}
        >
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="transferir-destino" className="text-caption font-medium">
              Unidade de destino
            </Label>
            <select
              id="transferir-destino"
              aria-label="Unidade de destino"
              className={`${campoClass} h-9`}
              value={destinoId === null ? '' : String(destinoId)}
              onChange={(e) => {
                setDestinoId(e.target.value === '' ? null : Number(e.target.value))
                setLocalId(null)
              }}
            >
              <option value="">Selecione a unidade</option>
              {destinos.map((u) => (
                <option key={u.id} value={u.id}>
                  {upper(u.nome)}
                </option>
              ))}
            </select>
          </div>

          <div role="group" aria-labelledby="transferir-local-rotulo" className="flex flex-col gap-1.5">
            <span id="transferir-local-rotulo" className="text-caption font-medium text-foreground">
              Local de destino
            </span>
            <LocalSelect unidadeId={destinoId} value={localId} onChange={setLocalId} />
          </div>

          <div className="flex flex-col gap-1.5">
            <Label htmlFor="transferir-motivo" className="text-caption font-medium">
              Motivo da transferência
            </Label>
            <textarea
              id="transferir-motivo"
              rows={3}
              maxLength={255}
              className={campoClass}
              value={motivo}
              onChange={(e) => setMotivo(e.target.value)}
              placeholder="Ex.: Mudança de setor"
            />
          </div>

          <DialogFooter className="gap-2 sm:gap-0">
            <Button type="button" variant="outline" size="sm" onClick={() => onOpenChange(false)}>
              Cancelar
            </Button>
            <Button type="submit" size="sm" loading={mutation.isPending} disabled={!podeConfirmar}>
              Confirmar transferência
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  )
}
