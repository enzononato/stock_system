import { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { createInventario } from '@/api/inventarios'
import { listUnidades } from '@/api/unidades'
import { Button } from '@/components/ui/button'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'

interface NovoInventarioDialogProps {
  open: boolean
  onOpenChange: (open: boolean) => void
}

const campoClass =
  'flex w-full rounded border border-input bg-surface-alt px-3 py-2 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring'

/** Abre um inventário para uma unidade ativa e vai direto para a tela de leitura. */
export function NovoInventarioDialog({ open, onOpenChange }: NovoInventarioDialogProps) {
  const navigate = useNavigate()
  const queryClient = useQueryClient()
  const [unidadeId, setUnidadeId] = useState('')
  const [observacao, setObservacao] = useState('')

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades', { includeInactive: false }],
    queryFn: () => listUnidades(),
  })

  const abrir = useMutation({
    mutationFn: () =>
      createInventario({
        unidade_id: Number(unidadeId),
        ...(observacao.trim() ? { observacao: observacao.trim() } : {}),
      }),
    onSuccess: (inventario) => {
      void queryClient.invalidateQueries({ queryKey: ['inventarios'] })
      onOpenChange(false)
      navigate(`/inventarios/${inventario.id}/leitura`)
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, 'Erro ao abrir o inventário.')),
  })

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-md">
        <DialogHeader>
          <DialogTitle>Novo inventário</DialogTitle>
          <DialogDescription>
            Escolha a unidade. Cada unidade pode ter um inventário aberto por vez.
          </DialogDescription>
        </DialogHeader>
        <form
          className="space-y-4"
          onSubmit={(e) => {
            e.preventDefault()
            if (unidadeId) abrir.mutate()
          }}
        >
          <label className="flex flex-col gap-1 text-caption text-muted-foreground">
            Unidade
            <select
              aria-label="Unidade do inventário"
              className={`${campoClass} h-9`}
              value={unidadeId}
              onChange={(e) => setUnidadeId(e.target.value)}
            >
              <option value="">Selecione uma unidade...</option>
              {unidades.map((u) => (
                <option key={u.id} value={u.id}>
                  {upper(u.nome)}
                </option>
              ))}
            </select>
          </label>
          <label className="flex flex-col gap-1 text-caption text-muted-foreground">
            Observação
            <textarea
              rows={3}
              maxLength={500}
              className={campoClass}
              placeholder="Opcional"
              value={observacao}
              onChange={(e) => setObservacao(e.target.value)}
            />
          </label>
          <DialogFooter>
            <Button type="button" variant="outline" onClick={() => onOpenChange(false)}>
              Cancelar
            </Button>
            <Button type="submit" loading={abrir.isPending} disabled={!unidadeId}>
              Abrir inventário
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  )
}
