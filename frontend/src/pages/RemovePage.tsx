import { useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { listItems, removeItem } from '@/api/items'
import { Button } from '@/components/ui/button'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { FileUpload } from '@/components/ui/FileUpload'
import { toast } from '@/components/ui/toast'
import { Trash2 } from 'lucide-react'

const REMOVAL_REASONS: Record<string, boolean> = {
  Roubo: true,
  Perda: false,
  Obsolescência: false,
  Doação: true,
  Venda: true,
}

export default function RemovePage() {
  const queryClient = useQueryClient()
  const [selectedId, setSelectedId] = useState('')
  const [reason, setReason] = useState('')
  const [attachment, setAttachment] = useState<File | null>(null)

  const { data: items = [] } = useQuery({ queryKey: ['items'], queryFn: () => listItems() })
  const disponiveis = items.filter(i => i.status === 'Disponível')

  const needsAttachment = reason ? REMOVAL_REASONS[reason] : false

  const mutation = useMutation({
    mutationFn: () => removeItem(Number(selectedId), reason, attachment ?? undefined),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] })
      setSelectedId('')
      setReason('')
      setAttachment(null)
      toast('Item removido do estoque.')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao remover item.'
      toast(msg, 'error')
    },
  })

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    if (!selectedId) { toast('Selecione um item.', 'error'); return }
    if (!reason) { toast('Selecione o motivo.', 'error'); return }
    if (needsAttachment && !attachment) { toast('Este motivo exige um comprovante.', 'error'); return }
    mutation.mutate()
  }

  return (
    <div className="max-w-xl space-y-6">
      <div>
        <h2 className="text-xl font-semibold text-red-700">Remover Equipamento</h2>
        <p className="text-sm text-slate-500">Remove permanentemente o item do estoque. Esta ação não pode ser desfeita.</p>
      </div>

      <form onSubmit={handleSubmit} className="bg-white rounded-xl border border-red-100 p-6 space-y-4">
        <div className="flex flex-col gap-1.5">
          <Label>Equipamento *</Label>
          <Select value={selectedId} onValueChange={setSelectedId} required>
            <SelectTrigger><SelectValue placeholder="Selecione um item" /></SelectTrigger>
            <SelectContent>
              {disponiveis.map(i => (
                <SelectItem key={i.id} value={String(i.id)}>
                  #{i.id} — {i.tipo} {i.brand} {i.model} ({i.revenda})
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>

        <div className="flex flex-col gap-1.5">
          <Label>Motivo da Remoção *</Label>
          <Select value={reason} onValueChange={setReason} required>
            <SelectTrigger><SelectValue placeholder="Selecione o motivo" /></SelectTrigger>
            <SelectContent>
              {Object.keys(REMOVAL_REASONS).map(r => <SelectItem key={r} value={r}>{r}</SelectItem>)}
            </SelectContent>
          </Select>
        </div>

        {reason && (
          <div className="flex flex-col gap-1.5">
            <Label>
              Comprovante {needsAttachment ? '(obrigatório)' : '(opcional)'}
            </Label>
            <FileUpload
              accept={{ 'application/pdf': ['.pdf'], 'image/*': ['.jpg', '.jpeg', '.png'] }}
              onFile={setAttachment}
              label="Arraste ou clique para anexar comprovante"
            />
          </div>
        )}

        <Button type="submit" variant="destructive" disabled={mutation.isPending} className="w-full">
          <Trash2 size={14} className="mr-2" />
          {mutation.isPending ? 'Removendo...' : 'Confirmar Remoção'}
        </Button>
      </form>
    </div>
  )
}
