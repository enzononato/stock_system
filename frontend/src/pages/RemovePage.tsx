import { useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { listItemsPaginated, removeItem } from '@/api/items'
import { Button } from '@/components/ui/button'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { SearchableSelect } from '@/components/ui/SearchableSelect'
import { FileUpload } from '@/components/ui/FileUpload'
import { toast } from '@/components/ui/toast'
import { useConstants } from '@/hooks/useConstants'
import { Trash2 } from 'lucide-react'

// Sem paginação nesta tela (filtra "Disponível" client-side a partir da
// lista completa) — usamos o teto de página do backend para não truncar em
// 50 itens (default de GET /api/items) como aconteceria chamando
// listItemsPaginated() sem limit.
const FETCH_ALL_LIMIT = 500

export default function RemovePage() {
  const queryClient = useQueryClient()
  const { removalReasons, removalReasonsAttachment, isLoading: constantsLoading } = useConstants()
  const [selectedId, setSelectedId] = useState('')
  const [reason, setReason] = useState('')
  const [attachment, setAttachment] = useState<File | null>(null)

  const { data } = useQuery({
    queryKey: ['items'],
    queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }),
  })
  const items = data?.items ?? []
  const disponiveis = items.filter(i => i.status === 'Disponível')

  // T2: o mapa de quais motivos exigem comprovante vem de /api/constants
  // (removalReasonsAttachment), não mais de um objeto hardcoded local.
  const needsAttachment = reason ? removalReasonsAttachment[reason] : false

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
    <div className="max-w-4xl space-y-6">
      <div>
        <h2 className="text-h1 text-destructive">Remover Equipamento</h2>
        <p className="text-caption mt-1">Remove permanentemente o item do estoque. Esta ação não pode ser desfeita.</p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-[1fr_260px] gap-5 items-start">
      <form onSubmit={handleSubmit} className="bg-card rounded-xl border border-destructive/25 p-6 space-y-4">
        <div className="flex flex-col gap-1.5">
          <Label>Equipamento *</Label>
          <SearchableSelect
            options={disponiveis.map((i) => ({
              value: String(i.id),
              label: `#${i.id} — ${i.tipo} ${i.brand || ''} ${i.model || ''}`,
              subtitle: [i.revenda, i.identificador].filter(Boolean).join(' • '),
            }))}
            value={selectedId}
            onValueChange={setSelectedId}
            placeholder="Selecione ou busque um equipamento..."
            searchPlaceholder="Buscar por ID, tipo, marca, modelo, patrimônio..."
          />
        </div>

        <div className="flex flex-col gap-1.5">
          <Label>Motivo da Remoção *</Label>
          <Select value={reason} onValueChange={setReason} required disabled={constantsLoading}>
            <SelectTrigger><SelectValue placeholder="Selecione o motivo" /></SelectTrigger>
            <SelectContent>
              {removalReasons.map(r => <SelectItem key={r} value={r}>{r}</SelectItem>)}
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

      <div className="rounded-lg border border-destructive/25 bg-destructive/5 p-4 space-y-2 sticky top-4">
        <h3 className="text-xs font-semibold uppercase tracking-wider text-destructive">Atenção</h3>
        <p className="text-[11px] text-muted-foreground leading-relaxed">
          A remoção é permanente e não pode ser desfeita. Use "Estorno" apenas para reverter operações recentes registradas por engano — para remoção definitiva de um item do acervo, esta é a tela correta.
        </p>
        <p className="text-[11px] text-muted-foreground leading-relaxed">
          Alguns motivos exigem comprovante anexado antes da confirmação.
        </p>
      </div>
      </div>
    </div>
  )
}
