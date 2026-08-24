import { useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { listItemsPaginated } from '@/api/items'
import {
  listPeripherals,
  listItemPeripherals,
  linkPeripheral,
  unlinkPeripheral,
  replacePeripheral,
  type Peripheral,
} from '@/api/peripherals'
import { Button } from '@/components/ui/button'
import { Label } from '@/components/ui/label'
import { Input } from '@/components/ui/input'
import { Badge } from '@/components/ui/badge'
import { FileUpload } from '@/components/ui/FileUpload'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { SearchableSelect } from '@/components/ui/SearchableSelect'
import { toast } from '@/components/ui/toast'
import { Link2, Unlink, RefreshCw } from 'lucide-react'

// Tipos de equipamento que aceitam periféricos. Regra própria desta tela
// (subconjunto curado, não um cadastro de domínio) — não é uma das listas
// duplicadas com o backend que a T2 pede para eliminar (revendas, setores,
// centros de custo, tipos de equipamento/periférico, motivos de remoção),
// por isso continua fixa aqui.
const LINK_ALLOWED_TYPES = ['Desktop', 'Notebook', 'Switch', 'Impressora']

// Sem paginação nesta tela (filtra os equipamentos linkáveis client-side a
// partir da lista completa) — usamos o teto de página do backend para não
// truncar em 50 itens (default de GET /api/items) como aconteceria chamando
// listItemsPaginated() sem limit.
const FETCH_ALL_LIMIT = 500

function PeripheralCard({
  peripheral,
  action,
  actionLabel,
  actionIcon,
  variant = 'default',
}: {
  peripheral: Peripheral
  action: () => void
  actionLabel: string
  actionIcon: React.ReactNode
  variant?: 'default' | 'destructive'
}) {
  return (
    <div className="flex items-center justify-between rounded-lg border bg-card px-4 py-3 shadow-sm">
      <div className="flex flex-col gap-0.5">
        <span className="text-sm font-medium text-foreground">
          {peripheral.tipo} — {peripheral.brand || '-'} {peripheral.model || ''}
        </span>
        <span className="text-xs text-muted-foreground">S/N: {peripheral.identificador || '-'}</span>
      </div>
      <div className="flex items-center gap-3">
        <Badge variant={peripheral.status === 'Disponível' ? 'success' : peripheral.status === 'Em Uso' ? 'warning' : 'danger'}>
          {peripheral.status}
        </Badge>
        <Button size="sm" variant={variant === 'destructive' ? 'destructive' : 'outline'} onClick={action}>
          {actionIcon}
          <span className="ml-1">{actionLabel}</span>
        </Button>
      </div>
    </div>
  )
}

export default function LinkPeripheralPage() {
  const queryClient = useQueryClient()
  const [selectedItemId, setSelectedItemId] = useState('')

  // Substituição
  const [replacingLinkId, setReplacingLinkId] = useState<number | null>(null)
  const [replacingOldId, setReplacingOldId] = useState<number | null>(null)
  const [replaceNewId, setReplaceNewId] = useState('')
  const [replaceReason, setReplaceReason] = useState('')
  const [replaceAttachment, setReplaceAttachment] = useState<File | null>(null)

  const { data } = useQuery({
    queryKey: ['items'],
    queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }),
  })
  const items = data?.items ?? []
  const linkableItems = items.filter(i => LINK_ALLOWED_TYPES.includes(i.tipo ?? ''))

  const { data: linkedPeripherals = [], refetch: refetchLinked } = useQuery({
    queryKey: ['item-peripherals', selectedItemId],
    queryFn: () => listItemPeripherals(Number(selectedItemId)),
    enabled: !!selectedItemId,
  })

  const { data: availablePeripherals = [] } = useQuery({
    queryKey: ['peripherals', 'Disponível'],
    queryFn: () => listPeripherals({ status: 'Disponível' }),
  })

  const linkMutation = useMutation({
    mutationFn: (pid: number) => linkPeripheral(Number(selectedItemId), pid),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['item-peripherals', selectedItemId] })
      queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      toast('Periférico vinculado!')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao vincular.'
      toast(msg, 'error')
    },
  })

  const unlinkMutation = useMutation({
    mutationFn: (linkId: number) => unlinkPeripheral(linkId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['item-peripherals', selectedItemId] })
      queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      toast('Periférico desvinculado.')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao desvincular.'
      toast(msg, 'error')
    },
  })

  const replaceMutation = useMutation({
    mutationFn: () =>
      replacePeripheral(
        Number(selectedItemId),
        replacingOldId!,
        Number(replaceNewId),
        replaceReason,
        replaceAttachment ?? undefined
      ),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['item-peripherals', selectedItemId] })
      queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      setReplacingLinkId(null)
      setReplacingOldId(null)
      setReplaceNewId('')
      setReplaceReason('')
      setReplaceAttachment(null)
      toast('Periférico substituído com sucesso!')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao substituir.'
      toast(msg, 'error')
    },
  })

  const selectedItem = items.find(i => String(i.id) === selectedItemId)

  return (
    <div className="space-y-6 max-w-4xl">
      <div>
        <h2 className="text-xl font-semibold">Vincular Periféricos</h2>
        <p className="text-sm text-muted-foreground">Associe periféricos a equipamentos como desktops, notebooks, switches e impressoras.</p>
      </div>

      {/* Seletor de equipamento */}
      <div className="bg-card rounded-xl border p-4 flex flex-col gap-2">
        <Label>Selecione o Equipamento</Label>
        <div className="max-w-md">
          <SearchableSelect
            options={linkableItems.map((i) => ({
              value: String(i.id),
              label: `#${i.id} — ${i.tipo} ${i.brand || ''} ${i.model || ''}`,
              subtitle: [i.revenda, i.identificador].filter(Boolean).join(' • '),
            }))}
            value={selectedItemId}
            onValueChange={setSelectedItemId}
            placeholder="Selecione ou busque um equipamento..."
            searchPlaceholder="Buscar por ID, tipo, marca, modelo, patrimônio..."
          />
        </div>
        {selectedItem && (
          <p className="text-sm text-muted-foreground">
            Status: <strong>{selectedItem.status}</strong> · Revenda: <strong>{selectedItem.revenda}</strong>
          </p>
        )}
      </div>

      {selectedItemId && (
        <div className="grid grid-cols-2 gap-6">
          {/* Periféricos vinculados */}
          <div className="space-y-3">
            <div className="flex items-center justify-between">
              <h3 className="font-medium text-foreground">Vinculados ({linkedPeripherals.length})</h3>
              <Button size="sm" variant="ghost" onClick={() => refetchLinked()}>
                <RefreshCw size={13} />
              </Button>
            </div>
            {linkedPeripherals.length === 0 ? (
              <p className="text-sm text-muted-foreground py-4 text-center border rounded-lg">Nenhum periférico vinculado.</p>
            ) : (
              <div className="space-y-2">
                {linkedPeripherals.map(p => (
                  <div key={p.link_id} className="flex flex-col gap-1">
                    <PeripheralCard
                      peripheral={p}
                      action={() => unlinkMutation.mutate(p.link_id!)}
                      actionLabel="Desvincular"
                      actionIcon={<Unlink size={13} />}
                      variant="destructive"
                    />
                    <button
                      className="text-xs text-blue-500 hover:underline text-left ml-1"
                      onClick={() => { setReplacingLinkId(p.link_id!); setReplacingOldId(p.id) }}
                    >
                      Substituir...
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Periféricos disponíveis */}
          <div className="space-y-3">
            <h3 className="font-medium text-foreground">Disponíveis ({availablePeripherals.length})</h3>
            {availablePeripherals.length === 0 ? (
              <p className="text-sm text-muted-foreground py-4 text-center border rounded-lg">Nenhum periférico disponível.</p>
            ) : (
              <div className="space-y-2 max-h-[420px] overflow-y-auto pr-1">
                {availablePeripherals.map(p => (
                  <PeripheralCard
                    key={p.id}
                    peripheral={p}
                    action={() => linkMutation.mutate(p.id)}
                    actionLabel="Vincular"
                    actionIcon={<Link2 size={13} />}
                  />
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Painel de substituição */}
      {replacingLinkId && replacingOldId && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-5 space-y-4">
          <h3 className="font-medium text-amber-800">Substituir Periférico #{replacingOldId}</h3>
          <div className="grid grid-cols-2 gap-4">
            <div className="flex flex-col gap-1.5">
              <Label>Novo Periférico (ID)</Label>
              <Select value={replaceNewId} onValueChange={setReplaceNewId}>
                <SelectTrigger><SelectValue placeholder="Selecione o substituto" /></SelectTrigger>
                <SelectContent>
                  {availablePeripherals.map(p => (
                    <SelectItem key={p.id} value={String(p.id)}>
                      #{p.id} — {p.tipo} {p.brand} {p.model}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>Motivo da Substituição *</Label>
              <Input value={replaceReason} onChange={e => setReplaceReason(e.target.value)} placeholder="Ex: Defeito, Upgrade..." />
            </div>
          </div>
          <FileUpload
            accept={{ 'application/pdf': ['.pdf'], 'image/*': ['.jpg', '.jpeg', '.png'] }}
            onFile={setReplaceAttachment}
            label="Comprovante (opcional)"
          />
          <div className="flex gap-3">
            <Button
              disabled={!replaceNewId || !replaceReason || replaceMutation.isPending}
              onClick={() => replaceMutation.mutate()}
            >
              {replaceMutation.isPending ? 'Substituindo...' : 'Confirmar Substituição'}
            </Button>
            <Button variant="ghost" onClick={() => { setReplacingLinkId(null); setReplacingOldId(null) }}>
              Cancelar
            </Button>
          </div>
        </div>
      )}
    </div>
  )
}
