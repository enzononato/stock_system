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

// Tipos de equipamento que aceitam periféricos
const LINK_ALLOWED_TYPES = ['Desktop', 'Notebook', 'Switch', 'Impressora']
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
    <div className="flex items-center justify-between border border-border rounded p-3 bg-surface hover:bg-surface-alt transition-colors">
      <div className="flex flex-col gap-0.5 min-w-0">
        <span className="text-body-sm font-medium text-foreground truncate">
          {peripheral.tipo} — {peripheral.brand || '—'} {peripheral.model || ''}
        </span>
        <span className="text-[11px] text-muted-foreground font-mono">
          S/N: {peripheral.identificador || '—'}
        </span>
      </div>
      <div className="flex items-center gap-3 shrink-0">
        <Badge
          variant={peripheral.status === 'Disponível' ? 'success' : peripheral.status === 'Em Uso' ? 'warning' : 'danger'}
          showDot
          className="text-[10px]"
        >
          {peripheral.status}
        </Badge>
        <Button
          size="sm"
          variant={variant === 'destructive' ? 'outline' : 'default'}
          onClick={action}
          className={variant === 'destructive' ? 'text-destructive border-border hover:bg-surface-alt text-xs h-7' : 'text-xs h-7'}
        >
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
  const linkableItems = items.filter((i) => LINK_ALLOWED_TYPES.includes(i.tipo ?? ''))

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
      void queryClient.invalidateQueries({ queryKey: ['item-peripherals', selectedItemId] })
      void queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      toast.success('Periférico vinculado!')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao vincular.'
      toast.error(msg)
    },
  })

  const unlinkMutation = useMutation({
    mutationFn: (linkId: number) => unlinkPeripheral(linkId),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['item-peripherals', selectedItemId] })
      void queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      toast.success('Periférico desvinculado.')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao desvincular.'
      toast.error(msg)
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
      void queryClient.invalidateQueries({ queryKey: ['item-peripherals', selectedItemId] })
      void queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      setReplacingLinkId(null)
      setReplacingOldId(null)
      setReplaceNewId('')
      setReplaceReason('')
      setReplaceAttachment(null)
      toast.success('Periférico substituído com sucesso!')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao substituir.'
      toast.error(msg)
    },
  })

  const selectedItem = items.find((i) => String(i.id) === selectedItemId)

  return (
    <div className="space-y-6 max-w-6xl mx-auto">
      {/* Header */}
      <div className="border-b border-border pb-4">
        <div className="flex items-center gap-2">
          <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
            Ativos de Suporte
          </span>
          <span className="text-muted-foreground">•</span>
          <span className="text-caption text-foreground font-medium">Vínculo de Periféricos</span>
        </div>
        <h1 className="text-heading-lg font-semibold tracking-tight text-foreground mt-0.5">
          Vincular Periféricos
        </h1>
        <p className="text-body-sm text-muted-foreground mt-1">
          Associe periféricos a equipamentos como desktops, notebooks, switches e impressoras.
        </p>
      </div>

      {/* Seletor de equipamento */}
      <div className="rounded border border-border bg-surface p-5 space-y-3">
        <Label className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
          Selecione o Equipamento Principal
        </Label>
        <div className="max-w-xl">
          <SearchableSelect
            options={linkableItems.map((i) => ({
              value: String(i.id),
              label: `#${i.id} — ${i.tipo} ${i.brand || ''} ${i.model || ''}`.trim(),
              subtitle: [i.revenda, i.identificador].filter(Boolean).join(' • '),
            }))}
            value={selectedItemId}
            onValueChange={setSelectedItemId}
            placeholder="Selecione ou busque um equipamento…"
            searchPlaceholder="Buscar por ID, tipo, marca, modelo, patrimônio…"
          />
        </div>
        {selectedItem && (
          <div className="flex items-center gap-4 text-caption text-muted-foreground pt-1">
            <span>
              Status: <strong className="text-foreground">{selectedItem.status}</strong>
            </span>
            <span>•</span>
            <span>
              Unidade: <strong className="text-foreground">{selectedItem.revenda || '—'}</strong>
            </span>
            {selectedItem.assigned_to && (
              <>
                <span>•</span>
                <span>
                  Alocado: <strong className="text-foreground">{selectedItem.assigned_to}</strong>
                </span>
              </>
            )}
          </div>
        )}
      </div>

      {selectedItemId && (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* Periféricos vinculados */}
          <div className="rounded border border-border bg-surface p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-border pb-3">
              <div>
                <h2 className="text-body font-semibold text-foreground">
                  Periféricos Vinculados ({linkedPeripherals.length})
                </h2>
                <p className="text-caption text-muted-foreground">Componentes conectados a este ativo.</p>
              </div>
              <Button size="sm" variant="ghost" onClick={() => void refetchLinked()} className="size-8 p-0">
                <RefreshCw className="size-3.5" />
              </Button>
            </div>

            {linkedPeripherals.length === 0 ? (
              <p className="text-body-sm text-muted-foreground py-12 text-center">Nenhum periférico vinculado.</p>
            ) : (
              <div className="space-y-2.5">
                {linkedPeripherals.map((p) => (
                  <div key={p.link_id} className="flex flex-col gap-1">
                    <PeripheralCard
                      peripheral={p}
                      action={() => unlinkMutation.mutate(p.link_id!)}
                      actionLabel="Desvincular"
                      actionIcon={<Unlink className="size-3" />}
                      variant="destructive"
                    />
                    <button
                      className="text-xs text-muted-foreground hover:text-foreground hover:underline text-left ml-1 cursor-pointer"
                      onClick={() => {
                        setReplacingLinkId(p.link_id!)
                        setReplacingOldId(p.id)
                      }}
                    >
                      Substituir por outro…
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Periféricos disponíveis */}
          <div className="rounded border border-border bg-surface p-5 space-y-4">
            <div className="border-b border-border pb-3">
              <h2 className="text-body font-semibold text-foreground">
                Periféricos Disponíveis ({availablePeripherals.length})
              </h2>
              <p className="text-caption text-muted-foreground">Itens em estoque livres para vinculação.</p>
            </div>

            {availablePeripherals.length === 0 ? (
              <p className="text-body-sm text-muted-foreground py-12 text-center">Nenhum periférico disponível.</p>
            ) : (
              <div className="space-y-2.5 max-h-[440px] overflow-y-auto pr-1">
                {availablePeripherals.map((p) => (
                  <PeripheralCard
                    key={p.id}
                    peripheral={p}
                    action={() => linkMutation.mutate(p.id)}
                    actionLabel="Vincular"
                    actionIcon={<Link2 className="size-3" />}
                  />
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Painel de substituição */}
      {replacingLinkId && replacingOldId && (
        <div className="rounded border border-border bg-surface-alt p-6 space-y-4">
          <div className="border-b border-border pb-2">
            <h3 className="text-body-lg font-semibold text-foreground">
              Substituição de Periférico #{replacingOldId}
            </h3>
            <p className="text-caption text-muted-foreground">
              Selecione o novo periférico que assumirá o vínculo e informe a justificativa.
            </p>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div className="flex flex-col gap-1.5">
              <Label className="text-caption font-medium">Novo Periférico Substituto *</Label>
              <Select value={replaceNewId} onValueChange={setReplaceNewId}>
                <SelectTrigger className="h-9">
                  <SelectValue placeholder="Selecione o substituto" />
                </SelectTrigger>
                <SelectContent>
                  {availablePeripherals.map((p) => (
                    <SelectItem key={p.id} value={String(p.id)}>
                      #{p.id} — {p.tipo} {p.brand} {p.model}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label className="text-caption font-medium">Motivo da Substituição *</Label>
              <Input
                value={replaceReason}
                onChange={(e) => setReplaceReason(e.target.value)}
                placeholder="Ex: Defeito, Upgrade, Avaria…"
                className="h-9"
              />
            </div>
          </div>

          <div className="pt-2">
            <FileUpload
              accept={{ 'application/pdf': ['.pdf'], 'image/*': ['.jpg', '.jpeg', '.png'] }}
              onFile={setReplaceAttachment}
              label="Comprovante / laudo (opcional)"
            />
          </div>

          <div className="flex justify-end gap-3 pt-2 border-t border-border">
            <Button
              variant="outline"
              size="sm"
              onClick={() => {
                setReplacingLinkId(null)
                setReplacingOldId(null)
              }}
            >
              Cancelar
            </Button>
            <Button
              size="sm"
              disabled={!replaceNewId || !replaceReason || replaceMutation.isPending}
              onClick={() => replaceMutation.mutate()}
            >
              {replaceMutation.isPending ? 'Substituindo…' : 'Confirmar Substituição'}
            </Button>
          </div>
        </div>
      )}
    </div>
  )
}
